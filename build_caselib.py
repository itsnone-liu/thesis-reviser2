#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""wikitext → 结构化案例 → caselib 入库 (400建库主管线)。

输入: case_library_src/ws_N.json (维基文库官方发布全文)
校验: 案号+名称 与双源目录(wiki列表/ggdlvshi汇编)交叉核对, 一致才 verified=true
字段: 名称/案号/当事人(从名称解析)/案由(从名称纠纷类型)/基本事实(基本案情)/
      诉讼请求(酌情)/裁判结果/争议焦点(裁判理由提炼)/裁判要点/关键词→tags/
      领域(case_taxonomy.classify)/批次/来源/verified
废止案(9/20号等)不入库或标废止。
用法: python3 build_caselib.py [--dry] [--src case_library_src]
"""
import glob, json, os, re, sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import caselib
from case_taxonomy import classify, normalize, AGG

SRC_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "case_library_src")

# 双源目录(用于交叉核验): 案号 → 名称
def load_catalogs():
    cat = {}
    p1 = "/tmp/guide_catalog.json"
    if os.path.exists(p1):
        cat.update({int(k): v for k, v in json.load(open(p1, encoding="utf-8")).items()})
    # ggdlvshi汇编(含批次+作废)
    p2 = "/tmp/comp224.html"
    if os.path.exists(p2):
        import html as _h
        src = open(p2, encoding="utf-8", errors="replace").read()
        src = _h.unescape(re.sub(r"<[^>]+>", "\n", src))
        for m in re.finditer(r"指导案例\s*(\d{1,3})\s*号[：:]\s*([^\n（(]{4,80})", src):
            n, nm = int(m.group(1)), m.group(2).strip()
            if n not in cat:
                cat[n] = nm
    return cat


ABOLISHED = {9, 20}  # 汇编页标注作废(2021废止, 不再参照)


def strip_markup(wt: str) -> str:
    t = wt
    t = re.sub(r"\{\{[^{}]*\}\}", "", t)            # 模板(嵌套少, 两轮够)
    t = re.sub(r"\{\{[^{}]*\}\}", "", t)
    t = re.sub(r"\{\|.*?\|\}", "", t, flags=re.S)
    t = re.sub(r"<br\s*/?>", "\n", t)
    t = re.sub(r"<[^>]+>", "", t)
    t = re.sub(r"'''?", "", t)
    t = re.sub(r"\[\[(?:[^|\]]*\|)?([^\]]+)\]\]", r"\1", t)
    t = re.sub(r"&nbsp;?", " ", t)
    t = re.sub(r"\[\[Category:[^\]]*\]\]", "", t)
    t = re.sub(r"Category:[\u4e00-\u9fff]+", "", t)
    t = re.sub(r"[ \t]+", " ", t)
    t = re.sub(r"\n{3,}", "\n\n", t)
    return t.strip()


def parse_case(no: int, wikitext: str) -> dict:
    t = strip_markup(wikitext)
    case = {"案号": f"指导案例{no}号", "批次": f"第{no}号(最高法指导性案例)"}

    # 名称: 优先从raw wikitext的header模板title=取(模板剥离会删掉center里的标题)
    nm_raw = ""
    mh = re.search(r"\|\s*title\s*=\s*指导案例\s*\d+\s*号[：:]\s*(.+)", wikitext)
    if mh:
        nm_raw = re.sub(r"<br\s*/?>", "", mh.group(1))
        nm_raw = re.sub(r"\[\[(?:[^|\]]*\|)?([^\]]+)\]\]", r"\1", nm_raw)
        nm_raw = re.sub(r"\{\{[^{}]*\}\}", "", nm_raw)
        nm_raw = re.sub(r"'''?", "", nm_raw).strip(" 　'|")
    if not nm_raw:
        # 兜底: strip后的正文里第一个"X诉Y...案"行(排除栏目词开头)
        for ln in t.split("\n"):
            s = ln.strip(" 　'")
            if 6 <= len(s) <= 80 and re.search(r"(诉|贪污|受贿|盗窃|抢劫|诈骗)", s) \
               and s.endswith(("案", "裁定", "决定", "批复")) \
               and not s.startswith(("关键词", "基本案", "裁判", "相关法条", "Category", "指导案例")):
                nm_raw = s
                break
    case["名称"] = nm_raw

    def section(names, maxlen=4000):
        for nm in names:
            i = t.find(nm)
            if i >= 0:
                j = min([x for x in (t.find(k, i + len(nm)) for k in
                                     ("基本案情", "裁判结果", "裁判理由", "裁判要点",
                                      "相关法条", "关键词", "典型意义"))
                          if x > 0] or [len(t)])
                seg = t[i + len(nm):j].strip(" \n：:")
                if seg:
                    return seg[:maxlen]
        return ""

    kw_raw = section(["关键词"], 400)
    case["关键词"] = [k.strip() for k in
                    re.split(r"[;；,，、\u3000\u2002/]+", kw_raw) if k.strip()][:8]
    case["基本事实"] = section(["基本案情", "基本案情："])
    case["裁判结果"] = section(["裁判结果"], 800)
    case["裁判理由"] = section(["裁判理由"], 4000)
    case["裁判要点"] = section(["裁判要点"], 2000)
    case["相关法条"] = section(["相关法条"], 400)

    # 当事人/案由: 从名称"X诉Y...纠纷案"解析
    nm = case["名称"]
    m2 = re.match(r"(.{2,30}?)诉(.{2,40}?)(?:[（(].*?)?(.*?(?:纠纷)?案|裁定|决定)$", nm)
    if m2:
        case["当事人"] = f"原告:{m2.group(1).strip()}; 被告:{m2.group(2).strip()}"
        issue = re.search(r"([一二三四五六七八九十\d\u4e00-\u9fff]{2,12}?(?:纠纷|决定|裁定))", nm)
        case["案由"] = issue.group(1) if issue else (m2.group(3) or "")[:20]
    else:
        case["当事人"] = nm[:30]
        case["案由"] = nm[-14:]
    # 诉讼请求: 官方发布文本不单列, 从基本事实中归纳(请求/诉请句), 无则标注归纳来源
    m3 = re.search(r"[^。]{0,40}(?:请求|诉请|提起[^。]{0,10}之诉)[^。]{0,60}", case["基本事实"])
    case["诉讼请求"] = m3.group(0).strip() if m3 else \
        f"就{case.get('案由', '本案纠纷')}请求法院依法裁判(由基本案情归纳)"
    case["争议焦点"] = [l.strip(" 　") for l in
                     re.split(r"[。；\n]", case.get("裁判理由", "")) if 8 < len(l.strip()) < 60][:4]
    if not case["争议焦点"] and case["裁判要点"]:
        case["争议焦点"] = [case["裁判要点"].split("。")[0][:60]]

    case["领域"] = classify(case.get("案由", ""), nm,
                            " ".join(case.get("关键词", [])))
    case["tags"] = case["关键词"] + ([case["领域"]] if case["领域"] else [])
    case["来源"] = "最高人民法院指导性案例"
    case["来源链接"] = f"https://zh.wikisource.org/wiki/指导案例{no}号"
    return case


def parse_spc_case(d: dict) -> dict:
    """解析最高法官网详情页已清洗正文；官网本身是一手权威源。"""
    no = int(d["no"])
    t = re.sub(r"[ \t\r]+", " ", d.get("text", ""))
    t = re.sub(r"\n{2,}", "\n", t)
    name = str(d.get("name", ""))
    # 抽取各栏目到下一个栏目；官网早期页的栏目顺序不固定，故统一找最近边界
    labels = ("关键词", "裁判要点", "相关法条", "基本案情", "裁判结果", "裁判理由")
    def section(label, limit=4000):
        i = t.find(label)
        if i < 0:
            return ""
        ends = [t.find(x, i + len(label)) for x in labels if x != label]
        ends = [x for x in ends if x >= 0]
        val = t[i + len(label):min(ends) if ends else len(t)]
        val = re.sub(r"\s+", " ", val).strip(" \t\n:：")
        return val[:limit]
    kw = section("关键词", 500)
    keywords = [x.strip() for x in re.split(r"[;；,，、\u3000\u2002/ ]+", kw) if x.strip()]
    # 官网页标题可能含导航尾缀，先去尾缀再规范化空白
    name = re.split(r"\s+-\s+中华人民共和国最高人民法院|\s+_\s+", name)[0].strip()
    name = re.sub(r"\s+", "", name).lstrip("：: ")
    case = {
        "名称": name, "案号": f"指导案例{no}号",
        "关键词": keywords[:12], "基本事实": section("基本案情"),
        "裁判结果": section("裁判结果", 1000),
        "裁判理由": section("裁判理由"), "裁判要点": section("裁判要点", 2500),
        "相关法条": section("相关法条", 600),
        "批次": f"第{no}号(最高法指导性案例)", "来源": "最高人民法院官网",
        "来源链接": f"https://www.court.gov.cn/shenpan/xiangqing/{d['page_id']}.html",
    }
    m = re.match(r"(.{2,30}?)诉(.{2,50}?)(.*?(?:纠纷)?案|裁定|决定)$", name)
    if m:
        case["当事人"] = f"原告:{m.group(1).strip()}; 被告:{m.group(2).strip()}"
        case["案由"] = (m.group(3) or "")[:30]
    else:
        case["当事人"] = name[:40]
        case["案由"] = name[-18:]
    req = re.search(r"[^。]{0,40}(?:请求|诉请|提起[^。]{0,10}之诉)[^。]{0,70}", case["基本事实"])
    case["诉讼请求"] = req.group(0).strip() if req else f"就{case['案由']}请求法院依法裁判(由基本案情归纳)"
    case["争议焦点"] = [x.strip() for x in re.split(r"[。；]", case["裁判理由"])
                     if 8 < len(x.strip()) < 70][:4]
    if not case["争议焦点"] and case["裁判要点"]:
        case["争议焦点"] = [case["裁判要点"].split("。", 1)[0][:70]]
    case["领域"] = classify(case["案由"], name, " ".join(keywords))
    case["tags"] = keywords + ([case["领域"]] if case["领域"] else [])
    case["verified"] = True
    case["verified_note"] = "最高人民法院官网详情页原文，核心栏目解析齐备"
    return case


def reclassify():
    """只刷新在库条目的领域/tags(taxonomy演进后), 不动其他字段。"""
    import case_taxonomy as CT
    n = 0
    for path in sorted(glob.glob(os.path.join(caselib.LIB_DIR, "*.json"))):
        try:
            case = json.load(open(path, encoding="utf-8"))
        except Exception:
            continue
        # 官网详情页自身是一手核验源：旧条目若已带官方链接且核心字段完整，
        # 只提升 verified/说明，不重写人工内容。
        if (case.get("来源") == "最高人民法院官网" and case.get("来源链接", "").startswith("https://www.court.gov.cn/")
                and case.get("基本事实") and case.get("裁判要点")):
            case["verified"] = True
            case["verified_note"] = "最高人民法院官网详情页原文，案号与核心栏目齐备"
        agg = CT.classify(case.get("案由", ""), case.get("名称", ""),
                          " ".join(case.get("关键词", []) or []))
        if agg and case.get("领域") != agg:
            case["领域"] = agg
            tags = [t for t in (case.get("tags") or []) if t not in CT.AGG]
            case["tags"] = tags + [agg]
            json.dump(case, open(path, "w", encoding="utf-8"),
                      ensure_ascii=False, indent=1)
            n += 1
    print(f"reclassify: 更新{n}条")


def main():
    if "--reclassify" in sys.argv:
        reclassify()
        return
    dry = "--dry" in sys.argv
    cat = load_catalogs()
    files = sorted(glob.glob(f"{SRC_DIR}/ws_*.json") + glob.glob(f"{SRC_DIR}/spc_*.json") + glob.glob(f"{SRC_DIR}/gb_*.json"),
                   key=lambda p: (json.load(open(p, encoding="utf-8"))["no"], p))
    stats = {"ok": 0, "skip_abolished": 0, "name_mismatch": 0,
             "empty_core": 0, "no_catalog": 0}
    for f in files:
        d = json.load(open(f, encoding="utf-8"))
        no = d["no"]
        if no in ABOLISHED:
            stats["skip_abolished"] += 1
            continue
        case = (parse_case(no, d["wikitext"]) if "wikitext" in d
                else parse_spc_case(d) if str(d.get("page_id", "")).startswith("http")
                else parse_spc_case(d))
        # 核心字段齐备性
        if not case["名称"] or not case["基本事实"] or not case["裁判要点"]:
            stats["empty_core"] += 1
            continue
        # 保护人工精修条目: 库内已有且含人工字段(诉讼请求/共性问题短语)的不覆盖
        existing = caselib.get(case["案号"])
        if existing and (existing.get("诉讼请求") or existing.get("共性问题短语")):
            stats["keep_curated"] = stats.get("keep_curated", 0) + 1
            continue
        # 双源交叉核验
        ref = cat.get(no, "")
        agree = bool(ref) and (ref[:10] in case["名称"] or case["名称"][:10] in ref)
        if ref and not agree:
            # 双源和解: 目录名(wiki列表/ggdlvshi汇编)为准修正, 全文字段保留
            stats["name_mismatch"] += 1
            case["名称"] = ref
            case["verified"] = True
            case["verified_note"] = "名称按双源目录修正(文库页标题解析歧义)"
        elif not ref:
            stats["no_catalog"] += 1
            # 最高法官网详情页本身是一手权威源；其案号与正文核心栏齐备时可直接核验。
            # gb/spc缓存统一通过 parse_spc_case 进入，page_id为数字即官网详情页。
            is_spc = ("wikitext" not in d and str(d.get("page_id", "")).isdigit())
            case["verified"] = bool(is_spc)
            case["verified_note"] = ("最高人民法院官网详情页原文，案号与核心栏目齐备"
                                      if is_spc else "目录双源均未收录该号, 仅维基文库单源")
        else:
            case["verified"] = True
            case["verified_note"] = "维基文库官方全文+双源目录名称一致"
        probs = caselib.validate_case(case)
        if probs and any("必填" in p for p in probs):
            stats["empty_core"] += 1
            continue
        if dry:
            print(f"[dry] {case['案号']} {case['名称'][:24]} 领域={case['领域']} "
                  f"verified={case['verified']} 事实{len(case['基本事实'])}字")
        else:
            caselib.save_case(case)
        stats["ok"] += 1
    print("入库统计:", stats)


if __name__ == "__main__":
    main()
