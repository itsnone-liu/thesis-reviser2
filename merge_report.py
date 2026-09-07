# -*- coding: utf-8 -*-
"""汇总复核器 (2026-09-07) — 合并确定性审计 + LLM 语义审计 → 最终验收报告。

LLM 报的 hard_flaws 做程序复核: 提取矛盾描述中的数字, 逐个在全文(docx XML)
验证存在性 — 矛盾双方数字都真实出现才算"确证", 否则标"存疑"(防片段幻觉)。

用法: python3 merge_report.py
产物: 论文修订档案/审计总表_0908.csv + 论文修订档案/审计总报告_0908.md
"""
import os, re, csv, json, zipfile
from collections import Counter
from docx import Document

BASE = os.path.dirname(os.path.abspath(__file__))
ROOT = "论文终版"
DEEP = "论文修订档案/审计报告_0908深度.json"
LLM = "论文修订档案/审计LLM_0908.json"
OUT_CSV = "论文修订档案/审计总表_0908.csv"
OUT_MD = "论文修订档案/审计总报告_0908.md"

_doc_cache = {}

def doc_text(rel):
    if rel in _doc_cache:
        return _doc_cache[rel]
    z = zipfile.ZipFile(os.path.join(ROOT, rel))
    txt = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>",
                             z.read("word/document.xml").decode("utf-8", "replace")))
    txt += "".join(c.text for t in Document(os.path.join(ROOT, rel)).tables
                   for row in t.rows for c in row.cells)
    _doc_cache[rel] = txt
    return txt


def verify_hard_flaws(rel, flaws):
    """对每条 hard_flaw: 提取数字, 逐个在全文验证存在性(数字边界匹配, 防'70'命中'170')。
    矛盾涉及的数字**全部**真实出现才算"确证", 否 → 存疑(防LLM幻觉数字)。"""
    txt = doc_text(rel)

    def num_in(n):
        n = n.rstrip("%")
        # 边界匹配: 前后都不能是数字/小数点(允许千分位逗号变体)
        variants = {n}
        if "." not in n and len(n) >= 4:
            variants.add(f"{int(n):,}")
        for v in variants:
            for m in re.finditer(re.escape(v), txt):
                s, e = m.start(), m.end()
                if (s == 0 or not (txt[s-1].isdigit() or txt[s-1] in ".,，")) and \
                   (e >= len(txt) or not (txt[e].isdigit() or txt[e] in ".,，")):
                    return True
        return False

    out = []
    for fl in flaws or []:
        nums = re.findall(r"\d+\.\d+|\d{3,4}|\d{1,2}%", fl)
        nums = [n for n in nums if len(n.rstrip("%")) >= 3 or "%" in n]
        miss = [n for n in nums[:8] if not num_in(n)]
        status = "存疑" if miss else "确证"
        out.append({"text": fl, "status": status, "missing": miss, "nums": nums[:8]})
    return out


def main():
    deep = {r["file"]: r for r in json.load(open(DEEP, encoding="utf-8"))}
    llm = json.load(open(LLM, encoding="utf-8"))
    rows, confirmed, suspect, warns = [], [], [], []
    for rel, dr in deep.items():
        lr = llm.get(rel, {})
        lverdict = lr.get("verdict", "未审")
        flaws_v = verify_hard_flaws(rel, lr.get("hard_flaws")) if lr.get("hard_flaws") else []
        n_conf = sum(1 for f in flaws_v if f["status"] == "确证")
        n_susp = sum(1 for f in flaws_v if f["status"] == "存疑")
        # 最终分级: 返修=确证硬伤; 复核=题目匹配可疑/跑题章/确定性❌/有存疑线索;
        # 通过=其余(字数略少/引用式样/无英文摘要等共性提示不压级, 见报告第三节)
        if n_conf:
            final = "返修"
        elif (isinstance(lr.get("topic_match"), (int, float)) and lr["topic_match"] <= 7) \
                or lr.get("offtopic") \
                or any(p.startswith("❌") for p in dr["problems"]) \
                or n_susp:
            final = "复核"
        else:
            final = "通过"
        row = {
            "文件": rel, "最终": final,
            "确定性": dr["verdict"], "确定性问题": "; ".join(dr["problems"]),
            "LLM": lverdict, "题目匹配": lr.get("topic_match", ""),
            "摘要质量": lr.get("abstract_quality", ""), "结论呼应": lr.get("conclusion_echo", ""),
            "确证硬伤": json.dumps([f["text"] for f in flaws_v if f["status"] == "确证"], ensure_ascii=False),
            "存疑线索": json.dumps([f["text"] for f in flaws_v if f["status"] == "存疑"], ensure_ascii=False),
            "弱提示": json.dumps(lr.get("weak_notes", []), ensure_ascii=False),
            "文献问题": json.dumps(lr.get("ref_issues", []), ensure_ascii=False),
            "跑题章": json.dumps(lr.get("offtopic", []), ensure_ascii=False),
            "总评": lr.get("summary", ""),
        }
        rows.append(row)
        if n_conf:
            confirmed.append(rel)
        elif n_susp or lverdict in ("fail", "warn"):
            suspect.append(rel)
        elif final == "复核":
            warns.append(rel)

    # ---- CSV ----
    with open(OUT_CSV, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)

    # ---- Markdown 报告 ----
    cnt = Counter(r["最终"] for r in rows)
    zero_cite = [r["文件"] for r in rows if "文献未被引用" in r["确定性问题"]]
    short = [r["文件"] for r in rows if "字略少" in r["确定性问题"]]
    lines = [
        "# 无源论文深度审计总报告 (2026-09-07/08)",
        "",
        f"审计范围: 25/26两届无过程文件论文 **{len(rows)} 篇** (有源新生成版 42 篇不在本次范围)",
        "",
        "## 总体结论",
        "",
        f"- **返修级(确证硬伤): {cnt.get('返修', 0)} 篇**",
        f"- **复核级(疑点待人工): {cnt.get('复核', 0)} 篇**",
        f"- **通过级: {cnt.get('通过', 0)} 篇**",
        "",
        "审计维度: 确定性9类(目录完整性/标题编号/表注表格对账/中英摘要/docx健康[修订痕迹/批注/TOC域/裸URL]/",
        "篇内重复段/引用文献对账/全篇字数/图表编号引用+空白图+占比) + LLM语义6项(题目匹配/跑题章/摘要质量/",
        "结论呼应/文献异常/硬伤) + hard_flaw数字全文复核。",
        "",
        "## 一、返修级清单(确证硬伤 — LLM报告的矛盾数字已全文验证存在)",
        "",
    ]
    for r in rows:
        if r["最终"] == "返修":
            lines.append(f"### {r['文件']}")
            for fl in json.loads(r["确证硬伤"]):
                lines.append(f"- ❌ {fl}")
            if r["跑题章"] != "[]":
                bits.append("跑题:" + "; ".join(x[:50] for x in json.loads(r["跑题章"]))[:120])
            if r["存疑线索"] != "[]":
                for fl in json.loads(r["存疑线索"]):
                    lines.append(f"- ⚠️存疑: {fl[:120]}")
            lines.append(f"- 总评: {r['总评']}")
            lines.append("")
    lines += ["## 二、复核级清单(疑点待人工抽查)", ""]
    for r in rows:
        if r["最终"] == "复核":
            bits = []
            tm = r.get("题目匹配")
            if isinstance(tm, (int, float)) and tm <= 7:
                bits.append(f"题目匹配{tm}分")
            if r["跑题章"] != "[]":
                bits.append("跑题:" + "; ".join(x[:50] for x in json.loads(r["跑题章"]))[:120])
            if r["存疑线索"] != "[]":
                bits.append("存疑:" + "; ".join(x[:60] for x in json.loads(r["存疑线索"]))[:150])
            if r["文献问题"] != "[]":
                bits.append("文献:" + json.loads(r["文献问题"])[0][:60])
            if r["确定性问题"] and "❌" in r["确定性问题"]:
                bits.append("确定性:" + r["确定性问题"][:60])
            lines.append(f"- {r['文件']} — {' | '.join(bits) if bits else r['总评']}")
    lines += ["", "## 三、全批共性特征(非个别缺陷, 供学位系统对账口径参考)", "",
              f"- 正文无编号式引用[N]且无作者-年份式引用: {len(zero_cite)} 篇 (有作者-年份式引用 78 篇)",
              f"- 正文字数略少于口径下限(8000/10000, 已含表格): {len(short)} 篇 — 自考专升本实际体量通常 8000-11000, 该项为边界提示",
              "- 无英文摘要为本批统一特征(抽样10篇9篇无), 未计问题", ""]
    open(OUT_MD, "w", encoding="utf-8").write("\n".join(lines))
    print(f"总表: {OUT_CSV}")
    print(f"报告: {OUT_MD}")
    print(f"返修 {cnt.get('返修',0)} | 复核 {cnt.get('复核',0)} | 通过 {cnt.get('通过',0)}")


if __name__ == "__main__":
    main()
