#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""公报案例入库：解析 crawl_gazette.py 抓取的 gbz_*.json 并写入案例库。

公报页面结构（官方全文转载）：
  标题(案件名)
  【裁判摘要】 …（→裁判要点）
  当事人段 / 诉讼经过（含各级案号）
  …经审理查明：事实…
  本院认为：…理由…
  判决/裁定如下：…结果…
入库门槛: 现代案号 + 裁判摘要 + 事实段>=200字 + 结果段非空。
领域软配额: 已达配额的领域跳过(避免挤占短板领域名额)。
"""
import glob
import json
import re
import sys

import caselib
import case_taxonomy as CT

SRC = "case_library_src/gbz_*.json"
NO_PAT = re.compile(r"（(?:19|20)\d{2}）[^，。；\s（]{1,20}号")
FILE_HINT = re.compile(r"关于.{0,30}(规定|决定|办法|安排|意见|解释|通知|条例|纪要|批复)")


def pick_no(t: str):
    """优先取正文中第一个出现的最高法案号(公报多转载最高法裁判), 无则取第一个。"""
    nos = NO_PAT.findall(t)
    if not nos:
        return ""
    for n in nos[:8]:
        if "最高法" in n:
            return n
    return nos[0]


def section(t, begin_pat, end_pats, limit=4000):
    m = re.search(begin_pat, t)
    if not m:
        return ""
    s0 = m.end()
    ends = [t.find(p, s0) for p in end_pats]
    ends = [e for e in ends if e > s0]
    seg = t[s0:min(ends) if ends else min(len(t), s0 + limit)]
    return re.sub(r"\s+", " ", seg).strip()[:limit]


def parse_gazette(d: dict):
    t = d.get("text", "")
    title = (d.get("title") or "").strip() or next(
        (l.strip() for l in t.split("\n") if 12 < len(l.strip()) < 80), "")
    if not title or FILE_HINT.search(title) or len(t) < 2000:
        return None
    no = pick_no(t)
    abstract = section(t, r"【?裁判摘要】?", ["再审申请人（", "上诉人（", "申请再审", "抗诉机关",
                                               "原告", "被告人", "原审", "当事人",
                                               "最高人民法院民事", "最高人民法院刑事",
                                               "最高人民法院行政", "最高人民法院执行",
                                               "最高人民法院赔偿", "审判机关"], 1200)
    if not no or len(abstract) < 40:
        return None
    # 事实段分层回退：公报正文是裁判文书全文，事实写法多样
    facts = (section(t, r"经(?:公开)?(?:审理|法庭|再审|二审|审查)?查明[:：]?", ["本院认为", "本院经审理认为", "本院经审查认为"], 4000)
             or section(t, r"(?:一审查明|原审查明|一审判决认定|原审判决认定|原审法院认定|再审查明)[:：]?", ["本院认为", "本院经审查认为"], 4000))
    if len(facts) < 200:
        # 兜底：当事人段之后 → 本院认为 之前的叙事段
        party_pat = re.compile(r"^(再审申请人|上诉人|被上诉人|申请人|被申请人|一审|二审|原审|原告|被告|被告人|第三人|抗诉机关|法定代表人|委托诉讼代理人|诉讼代理人|负责人|指定辩护人|附带民事诉讼)[^\n]{0,80}$", re.M)
        ps = list(party_pat.finditer(t))
        think = re.search(r"本院(?:经(?:审理|审查))?认为[:：]?", t)
        if ps and think:
            facts_start = ps[-1].end()
            facts = re.sub(r"\s+", " ", t[facts_start:think.start()]).strip()[:4000]
    reason = section(t, r"本院(?:经审理|经审查)?认为[:：]?", ["判决如下", "裁定如下", "综上"], 3000)
    result = section(t, r"(?:判决|裁定)(?:如下)?[:：]?", ["审判长", "本判决为终审", "二〇", "二０",
                                                          "书记员", "书 记 员", "如不服本判决"], 800)
    if len(facts) < 200 or not result.strip():
        return None
    kw = re.findall(r"[\u4e00-\u9fff]{2,6}(?:纠纷|争议|罪|合同|侵权)", title)[:6]
    m = re.match(r"(.{2,40}?)(?:与|诉|不服)(.{2,40}?)(.*?)案$", title)
    parties = (f"原告:{m.group(1)}; 被告:{m.group(2)}" if m else title[:40])
    cause = next((k for k in kw if "纠纷" in k or "罪" in k), title[-16:])
    case = {
        "名称": title, "案号": no,
        "关键词": list(dict.fromkeys(kw))[:8],
        "基本事实": facts, "裁判结果": result,
        "裁判理由": reason, "裁判要点": abstract,
        "相关法条": "",
        "当事人": parties, "案由": cause,
        "诉讼请求": f"就{cause}请求裁判(由公报全文归纳)",
        "争议焦点": [x.strip() for x in re.split(r"[。；]", reason) if 10 < len(x.strip()) < 70][:3]
                   or [abstract.split("。")[0][:70]],
        "来源": "最高人民法院公报",
        "来源链接": d.get("url", ""),
        "来源存档": f"http://web.archive.org/web/{d.get('ts','')}/" + d.get("url", ""),
    }
    case["领域"] = CT.classify(cause, title, " ".join(case["关键词"]))
    case["tags"] = case["关键词"] + ([case["领域"]] if case["领域"] else [])
    case["verified"] = True
    case["verified_note"] = "最高人民法院公报全文(裁判摘要与案号齐备)"
    return case


def main():
    dry = "--dry" in sys.argv
    cnt = {"ok": 0, "nosec": 0, "nofield": 0, "quota": 0, "dupe": 0}
    counts = CT.stats(caselib.load_library())["by_agg"]
    for f in sorted(glob.glob(SRC)):
        try:
            d = json.load(open(f, encoding="utf-8"))
        except Exception:
            continue
        c = parse_gazette(d)
        if not c:
            cnt["nosec"] += 1
            continue
        if not c["领域"] or not all([c["名称"], c["基本事实"], c["裁判要点"], c["裁判结果"]]):
            cnt["nofield"] += 1
            continue
        if caselib.get(c["案号"]):
            cnt["dupe"] += 1
            continue
        agg = c["领域"]
        quota = CT.AGG.get(agg, (0, []))[0]
        have = counts.get(agg, {}).get("verified", 0)
        if have >= quota and "--no-quota" not in sys.argv:  # 已达配额领域让位短板领域
            cnt["quota"] += 1
            continue
        probs = caselib.validate_case(c)
        if any("必填" in p for p in probs):
            cnt["nofield"] += 1
            continue
        if dry:
            print(f"[dry] {c['案号']} {c['名称'][:30]} {agg} 要点{len(c['裁判要点'])}")
        else:
            caselib.save_case(c)
        counts.setdefault(agg, {})["verified"] = have + 1
        cnt["ok"] += 1
    print("公报入库:", cnt)


if __name__ == "__main__":
    main()
