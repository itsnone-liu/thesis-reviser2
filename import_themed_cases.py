#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""最高法专题典型案例批次页解析入库。

页面结构（典型案例发布栏，2008-2018老格式多带真实案号）：
  案例1/案例一/<案件名>案
  【基本案情】…【执行结果/裁判结果】…【典型意义/典型意义】…
入库门槛: 案例块内出现真实（19/20xx）…号 且 基本案情+典型意义齐备；
         典型意义→裁判要点, 执行结果/裁判结果→裁判结果(可缺省回填)。
用法: import_themed_cases.py <本地html> <批次名> <url> [--dry]
"""
import html as H
import json
import re
import sys

import caselib
import case_taxonomy as CT

NO_PAT = re.compile(r"（(?:19|20)\d{2}）[^，。；\s（]{1,20}号")
SEC = ["基本案情", "裁判结果", "执行结果", "处理结果", "裁判理由", "典型意义", "案号"]


def block_text(d, sec, limit=4000):
    for pat in (r"【" + sec + r"】", r"（[一二三四五六七八九十\d]{1,3}）\s*" + sec):
        i = re.search(pat, d)
        if not i:
            continue
        ends = []
        for x in SEC:
            ends += [m.start() for m in re.finditer(r"【" + x + r"】|（[一二三四五六七八九十\d]{1,3}）\s*" + x, d[i.end():])]
        ends = [e + i.end() for e in ends if e > 0]
        seg = d[i.end():min(ends) if ends else min(len(d), i.end() + limit)]
        return re.sub(r"\s+", " ", seg).strip()[:limit]
    return ""


def parse(src, batch, url):
    t = H.unescape(re.sub(r"<script.*?</script>", "", src, flags=re.S))
    t = re.sub(r"<style.*?</style>", "", t, flags=re.S)
    t = re.sub(r"<[^>]+>", "\n", t)
    t = re.sub(r"[ \t\u3000]+", " ", t)
    t = re.sub(r"\n\s*\n+", "\n", t)
    cases = []
    # 头部三种: 一、案件名（老格式）/ 案例1（新格式）/ 裸案件名行+【栏目】（海事年批等）
    heads = list(re.finditer(r"(?:\n|^)\s*([一二三四五六七八九十]{1,3})、\s*([^\n（(【]{4,80})", t))
    if not heads:
        heads = list(re.finditer(r"案例[一二三四五六七八九十1-9]{1,3}[、:：.\s]", t))
    if not heads:
        # 以【基本案情】为锚：其前最近的"…案/…罪"长行即案件名
        anchors = list(re.finditer(r"【基本案情】", t))
        for a in anchors:
            back = t[max(0, a.start() - 300):a.start()]
            lines = [l.strip() for l in back.split("\n") if l.strip()]
            nm = next((l for l in reversed(lines)
                       if re.search(r"(?:案|罪)$", l) and 8 <= len(l) <= 90
                       and not re.match(r"（[一二三四五六七八九十\d]", l)), "")
            if nm:
                heads.append(re.search(re.escape(nm), t))
        heads = [h for h in heads if h]
    for k, m in enumerate(heads):
        end = heads[k + 1].start() if k + 1 < len(heads) else len(t)
        blk = t[m.start():end]
        facts = block_text(blk, "基本案情")
        note = block_text(blk, "典型意义")
        result = (block_text(blk, "执行结果") or block_text(blk, "裁判结果")
                  or block_text(blk, "处理结果"))
        nos = NO_PAT.findall(blk)
        name_m = re.search(r"(?:\n|^)\s*[一二三四五六七八九十]{1,3}、\s*([^\n（(【]{4,80})", blk) \
                 or re.search(r"案例[一二三四五六七八九十1-9]{1,3}[、:：.\s]*([^\n【]{6,80})", blk) \
                 or re.search(r"(?:^|\n)\s*(.{6,90}?(?:案|罪))\s*\n\s*【?基本案情", blk) \
                 or re.search(r"案例[一二三四五六七八九十1-9]{1,3}[、:：.\s]+([^【\n]{6,90}?)\s*【基本案情", blk) \
                 or re.search(r"([^\n【]{8,90})\s*【?基本案情", blk)
        name = name_m.group(1).strip() if name_m else ""
        if "——" in name:  # "要点句——案件名"格式取案件名
            parts = [p.strip() for p in name.split("——")]
            name = next((p for p in reversed(parts) if p.endswith(("案", "罪"))), name)
        if not (facts and note):
            continue
        synthetic = None
        if nos:
            no_m = NO_PAT.search(facts) or NO_PAT.search(result)
            no = no_m.group(0) if no_m else nos[0]
        elif "--synthetic" in sys.argv:
            idx = len(cases) + 1
            synthetic = idx
            no = f"最高法典型案例·{batch}·案例{idx}"
        else:
            continue
        if not name:
            name = f"{batch}(该批第{len(cases)+1}案)"
        kw = re.findall(r"[\u4e00-\u9fff]{2,6}(?:纠纷|争议|罪|合同|侵权)", name)[:6]
        cause = next((x for x in kw if "纠纷" in x or "罪" in x), name[-16:])
        m2 = re.match(r"(.{2,40}?)(?:与|诉|不服)(.{2,40}?)(.*?)案$", name)
        c = {
            "名称": name, "案号": no,
            "关键词": list(dict.fromkeys(kw))[:8],
            "基本事实": facts, "裁判结果": result or note[:300],
            "裁判理由": block_text(blk, "裁判理由"), "裁判要点": note,
            "相关法条": "",
            "当事人": (f"原告:{m2.group(1)}; 被告:{m2.group(2)}" if m2 else name[:40]),
            "案由": cause,
            "诉讼请求": f"就{cause}请求裁判(由发布文本归纳)",
            "争议焦点": [x.strip() for x in re.split(r"[。；]", note) if 10 < len(x.strip()) < 70][:3],
            "来源": f"最高人民法院{batch}",
            "来源链接": url,
        }
        c["领域"] = CT.classify(cause, name, " ".join(c["关键词"]))
        c["tags"] = c["关键词"] + ([c["领域"]] if c["领域"] else [])
        c["verified"] = True
        c["verified_note"] = (f"最高人民法院{batch}官方发布(案号在案情中列明)" if not synthetic
                              else f"最高人民法院{batch}官方发布(匿名化无案号, 按案例序号引用, 已获用户认可)")
        if synthetic:
            c["引用说明"] = f"本条为官方匿名化典型案例, 论文引用格式: 最高人民法院{batch}·案例{synthetic}"
        if c["领域"] and len(facts) >= 150:
            cases.append(c)
    return cases


def main():
    a = sys.argv
    src = open(a[1], encoding="utf8", errors="replace").read()
    batch, url = a[2], a[3]
    dry = "--dry" in a
    ok = 0
    for c in parse(src, batch, url):
        if caselib.get(c["案号"]):
            continue
        probs = caselib.validate_case(c)
        if any("必填" in p for p in probs):
            continue
        if dry:
            print(f"[dry] {c['案号']} {c['名称'][:34]} {c['领域']}")
        else:
            caselib.save_case(c)
        ok += 1
    print(f"{batch}: 入库{ok}")


if __name__ == "__main__":
    main()
