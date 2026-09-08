#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""从最高法(含知识产权法庭 ipc.court.gov.cn)批次发布页全文解析指导性案例。

页面结构（官方发布通稿附全文）：
  指导性案例26X号
  <案件名>
  （最高人民法院审判委员会讨论通过YYYY年M月D日发布）
  　关键词 …
  　裁判要点 …
  　相关法条 …
  　基本案情 …
  　裁判结果 …
  　裁判理由 …
输出到 case_library_src/gb_<no>.json，供 build_caselib.py 统一入库。
"""
import html, json, re, sys

SECTIONS = ["关键词", "裁判要点", "相关法条", "基本案情", "裁判结果", "裁判理由"]


def strip_tags(src: str) -> str:
    t = html.unescape(re.sub(r"<script.*?</script>", "", src, flags=re.S))
    t = re.sub(r"<style.*?</style>", "", t, flags=re.S)
    t = re.sub(r"<[^>]+>", "\n", t)
    return re.sub(r"\n{2,}", "\n", t)


def parse_batch(text: str, base_url: str = ""):
    """返回 [{no,name,published,关键词,裁判要点,相关法条,基本案情,裁判结果,裁判理由}]"""
    # 执行专题批次变体标签归一化（执行实施/监督要点、执行结果、执行理由）
    text = re.sub(r"执行(?:实施|监督)要点", "裁判要点", text)
    text = re.sub(r"执行结果", "裁判结果", text)
    text = re.sub(r"执行理由", "裁判理由", text)
    out = []
    # 案例块起点：指导性案例N号 后紧跟案件名行
    heads = list(re.finditer(r"指导性?案例\s*(\d{1,3})\s*号\s*\n", text))
    for idx, m in enumerate(heads):
        no = int(m.group(1))
        end = heads[idx + 1].start() if idx + 1 < len(heads) else len(text)
        block = text[m.start():end]
        # 名称：紧跟头部的第一行非空文本（去掉全角空格）
        after = text[m.end():end]
        lines = [l.strip().replace("\u3000", "") for l in after.split("\n") if l.strip()]
        if not lines:
            continue
        name = lines[0]
        # 发布日期
        pub = ""
        pm = re.search(r"（最高人民法院审判委员会讨论通过(.+?)发布）", block)
        if pm:
            pub = pm.group(1)
        # 各栏目：按“关键词/裁判要点/…”标签切段
        case = {"no": no, "name": name, "published": pub, "url": base_url}
        positions = []
        for sec in SECTIONS:
            for sm in re.finditer(rf"(?:^|\n)\s*{sec}\s*(?:[:：]|(?=\n)|(?=\s*/))", block):
                positions.append((sm.start(), sm.end(), sec))
        positions.sort()
        for k, (s0, s1, sec) in enumerate(positions):
            e0 = positions[k + 1][0] if k + 1 < len(positions) else len(block)
            body = block[s1:e0]
            # 拼接并压缩空白，保留段落分隔
            body = re.sub(r"[ \t\u3000]+", "", body)
            body = re.sub(r"\n+", "\n", body).strip("\n")
            case[sec] = body
        if case.get("基本案情") and case.get("裁判要点"):
            out.append(case)
    return out


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return
    src = open(sys.argv[1], encoding="utf-8", errors="replace").read()
    base_url = sys.argv[2] if len(sys.argv) > 2 else ""
    cases = parse_batch(strip_tags(src), base_url)
    for c in cases:
        fn = f"case_library_src/gb_{c['no']}.json"
        json.dump(c, open(fn, "w"), ensure_ascii=False, indent=1)
        print(f"OK {c['no']} {c['name'][:40]} 案情{len(c.get('基本案情',''))}字 要点{len(c.get('裁判要点',''))}字")
    print("parsed", len(cases))


if __name__ == "__main__":
    main()
