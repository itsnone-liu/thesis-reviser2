# -*- coding: utf-8 -*-
"""
rescue_refs.py — 空承诺/裸意图的确定性补偿器(显式调用,审计层只报告不改写)
==========================================================================
对 audit_txt 报出的可补偿问题自动构造标签:
  R1 裸标题+括号参数(严涛4.4形态): 标题行+参数行 → <drawing/> 标签(替换原两行)
  R2 指代句含完整图描述: "……图中标注A、B、C" 且同节无锚点 → 在指代段后构造标签
     (title=从节名推断, description=指代句上下文)——仅当上下文素材充分,否则跳过
不编造: description 必须来自原文文字,无素材就不补偿(留给人工)。
"""
import re

RE_BARE_TITLE2 = re.compile(r"^([\u4e00-\u9fa5]{2,14}(?:示意图|布置图|剖面图|大样图|流程图|横道图))$", re.M)
RE_PAREN2 = re.compile(r"^\((.+)\)\s*$", re.M)


def rescue(txt: str) -> tuple:
    """返回 (新文本, 动作清单)。只做有素材的补偿。"""
    actions = []
    lines = txt.splitlines(keepends=True)
    out = []
    i = 0
    while i < len(lines):
        line = lines[i].rstrip("\n")
        m = RE_BARE_TITLE2.match(line.strip())
        if m and i + 1 < len(lines):
            nxt = lines[i + 1].strip()
            pm = RE_PAREN2.match(nxt)
            if pm and "图中" in pm.group(1):
                title = m.group(1)
                desc = pm.group(1)
                ch = _chapter_num(txt, sum(len(l) for l in out))
                seq = _next_seq(txt, ch)
                cap = f"图{ch}-{seq} {title}({desc})" if len(desc) <= 60 else f"图{ch}-{seq} {title}"
                tag = (f'<drawing id="r{ch}{seq}" type="工程示意图" title="{cap}" '
                       f'description="{desc[:180]}" />' + "\n")
                out.append(tag)
                actions.append(f"R1: '{title}' 裸标题+括号 → {cap[:40]}")
                i += 2
                continue
        out.append(lines[i])
        i += 1
    return "".join(out), actions


def _chapter_num(txt: str, pos: int) -> str:
    heads = [(m.start(), m.group(1)) for m in re.finditer(r"第\s*(\d+)\s*章", txt)]
    ch = "1"
    for p, n in heads:
        if p <= pos:
            ch = n
    return ch


def _next_seq(txt: str, ch: str) -> int:
    seqs = [int(s) for s in re.findall(rf"图\s?{ch}[-–](\d{{1,3}})", txt)]
    return (max(seqs) + 1) if seqs else 1


if __name__ == "__main__":
    import sys
    t = open(sys.argv[1], encoding="utf-8").read()
    t2, acts = rescue(t)
    for a in acts:
        print("补偿:", a)
    if acts and len(sys.argv) > 2 and sys.argv[2] == "--write":
        open(sys.argv[1], "w", encoding="utf-8").write(t2)
        print("已写回")
