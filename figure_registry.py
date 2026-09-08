# -*- coding: utf-8 -*-
"""图表编号 registry — 正文引用与标签编号的统一对账/重编号。

解决的问题(2026-09-08实锤): 正文引用"图5-1/表5-1"(章节式), 但标签用
id="1".."12" 顺序编号、title 不含编号 → 渲染出来的图注是"表1/图1",
正文引用全部悬空。两套编号体系没有对账方, 谁也不知道谁错。

契约:
  scheme = "chapter"     正文存在章节式引用(图C-N) → 渲染编号按"所在章+章内序号"
  scheme = "sequential"  正文只有简单式(图N)或无引用 → 维持全篇顺序编号

build_registry(text, els) → dict:
  {
    "scheme": "chapter" | "sequential",
    "labels": {start_pos: "5-1"},        # 每个标签的规范编号(供渲染端图注前缀)
    "kind_by_pos": {start_pos: "chart|table|drawing"},
    "body_refs": {"图": [(5,1),...], "表": [...]},
    "unbound_refs": ["图5-1", ...],      # 引用了但任何章都没有对应标签
    "tag_labels": {"图": ["5-1", ...]},  # 实际分配出的编号
  }
纯函数, 只读文本; 渲染端负责消费 labels, 审计端消费 unbound_refs。
"""
import re

_CH_HEAD = re.compile(r"^第\s*(\d{1,2})\s*章", re.M)
_TAG = re.compile(r"<(?:drawing|chart|table|graphic)\b[^>]*/?>", re.S)
_REF_CH = re.compile(r"([图表])\s?(\d{1,2})[-–](\d{1,3})")
_REF_SIMPLE = re.compile(r"([图表])\s?(\d{1,2})(?![\d\-–])")


def _body_only(text: str) -> str:
    """剥掉标签内部属性后做引用统计: title里的'图5-1'不是正文引用。"""
    return _TAG.sub(" ", text)


def build_registry(text: str, els) -> dict:
    """els: [(start_pos, end_pos, kind, payload)] — tolerant_extract_* 的产物。"""
    body = _body_only(text)
    refs_ch = [(m.group(1), int(m.group(2)), int(m.group(3)))
               for m in _REF_CH.finditer(body)]
    scheme = "chapter" if refs_ch else "sequential"

    # 章节边界: 每个位置 → 最近的前置章号
    heads = [(m.start(), int(m.group(1))) for m in _CH_HEAD.finditer(text)]

    def chapter_of(pos: int) -> int:
        ch = 0
        for hp, cn in heads:
            if hp < pos:
                ch = cn
            else:
                break
        return ch

    labels, kind_by_pos = {}, {}
    counters = {}  # (kind_base, chapter) → k
    tag_labels = {"图": [], "表": []}
    for s, _e, kind, _payload in els:
        base = "表" if kind == "table" else "图"
        ch = chapter_of(s)
        if scheme == "chapter" and ch >= 1:
            k = counters.get((base, ch), 0) + 1
            counters[(base, ch)] = k
            lab = f"{ch}-{k}"
        else:
            k = counters.get((base, 0), 0) + 1
            counters[(base, 0)] = k
            lab = str(k)
        labels[s] = lab
        kind_by_pos[s] = kind
        tag_labels[base].append(lab)

    # 未绑定引用: 章节式引用 (C,N) 在分配出的编号中找不到
    unbound = []
    for kind, c, n in refs_ch:
        if f"{c}-{n}" not in tag_labels.get(kind, []):
            unbound.append(f"{kind}{c}-{n}")

    return {
        "scheme": scheme,
        "labels": labels,
        "kind_by_pos": kind_by_pos,
        "body_refs": {"图": sorted({(c, n) for k, c, n in refs_ch if k == "图"}),
                      "表": sorted({(c, n) for k, c, n in refs_ch if k == "表"})},
        "unbound_refs": sorted(set(unbound)),
        "tag_labels": tag_labels,
    }


def registry_summary(reg: dict) -> str:
    if not reg:
        return ""
    return (f"编号方案={reg['scheme']} 图{len(reg['tag_labels']['图'])}"
            f"/表{len(reg['tag_labels']['表'])}"
            + (f" 未绑定{len(reg['unbound_refs'])}" if reg["unbound_refs"] else ""))
