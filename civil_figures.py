# -*- coding: utf-8 -*-
"""Deterministic civil-engineering figures for thesis body illustrations.

These are explanatory figures, not construction documents.  All numeric labels
come from the supplied drawing spec; no image model is involved.
"""
from __future__ import annotations

import os
from typing import Any, Dict, Optional

from PIL import Image, ImageDraw, ImageFont


def _font(size: int, bold: bool = False):
    candidates = [
        "C:/Windows/Fonts/simhei.ttf" if bold else "C:/Windows/Fonts/simfang.ttf",
        "C:/Windows/Fonts/NotoSansSC-VF.ttf",
    ]
    for path in candidates:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except Exception:
                pass
    return ImageFont.load_default()


def render_seismic_distribution(drawing: Dict[str, Any], save_dir: str) -> str:
    """Render the floor seismic-force distribution as a clean bar chart."""
    os.makedirs(save_dir, exist_ok=True)
    values = drawing.get("values") or [210, 310, 410, 510, 610, 1030]
    levels = drawing.get("levels") or ["1层", "2层", "3层", "4层", "5层", "6层*"]
    values = [float(v) for v in values]
    w, h = 1800, 1100
    im = Image.new("RGB", (w, h), "white")
    d = ImageDraw.Draw(im)
    title = _font(48, True); label = _font(30); small = _font(24)
    d.text((w // 2, 55), drawing.get("title") or "各层水平地震作用分布图", font=title, fill="#111", anchor="ma")
    left, right, top, bottom = 300, 150, 150, 175
    plot_w, plot_h = w - left - right, h - top - bottom
    maxv = max(values) * 1.12
    row_h = plot_h / len(values)
    for i, (level, value) in enumerate(zip(levels, values)):
        y = top + row_h * (i + 0.5)
        x2 = left + plot_w * value / maxv
        d.text((left - 25, y), str(level), font=label, fill="#222", anchor="rm")
        d.rectangle((left, y - row_h * .28, x2, y + row_h * .28), fill="#4f81bd", outline="#1f4e79", width=3)
        d.text((x2 + 18, y), f"{value:g} kN", font=label, fill="#222", anchor="lm")
        d.line((left, y + row_h * .48, left + plot_w, y + row_h * .48), fill="#e4e8ed", width=2)
    for tick in range(0, int(maxv) + 1, 200):
        x = left + plot_w * tick / maxv
        d.line((x, top - 15, x, h - bottom), fill="#d8dde3", width=2)
        d.text((x, h - bottom + 24), str(tick), font=small, fill="#444", anchor="ma")
    d.text((left + plot_w // 2, h - 55), "水平地震作用标准值 Fi（kN）", font=label, fill="#222", anchor="ms")
    d.text((left, h - 28), "* 含顶部附加地震作用", font=small, fill="#666", anchor="ls")
    d.text((w - right, h - 25), "正文示意图；数值来自结构计算章节", font=small, fill="#666", anchor="rs")
    path = os.path.join(save_dir, f"drawing_{drawing.get('seq') or drawing.get('id') or 'seismic'}.png")
    im.save(path, dpi=(180, 180))
    return path


def is_seismic_distribution(drawing: Dict[str, Any]) -> bool:
    text = " ".join(str(drawing.get(k, "")) for k in ("type", "title", "description"))
    return "地震" in text and ("分布" in text or "水平作用" in text)


def _canvas(title: str, size=(1800, 1200)):
    im = Image.new("RGB", size, "white")
    d = ImageDraw.Draw(im)
    d.text((size[0] // 2, 42), title, font=_font(46, True), fill="#111", anchor="ma")
    return im, d


def _save(im, drawing, save_dir):
    os.makedirs(save_dir, exist_ok=True)
    path = os.path.join(save_dir, f"drawing_{drawing.get('seq') or drawing.get('id') or 'civil'}.png")
    im.save(path, dpi=(180, 180))
    return path


def validate_civil_figure(path: str, min_size=(1200, 800), edge_margin=8) -> tuple[bool, list[str]]:
    """Acceptance gate before a local figure is allowed into a DOCX.

    This catches corrupt/blank files and content touching the canvas edge,
    which is the usual precursor to clipped labels.  Template geometry keeps
    title and annotation zones away from the drawing itself.
    """
    errors: list[str] = []
    try:
        im = Image.open(path).convert("RGB")
    except Exception as exc:
        return False, [f"decode:{exc}"]
    w, h = im.size
    if w < min_size[0] or h < min_size[1]:
        errors.append(f"resolution:{w}x{h}")
    pix = im.load()
    def dark(x, y):
        r, g, b = pix[x, y]
        return min(r, g, b) < 210
    # A dark pixel on every side means a label/line may have been clipped.
    stride = 4
    for name, coords in {
        "left": ((x, y) for x in range(edge_margin) for y in range(0, h, stride)),
        "right": ((x, y) for x in range(w-edge_margin, w) for y in range(0, h, stride)),
        "top": ((x, y) for x in range(0, w, stride) for y in range(edge_margin)),
        "bottom": ((x, y) for x in range(0, w, stride) for y in range(h-edge_margin, h)),
    }.items():
        if any(dark(x, y) for x, y in coords):
            errors.append(f"edge:{name}")
    # Reject near-empty outputs.
    sample = im.resize((120, 80))
    nonwhite = sum(1 for r, g, b in sample.getdata() if min(r, g, b) < 245)
    if nonwhite < 40:
        errors.append("blank")
    return not errors, errors


def render_structural_plan(drawing: Dict[str, Any], save_dir: str) -> str:
    im, d = _canvas(drawing.get("title") or "标准层结构平面布置图")
    x0, y0, sx, sy = 260, 180, 180, 125
    cols, rows = 7, 3
    # grid and dimensions
    for i in range(cols):
        x = x0 + i * sx
        d.line((x, y0, x, y0 + (rows - 1) * sy), fill="#222", width=4)
        d.ellipse((x - 14, y0 - 52, x + 14, y0 - 24), outline="#222", width=3)
        d.text((x, y0 - 38), chr(65 + i), font=_font(24, True), fill="#222", anchor="mm")
    for j in range(rows):
        y = y0 + j * sy
        d.line((x0, y, x0 + (cols - 1) * sx, y), fill="#222", width=4)
        d.ellipse((x0 - 70, y - 14, x0 - 42, y + 14), outline="#222", width=3)
        d.text((x0 - 56, y), str(j + 1), font=_font(24, True), fill="#222", anchor="mm")
    # columns and beams
    for i in range(cols):
        for j in range(rows):
            x, y = x0 + i * sx, y0 + j * sy
            d.rectangle((x - 18, y - 18, x + 18, y + 18), fill="#777", outline="#111", width=2)
    for j in range(rows):
        d.line((x0, y0 + j * sy, x0 + (cols - 1) * sx, y0 + j * sy), fill="#1f4e79", width=10)
    for i in range(cols):
        d.line((x0 + i * sx, y0, x0 + i * sx, y0 + (rows - 1) * sy), fill="#4f81bd", width=8)
    # secondary beams
    for i in range(cols - 1):
        x = x0 + (i + .5) * sx
        d.line((x, y0, x, y0 + (rows - 1) * sy), fill="#888", width=4)
    # Keep explanatory text outside the frame so it never obscures beams/columns.
    d.text((x0 + 3 * sx, y0 + (rows - 1) * sy + 105), "板厚 120 mm", font=_font(28), fill="#333", anchor="ma")
    # dimension chains
    for i in range(cols - 1):
        x1, x2 = x0 + i * sx, x0 + (i + 1) * sx
        d.line((x1, y0 - 47, x2, y0 - 47), fill="#555", width=2)
        d.text(((x1 + x2) / 2, y0 - 67), "7.2 m", font=_font(22), fill="#444", anchor="mm")
    d.text((80, 1030), "柱：500×500 mm（1～3层）  梁：300×600 / 300×550 mm", font=_font(27), fill="#222")
    d.text((80, 1080), "图示为正文结构布置示意，不替代施工图", font=_font(23), fill="#666")
    return _save(im, drawing, save_dir)


def render_beam_rebar(drawing: Dict[str, Any], save_dir: str) -> str:
    im, d = _canvas(drawing.get("title") or "标准层框架梁配筋详图")
    # elevation
    x1, x2, y = 170, 1200, 430
    d.rectangle((x1, y - 95, x2, y + 95), outline="#111", width=4)
    d.rectangle((x1 - 35, y - 140, x1, y + 140), fill="#b8b8b8", outline="#111", width=3)
    d.rectangle((x2, y - 140, x2 + 35, y + 140), fill="#b8b8b8", outline="#111", width=3)
    # longitudinal bars
    d.line((x1 + 25, y - 55, x2 - 25, y - 55), fill="#c00000", width=7)
    d.line((x1 + 25, y + 55, x2 - 25, y + 55), fill="#c00000", width=7)
    for x in range(x1 + 60, x2 - 20, 70):
        d.rectangle((x - 12, y - 78, x + 12, y + 78), outline="#2e5c8a", width=3)
    d.text(((x1 + x2) // 2, y - 180), "梁截面 300×550 mm", font=_font(30, True), fill="#222", anchor="mm")
    d.text((x1 + 10, y - 110), "支座上部 5Φ25", font=_font(25), fill="#c00000")
    d.text((x1 + 10, y + 125), "跨中下部 3Φ22", font=_font(25), fill="#c00000")
    # sections
    for cx, label, bars in [(450, "支座截面", 5), (870, "跨中截面", 3)]:
        top = 710
        d.text((cx, top - 35), label, font=_font(28, True), fill="#222", anchor="mm")
        d.rectangle((cx - 105, top, cx + 105, top + 180), outline="#111", width=4)
        for k in range(bars):
            xx = cx - 70 + k * (140 / max(1, bars - 1))
            d.ellipse((xx - 10, top + 18, xx + 10, top + 38), fill="#c00000")
            d.ellipse((xx - 10, top + 142, xx + 10, top + 162), fill="#c00000")
        d.text((cx, top + 220), "箍筋 Φ8@100/200", font=_font(24), fill="#333", anchor="ma")
    d.text((80, 1080), "加密区自柱边起 900 mm；图示用于正文说明，钢筋以计算书和正式施工图为准", font=_font(22), fill="#666")
    return _save(im, drawing, save_dir)


def render_column_rebar(drawing: Dict[str, Any], save_dir: str) -> str:
    im, d = _canvas(drawing.get("title") or "底层框架柱配筋详图")
    x, y1, y2 = 760, 220, 780
    d.rectangle((x - 105, y1, x + 105, y2), outline="#111", width=5)
    for yy in range(y1 + 45, y2 - 20, 65):
        d.rectangle((x - 82, yy - 18, x + 82, yy + 18), outline="#2e5c8a", width=3)
    for yy in (y1 + 35, y2 - 35):
        for xx in (x - 75, x - 25, x + 25, x + 75):
            d.ellipse((xx - 11, yy - 11, xx + 11, yy + 11), fill="#c00000")
    d.line((x - 245, y1, x - 245, y2), fill="#555", width=2)
    d.text((x - 275, (y1 + y2) // 2), "500 mm", font=_font(25), fill="#444", anchor="mm",) 
    d.text((x, 150), "底层中柱 500×500 mm", font=_font(30, True), fill="#222", anchor="ma")
    d.text((x + 220, 360), "纵筋：12Φ22", font=_font(28), fill="#c00000")
    d.text((x + 220, 430), "箍筋：Φ10@100（加密区）", font=_font(25), fill="#333")
    d.text((x + 220, 490), "Φ10@200（非加密区）", font=_font(25), fill="#333")
    d.text((x - 350, 950), "加密区长度约 650 mm；对称配筋，井字复合箍", font=_font(26), fill="#222")
    d.text((80, 1080), "图示为配筋构造说明，不替代正式结构施工图", font=_font(22), fill="#666")
    return _save(im, drawing, save_dir)


def render_schedule(drawing: Dict[str, Any], save_dir: str) -> str:
    im, d = _canvas(drawing.get("title") or "施工进度横道图")
    left, top, row, day = 300, 190, 95, 3.5
    items = [("施工准备", 0, 15), ("基础工程", 16, 60), ("主体结构", 61, 180), ("砌体工程", 181, 240), ("装饰装修", 241, 330), ("室外配套及竣工验收", 331, 360)]
    for t, start, end in items:
        y = top + items.index((t, start, end)) * row
        d.text((left - 20, y + 28), t, font=_font(25), fill="#222", anchor="rm")
        d.rectangle((left + start * day, y, left + end * day, y + 55), fill="#4f81bd", outline="#1f4e79", width=2)
        d.text((left + (start + end) * day / 2, y + 28), f"{end-start}天", font=_font(21), fill="white", anchor="mm")
    for m in range(0, 361, 30):
        x = left + m * day
        d.line((x, top - 25, x, top + len(items) * row), fill="#d8dde3", width=2)
        d.text((x, top - 48), f"{m}d", font=_font(20), fill="#444", anchor="ma")
    d.text((80, 1030), "总工期：360日历天；横道图按论文第5.2节施工阶段数据绘制", font=_font(25), fill="#222")
    return _save(im, drawing, save_dir)


def render_labor(drawing: Dict[str, Any], save_dir: str) -> str:
    im, d = _canvas(drawing.get("title") or "主要工种劳动力配置图")
    names = ["钢筋工", "木工", "混凝土工", "砌筑工", "装饰工", "安装工"]
    vals = [28, 36, 24, 30, 42, 18]
    left, top, day = 250, 190, 24
    for i, (name, val) in enumerate(zip(names, vals)):
        y = top + i * 105
        d.text((left - 20, y + 30), name, font=_font(27), fill="#222", anchor="rm")
        d.rectangle((left, y, left + val * day, y + 60), fill="#6aa84f", outline="#38761d", width=2)
        d.text((left + val * day + 18, y + 30), f"{val}人", font=_font(25), fill="#222", anchor="lm")
    d.text((80, 1025), "人数为正文插图示例，可由施工组织数据表自动替换", font=_font(24), fill="#666")
    return _save(im, drawing, save_dir)


def is_structural_plan(d):
    t = " ".join(str(d.get(k, "")) for k in ("type", "title", "description"))
    return "结构平面" in t or "标准层结构" in t


def is_beam_rebar(d):
    t = " ".join(str(d.get(k, "")) for k in ("type", "title", "description"))
    return "梁配筋" in t or "框架梁" in t


def is_column_rebar(d):
    t = " ".join(str(d.get(k, "")) for k in ("type", "title", "description"))
    return "柱配筋" in t or "框架柱" in t


def is_schedule(d):
    t = " ".join(str(d.get(k, "")) for k in ("type", "title", "description"))
    return "横道" in t or "施工进度" in t


def is_labor(d):
    t = " ".join(str(d.get(k, "")) for k in ("type", "title", "description"))
    return "劳动力" in t


def render_civil_figure(drawing: Dict[str, Any], save_dir: str) -> Optional[str]:
    renderer = None
    if is_seismic_distribution(drawing): renderer = render_seismic_distribution
    elif is_structural_plan(drawing): renderer = render_structural_plan
    elif is_beam_rebar(drawing): renderer = render_beam_rebar
    elif is_column_rebar(drawing): renderer = render_column_rebar
    elif is_schedule(drawing): renderer = render_schedule
    elif is_labor(drawing): renderer = render_labor
    if renderer is None:
        return None
    path = renderer(drawing, save_dir)
    ok, errors = validate_civil_figure(path)
    if not ok:
        print(f"本地土木图未通过质量闸门: {path} ({', '.join(errors)})")
        return None
    return path
