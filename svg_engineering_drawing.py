# -*- coding: utf-8 -*-
"""SVG engineering drawing backend for mechanical thesis.

Replaces the FreeCAD headless backend with a lightweight, deterministic SVG
pipeline. Supports professional engineering-style templates for fixtures,
molds, force analysis, assembly, flow, motion and parameter drawings.

Outputs are self-contained SVG files converted to PNG for DOCX insertion.
"""
from __future__ import annotations

import json
import os
import re
import subprocess
from typing import Any, Dict, List, Optional, Tuple


WIDTH = 1800
HEIGHT = 1200
MARGIN = 60
BLACK = "#000000"
GRAY = "#555555"
LIGHT = "#f0f0f0"
BLUE = "#2e5c8a"
RED = "#c00000"
GREEN = "#386e32"
PURPLE = "#6e4f9e"
ORANGE = "#a56b18"
BG = "#ffffff"
SECTION_HATCH = "#d0d0d0"


def _escape(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _load_fonts() -> str:
    return "WenQuanYi Micro Hei, WenQuanYi Zen Hei, Noto Sans CJK SC, SimHei, Source Han Sans SC, DejaVu Sans, Liberation Sans, sans-serif"


def _font(size: int, bold: bool = False, fill: str = BLACK, anchor: str = "start") -> str:
    weight = "bold" if bold else "normal"
    return (f'font-family="{_load_fonts()}" font-size="{size}" '
            f'font-weight="{weight}" fill="{fill}" text-anchor="{anchor}" '
            f'dominant-baseline="middle"')


def _rect(x: float, y: float, w: float, h: float, fill: str = "none",
          stroke: str = BLACK, stroke_width: float = 2.0, dash: str = None,
          rx: float = 0, ry: float = 0) -> str:
    attrs = [f'x="{x}"', f'y="{y}"', f'width="{w}"', f'height="{h}"',
             f'fill="{fill}"', f'stroke="{stroke}"', f'stroke-width="{stroke_width}"']
    if rx:
        attrs.append(f'rx="{rx}"')
        attrs.append(f'ry="{ry or rx}"')
    if dash:
        attrs.append(f'stroke-dasharray="{dash}"')
    return "<rect " + " ".join(attrs) + "/>"


def _line(x1: float, y1: float, x2: float, y2: float, stroke: str = BLACK,
          stroke_width: float = 2.0, dash: str = None, marker_end: bool = False,
          marker_start: bool = False) -> str:
    attrs = [f'x1="{x1}"', f'y1="{y1}"', f'x2="{x2}"', f'y2="{y2}"',
             f'stroke="{stroke}"', f'stroke-width="{stroke_width}"', 'fill="none"']
    if dash:
        attrs.append(f'stroke-dasharray="{dash}"')
    if marker_end:
        attrs.append('marker-end="url(#arrow)"')
    if marker_start:
        attrs.append('marker-start="url(#arrow_start)"')
    return "<line " + " ".join(attrs) + "/>"


def _circle(cx: float, cy: float, r: float, fill: str = "none", stroke: str = BLACK,
            stroke_width: float = 2.0) -> str:
    return f'<circle cx="{cx}" cy="{cy}" r="{r}" fill="{fill}" stroke="{stroke}" stroke-width="{stroke_width}"/>'


def _ellipse(cx: float, cy: float, rx: float, ry: float, fill: str = "none",
             stroke: str = BLACK, stroke_width: float = 2.0) -> str:
    return f'<ellipse cx="{cx}" cy="{cy}" rx="{rx}" ry="{ry}" fill="{fill}" stroke="{stroke}" stroke-width="{stroke_width}"/>'


def _text(x: float, y: float, text: str, size: int = 20, bold: bool = False,
          fill: str = BLACK, anchor: str = "start", rotate: float = None) -> str:
    content = _escape(str(text))
    attrs = f'x="{x}" y="{y}" {_font(size, bold, fill, anchor)}'
    if rotate:
        attrs += f' transform="rotate({rotate} {x} {y})"'
    return f'<text {attrs}>{content}</text>'


def _polygon(points: List[Tuple[float, float]], fill: str = "none", stroke: str = BLACK,
             stroke_width: float = 2.0) -> str:
    pts = " ".join(f"{x},{y}" for x, y in points)
    return f'<polygon points="{pts}" fill="{fill}" stroke="{stroke}" stroke-width="{stroke_width}"/>'


def _path(d: str, fill: str = "none", stroke: str = BLACK, stroke_width: float = 2.0,
          dash: str = None) -> str:
    attrs = [f'd="{d}"', f'fill="{fill}"', f'stroke="{stroke}"', f'stroke-width="{stroke_width}"']
    if dash:
        attrs.append(f'stroke-dasharray="{dash}"')
    return f'<path {" ".join(attrs)}/>'


def _hatch_rect(x: float, y: float, w: float, h: float, spacing: int = 10,
                stroke: str = SECTION_HATCH, stroke_width: float = 0.8) -> str:
    """Clip diagonal hatching inside a rectangle using a clipPath."""
    cid = f"hatch_{int(x)}_{int(y)}_{int(w)}_{int(h)}"
    clip = f'<clipPath id="{cid}"><rect x="{x}" y="{y}" width="{w}" height="{h}"/></clipPath>'
    lines = []
    start = int(x - h - spacing)
    end = int(x + w + spacing)
    for i in range(start, end, spacing):
        lines.append(f'<line x1="{i}" y1="{y}" x2="{i + h}" y2="{y + h}" stroke="{stroke}" stroke-width="{stroke_width}" clip-path="url(#{cid})"/>')
    return clip + "\n" + "\n".join(lines)


def _center_line(x1: float, y1: float, x2: float, y2: float) -> str:
    return _line(x1, y1, x2, y2, stroke=GRAY, stroke_width=1.0, dash="10 3 2 3")


def _dimension(x1: float, y1: float, x2: float, y2: float, text: str,
               offset: float = 25, text_offset: float = 8) -> List[str]:
    """Draw a dimension line with extension lines and arrows."""
    out = []
    dx = x2 - x1
    dy = y2 - y1
    length = (dx * dx + dy * dy) ** 0.5
    if length < 1:
        return out
    ux, uy = dx / length, dy / length
    px, py = -uy, ux
    # Extension lines
    out.append(_line(x1 - px * 10, y1 - py * 10, x1 + px * offset, y1 + py * offset, stroke=BLACK, stroke_width=1.2))
    out.append(_line(x2 - px * 10, y2 - py * 10, x2 + px * offset, y2 + py * offset, stroke=BLACK, stroke_width=1.2))
    # Dimension line with arrows
    mid_x = (x1 + x2) / 2 + px * offset
    mid_y = (y1 + y2) / 2 + py * offset
    out.append(_line(x1 + px * offset, y1 + py * offset, x2 + px * offset, y2 + py * offset,
                     stroke=BLACK, stroke_width=1.2, marker_start=True, marker_end=True))
    # Text
    out.append(_text(mid_x + px * text_offset, mid_y + py * text_offset, text, size=18, anchor="middle"))
    return out


def _balloon(x: float, y: float, num: int, leader_x: float, leader_y: float) -> List[str]:
    out = []
    out.append(_line(x, y, leader_x, leader_y, stroke=BLACK, stroke_width=1.2))
    out.append(_circle(x, y, 14, fill="white", stroke=BLACK, stroke_width=1.5))
    out.append(_text(x, y, str(num), size=16, bold=True, anchor="middle"))
    return out


def _drawing_blob(drawing: Dict[str, Any]) -> str:
    parts = []
    for key in ("type", "title", "description", "template", "kind", "category",
                "scene", "layout", "mech_type", "mech_object", "working_condition",
                "technical_params", "critical_params", "parameters", "dimensions",
                "forces", "annotations", "notes", "structured", "meta", "payload",
                "steps", "flow_steps", "motion_steps", "parts", "labels"):
        val = drawing.get(key)
        if val not in (None, "", [], {}):
            if isinstance(val, (dict, list, tuple, set)):
                parts.append(json.dumps(val, ensure_ascii=False))
            else:
                parts.append(str(val))
    return " ".join(parts)


def _extract_numbers(blob: str) -> List[str]:
    found = []
    patterns = [
        r'[φΦØ]\s*\d+(?:\.\d+)?\s*(?:mm|cm|m|μm|°|N|kN|MPa|rpm|r/min|s|min|%)?',
        r'\d+(?:\.\d+)?\s*(?:mm|cm|m|μm|°|N|kN|MPa|rpm|r/min|s|min|%)?',
    ]
    for pat in patterns:
        for m in re.finditer(pat, blob, flags=re.IGNORECASE):
            item = m.group(0).strip()
            if item and item not in found and len(item) > 1:
                found.append(item)
    return found[:8]


def _collect_annotations(drawing: Dict[str, Any]) -> List[str]:
    lines = []
    numeric = drawing.get("numeric_annotations")
    if isinstance(numeric, list):
        for item in numeric:
            text = str(item).strip()
            if text and text not in lines:
                lines.append(text)
    for key in ("technical_params", "critical_params", "parameters", "dimensions", "forces", "annotations"):
        val = drawing.get(key)
        if isinstance(val, dict):
            for k, v in val.items():
                text = str(v).strip()
                if text and text not in lines:
                    lines.append(f"{k}：{text}")
        elif isinstance(val, list):
            for item in val:
                if isinstance(item, dict):
                    name = str(item.get("name") or item.get("key") or "").strip()
                    raw = str(item.get("raw") or item.get("value") or "").strip()
                    if name and raw and f"{name}：{raw}" not in lines:
                        lines.append(f"{name}：{raw}")
                elif isinstance(item, str) and item.strip() and item not in lines:
                    lines.append(item.strip())
    blob = _drawing_blob(drawing)
    for num in _extract_numbers(blob):
        if num not in lines and len([l for l in lines if num in l]) == 0:
            lines.append(num)
    out = []
    for l in lines:
        if l not in out:
            out.append(l)
    return out[:8]


def _pick_template(drawing: Dict[str, Any]) -> str:
    explicit_type = re.sub(r"\s+", "", str(drawing.get("type", ""))).lower()
    blob = _drawing_blob(drawing).lower()
    if explicit_type in ("结构图", "总体布局图", "零件图"):
        if "夹具" in blob:
            return "fixture"
        if any(k in blob for k in ("模具", "冲压", "凸模", "凹模", "弯曲模", "冲裁模")):
            return "mold"
        return "structure"
    if "夹具" in blob and ("结构图" in blob or explicit_type == "结构图"):
        return "fixture"
    if any(k in blob for k in ("模具", "冲压", "凸模", "凹模", "弯曲模", "冲裁模")):
        return "mold"
    if explicit_type in ("装配示意图", "装配图", "爆炸图"):
        return "assembly"
    if explicit_type in ("受力分析图", "载荷图"):
        return "force"
    if explicit_type in ("流程图", "工艺图", "时序图"):
        return "flow"
    if explicit_type in ("运动过程图", "轨迹图"):
        return "motion"
    if explicit_type in ("参数图", "尺寸图", "标注图"):
        return "parameter"
    if any(k in blob for k in ("装配", "爆炸", "分解")):
        return "assembly"
    if any(k in blob for k in ("受力", "载荷", "应力")):
        return "force"
    if any(k in blob for k in ("流程", "步骤", "时序")):
        return "flow"
    if any(k in blob for k in ("运动", "轨迹", "位移")):
        return "motion"
    if any(k in blob for k in ("参数", "尺寸", "标注")):
        return "parameter"
    if "夹具" in blob:
        return "fixture"
    return "structure"


def _title_block(title: str) -> List[str]:
    tb_x = WIDTH - 420
    tb_y = HEIGHT - 220
    tb_w = 360
    tb_h = 160
    out = []
    out.append(_rect(tb_x, tb_y, tb_w, tb_h, fill="none", stroke=BLACK, stroke_width=1.5))
    row_h = 32
    for i in range(1, 6):
        out.append(_line(tb_x, tb_y + i * row_h, tb_x + tb_w, tb_y + i * row_h, stroke_width=1))
    col_w = 80
    out.append(_line(tb_x + col_w, tb_y, tb_x + col_w, tb_y + tb_h, stroke_width=1))
    out.append(_line(tb_x + col_w + 120, tb_y, tb_x + col_w + 120, tb_y + tb_h, stroke_width=1))
    out.append(_line(tb_x + col_w + 200, tb_y, tb_x + col_w + 200, tb_y + tb_h, stroke_width=1))
    labels = [
        (tb_x + col_w / 2, tb_y + 16, "图样名称"), (tb_x + col_w + 60, tb_y + 16, title),
        (tb_x + col_w / 2, tb_y + 48, "图样代号"), (tb_x + col_w + 60, tb_y + 48, "MJ-01"),
        (tb_x + col_w / 2, tb_y + 80, "材料"), (tb_x + col_w + 60, tb_y + 80, "45钢 / HT200"),
        (tb_x + col_w / 2, tb_y + 112, "比例"), (tb_x + col_w + 60, tb_y + 112, "1:1"),
        (tb_x + col_w / 2, tb_y + 144, "件数"), (tb_x + col_w + 60, tb_y + 144, "1"),
        (tb_x + col_w + 160, tb_y + 16, "设计"), (tb_x + col_w + 280, tb_y + 16, "Auto"),
        (tb_x + col_w + 160, tb_y + 48, "审核"), (tb_x + col_w + 280, tb_y + 48, "——"),
        (tb_x + col_w + 160, tb_y + 80, "批准"), (tb_x + col_w + 280, tb_y + 80, "——"),
        (tb_x + col_w + 160, tb_y + 112, "日期"), (tb_x + col_w + 280, tb_y + 112, "2024-06"),
    ]
    for x, y, t in labels:
        out.append(_text(x, y, t, size=15, anchor="middle"))
    return out


def _header(title: str, subtitle: str, annotations: List[str]) -> List[str]:
    out = []
    out.append(_rect(MARGIN, MARGIN, WIDTH - 2 * MARGIN, HEIGHT - 2 * MARGIN, fill="none", stroke=BLACK, stroke_width=2.5))
    out.append(_rect(MARGIN + 10, MARGIN + 10, WIDTH - 2 * MARGIN - 20, HEIGHT - 2 * MARGIN - 20, fill="none", stroke=BLACK, stroke_width=1.2, dash="8 6"))
    out.append(_text(WIDTH / 2, 105, title, size=34, bold=True, anchor="middle"))
    if subtitle:
        out.append(_text(90, 150, subtitle, size=18, fill=GRAY))
    if annotations:
        out.append(_text(90, 180, "主要参数：" + "；".join(annotations[:3]), size=16, fill=GRAY))
    return out


def _defs() -> str:
    return """<defs>
  <marker id="arrow" viewBox="0 0 10 10" refX="8.5" refY="5" markerWidth="7" markerHeight="7" orient="auto" markerUnits="strokeWidth">
    <path d="M 0 0 L 10 5 L 0 10 z" fill="#000000"/>
  </marker>
  <marker id="arrow_start" viewBox="0 0 10 10" refX="1.5" refY="5" markerWidth="7" markerHeight="7" orient="auto" markerUnits="strokeWidth">
    <path d="M 10 0 L 0 5 L 10 10 z" fill="#000000"/>
  </marker>
</defs>"""


# ==================== Fixture drawing ====================

def _render_fixture(drawing: Dict[str, Any], annotations: List[str]) -> List[str]:
    """Professional milling fixture front view."""
    out = []
    cx, cy = WIDTH // 2, HEIGHT // 2 + 40

    # Base plate (夹具体) - cast iron with hatching
    base_w, base_h = 600, 90
    base_x, base_y = cx - base_w / 2, cy + 80
    out.append(_rect(base_x, base_y, base_w, base_h, fill="#e8e8e8", stroke=BLACK, stroke_width=2.5))
    out.append(_hatch_rect(base_x, base_y, base_w, base_h, spacing=12, stroke="#999"))
    out.append(_text(cx, base_y + base_h / 2, "夹具体 HT200", size=18, anchor="middle"))

    # T-slot on base (2 slots)
    slot_w, slot_h = 40, 15
    for sx in [cx - 180, cx + 140]:
        out.append(_rect(sx, base_y + base_h - slot_h, slot_w, slot_h, fill="#ccc", stroke=BLACK, stroke_width=1.5))

    # Workpiece (矩形工件) standing vertically
    wp_w, wp_h = 120, 180
    wp_x, wp_y = cx - wp_w / 2 - 40, base_y - wp_h
    out.append(_rect(wp_x, wp_y, wp_w, wp_h, fill="#fff8e1", stroke=BLACK, stroke_width=2.5))
    out.append(_text(wp_x + wp_w / 2, wp_y + wp_h / 2, "工件", size=18, anchor="middle"))

    # V-block on left
    vb_x = wp_x - 80
    vb_y = base_y - 60
    vb_points = [(vb_x, vb_y), (vb_x + 60, vb_y), (vb_x + 60, vb_y + 60), (vb_x + 30, vb_y + 60), (vb_x + 30, vb_y + 20), (vb_x, vb_y + 60)]
    out.append(_polygon(vb_points, fill="#e0e0e0", stroke=BLACK, stroke_width=2.0))
    out.append(_hatch_rect(vb_x, vb_y, 60, 60, spacing=8, stroke="#aaa"))
    out.append(_text(vb_x + 30, vb_y + 40, "V形块", size=14, anchor="middle"))

    # Locating pin (cylinder) on right side
    pin_x = wp_x + wp_w + 35
    pin_y = base_y - 90
    pin_h = 70
    out.append(_rect(pin_x - 10, pin_y, 20, pin_h, fill="#d0d0d0", stroke=BLACK, stroke_width=2.0))
    out.append(_hatch_rect(pin_x - 10, pin_y, 20, pin_h, spacing=5, stroke="#aaa"))
    out.append(_circle(pin_x, pin_y, 8, fill="none", stroke=BLACK, stroke_width=1.5))
    out.append(_text(pin_x, pin_y - 20, "定位销", size=14, anchor="middle"))

    # Clamping plate above workpiece
    clamp_w, clamp_h = 160, 25
    clamp_x = wp_x - 20
    clamp_y = wp_y - 30
    out.append(_rect(clamp_x, clamp_y, clamp_w, clamp_h, fill="#d0d0d0", stroke=BLACK, stroke_width=2.0))
    out.append(_hatch_rect(clamp_x, clamp_y, clamp_w, clamp_h, spacing=8, stroke="#aaa"))
    out.append(_text(clamp_x + clamp_w / 2, clamp_y + clamp_h / 2, "压板", size=16, anchor="middle"))

    # Clamping bolt
    bolt_x = clamp_x + clamp_w / 2
    bolt_y_top = clamp_y - 60
    out.append(_line(bolt_x, clamp_y, bolt_x, bolt_y_top, stroke=BLACK, stroke_width=3.0))
    out.append(_rect(bolt_x - 12, bolt_y_top - 10, 24, 20, fill="#666", stroke=BLACK, stroke_width=1.5))
    out.append(_text(bolt_x + 20, bolt_y_top - 15, "M12 螺栓", size=14, anchor="start"))

    # Milling cutter on top-right
    cutter_x = wp_x + wp_w + 30
    cutter_y = wp_y - 100
    out.append(_rect(cutter_x - 15, cutter_y, 30, 80, fill="#b0b0b0", stroke=BLACK, stroke_width=2.0))
    for i in range(4):
        out.append(_line(cutter_x - 15, cutter_y + i * 20, cutter_x + 15, cutter_y + i * 20 + 10, stroke=BLACK, stroke_width=1.0))
    out.append(_text(cutter_x, cutter_y - 15, "铣刀", size=14, anchor="middle"))

    # Center line
    out.append(_center_line(cx, base_y - 250, cx, base_y + base_h + 30))

    # Dimensions
    out.extend(_dimension(wp_x, wp_y, wp_x + wp_w, wp_y, "120"))
    out.extend(_dimension(wp_x, wp_y + wp_h, wp_x, wp_y, "180"))
    out.extend(_dimension(base_x, base_y, base_x + base_w, base_y, "600", offset=45))

    # Balloons / part numbers
    out.extend(_balloon(base_x + 60, base_y - 40, 1, base_x + 60, base_y - 80))
    out.extend(_balloon(wp_x + wp_w / 2, wp_y - 50, 2, wp_x + wp_w / 2, wp_y - 90))
    out.extend(_balloon(pin_x, pin_y + pin_h + 20, 3, pin_x + 40, pin_y + pin_h + 50))
    out.extend(_balloon(clamp_x + clamp_w / 2, clamp_y - 10, 4, clamp_x + clamp_w / 2 + 60, clamp_y - 50))

    # Part list table (simple left side)
    out.append(_text(90, 780, "明细栏", size=18, bold=True))
    part_items = [
        ("1", "夹具体", "HT200", "1"),
        ("2", "工件", "45钢", "1"),
        ("3", "定位销", "T8A", "1"),
        ("4", "压板", "45钢", "1"),
    ]
    for i, (num, name, mat, qty) in enumerate(part_items):
        y = 815 + i * 30
        out.append(_rect(90, y - 12, 400, 28, fill=LIGHT, stroke=BLACK, stroke_width=1.0))
        out.append(_text(105, y, f"{num}", size=14, anchor="middle"))
        out.append(_text(150, y, name, size=14))
        out.append(_text(250, y, mat, size=14))
        out.append(_text(350, y, f"数量 {qty}", size=14))

    # Notes
    out.append(_text(90, 1040, "技术要求：1.未注倒角 C1；2.定位销与工件孔配合 H7/g6；3.夹紧后定位面贴合率≥80%", size=16, fill=GRAY))
    return out


# ==================== Mold drawing ====================

def _render_mold(drawing: Dict[str, Any], annotations: List[str]) -> List[str]:
    """Professional stamping mold front sectional view."""
    out = []
    cx, cy = WIDTH // 2, HEIGHT // 2 + 20

    # Upper die shoe
    upper_w, upper_h = 520, 70
    upper_x, upper_y = cx - upper_w / 2, cy - 220
    out.append(_rect(upper_x, upper_y, upper_w, upper_h, fill="#e0e0e0", stroke=BLACK, stroke_width=2.5))
    out.append(_hatch_rect(upper_x, upper_y, upper_w, upper_h, spacing=10, stroke="#999"))
    out.append(_text(cx, upper_y + upper_h / 2, "上模座", size=18, anchor="middle"))

    # Punch plate
    pp_w, pp_h = 360, 50
    pp_x, pp_y = cx - pp_w / 2, upper_y + upper_h
    out.append(_rect(pp_x, pp_y, pp_w, pp_h, fill="#d0d0d0", stroke=BLACK, stroke_width=2.2))
    out.append(_hatch_rect(pp_x, pp_y, pp_w, pp_h, spacing=8, stroke="#aaa"))
    out.append(_text(cx, pp_y + pp_h / 2, "凸模固定板", size=16, anchor="middle"))

    # Punch (trapezoid)
    punch_top_w = 140
    punch_bot_w = 120
    punch_h = 60
    p_top_y = pp_y + pp_h
    p_bot_y = p_top_y + punch_h
    punch_pts = [
        (cx - punch_top_w / 2, p_top_y),
        (cx + punch_top_w / 2, p_top_y),
        (cx + punch_bot_w / 2, p_bot_y),
        (cx - punch_bot_w / 2, p_bot_y),
    ]
    out.append(_polygon(punch_pts, fill="#b0b0b0", stroke=BLACK, stroke_width=2.0))
    out.append(_text(cx, p_top_y + punch_h / 2, "凸模", size=16, anchor="middle"))

    # Stripper plate
    sp_w, sp_h = 400, 30
    sp_x, sp_y = cx - sp_w / 2, p_bot_y
    out.append(_rect(sp_x, sp_y, sp_w, sp_h, fill="#f0f0f0", stroke=BLACK, stroke_width=2.0))
    out.append(_text(cx, sp_y + sp_h / 2, "卸料板", size=14, anchor="middle"))

    # Blank / workpiece
    blank_w, blank_h = 180, 40
    blank_x, blank_y = cx - blank_w / 2, sp_y + sp_h
    out.append(_rect(blank_x, blank_y, blank_w, blank_h, fill="#fff8e1", stroke=RED, stroke_width=2.0))
    out.append(_text(cx, blank_y + blank_h / 2, "板料", size=14, anchor="middle", fill=RED))

    # Die plate
    dp_w, dp_h = 360, 55
    dp_x, dp_y = cx - dp_w / 2, blank_y + blank_h
    out.append(_rect(dp_x, dp_y, dp_w, dp_h, fill="#d0d0d0", stroke=BLACK, stroke_width=2.2))
    out.append(_hatch_rect(dp_x, dp_y, dp_w, dp_h, spacing=8, stroke="#aaa"))
    out.append(_text(cx, dp_y + dp_h / 2, "凹模固定板", size=16, anchor="middle"))

    # Die cavity (V-shape cutout)
    die_top_w = 140
    die_bot_w = 120
    die_h = 45
    d_top_y = dp_y + dp_h
    d_bot_y = d_top_y + die_h
    die_pts = [
        (cx - die_top_w / 2, d_top_y),
        (cx + die_top_w / 2, d_top_y),
        (cx + die_bot_w / 2, d_bot_y),
        (cx - die_bot_w / 2, d_bot_y),
    ]
    out.append(_polygon(die_pts, fill="#ffffff", stroke=BLACK, stroke_width=2.0))
    out.append(_text(cx, d_top_y + die_h / 2, "凹模", size=14, anchor="middle"))

    # Lower die shoe
    lower_w, lower_h = 520, 80
    lower_x, lower_y = cx - lower_w / 2, d_bot_y
    out.append(_rect(lower_x, lower_y, lower_w, lower_h, fill="#e0e0e0", stroke=BLACK, stroke_width=2.5))
    out.append(_hatch_rect(lower_x, lower_y, lower_w, lower_h, spacing=10, stroke="#999"))
    out.append(_text(cx, lower_y + lower_h / 2, "下模座", size=18, anchor="middle"))

    # Guide pillars (left and right)
    for px in (cx - 210, cx + 210):
        out.append(_rect(px - 12, upper_y + 20, 24, 280, fill="#f5f5f5", stroke=BLACK, stroke_width=2.0))
        out.append(_circle(px, upper_y + 35, 6, fill="none", stroke=BLACK, stroke_width=1.5))
        out.append(_circle(px, lower_y - 25, 6, fill="none", stroke=BLACK, stroke_width=1.5))
        out.append(_text(px, cy - 20, "导柱", size=14, anchor="middle", rotate=-90))

    # Center line
    out.append(_center_line(cx, upper_y - 30, cx, lower_y + lower_h + 30))
    out.append(_center_line(cx - 300, cy, cx + 300, cy))

    # Dimensions
    out.extend(_dimension(upper_x, upper_y, upper_x + upper_w, upper_y, "520", offset=40))
    out.extend(_dimension(upper_x + upper_w + 30, upper_y, upper_x + upper_w + 30, lower_y + lower_h, "H", offset=0))

    # Notes
    out.append(_text(90, 1050, "技术要求：1.凸模与凹模间隙均匀 0.05mm；2.上下模座平行度≤0.05mm；3.导柱与导套配合H7/h6", size=16, fill=GRAY))
    return out


# ==================== Force diagram ====================

def _render_force(drawing: Dict[str, Any], annotations: List[str]) -> List[str]:
    out = []
    cx, cy = WIDTH // 2, HEIGHT // 2 + 40
    # Workpiece body
    body_w, body_h = 320, 180
    body_x, body_y = cx - body_w / 2, cy - body_h / 2
    out.append(_rect(body_x, body_y, body_w, body_h, fill="#fff8f8", stroke=BLACK, stroke_width=2.5, rx=8))
    out.append(_text(cx, cy + 5, "工件 / 夹具体", size=22, bold=True, anchor="middle"))

    # Arrows and labels
    forces = annotations[:4] or ["F1", "F2", "F3", "F4"]
    out.append(_line(body_x - 120, cy, body_x, cy, stroke=RED, stroke_width=3.0, marker_end=True))
    out.append(_text(body_x - 60, cy - 18, forces[0], size=20, bold=True, fill=RED, anchor="middle"))

    out.append(_line(body_x + body_w + 120, cy, body_x + body_w, cy, stroke=RED, stroke_width=3.0, marker_end=True))
    out.append(_text(body_x + body_w + 60, cy - 18, forces[1] if len(forces) > 1 else "F2", size=20, bold=True, fill=RED, anchor="middle"))

    out.append(_line(cx, body_y - 90, cx, body_y, stroke=RED, stroke_width=3.0, marker_end=True))
    out.append(_text(cx + 18, body_y - 45, forces[2] if len(forces) > 2 else "F3", size=20, bold=True, fill=RED, anchor="start"))

    out.append(_line(cx, body_y + body_h + 90, cx, body_y + body_h, stroke=RED, stroke_width=3.0, marker_end=True))
    out.append(_text(cx + 18, body_y + body_h + 45, forces[3] if len(forces) > 3 else "F4", size=20, bold=True, fill=RED, anchor="start"))

    # Coordinate axes
    ox, oy = body_x - 200, body_y + body_h + 80
    out.append(_line(ox, oy, ox + 80, oy, stroke=BLACK, stroke_width=2.0, marker_end=True))
    out.append(_line(ox, oy, ox, oy - 80, stroke=BLACK, stroke_width=2.0, marker_end=True))
    out.append(_text(ox + 90, oy, "X", size=16, anchor="middle"))
    out.append(_text(ox, oy - 90, "Y", size=16, anchor="middle"))

    # Notes
    out.append(_text(90, 1000, "受力分析图：标注各力大小、方向与作用点", size=18, fill=GRAY))
    return out


# ==================== Assembly / explosion ====================

def _render_assembly(drawing: Dict[str, Any], annotations: List[str]) -> List[str]:
    out = []
    cx, cy = WIDTH // 2, HEIGHT // 2 + 40
    # Central assembly
    out.append(_rect(cx - 100, cy - 60, 200, 120, fill="#f4f9f3", stroke=GREEN, stroke_width=3.0, rx=8))
    out.append(_text(cx, cy + 5, "总成", size=28, bold=True, fill=GREEN, anchor="middle"))

    parts = ["夹具体", "定位销", "压板", "螺栓", "V形块"]
    positions = [(cx - 320, cy - 140), (cx + 260, cy - 140), (cx - 320, cy + 80), (cx + 260, cy + 80), (cx - 60, cy + 180)]
    for (x, y), label in zip(positions, parts):
        out.append(_rect(x, y, 150, 70, fill=LIGHT, stroke=GREEN, stroke_width=2.2, rx=8))
        out.append(_text(x + 75, y + 35, label, size=18, anchor="middle"))
        out.append(_line(x + 75, y + 35, cx, cy, stroke=GRAY, stroke_width=1.8, dash="6 5"))
        out.append(_line(x + 75, y + 35, cx, cy, stroke=GREEN, stroke_width=2.0, marker_end=True))
    return out


# ==================== Flow / process ====================

def _render_flow(drawing: Dict[str, Any], annotations: List[str]) -> List[str]:
    out = []
    steps = annotations[:5] or ["装夹工件", "启动铣刀", "进给切削", "退刀", "松开取件"]
    x0, y0, w, h, gap = 140, 460, 220, 90, 70
    for i, step in enumerate(steps):
        x = x0 + i * (w + gap)
        y = y0 + (i % 2) * 30
        out.append(_rect(x, y, w, h, fill=BG, stroke=BLUE, stroke_width=2.5, rx=10))
        out.append(_text(x + w / 2, y + h / 2 - 10, f"{i+1}", size=18, bold=True, anchor="middle", fill=BLUE))
        out.append(_text(x + w / 2, y + h / 2 + 15, step, size=16, anchor="middle"))
        if i < len(steps) - 1:
            out.append(_line(x + w, y + h / 2, x + w + gap, y + h / 2, stroke=BLUE, stroke_width=2.2, marker_end=True))
    return out


# ==================== Motion / trajectory ====================

def _render_motion(drawing: Dict[str, Any], annotations: List[str]) -> List[str]:
    out = []
    pts = [(180, 720), (380, 650), (580, 580), (800, 500), (1020, 420), (1260, 350), (1500, 280)]
    for i in range(len(pts) - 1):
        out.append(_line(pts[i][0], pts[i][1], pts[i+1][0], pts[i+1][1], stroke=PURPLE, stroke_width=2.6, marker_end=True))
    for i, (x, y) in enumerate(pts[::2]):
        out.append(_rect(x - 24, y - 24, 48, 48, fill=BG, stroke=PURPLE, stroke_width=2.0, rx=4))
        out.append(_text(x, y, str(i + 1), size=16, bold=True, fill=PURPLE, anchor="middle"))
    out.append(_rect(640, 320, 260, 90, fill="#f7f5fc", stroke=PURPLE, stroke_width=2.5, rx=8))
    out.append(_text(770, 365, "运动主体", size=22, bold=True, fill=PURPLE, anchor="middle"))
    return out


# ==================== Parameter / dimension ====================

def _render_parameter(drawing: Dict[str, Any], annotations: List[str]) -> List[str]:
    out = []
    # Main part
    cx, cy = WIDTH // 2, HEIGHT // 2 + 40
    out.append(_rect(cx - 240, cy - 120, 480, 240, fill="#fffdf5", stroke=BLACK, stroke_width=2.6, rx=6))
    out.append(_text(cx, cy + 5, "主体零件", size=26, bold=True, anchor="middle"))

    # Callouts
    dims = annotations[:4] or ["L", "W", "H", "φd"]
    callouts = [
        ((cx - 240, cy - 120), (cx - 360, cy - 180), dims[0]),
        ((cx + 240, cy - 120), (cx + 360, cy - 180), dims[1] if len(dims) > 1 else "W"),
        ((cx + 240, cy + 120), (cx + 360, cy + 180), dims[2] if len(dims) > 2 else "H"),
        ((cx - 240, cy + 120), (cx - 360, cy + 180), dims[3] if len(dims) > 3 else "φd"),
    ]
    for s, e, lab in callouts:
        out.append(_line(s[0], s[1], e[0], e[1], stroke=ORANGE, stroke_width=2.0, marker_end=True))
        out.append(_rect(e[0] - 50, e[1] - 15, 100, 30, fill="#fff4d6", stroke=ORANGE, stroke_width=2.0, rx=8))
        out.append(_text(e[0], e[1], lab, size=18, anchor="middle"))
    return out


# ==================== Structure (generic) ====================

def _render_structure(drawing: Dict[str, Any], annotations: List[str]) -> List[str]:
    out = []
    cx, cy = WIDTH // 2, HEIGHT // 2 + 40
    # Base
    out.append(_rect(cx - 300, cy + 100, 600, 100, fill="#e8e8e8", stroke=BLACK, stroke_width=2.5, rx=6))
    out.append(_text(cx, cy + 150, "底座 / 支撑", size=20, anchor="middle"))
    # Middle
    out.append(_rect(cx - 260, cy - 60, 520, 140, fill="#f8f9fa", stroke=BLACK, stroke_width=2.5, rx=6))
    out.append(_text(cx, cy + 15, "核心机构", size=22, bold=True, anchor="middle"))
    # Top
    out.append(_rect(cx - 220, cy - 170, 440, 90, fill=LIGHT, stroke=BLACK, stroke_width=2.5, rx=6))
    out.append(_text(cx, cy - 130, "上层模块", size=20, anchor="middle"))
    # Connections
    out.append(_line(cx, cy - 80, cx, cy - 60, stroke=BLUE, stroke_width=2.2, marker_end=True))
    out.append(_line(cx, cy + 80, cx, cy + 100, stroke=BLUE, stroke_width=2.2, marker_end=True))
    out.append(_center_line(cx, cy - 220, cx, cy + 240))
    return out


# ==================== Main render ====================

def generate_engineering_svg(drawing: Dict[str, Any], out_path: str) -> str:
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    title = drawing.get("title") or drawing.get("description") or drawing.get("mech_object") or "机械工程图"
    template = _pick_template(drawing)
    annotations = _collect_annotations(drawing)
    subtitle = f"类型：{drawing.get('type', '工程图')}  |  模板：{template}"

    parts = [
        f'<?xml version="1.0" encoding="UTF-8"?>',
        # 宽高禁止带物理单位(mm)：ImageMagick 会按 -density 把 1800mm 换算成上万像素
        # (1800mm≈70.9in，density 150 → 10629px 宽=7500万像素)，后续 PIL 逐像素
        # 空白审计(_is_blank_image)会分配数 GB Python 对象，小内存机器直接 OOM。
        # 无单位时 1 用户单位=1 像素，输出即 WIDTH×HEIGHT。
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{WIDTH}" height="{HEIGHT}" viewBox="0 0 {WIDTH} {HEIGHT}">',
        _defs(),
    ]
    parts.extend(_header(title, subtitle, annotations))
    parts.extend(_title_block(title))

    if template == "fixture":
        parts.extend(_render_fixture(drawing, annotations))
    elif template == "mold":
        parts.extend(_render_mold(drawing, annotations))
    elif template == "force":
        parts.extend(_render_force(drawing, annotations))
    elif template == "assembly":
        parts.extend(_render_assembly(drawing, annotations))
    elif template == "flow":
        parts.extend(_render_flow(drawing, annotations))
    elif template == "motion":
        parts.extend(_render_motion(drawing, annotations))
    elif template == "parameter":
        parts.extend(_render_parameter(drawing, annotations))
    else:
        parts.extend(_render_structure(drawing, annotations))

    parts.append('</svg>')

    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(parts))
    return out_path


def _svg_to_png(svg_path: str, png_path: str) -> str:
    cairosvg = None
    for cand in ("cairosvg", "/root/.hermes/hermes-agent/venv/bin/cairosvg"):
        import shutil
        if shutil.which(cand):
            cairosvg = cand
            break
    if cairosvg:
        try:
            subprocess.run([cairosvg, svg_path, "-o", png_path], check=True, capture_output=True, timeout=30)
            if os.path.exists(png_path):
                return png_path
        except Exception:
            pass
    try:
        # 不传 -density：SVG 宽高已是无单位像素，1用户单位=1像素输出；
        # 传 density 会按物理尺寸放大栅格化，曾产出 10629x7085 巨图拖垮小内存机器。
        subprocess.run(["convert", "-background", "white", svg_path, png_path],
                       check=True, capture_output=True, timeout=120)
        if os.path.exists(png_path):
            _cap_png_size(png_path, max_width=2600, max_height=2000)
            return png_path
    except Exception:
        pass
    raise RuntimeError("无法将SVG转换为PNG")


def _cap_png_size(png_path: str, max_width: int = 2600, max_height: int = 2000) -> None:
    """防御性尺寸封顶：任何来源的图纸PNG超过上限就等比缩小，防止巨图进入后续管线。"""
    try:
        from PIL import Image
        with Image.open(png_path) as im:
            w, h = im.size
            if w <= max_width and h <= max_height:
                return
            scale = min(max_width / w, max_height / h)
            im = im.convert("RGB")
            im = im.resize((max(1, int(w * scale)), max(1, int(h * scale))), Image.LANCZOS)
            im.save(png_path)
            print(f"  [尺寸守卫] {os.path.basename(png_path)} {w}x{h} → {im.size[0]}x{im.size[1]}")
    except Exception as exc:
        print(f"  [尺寸守卫] 失败(保留原图): {exc}")


def generate_engineering_png(drawing: Dict[str, Any], save_dir: str) -> Optional[str]:
    os.makedirs(save_dir, exist_ok=True)
    key = drawing.get("seq") or drawing.get("id") or re.sub(r"[^\w\u4e00-\u9fff]+", "_", str(drawing.get("title") or "drawing")).strip("_") or "drawing"
    stem = f"drawing_{key}"
    svg_path = os.path.join(save_dir, f"{stem}.svg")
    png_path = os.path.join(save_dir, f"{stem}.png")
    try:
        generate_engineering_svg(drawing, svg_path)
        _svg_to_png(svg_path, png_path)
        return png_path
    except Exception as e:
        print(f"SVG工程图生成失败: {e}")
        return None


if __name__ == "__main__":
    test = {
        "id": "1",
        "seq": 1,
        "type": "结构图",
        "title": "专用铣削夹具总体结构图",
        "description": "夹具体、定位销、V形块、压板、螺栓、工件装配",
    }
    out = generate_engineering_png(test, "/tmp/mech_svg_test3")
    print("PNG:", out)
