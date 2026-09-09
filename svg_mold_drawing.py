# -*- coding: utf-8 -*-
"""Standalone lightweight SVG mold drawing generator.

This module does not rely on FreeCAD or cairosvg. It generates a clean,
engineering-style mold drawing as a self-contained SVG file, suitable for
viewing in a browser or converting to PNG later.
"""
import os


def generate_mold_drawing_svg(
    out_path: str,
    title: str = "冲压模具总装图",
    mold_object: str = "U形弯曲模",
    width: int = 1600,
    height: int = 1100,
) -> str:
    """Generate a schematic mold drawing SVG and return the file path."""
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)

    svg = []
    # 宽高禁止带物理单位(mm)：ImageMagick 按 -density 换算会产出7500万像素巨图，触发 _is_blank_image 数GB分配直接OOM(2026-09-09汤圆事故)
    svg.append(f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" viewBox="0 0 {width} {height}">')
    svg.append('<defs>')
    svg.append('  <marker id="arrow" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="7" markerHeight="7" orient="auto" markerUnits="strokeWidth">')
    svg.append('    <path d="M 0 0 L 10 5 L 0 10 z" fill="#000"/>')
    svg.append('  </marker>')
    svg.append('  <marker id="arrow_start" viewBox="0 0 10 10" refX="2" refY="5" markerWidth="7" markerHeight="7" orient="auto" markerUnits="strokeWidth">')
    svg.append('    <path d="M 10 0 L 0 5 L 10 10 z" fill="#000"/>')
    svg.append('  </marker>')
    svg.append('</defs>')

    # Drawing border
    svg.append(f'<rect x="60" y="60" width="{width-120}" height="{height-120}" fill="none" stroke="#000" stroke-width="2.5"/>')
    # Inner dashed border
    svg.append(f'<rect x="72" y="72" width="{width-144}" height="{height-144}" fill="none" stroke="#000" stroke-width="1.2" stroke-dasharray="8 6"/>')

    # Title block in lower-right corner (Chinese engineering title block style)
    tb_x = width - 420
    tb_y = height - 260
    tb_w = 360
    tb_h = 200
    svg.append(f'<rect x="{tb_x}" y="{tb_y}" width="{tb_w}" height="{tb_h}" fill="none" stroke="#000" stroke-width="1.5"/>')
    # Title block grid
    rows = [40, 40, 40, 40, 40]
    cols = [80, 120, 80, 80]
    y = tb_y
    for rh in rows:
        svg.append(f'<line x1="{tb_x}" y1="{y+rh}" x2="{tb_x+tb_w}" y2="{y+rh}" stroke="#000" stroke-width="1"/>')
        y += rh
    x = tb_x
    for cw in cols:
        svg.append(f'<line x1="{x+cw}" y1="{tb_y}" x2="{x+cw}" y2="{tb_y+tb_h}" stroke="#000" stroke-width="1"/>')
        x += cw

    # Title block labels
    labels = [
        (tb_x + 40, tb_y + 28, "图样名称"), (tb_x + 140, tb_y + 28, title),
        (tb_x + 40, tb_y + 68, "图样代号"), (tb_x + 140, tb_y + 68, "MJ-01"),
        (tb_x + 40, tb_y + 108, "材料"), (tb_x + 140, tb_y + 108, "Cr12MoV / 45钢"),
        (tb_x + 40, tb_y + 148, "比例"), (tb_x + 140, tb_y + 148, "1:1"),
        (tb_x + 40, tb_y + 188, "件数"), (tb_x + 140, tb_y + 188, "1"),
        (tb_x + 200, tb_y + 28, "设计"), (tb_x + 280, tb_y + 28, "Auto"),
        (tb_x + 200, tb_y + 68, "审核"), (tb_x + 280, tb_y + 68, "——"),
        (tb_x + 200, tb_y + 108, "批准"), (tb_x + 280, tb_y + 108, "——"),
        (tb_x + 200, tb_y + 148, "日期"), (tb_x + 280, tb_y + 148, "2024-06"),
    ]
    for x, y, text in labels:
        svg.append(f'<text x="{x}" y="{y}" font-family="WenQuanYi Micro Hei, SimHei, sans-serif" font-size="16" fill="#000" text-anchor="middle" dominant-baseline="middle">{text}</text>')

    # Main title
    svg.append(f'<text x="{width/2}" y="120" font-family="WenQuanYi Micro Hei, SimHei, sans-serif" font-size="34" font-weight="bold" fill="#000" text-anchor="middle">{title}</text>')
    svg.append(f'<text x="100" y="155" font-family="WenQuanYi Micro Hei, SimHei, sans-serif" font-size="18" fill="#555">对象：{mold_object}</text>')

    # Center the mold assembly in the drawing area
    cx, cy = width // 2, (height - 260) // 2 + 60
    # Upper die shoe
    upper_shoe = (cx - 260, cy - 220, 520, 70)
    # Lower die shoe
    lower_shoe = (cx - 260, cy + 150, 520, 80)
    # Punch plate (upper)
    punch_plate = (cx - 150, cy - 150, 300, 55)
    # Die plate (lower)
    die_plate = (cx - 150, cy + 95, 300, 55)
    # Stripper plate
    stripper = (cx - 170, cy - 95, 340, 35)
    # Workpiece / blank
    blank = (cx - 90, cy - 55, 180, 90)

    # Guide pillars (left and right)
    pillar_left = (cx - 220, cy - 140, 28, 290)
    pillar_right = (cx + 192, cy - 140, 28, 290)

    # Center lines (dash-dot)
    dash_dot = '10 4 2 4'
    svg.append(f'<line x1="{cx}" y1="{cy-260}" x2="{cx}" y2="{cy+260}" stroke="#000" stroke-width="1.2" stroke-dasharray="{dash_dot}"/>')
    svg.append(f'<line x1="{cx-320}" y1="{cy}" x2="{cx+320}" y2="{cy}" stroke="#000" stroke-width="1.2" stroke-dasharray="{dash_dot}"/>')

    # Draw parts with hatching for metal sections
    def draw_part(x, y, w, h, fill, stroke_w=2.2, label=None, label_y=None):
        svg.append(f'<rect x="{x}" y="{y}" width="{w}" height="{h}" fill="{fill}" stroke="#000" stroke-width="{stroke_w}"/>')
        # Section hatching (45-degree lines)
        hatch_spacing = 12
        for i in range(-int(h / hatch_spacing) - 10, int(w / hatch_spacing) + 10):
            x1 = x + i * hatch_spacing
            y1 = y
            x2 = x1 + h
            y2 = y + h
            # Clip to rect using a simple line with stroke-dasharray trick not possible; just draw clipped lines via SVG line with reasonable bounds
            # Use a mask-free approach: draw only lines whose segment intersects the rectangle.
            # For simplicity, draw full diagonal lines and rely on rect fill overlay? No, lines on top.
            # Better: draw short segments inside the rect.
            for j in range(0, int(h) + 1, hatch_spacing):
                px = x1 + j
                py = y1 + j
                if x <= px <= x + w and y <= py <= y + h:
                    seg_len = min(w, h)
                    svg.append(f'<line x1="{px}" y1="{py}" x2="{px + seg_len}" y2="{py + seg_len}" stroke="#000" stroke-width="0.6" opacity="0.35"/>')
        if label:
            ly = label_y or (y + h / 2)
            svg.append(f'<text x="{x + w/2}" y="{ly}" font-family="WenQuanYi Micro Hei, SimHei, sans-serif" font-size="18" fill="#000" text-anchor="middle" dominant-baseline="middle">{label}</text>')

    draw_part(*upper_shoe, "#e8e8e8", label="上模座")
    draw_part(*punch_plate, "#d0d0d0", label="凸模固定板")
    draw_part(*stripper, "#f0f0f0", stroke_w=1.8, label="卸料板")
    # Blank (workpiece) in red-ish outline
    svg.append(f'<rect x="{blank[0]}" y="{blank[1]}" width="{blank[2]}" height="{blank[3]}" fill="#fff8e1" stroke="#c00" stroke-width="2"/>')
    svg.append(f'<text x="{blank[0] + blank[2]/2}" y="{blank[1] + blank[3]/2}" font-family="WenQuanYi Micro Hei, SimHei, sans-serif" font-size="16" fill="#c00" text-anchor="middle" dominant-baseline="middle">板料</text>')
    draw_part(*die_plate, "#d0d0d0", label="凹模固定板")
    draw_part(*lower_shoe, "#e8e8e8", label="下模座")

    # Guide pillars (simple cylinders with section fill)
    for px, py, pw, ph in [pillar_left, pillar_right]:
        svg.append(f'<rect x="{px}" y="{py}" width="{pw}" height="{ph}" fill="#f5f5f5" stroke="#000" stroke-width="2"/>')
        # Add a small circle top/bottom to represent dowel holes
        svg.append(f'<circle cx="{px + pw/2}" cy="{py + 20}" r="8" fill="none" stroke="#000" stroke-width="1.5"/>')
        svg.append(f'<circle cx="{px + pw/2}" cy="{py + ph - 20}" r="8" fill="none" stroke="#000" stroke-width="1.5"/>')
    svg.append(f'<text x="{pillar_left[0] + pillar_left[2]/2}" y="{pillar_left[1] + pillar_left[3]/2}" font-family="WenQuanYi Micro Hei, SimHei, sans-serif" font-size="16" fill="#000" text-anchor="middle" dominant-baseline="middle" transform="rotate(-90 {pillar_left[0] + pillar_left[2]/2} {pillar_left[1] + pillar_left[3]/2})">导柱</text>')

    # Punch (trapezoid) inside stripper
    punch_pts = [
        (cx - 50, cy - 95),
        (cx + 50, cy - 95),
        (cx + 40, cy - 55),
        (cx - 40, cy - 55),
    ]
    pts_str = " ".join(f"{x},{y}" for x, y in punch_pts)
    svg.append(f'<polygon points="{pts_str}" fill="#b0b0b0" stroke="#000" stroke-width="2"/>')
    svg.append(f'<text x="{cx}" y="{cy - 78}" font-family="WenQuanYi Micro Hei, SimHei, sans-serif" font-size="14" fill="#000" text-anchor="middle">凸模</text>')

    # Die cavity (V/U shape) inside die plate
    die_pts = [
        (cx - 55, cy + 95),
        (cx + 55, cy + 95),
        (cx + 45, cy + 150),
        (cx - 45, cy + 150),
    ]
    pts_str = " ".join(f"{x},{y}" for x, y in die_pts)
    svg.append(f'<polygon points="{pts_str}" fill="#ffffff" stroke="#000" stroke-width="2"/>')
    svg.append(f'<text x="{cx}" y="{cy + 128}" font-family="WenQuanYi Micro Hei, SimHei, sans-serif" font-size="14" fill="#000" text-anchor="middle">凹模</text>')

    # Dimension annotations
    # Overall height
    dim_x = cx + 300
    svg.append(f'<line x1="{dim_x}" y1="{upper_shoe[1]}" x2="{dim_x}" y2="{lower_shoe[1] + lower_shoe[3]}" stroke="#000" stroke-width="1.2"/>')
    svg.append(f'<line x1="{dim_x - 6}" y1="{upper_shoe[1]}" x2="{dim_x + 6}" y2="{upper_shoe[1]}" stroke="#000" stroke-width="1.2"/>')
    svg.append(f'<line x1="{dim_x - 6}" y1="{lower_shoe[1] + lower_shoe[3]}" x2="{dim_x + 6}" y2="{lower_shoe[1] + lower_shoe[3]}" stroke="#000" stroke-width="1.2"/>')
    svg.append(f'<text x="{dim_x + 12}" y="{cy}" font-family="WenQuanYi Micro Hei, SimHei, sans-serif" font-size="18" fill="#000" text-anchor="start" dominant-baseline="middle">H = 450 mm</text>')

    # Width dimension
    dim_y = lower_shoe[1] + lower_shoe[3] + 40
    svg.append(f'<line x1="{lower_shoe[0]}" y1="{dim_y}" x2="{lower_shoe[0] + lower_shoe[2]}" y2="{dim_y}" stroke="#000" stroke-width="1.2" marker-end="url(#arrow)" marker-start="url(#arrow_start)"/>')
    svg.append(f'<line x1="{lower_shoe[0]}" y1="{lower_shoe[1] + lower_shoe[3]}" x2="{lower_shoe[0]}" y2="{dim_y}" stroke="#000" stroke-width="1" stroke-dasharray="4 4"/>')
    svg.append(f'<line x1="{lower_shoe[0] + lower_shoe[2]}" y1="{lower_shoe[1] + lower_shoe[3]}" x2="{lower_shoe[0] + lower_shoe[2]}" y2="{dim_y}" stroke="#000" stroke-width="1" stroke-dasharray="4 4"/>')
    svg.append(f'<text x="{cx}" y="{dim_y + 8}" font-family="WenQuanYi Micro Hei, SimHei, sans-serif" font-size="18" fill="#000" text-anchor="middle" dominant-baseline="hanging">W = 520 mm</text>')

    # Thickness annotations (punch plate)
    svg.append(f'<line x1="{punch_plate[0] + punch_plate[2] + 30}" y1="{punch_plate[1]}" x2="{punch_plate[0] + punch_plate[2] + 30}" y2="{punch_plate[1] + punch_plate[3]}" stroke="#000" stroke-width="1.2" marker-end="url(#arrow)" marker-start="url(#arrow_start)"/>')
    svg.append(f'<text x="{punch_plate[0] + punch_plate[2] + 42}" y="{punch_plate[1] + punch_plate[3]/2}" font-family="WenQuanYi Micro Hei, SimHei, sans-serif" font-size="16" fill="#000" text-anchor="start" dominant-baseline="middle">55</text>')

    # Note block
    note_y = 240
    notes = [
        "技术要求：",
        "1. 未注倒角 C1；",
        "2. 未注圆角 R2；",
        "3. 上下模座平行度 ≤ 0.05 mm；",
        "4. 凸模与凹模间隙均匀，单面间隙 0.05 mm。",
    ]
    for i, note in enumerate(notes):
        svg.append(f'<text x="100" y="{note_y + i * 28}" font-family="WenQuanYi Micro Hei, SimHei, sans-serif" font-size="16" fill="#000">{note}</text>')

    svg.append('</svg>')

    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(svg))
    return out_path


if __name__ == "__main__":
    out = generate_mold_drawing_svg(
        "/root/project/workspace/thesis-reviser/output/mold_drawing.svg",
        title="U形弯曲模总装图",
        mold_object="U形弯曲模",
    )
    print(f"SVG mold drawing saved to: {out}")
