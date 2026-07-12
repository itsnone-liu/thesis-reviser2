# -*- coding: utf-8 -*-
"""Lightweight SVG backend for mechanical engineering diagrams.

Replaces the FreeCAD headless backend with a deterministic, dependency-light
SVG pipeline. Engineering drawings (structure, force, assembly, flow, motion,
parameter, mold) are rendered as SVG and converted to PNG for DOCX insertion.

Effect / rendered images for non-engineering purposes still go through the GPT
image API in core.py.
"""
from __future__ import annotations

import os
import re
from typing import Any, Dict, Optional

try:
    from svg_engineering_drawing import generate_engineering_png as _svg_generate_png
except Exception as exc:  # pragma: no cover
    _svg_generate_png = None
    print(f"[mechanical_cad] SVG backend import failed: {exc}")


_MECH_HINTS = (
    "mechanical", "机械", "夹具", "机构", "零件", "部件", "装配", "受力", "载荷",
    "应力", "行程", "参数", "尺寸", "标注", "公差", "motion", "force", "assembly",
    "parameter", "运动", "位移", "传动", "送料", "搬运", "升降", "动作", "工装",
)

_EFFECT_HINTS = (
    "效果图", "外观", "渲染", "场景", "展示图", "概念图", "室内", "空间", "材质", "软装",
)


def _drawing_blob(drawing: Dict[str, Any]) -> str:
    parts = []
    for key in (
        "type", "title", "description", "template", "kind", "category",
        "scene", "layout", "mech_type", "mech_object", "working_condition",
        "technical_params", "critical_params", "parameters", "dimensions",
        "forces", "annotations", "notes", "structured", "meta", "payload",
        "steps", "flow_steps", "motion_steps", "parts", "labels",
    ):
        val = drawing.get(key)
        if val not in (None, "", [], {}):
            if isinstance(val, (dict, list, tuple, set)):
                parts.append(str(val))
            else:
                parts.append(str(val))
    return " ".join(parts)


def _is_mechanical_candidate(drawing: Dict[str, Any]) -> bool:
    if any(drawing.get(k) not in (None, "", [], {}) for k in (
        "mech_type", "mech_object", "working_condition", "technical_params",
        "critical_params", "parameters", "dimensions", "forces",
    )):
        return True
    blob = _drawing_blob(drawing).lower()
    if any(k in blob for k in _EFFECT_HINTS):
        return False
    return any(k.lower() in blob for k in _MECH_HINTS)


def _safe_name(text: str) -> str:
    text = re.sub(r"[^\w\u4e00-\u9fff]+", "_", text or "drawing")
    text = re.sub(r"_+", "_").strip("_")
    return text or "drawing"


def freecad_available() -> bool:
    """SVG backend is always available once imported."""
    return _svg_generate_png is not None


def generate_engineering_png(drawing: Dict[str, Any], save_dir: str) -> Optional[str]:
    """Generate an engineering-style PNG using the SVG backend."""
    if not drawing or not _is_mechanical_candidate(drawing):
        return None
    if _svg_generate_png is None:
        return None
    return _svg_generate_png(drawing, save_dir)


def generate_mechanical_cad_image(drawing: Dict[str, Any], save_dir: str) -> Optional[str]:
    """Alias for the SVG engineering drawing backend."""
    return generate_engineering_png(drawing, save_dir)
