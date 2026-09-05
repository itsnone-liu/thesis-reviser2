# -*- coding: utf-8 -*-
"""Parameter consistency gate for civil-engineering thesis figures.

The gate is deliberately conservative: it only compares values when the
article labels them as the same engineering parameter (story height, member
section, slab thickness, etc.).  Unlabelled dimensions are not treated as a
conflict because a civil paper legitimately contains many different lengths.
"""
from __future__ import annotations

import re
from collections import defaultdict
from typing import Any, Dict, List


_NUM = r"(\d+(?:\.\d+)?)"


def _norm(value: Any) -> str:
    return re.sub(r"\s+", "", str(value or "")).replace("×", "x").lower()


def _add(facts: Dict[str, List[str]], key: str, value: str) -> None:
    value = _norm(value)
    if value and value not in facts[key]:
        facts[key].append(value)


def extract_civil_facts(text: str) -> Dict[str, List[str]]:
    """Extract repeated, semantically-labelled civil design parameters."""
    facts: Dict[str, List[str]] = defaultdict(list)
    patterns = {
        "floor_count": rf"(?:总层数|地上(?:建筑)?层数|共)(?:为|：|:)?\s*{_NUM}\s*层",
        "story_height": rf"(?:层高|标准层高|首层层高)(?:为|：|:)?\s*{_NUM}\s*m",
        "slab_thickness": rf"(?:板厚|楼板厚度|屋面板厚)(?:为|：|:)?\s*{_NUM}\s*mm",
        "column_section": r"(?:柱截面|框架柱截面|柱尺寸)(?:为|：|:)?\s*(\d+(?:\.\d+)?\s*[x×]\s*\d+(?:\.\d+)?)\s*mm",
        "beam_section": r"(?:梁截面|框架梁截面|梁尺寸)(?:为|：|:)?\s*(\d+(?:\.\d+)?\s*[x×]\s*\d+(?:\.\d+)?)\s*mm",
        "concrete_grade": r"(?:混凝土强度等级|混凝土等级|采用混凝土)(?:为|：|:)?\s*(C\d+)",
        "rebar_grade": r"(?:钢筋等级|纵筋等级|箍筋等级|采用钢筋)(?:为|：|:)?\s*(HRB\d+|HPB\d+|HRB\s*\d+)",
        "building_height": rf"(?:建筑高度|结构高度)(?:为|：|:)?\s*{_NUM}\s*m",
        "construction_duration": rf"(?:总工期|计划工期|施工工期)(?:为|：|:)?\s*{_NUM}\s*(?:天|日|d)",
    }
    for key, pattern in patterns.items():
        for match in re.finditer(pattern, text, re.I):
            _add(facts, key, match.group(1))

    # Horizontal seismic-force charts often repeat Fi values.  Keep the
    # sequence as one field so it can be compared with a drawing spec.
    seismic = re.findall(r"(?:F\s*[iI]|水平地震作用)[^\d]{0,12}(\d+(?:\.\d+)?)\s*kN", text)
    if seismic:
        facts["seismic_force_values"] = list(dict.fromkeys(_norm(x) for x in seismic))
    return dict(facts)


def validate_civil_spec(spec: Dict[str, Any]) -> Dict[str, Any]:
    errors: List[Dict[str, Any]] = []
    warnings: List[Dict[str, Any]] = []
    for key, value in spec.items():
        if key in {"resolved_conflicts", "resolution_policy", "source_confidence"}:
            continue
        values = value if isinstance(value, (list, tuple, set)) else [value]
        normalized = sorted({_norm(v) for v in values if _norm(v)})
        if len(normalized) > 1:
            errors.append({"key": key, "values": normalized, "message": "same civil parameter has conflicting values"})
    if spec.get("source_confidence") == "low":
        warnings.append({"key": "source_confidence", "message": "civil spec was resolved from conflicting prose"})
    return {"ok": not errors, "errors": errors, "warnings": warnings, "spec": spec}


def gate_civil_drawing(spec: Dict[str, Any] | None = None, text: str = "") -> Dict[str, Any]:
    """Gate a complete civil drawing batch before any image is rendered."""
    spec = spec or {}
    spec_report = validate_civil_spec(spec)
    facts = extract_civil_facts(text) if text else {}
    errors = list(spec_report["errors"])
    warnings = list(spec_report["warnings"])
    resolved = set(spec.get("resolved_conflicts") or [])
    for key, values in facts.items():
        # A seismic distribution is intentionally a sequence of different Fi
        # values, not a contradiction between two statements.
        if key == "seismic_force_values":
            continue
        if len(values) <= 1:
            continue
        if key in resolved:
            warnings.append({"key": key, "message": "explicitly resolved: " + ", ".join(values)})
        else:
            errors.append({"key": key, "values": values, "message": "same parameter appears with multiple values"})
    return {"ok": not errors, "errors": errors, "warnings": warnings, "spec": spec, "facts": facts}


def format_civil_gate_report(report: Dict[str, Any]) -> str:
    lines = ["土木参数一致性闸门：" + ("通过" if report.get("ok") else "阻断")]
    for item in report.get("errors", []):
        lines.append(f"- {item.get('key', 'unknown')}: {', '.join(map(str, item.get('values', [])))}；{item.get('message', '')}")
    for item in report.get("warnings", []):
        lines.append(f"- 警告：{item.get('message', '')}")
    return "\n".join(lines)
