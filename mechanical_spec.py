# -*- coding: utf-8 -*-
"""
mechanical_spec.py — 机械 master-spec 锁定与归一化
=================================================
为机械类画像提供稳定、可复用的 canonical machine_spec。

目标：
- 将松散机械画像归一为统一 schema
- 明确 load / precision / stroke / material / working_condition / critical_params
- 对下游提供稳定只读访问函数
- 保留旧字段，确保既有 profile 生成调用方兼容
"""
from __future__ import annotations

import copy
import json
import re
from types import MappingProxyType
from typing import Any, Dict, Iterable, Mapping, MutableMapping, Optional

SCHEMA_ID = "mechanical.master_spec.v1"

_DEFAULT_MECH_TYPE = "夹具"
_DEFAULT_LOAD = {"value": 500, "unit": "N", "raw": "500N"}
_DEFAULT_PRECISION = {"value": 0.1, "unit": "mm", "raw": "0.1mm"}
_DEFAULT_STROKE = {"value": 150, "unit": "mm", "raw": "150mm"}
_DEFAULT_MATERIAL = {"value": "45钢调质", "unit": "", "raw": "45钢调质"}
_DEFAULT_WORKING_CONDITION = {"value": "中批量加工，重载间歇工况", "unit": "", "raw": "中批量加工，重载间歇工况"}


def _norm_text(value: Any, default: str = "") -> str:
    if value is None:
        return default
    text = str(value).strip()
    return text or default


def _parse_quantity(value: Any, default_unit: str = "") -> Dict[str, Any]:
    raw = _norm_text(value)
    if not raw:
        return {"value": None, "unit": default_unit, "raw": ""}

    # 常见形态：500N、0.1mm、150 mm、约500N、≈0.1mm
    m = re.search(r'([-+]?\d+(?:\.\d+)?)\s*([a-zA-Z\u4e00-\u9fa5/%°μΩ]+)?', raw)
    if not m:
        return {"value": raw, "unit": default_unit, "raw": raw}

    num_str = m.group(1)
    unit = (m.group(2) or default_unit or "").strip()
    try:
        num = float(num_str)
        if num.is_integer():
            num = int(num)
        return {"value": num, "unit": unit, "raw": raw}
    except Exception:
        return {"value": raw, "unit": unit or default_unit, "raw": raw}


def _copy_item(item: Mapping[str, Any]) -> Dict[str, Any]:
    return {k: copy.deepcopy(v) for k, v in item.items()}


def _freeze(value: Any):
    """递归冻结为只读 MappingProxyType / tuple。"""
    if isinstance(value, Mapping):
        frozen = {k: _freeze(v) for k, v in value.items()}
        return MappingProxyType(frozen)
    if isinstance(value, list):
        return tuple(_freeze(v) for v in value)
    if isinstance(value, tuple):
        return tuple(_freeze(v) for v in value)
    return value


def _pick_mech_type(profile: Mapping[str, Any]) -> str:
    text = _norm_text(profile.get("mech_type") or profile.get("mechanical_type"), "")
    if text:
        return text
    title = _norm_text(profile.get("title"), "")
    if "夹具" in title:
        return "夹具"
    if "传动" in title:
        return "传动"
    if "送料" in title:
        return "送料"
    if "搬运" in title:
        return "搬运"
    if "升降" in title:
        return "升降"
    if "优化" in title:
        return "结构优化"
    return _DEFAULT_MECH_TYPE


def _pick_mech_object(profile: Mapping[str, Any]) -> str:
    return _norm_text(profile.get("mech_object") or profile.get("company") or profile.get("object"), "某机械对象")


def _pick_working_condition(profile: Mapping[str, Any]) -> Dict[str, Any]:
    working = profile.get("working_condition")
    if isinstance(working, Mapping):
        value = _norm_text(working.get("value") or working.get("text") or working.get("raw"), "")
        unit = _norm_text(working.get("unit"), "")
        raw = _norm_text(working.get("raw"), value)
        if value or raw:
            return {"value": value or raw, "unit": unit, "raw": raw or value}
    return _parse_quantity(_norm_text(working, _DEFAULT_WORKING_CONDITION["raw"]), "")


def _pick_material(profile: Mapping[str, Any]) -> Dict[str, Any]:
    material = profile.get("material", profile.get("materials"))
    if isinstance(material, Mapping):
        value = _norm_text(material.get("value") or material.get("text") or material.get("raw"), "")
        unit = _norm_text(material.get("unit"), "")
        raw = _norm_text(material.get("raw"), value)
        if value or raw:
            return {"value": value or raw, "unit": unit, "raw": raw or value}
    raw = _norm_text(material, _DEFAULT_MATERIAL["raw"])
    return {"value": raw, "unit": "", "raw": raw}


def _pick_param(profile: Mapping[str, Any], key: str, default: Dict[str, Any], unit: str) -> Dict[str, Any]:
    params = profile.get("technical_params")
    if isinstance(params, Mapping):
        raw = params.get(key)
        if raw is not None:
            parsed = _parse_quantity(raw, unit)
            if parsed.get("value") is not None or parsed.get("raw"):
                return parsed
    raw = profile.get(key)
    if raw is not None:
        parsed = _parse_quantity(raw, unit)
        if parsed.get("value") is not None or parsed.get("raw"):
            return parsed
    return copy.deepcopy(default)


def _build_critical_params(profile: Mapping[str, Any], load: Dict[str, Any], precision: Dict[str, Any], stroke: Dict[str, Any], material: Dict[str, Any], working_condition: Dict[str, Any]) -> list:
    critical = []
    base_items = [
        ("load", "载荷", load, "technical_params.载荷"),
        ("precision", "精度", precision, "technical_params.精度"),
        ("stroke", "行程", stroke, "technical_params.行程"),
        ("material", "材料", material, "materials"),
        ("working_condition", "工作条件", working_condition, "working_condition"),
    ]
    for key, name, item, source in base_items:
        critical.append({
            "key": key,
            "name": name,
            "value": copy.deepcopy(item.get("value")),
            "unit": item.get("unit", ""),
            "raw": item.get("raw", ""),
            "source": source,
        })

    # 额外技术参数：保留原始信息，供章节计算/校核继续引用
    technical_params = profile.get("technical_params")
    if isinstance(technical_params, Mapping):
        for name, raw in technical_params.items():
            if str(name).strip() in {"载荷", "精度", "行程"}:
                continue
            parsed = _parse_quantity(raw, "")
            critical.append({
                "key": f"technical_params.{name}",
                "name": _norm_text(name),
                "value": parsed.get("value"),
                "unit": parsed.get("unit", ""),
                "raw": parsed.get("raw", _norm_text(raw)),
                "source": f"technical_params.{name}",
            })

    calc_items = profile.get("calc_items")
    if isinstance(calc_items, Iterable) and not isinstance(calc_items, (str, bytes)):
        for idx, item in enumerate(calc_items, 1):
            critical.append({
                "key": f"calc_items.{idx}",
                "name": f"校核项{idx}",
                "value": _norm_text(item),
                "unit": "",
                "raw": _norm_text(item),
                "source": f"calc_items[{idx - 1}]",
            })

    data_hints = profile.get("data_hints")
    if isinstance(data_hints, Iterable) and not isinstance(data_hints, (str, bytes)):
        for idx, item in enumerate(data_hints, 1):
            critical.append({
                "key": f"data_hints.{idx}",
                "name": f"数据线索{idx}",
                "value": _norm_text(item),
                "unit": "",
                "raw": _norm_text(item),
                "source": f"data_hints[{idx - 1}]",
            })

    return critical


def _build_legacy_profile(profile: Mapping[str, Any], spec: Dict[str, Any]) -> Dict[str, Any]:
    legacy = dict(profile)
    legacy["major"] = _norm_text(legacy.get("major"), "机械设计")
    legacy["title"] = _norm_text(legacy.get("title"), "")
    legacy["company"] = _pick_mech_object(profile)
    legacy["industry"] = _norm_text(legacy.get("industry"), "机械设计")
    legacy["mech_object"] = spec["mech_object"]
    legacy["mech_type"] = spec["mech_type"]
    legacy["working_condition"] = spec["working_condition"]["raw"] or spec["working_condition"]["value"]
    legacy["technical_params"] = {
        "载荷": spec["load"]["raw"],
        "行程": spec["stroke"]["raw"],
        "精度": spec["precision"]["raw"],
    }
    legacy["materials"] = spec["material"]["raw"]
    legacy["load"] = spec["load"]["raw"]
    legacy["precision"] = spec["precision"]["raw"]
    legacy["stroke"] = spec["stroke"]["raw"]
    legacy["material"] = spec["material"]["raw"]
    legacy["critical_params"] = copy.deepcopy(spec["critical_params"])
    legacy["machine_spec"] = copy.deepcopy(spec)
    legacy["machine_spec_schema"] = SCHEMA_ID
    return legacy


def build_mechanical_master_spec(profile: Mapping[str, Any], immutable: bool = False) -> Mapping[str, Any]:
    """
    将松散机械画像归一为 canonical master-spec。

    返回值默认是普通 dict，便于 JSON 序列化；immutable=True 时返回只读 MappingProxyType。
    """
    profile = profile or {}
    mech_object = _pick_mech_object(profile)
    mech_type = _pick_mech_type(profile)
    title = _norm_text(profile.get("title"), "")

    load = _pick_param(profile, "载荷", _DEFAULT_LOAD, "N")
    precision = _pick_param(profile, "精度", _DEFAULT_PRECISION, "mm")
    stroke = _pick_param(profile, "行程", _DEFAULT_STROKE, "mm")
    material = _pick_material(profile)
    working_condition = _pick_working_condition(profile)
    critical_params = _build_critical_params(profile, load, precision, stroke, material, working_condition)

    spec = {
        "schema_id": SCHEMA_ID,
        "locked": True,
        "paper_title": title,
        "mech_object": mech_object,
        "mech_type": mech_type,
        "load": load,
        "precision": precision,
        "stroke": stroke,
        "material": material,
        "working_condition": working_condition,
        "critical_params": critical_params,
        "units": {
            "load": load.get("unit", "N") or "N",
            "precision": precision.get("unit", "mm") or "mm",
            "stroke": stroke.get("unit", "mm") or "mm",
        },
        "source_fields": {
            "mech_object": "mech_object/company",
            "mech_type": "mech_type/mechanical_type",
            "load": "technical_params.载荷",
            "precision": "technical_params.精度",
            "stroke": "technical_params.行程",
            "material": "materials/material",
            "working_condition": "working_condition",
        },
    }

    if immutable:
        return _freeze(spec)
    return spec


def lock_mechanical_master_spec(profile: Mapping[str, Any]) -> Mapping[str, Any]:
    """稳定只读入口：下游代码调用此函数获取机械 master-spec。"""
    return build_mechanical_master_spec(profile, immutable=True)


def attach_mechanical_master_spec(profile: Mapping[str, Any]) -> Dict[str, Any]:
    """
    为既有 profile 附加 canonical machine_spec，同时保留旧字段。
    返回可 JSON 序列化的普通 dict。
    """
    spec = build_mechanical_master_spec(profile, immutable=False)
    return _build_legacy_profile(profile, spec)


def machine_spec_to_json(spec: Mapping[str, Any]) -> str:
    """便于调试/落盘的 JSON 字符串。"""
    return json.dumps(spec, ensure_ascii=False, indent=2)

