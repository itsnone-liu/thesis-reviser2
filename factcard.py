# -*- coding: utf-8 -*-
"""
factcard.py — 生成循环的事实卡与章节检查点（零LLM，确定性）
================================================================
解决长文参数漂移：章节级生成时把已确立的工程参数作为显式约束注入
后续 prompt（"查表"替代"回忆"），每章生成后跑一致性闸门，冲突带
反馈重生成本章（修复半径=一章）。

用法(生成循环内):
    from factcard import build_fact_card, checkpoint_chapter
    card = build_fact_card(accumulated_text, paper_type)   # 注入 prompt 尾部
    errs = checkpoint_chapter(accumulated_text + chapter_text, paper_type)
"""
import re

_CARD_HEADER = "\n\n【本工程已确立参数(全文一致,必须严格遵守,禁止出现不同值)】\n"
_CIV_LABELS = {
    "floor_count": "层数", "story_height_first": "首层层高(m)",
    "story_height_standard": "标准层层高(m)", "building_height": "建筑高度(m)",
    "construction_duration": "总工期(天)", "slab_thickness_floor": "楼板厚(mm)",
    "slab_thickness_roof": "屋面板厚(mm)", "concrete_grade": "混凝土等级",
    "rebar_grade": "钢筋等级", "column_section": "柱截面(mm)",
    "beam_section": "梁截面(mm)",
}
# 多值参数不进卡(分部位差异正常),只锁单值确立的
_SKIP_MULTI = {"concrete_grade", "rebar_grade", "column_section", "beam_section",
               "slab_thickness_floor", "slab_thickness_roof"}


def build_fact_card(accumulated: str, paper_type: str) -> str:
    """从已生成文本抽取参数卡。土木用 civil_consistency 抽取器;
    其他类型暂无抽取器,返回空串(无卡≠错,只是不注入)。"""
    if not accumulated or len(accumulated) < 200:
        return ""  # 首章/摘要太短,无从确立
    try:
        if "土木" in (paper_type or "") or "civil" in (paper_type or "").lower():
            from civil_consistency import extract_civil_facts
            facts = extract_civil_facts(accumulated)
            rows = []
            for k, label in _CIV_LABELS.items():
                vals = facts.get(k) or []
                if len(vals) == 1:
                    rows.append(f"- {label}: {vals[0]}")
                elif len(vals) > 1 and k not in _SKIP_MULTI:
                    # 全局唯一型参数出现多值本身就是漂移,取首值并标注
                    rows.append(f"- {label}: {vals[0]} (注意:前文曾出现{','.join(vals[1:])},以本值为准)")
            if rows:
                return _CARD_HEADER + "\n".join(rows) + "\n"
        elif "机械" in (paper_type or "") or "mech" in (paper_type or "").lower():
            from mechanical_consistency import resolve_article_spec
            spec = resolve_article_spec(accumulated)
            rows = [f"- {k}: {v}" for k, v in spec.items()
                    if isinstance(v, (str, int, float)) and v and k in (
                        "part_material", "clamp_force", "fixture_material")]
            # 0908加固: 机械正文叙事参数也进卡(machine_spec锁设计参数,正文数字仍可漂移)
            rows += _generic_param_rows(accumulated, skip={"夹紧力", "主轴转速", "电机功率"})
            if rows:
                return _CARD_HEADER + "\n".join(rows) + "\n"
        else:
            # 0908一般性加固: 通用参数卡(管理/设计/法学等) — 复用tagguard跨专业参数字典,
            # 从已生成正文回抽参数, 多值取首现值并标注漂移(与civil同构的"查表替代回忆")
            rows = _generic_param_rows(accumulated)
            if rows:
                return _CARD_HEADER + "\n".join(rows) + "\n"
    except Exception as exc:
        # 生成质量闸门不能把抽取异常伪装成“无事实”。调用方必须记录并决定是否阻断。
        raise RuntimeError(f"factcard构建失败: {type(exc).__name__}: {exc}") from exc
    return ""


def _generic_param_rows(accumulated: str, skip: set = None, max_rows: int = 14) -> list:
    """通用参数抽取(零LLM): 扫tagguard._PARAM_DICT各专业参数在accumulated中的取值,
    单值→直接列; 多值→取首现值+标注曾出现的其他值(提示LLM以首值为准)。
    skip: 已由专业抽取器覆盖的参数名(避免重复行)。"""
    skip = skip or set()
    try:
        from tagguard import _PARAM_DICT, _parse_num
    except Exception:
        return []
    seen = {}
    for pname, pat in _PARAM_DICT:
        if pname in skip:
            continue
        for m in re.finditer(pat, accumulated):
            try:
                raw = m.group(2) if pname.startswith("通用") else m.group(1)
            except Exception:
                continue
            try:
                v = _parse_num(raw)
            except Exception:
                v = None
            if v is None:
                v = raw
            seen.setdefault(pname, []).append(f"{v:g}" if isinstance(v, float) else str(v))
    rows = []
    for pname, vals in seen.items():
        uniq = list(dict.fromkeys(vals))
        if len(uniq) == 1:
            rows.append(f"- {pname}: {uniq[0]}")
        elif len(uniq) <= 4:
            rows.append(f"- {pname}: {uniq[0]} (注意:前文曾出现{','.join(uniq[1:])},以本值为准)")
        # >4个不同值视为分对象参数(如多企业/多方案对比),不锁
    return rows[:max_rows]


def checkpoint_chapter(accumulated: str, paper_type: str) -> list:
    """章节边界检查点: 返回需反馈重写的冲突描述(空=通过)。"""
    try:
        if "土木" in (paper_type or "") or "civil" in (paper_type or "").lower():
            from civil_consistency import gate_civil_drawing
            rep = gate_civil_drawing(text=accumulated)
            return [f"{e.get('key')}: {','.join(map(str, e.get('values', [])))}"
                    for e in rep.get("errors", [])]
        if "机械" in (paper_type or "") or "mech" in (paper_type or "").lower():
            from mechanical_consistency import resolve_article_spec, validate_mechanical_spec
            rep = validate_mechanical_spec(resolve_article_spec(accumulated))
            return [f"{e.get('key')}: {e.get('message', '')}"[:60]
                    for e in rep.get("errors", [])]
        # 管理/设计/法学等没有专业工程抽取器时，仍执行通用同名参数冲突检查。
        # 只报告（不自动改写），避免把年度变化、方案对比误当成固定参数。
        from tagguard import check_numeric_consistency
        _unused, rep = check_numeric_consistency(accumulated, auto_fix=False)
        return [f"{c.get('name')}: {c.get('values')}"[:100]
                for c in (rep or {}).get('conflicts', [])]
    except Exception as exc:
        # 检查异常不能伪装成“无冲突”；上层据此阻断或标记 unavailable。
        raise RuntimeError(f"章节一致性检查失败: {type(exc).__name__}: {exc}") from exc
    return []
