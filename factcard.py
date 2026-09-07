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
            if rows:
                return _CARD_HEADER + "\n".join(rows) + "\n"
    except Exception:
        pass
    return ""


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
    except Exception:
        pass
    return []
