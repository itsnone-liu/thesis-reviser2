# -*- coding: utf-8 -*-
"""
机械设计类论文提示词模板
按章节分类，重点解决：
1. 机械设计论文结构模板化，便于程序生成
2. 文字与普通示意图/图片一致
3. 计算与校核章节必须有参数、公式和结论回扣
"""
import sys, os, json
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from core import MC_WORD_LIMITS

FMT_RULE = """
【格式铁律 - 必须严格遵守】
1. 严禁使用任何markdown符号：# ** * ` > | --- |
2. 严禁使用加粗标记，所有文字统一纯文本
3. 段落之间用换行分隔，不要用空行
4. 正文中涉及的结构维度、对比项、分类或参数项，要保持完整，不得为了简化删减
5. <drawing/> 标签必须与正文分离，单独成行，不能混在句子中
6. <table/> 标签必须包含完整属性，格式如下：
   对比表（5列表头 -> data每行4个值）：
   <table id="1" title="方案对比表" header="方案,定位精度,夹紧可靠性,操作便捷性,适用性" rows="方案一,方案二,方案三" data="高|中|低|好;中|高|高|一般;低|高|中|差" data_source="设计分析"/>
   
   参数表（4列表头 -> data每行3个值：符号|数值|单位）：
   <table id="2" title="铣削力参数表" header="参数名称,符号,数值,单位" rows="主切削力,进给力,背向力" data="Fc|1200|N;Ff|400|N;Fp|600|N" data_source="设计计算"/>
   
   注意：
   - data 中每行用分号 ; 分隔
   - 每行内的单元格用竖线 | 分隔
   - data 中每行数据列数必须等于 header 列数减 1（因为第一列由 rows 提供）
   - 严禁 data 中每行只有1-2个值，必须使用竖线 | 分隔成 (header列数-1) 个值
"""

SCENARIO_LOCK_RULE = """
【场景锁定 - 必须严格遵守】
1. 全文只能围绕当前题目中的同一个机械对象展开，不得中途切换到其他设备、机构、工艺或应用场景
2. 章节1~6必须前后呼应，所有零件、机构、受力、计算和图纸都要服务于同一套方案
3. 不得把当前题目写成另一个机械课题，例如卡盘/顶尖/三爪卡盘/尾座/车床主轴/铣床导向等无关对象
4. 不得同时混用两套不同工艺链；如果题目是夹具设计，就始终写夹具，不要切换成传动、送料、搬运或升降系统
5. 图纸标题、图纸描述、正文术语必须使用同一套对象命名，不能前后不一致
"""

DRAWING_RULE = """
【图纸规则 - 必须严格遵守】
1. 机械类图纸使用 <drawing/> 标签，表示普通图片形式的示意图/结构图/装配图/流程图
2. 设计内容描述后，如需配图，必须紧跟对应的 <drawing/> 标签
3. <drawing/> 标签必须单独成行输出，且必须严格使用以下格式，不得夹杂正文：
   <drawing id="1" type="结构图" title="图纸标题" description="图纸描述"/>
4. 标签前后不要附加任何说明文字、标点或换行内正文，不要出现 `<drawing/>` 后再接 `/`、`>`、中文说明或其他字符
5. 每个标签四个属性必须齐全：id、type、title、description，缺任何一个都不要输出半成品标签
6. 图纸类型优先使用：原理图、总体布局图、结构图、零件图、装配示意图、受力分析图、传动简图、运动过程图
7. 继续教育论文不要求高精CAD，图纸可按普通图片呈现，但必须与正文一致
8. 不要把明显需要图纸的内容只写成文字说明；如果这一段对应的图暂时写不完整，就跳过，程序会按上下文再补
9. 同一章的 `<drawing/>` 标签尽量按出现顺序编号，从1开始递增，且不要重复、不要跳行混写
10. 图纸标题和 description 必须具体、可测量、可校核，至少包含尺寸、数量、角度、载荷、行程、转速、孔径、间距或材料参数中的一类；禁止只写“示意图”“概览图”
11. 正确示例：
<drawing id="1" type="结构图" title="夹具总体结构图" description="展示夹具体、定位元件、夹紧元件和导向元件的装配关系"/>
错误示例：
<drawing/ 图1-1 夹具总体结构图，展示装配关系/>  # 禁止
"""


CONSISTENCY_RULE = """
【数据一致性铁律 - 必须严格遵守】
1. 全文所有图纸、表格、正文中的同一参数必须保持完全一致，禁止前后矛盾
2. 必须严格使用 machine_spec 中给出的参数值，不得随意编造新数值：
   - 载荷：{load}
   - 精度：{precision}
   - 行程：{stroke}
   - 材料：{material}
   - 工作条件：{working_condition}
   - 关键参数：{critical_params}
3. 如果 machine_spec 已给出参数，所有章节必须沿用，禁止改写
4. 如果 machine_spec 未给出参数，必须在第1章首次出现时确定，并在后续所有章节、图纸、表格中保持一致
5. 每个 <drawing/> 标签的 description 必须包含至少2项 machine_spec 中的关键参数（载荷、精度、行程、材料、尺寸等）
6. 同一零件在不同章节的名称、尺寸、数量、材料必须完全一致
7. 计算与校核章节的数值必须引用前文已确定的参数，不得重新发明
"""

CALC_RULE = """
【计算与校核规则 - 必须遵守】
1. 第5章必须有参数、公式或计算过程，不得只写纯结论
2. 关键参数要和正文结构一致，不能出现图是一个方案、计算却按另一个方案写
3. 至少给出1个核心计算链，并在结论里回扣结果
4. 机械类可以使用表格呈现参数、公式、结果，但不要用冗长的markdown表
"""


def mech_type_hint(mech_type: str) -> str:
    mech_type = (mech_type or "").strip()
    if "夹具" in mech_type:
        return "定位、夹紧、夹具体、受力、定位误差"
    if "传动" in mech_type:
        return "动力输入、传动链、传动比、功率、扭矩"
    if "送料" in mech_type:
        return "送料原理、驱动机构、导向机构、节拍、推送"
    if "搬运" in mech_type:
        return "动作流程、搬运路径、执行机构、稳定性、定位"
    if "升降" in mech_type:
        return "升降原理、导向方式、承载能力、安全性、驱动方式"
    if "优化" in mech_type:
        return "原结构问题、优化目标、对比方案、改进效果、适用性"
    return "总体方案、关键部件、工作原理、参数选取、校核"


def fixed_scene_hint(profile: dict) -> str:
    """给出全文唯一工艺场景，避免章节间漂移。"""
    mech_type = profile.get("mech_type", "") or profile.get("mechanical_type", "")
    mech_object = profile.get("mech_object") or profile.get("company", "") or "某零件"
    if "夹具" in mech_type:
        return f"全文统一为：{mech_object}的铣削工序专用夹具设计"
    if "传动" in mech_type:
        return f"全文统一为：{mech_object}的动力传动系统设计"
    if "送料" in mech_type:
        return f"全文统一为：{mech_object}的送料机构设计"
    if "搬运" in mech_type:
        return f"全文统一为：{mech_object}的搬运机构设计"
    if "升降" in mech_type:
        return f"全文统一为：{mech_object}的升降机构设计"
    if "优化" in mech_type:
        return f"全文统一为：{mech_object}的结构优化设计"
    return f"全文统一为：{mech_object}的机械设计"


def _spec_to_text(machine_spec) -> str:
    if not machine_spec:
        return ""
    if isinstance(machine_spec, str):
        return machine_spec.strip()
    try:
        return json.dumps(machine_spec, ensure_ascii=False, indent=2)
    except Exception:
        return str(machine_spec)


MECH_TYPE_MAP = {
    "夹具": {
        "chapter4": {"4.1": "定位方案", "4.2": "夹紧机构", "4.3": "夹具体", "4.4": "受力分析"},
        "content_points": ["定位元件", "夹紧元件", "夹具体", "对刀导向", "定位误差"],
        "drawing_types": ["定位原理图", "夹紧结构图", "夹具体零件图", "装配示意图"],
        "calc_items": ["夹紧力校核", "定位误差计算", "夹具体强度校核", "螺栓连接强度"],
    },
    "传动": {
        "chapter4": {"4.1": "动力源选型", "4.2": "传动链设计", "4.3": "关键传动件", "4.4": "润滑与维护"},
        "content_points": ["电机/动力源", "传动轴", "齿轮/带轮/链轮", "联轴器/离合器", "轴承"],
        "drawing_types": ["传动简图", "动力源布局图", "零件图（轴/齿轮）", "装配图"],
        "calc_items": ["传动比分配", "齿轮强度校核", "轴强度校核", "轴承寿命计算"],
    },
    "送料": {
        "chapter4": {"4.1": "送料机构", "4.2": "驱动机构", "4.3": "导向与定位", "4.4": "节拍分析"},
        "content_points": ["推送机构", "料仓/料道", "导向导轨", "分料机构", "传感器"],
        "drawing_types": ["送料机构运动图", "驱动结构图", "导向结构图", "节拍流程图"],
        "calc_items": ["推送力校核", "驱动功率计算", "节拍时间链校核", "导向摩擦校核"],
    },
    "搬运": {
        "chapter4": {"4.1": "执行机构", "4.2": "运动路径", "4.3": "定位机构", "4.4": "稳定性分析"},
        "content_points": ["机械手/吸盘/夹爪", "移动模组", "缓冲机构", "定位锁紧", "机架"],
        "drawing_types": ["搬运路径简图", "执行机构结构图", "定位结构图", "动作时序图"],
        "calc_items": ["搬运力矩校核", "定位精度计算", "速度/加速度校核", "结构稳定性校核"],
    },
    "升降": {
        "chapter4": {"4.1": "升降机构", "4.2": "导向机构", "4.3": "驱动机构", "4.4": "安全防护"},
        "content_points": ["升降平台/叉臂", "导向柱/导轨", "丝杠/液压缸/链条", "限位开关", "防坠落装置"],
        "drawing_types": ["升降原理图", "机构结构图", "安装布局图", "受力分析图"],
        "calc_items": ["承载能力校核", "驱动机构强度校核", "导向摩擦力校核", "倾覆稳定性计算"],
    },
    "结构优化": {
        "chapter4": {"4.1": "原结构问题", "4.2": "优化方案", "4.3": "关键改进", "4.4": "优化效果"},
        "content_points": ["薄弱环节", "改进件", "连接方式", "方案对比", "关键约束"],
        "drawing_types": ["原结构问题示意图", "优化后结构图", "对比效果图", "改进部位细部图"],
        "calc_items": ["原方案强度校核", "优化后强度校核", "减重率计算", "疲劳寿命对比"],
    },
}




def _extract_machine_spec_values(machine_spec):
    """Extract key values for consistency rule formatting."""
    values = {
        "load": "",
        "precision": "",
        "stroke": "",
        "material": "",
        "working_condition": "",
        "critical_params": "",
    }
    if not machine_spec:
        return values
    if isinstance(machine_spec, str):
        try:
            import json
            machine_spec = json.loads(machine_spec)
        except Exception:
            return values
    def get_raw(key):
        item = machine_spec.get(key)
        if isinstance(item, dict):
            return item.get("raw", item.get("value", ""))
        return str(item) if item not in (None, "") else ""
    values["load"] = get_raw("load")
    values["precision"] = get_raw("precision")
    values["stroke"] = get_raw("stroke")
    values["material"] = get_raw("material")
    values["working_condition"] = get_raw("working_condition")
    critical = machine_spec.get("critical_params", [])
    if isinstance(critical, list):
        parts = []
        for item in critical[:6]:
            if isinstance(item, dict):
                raw = item.get("raw") or item.get("value")
                name = item.get("name") or item.get("key")
                if raw and name:
                    parts.append(f"{name}={raw}")
        values["critical_params"] = "，".join(parts)
    return values


def build_outline(profile: dict) -> dict:
    """机械设计类固定大纲（可按题型微调）"""
    mech_type = profile.get("mech_type", "") or profile.get("mechanical_type", "")
    base = {
        "chapter1_overview": {"1.1": "设计背景", "1.2": "研究意义", "1.3": "设计任务与要求"},
        "chapter2_principle": {"2.1": "工作原理", "2.2": "方案比选", "2.3": "关键问题分析"},
        "chapter3_general": {"3.1": "总体方案", "3.2": "总体结构", "3.3": "主要参数初定"},
        "chapter5_calc": {"5.1": "参数计算", "5.2": "强度校核", "5.3": "结果分析"},
        "chapter6_conclusion": {"6.1": "设计成果", "6.2": "不足之处", "6.3": "改进方向"},
    }
    cfg = None
    for k, v in MECH_TYPE_MAP.items():
        if k in mech_type:
            cfg = v
            break
    base["chapter4_key"] = (cfg or MECH_TYPE_MAP["夹具"])["chapter4"]
    return base


def build_machine_spec_prompt(profile: dict, outline: dict = None) -> str:
    """生成机械论文全局锁定 machine_spec 的提示词。"""
    title = profile.get("title", "")
    mech_object = profile.get("mech_object") or profile.get("company", "")
    mech_type = profile.get("mech_type", "") or profile.get("mechanical_type", "")
    scene_hint = fixed_scene_hint(profile)
    cfg = None
    for k, v in MECH_TYPE_MAP.items():
        if k in mech_type:
            cfg = v
            break
    cfg = cfg or MECH_TYPE_MAP["夹具"]
    outline_text = ""
    if outline:
        outline_text = json.dumps(outline, ensure_ascii=False)
    return f"""你是一名机械设计论文的全局约束建模专家。请基于题目与大纲生成唯一且锁定的 machine_spec。

题目：{title}
设计对象：{mech_object}
机械类型：{mech_type}
{scene_hint}
推荐关注点：{", ".join(cfg["content_points"])}
重点校核项：{", ".join(cfg["calc_items"])}

大纲参考：
{outline_text}

输出要求：
1. 只输出严格JSON，不要输出任何解释文字
2. machine_spec 必须是整篇论文唯一约束来源，后续所有章节都必须遵守
3. 必须明确锁定同一机械对象、同一工艺场景、同一机构链、同一命名体系
4. 必须给出章节级 focus、固定术语、禁止术语、图纸要求、校验键
5. 必须包含数值化要求，尤其是图纸描述必须具体、可测量、可校核，禁止“示意”“概览”“大概”等空泛表述

JSON字段建议至少包含：
{{
  "paper_title": "...",
  "mech_object": "...",
  "mech_type": "...",
  "scene_lock": "...",
  "scope_boundary": ["..."],
  "required_terms": ["..."],
  "forbidden_terms": ["..."],
  "fixed_entities": ["..."],
  "chapter_focus": {{
    "概述": "...",
    "原理与方案分析": "...",
    "总体设计": "...",
    "关键部件设计": "...",
    "计算与校核": "...",
    "总结": "..."
  }},
  "figure_policy": {{
    "required": true,
    "must_be_specific": true,
    "must_include_numeric": true,
    "forbidden_generic_words": ["示意图", "概览图", "示意", "概况图"]
  }},
  "consistency_keys": ["scene_lock", "mech_object", "mech_type", "fixed_entities", "chapter_focus", "figure_policy"]
}}
"""


def build_chapter_prompt(chapter_name: str, chapter_num: int, profile: dict, outline: dict = None, machine_spec=None) -> str:
    """构造机械设计类单章节提示词"""
    title = profile.get("title", "")
    mech_object = profile.get("mech_object") or profile.get("company", "")
    mech_type = profile.get("mech_type", "") or profile.get("mechanical_type", "")
    scene_hint = fixed_scene_hint(profile)
    word_limit = MC_WORD_LIMITS.get(chapter_name, 1000)
    machine_spec_text = _spec_to_text(machine_spec)
    machine_spec_values = _extract_machine_spec_values(machine_spec)

    cfg = None
    for k, v in MECH_TYPE_MAP.items():
        if k in mech_type:
            cfg = v
            break
    cfg = cfg or MECH_TYPE_MAP["夹具"]
    outline_text = ""
    if outline:
        key_map = {
            "概述": "chapter1_overview",
            "原理与方案分析": "chapter2_principle",
            "总体设计": "chapter3_general",
            "关键部件设计": "chapter4_key",
            "计算与校核": "chapter5_calc",
            "总结": "chapter6_conclusion",
        }
        ch_key = key_map.get(chapter_name)
        if ch_key and ch_key in outline:
            outline_text = "\n".join([f"  {k}: {v}" for k, v in outline[ch_key].items()])

    special_rules = {
        "概述": f"""
【概述特殊要求】
- 小标题格式：1.1、1.2、1.3
- 写清设计背景、任务与要求
- 设计对象：{mech_object}
- {scene_hint}
- 严格遵守下方 machine_spec，不得新增其他机械对象、机构链或工艺场景
- 本章如输出图纸，必须具体、可测量、可校核，禁止“示意图”“概览图”等空泛表述
""",
        "原理与方案分析": f"""
【原理与方案分析特殊要求】
- 小标题格式：2.1、2.2、2.3
- 围绕{mech_type}的工作原理和方案比选展开
- {scene_hint}
- 方案对比可用<table/>，配图可用<drawing/>，不要只写文字
- <drawing/>标签必须单独成行、严格闭合，不得与正文混写
- 图纸标题和描述必须出现具体数值、尺寸、数量、角度、行程、速度、载荷或传动比中的至少一类信息
""",
        "总体设计": f"""
【总体设计特殊要求】
- 小标题格式：3.1、3.2、3.3
- 讲清总体方案、总体结构、主要参数初定
- {scene_hint}
- 总体布局图或系统示意图建议使用<drawing/>标签
- <drawing/>标签必须单独成行、严格闭合，不得与正文混写
- 图纸必须明确总体尺寸、布置关系、部件数量或参数范围，禁止只写“总体示意”
""",
        "关键部件设计": f"""
【关键部件设计特殊要求】
- 小标题格式：4.1、4.2、4.3、4.4
- 按{mech_type_hint(mech_type)}展开
- 重点内容：{", ".join(cfg["content_points"])}
- {scene_hint}
- 每个关键部件描述后如需配图，必须紧跟<drawing/>标签
- 图纸类型优先使用：{", ".join(cfg["drawing_types"])}
- <drawing/>标签必须单独成行、严格闭合，不得与正文混写
- 图纸描述必须具体到零件尺寸、材料、安装位置、孔径、间距、受力或运动参数；禁止“关键部件示意图”这类泛化标题
""",
        "计算与校核": f"""
【计算与校核特殊要求】
- 小标题格式：5.1、5.2、5.3
- 必须写参数来源、计算过程、校核结论
- 计算表格可以用<table/>标签呈现
- 重点校核项：{", ".join(cfg["calc_items"])}
- {scene_hint}
- 不要在本章强行增加新的主图纸，优先引用前文图号
""",
        "总结": """
【总结特殊要求】
- 300-500字总结
- 无标题编号，不要写"第6章"
""",
    }

    sp = special_rules.get(chapter_name, "")
    outline_block = f"\n大纲建议：\n{outline_text}" if outline_text else ""
    machine_spec_block = f"\nmachine_spec（必须遵守，禁止改写）：\n{machine_spec_text}" if machine_spec_text else ""
    prompt = f"""你是一名机械设计类学术论文写作专家。写《{title}》第{chapter_num}章 {chapter_name}。
字数要求：{word_limit}字以内。
{FMT_RULE}
{SCENARIO_LOCK_RULE}
{DRAWING_RULE}
{CALC_RULE}
{sp}
{machine_spec_block}
{outline_block}

输出格式必须严格遵守：
1. 先输出正文，正文中如需要图纸则使用 <drawing/> 标签单独成行
2. 正文结束后，必须输出机器可读 JSON 块，且只能有一个，格式如下：
[[MECH_JSON]]
{{"chapter_num":{chapter_num},"chapter_name":"{chapter_name}","text_metadata":{{...}},"facts_used":[...],"figures":[...],"consistency_keys":[...]}}
[[/MECH_JSON]]
3. JSON 必须是严格 JSON，不能有注释、尾逗号或 markdown
4. text_metadata 必须包含 sections、word_target、machine_spec_keys_used、scene_lock、figure_policy
5. facts_used 必须列出本章用到的 machine_spec 事实、参数、假设、公式来源或校核依据
6. figures 必须列出本章出现的图纸信息；图纸标题和描述必须具体、数值化、可校核，禁止 generic / 示例 / 示意 / 概览 之类空泛词
7. consistency_keys 必须列出本章用于与 machine_spec 对照的一组键名
8. 如果没有图纸，figures 仍然必须是空数组 []，不能省略

只输出上述内容，不要输出"第{chapter_num}章"字样，不要输出其他说明。"""
    return prompt


def build_abstract_prompt(profile: dict) -> str:
    scene_hint = fixed_scene_hint(profile)
    return f"""写机械设计类论文摘要，400字左右。
题目：{profile.get('title', '')}
设计对象：{profile.get('mech_object') or profile.get('company', '')}
机械类型：{profile.get('mech_type', '') or profile.get('mechanical_type', '')}
核心问题：{profile.get('core_problems', [])}
{scene_hint}
{FMT_RULE}
{SCENARIO_LOCK_RULE}
只输出摘要正文。"""


def build_keywords_prompt(profile: dict) -> str:
    scene_hint = fixed_scene_hint(profile)
    return f"""生成3-5个机械设计类论文关键词，用分号隔开。
机械对象：{profile.get('mech_object') or profile.get('company', '')}
机械类型：{profile.get('mech_type', '') or profile.get('mechanical_type', '')}
{scene_hint}
{SCENARIO_LOCK_RULE}
只输出关键词，如：关键词1；关键词2；关键词3"""


def build_references_prompt(profile: dict) -> str:
    return f"""生成12-15条机械设计类论文规范参考文献，类型包括书籍[M]、期刊论文[J]、学位论文[D]。
主题：{profile.get('title', '')}、{profile.get('mech_type', '') or profile.get('mechanical_type', '')}相关理论。
每条占一行，格式如：[1] 作者. 书名[M]. 出版社, 年份。"""


def build_drawing_analysis_prompt(full_text: str, profile: dict) -> str:
    mech_type = profile.get("mech_type", "") or profile.get("mechanical_type", "")
    scene_hint = fixed_scene_hint(profile)
    return f"""你是一名机械设计论文图纸标注专家。分析以下论文全文，完成两项任务：

任务A：查找是否有总结性表述，如“本文共X张图：图1...、图2...”等
任务B：查找各章节描述机械结构/原理/部件但缺少<drawing/>标签的段落

机械类型：{mech_type}
机械对象：{profile.get('mech_object') or profile.get('company', '')}
{scene_hint}
{SCENARIO_LOCK_RULE}

图纸类型可包括：
原理图、总体布局图、结构图、零件图、装配示意图、受力分析图、传动简图、运动过程图

硬约束：
1. 不要输出半成品 drawing 标签，不要输出 `<drawing/ xxx`、`<drawing/>/`、`<drawing` 独行、或属性不全的标签
2. 需要补图时，直接给出完整的 replacement 文本，replacement 中每个 drawing 标签必须独立成行且格式严格正确
3. 如果无法保证格式正确，宁可不补该图，也不要输出错误标签

论文全文：
{full_text[:8000]}

输出严格JSON格式：
{{
    "has_summary": true/false,
    "summary_text": "如果有总结表述，原文",
    "existing_drawings": ["已存在的drawing的标题列表"],
    "missing_drawings": [
        {{"type": "结构图", "title": "建议标题", "description": "建议描述", "context": "对应段落前100字"}}
    ],
    "action": "修改建议: add_missing / keep_existing",
    "modified_text_snippets": [
        {{"original": "原文片段", "replacement": "替换为带标签的内容"}}
    ]
}}

只输出JSON。"""
