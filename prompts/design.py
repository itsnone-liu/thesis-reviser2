# -*- coding: utf-8 -*-
"""
设计类论文提示词模板
按章节分类，重点解决：
1. 设计策略章节必须输出 <drawing/> 标签
2. 设计类术语和流程规范
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from core import SJ_WORD_LIMITS

FMT_RULE = """
【格式铁律 - 必须严格遵守】
1. 严禁使用正文markdown符号：# ** * ` > ---；但结构化标签属性内允许使用协议分隔符 |、|| 和 ;，不得将标签拆成多行
2. 严禁使用加粗标记，所有文字统一纯文本
3. 段落之间用换行分隔，不要用空行
4. 【图表/表格维度完整性】正文中涉及的每个结构维度、对比项、分类是否都有<chart/>或<table/>覆盖？→ 缺则补，不得为简化删维度
5. 设计类图纸总数控制在5-8张，优先集中在第3章和第4章；第1章、第2章、第5章原则上不输出drawing
6. 设计类不要为每个细节都配图，只保留关键总图、立面图、平面图、节点图、效果图
7. 需要输出图表或表格时，只能使用一行一个的标准自闭合标签，格式必须完整，例如<chart id="1" title="..." type="..." x="..." y="..."/>或<table id="1" title="..." header="..." rows="..." data="..."/>。
8. 严禁输出HTML表格片段、<tr>、<td>、<th>、<caption>等任何表格源码，也不要把说明文字拼接在标签后面。
9. 严禁输出类似“摘要：”或“摘要:”这样的前缀，摘要正文直接起写。
10. <chart/>只能表达纯数字序列，x写分类名称，y只写数字并用逗号分隔，单位单独写在unit属性里；如果不是纯数字序列，改用<table/>。
11. 设计类正文只允许两类表格模板：对比型表格和评价/统计型表格。对比型表格用于多对象对照，header第一列使用“对比项”或“指标”；评价/统计型表格用于调研评分、满意度、频次统计，header第一列使用“评价维度”或“统计维度”。两类表格的rows必须是实际行名，不得写成数字占位。
12. `<table/>` 的数据必须严格按列组织：header 有几列，data 每一行就必须提供相应数量的单元值；row name 只放在 rows 里，不要混进 data；多对象对比时，data 用 `;` 分行、用 `|` 分列，不能把多个对象的值挤进第二列。
"""

DRAWING_CHECKLIST = """
【图纸标签自检 - 逐项核对】
1. 每一个设计描述后面是否紧跟着<drawing/>标签？→ 缺则补
2. <drawing/>的tag是否已正确闭合（/>）？→ 否则修正
3. id属性是否用纯数字（1,2,3...）？→ 否则修正
"""


def build_chapter_prompt(chapter_name: str, chapter_num: int, profile: dict, outline: dict = None) -> str:
    """构造设计类单章节提示词"""
    title = profile.get("title", "")
    design_object = profile.get("design_object", "")
    design_type = profile.get("design_type", "")
    context = profile.get("context", "")
    word_limit = SJ_WORD_LIMITS.get(chapter_name, 1000)

    # 子标题内容（如果有大纲）
    sub_outline = ""
    if outline:
        ch_key = {
            "绪论": "chapter1_intro",
            "理论基础": "chapter2_theory",
            "问题发现与分析": "chapter3_problem",
            "设计策略与方案": "chapter4_design",
            "总结与反思": "chapter5_conclusion",
        }.get(chapter_name)
        if ch_key and ch_key in outline:
            ch_data = outline[ch_key]
            sub_outline = "\n".join([f"  {k}: {v}" for k, v in ch_data.items()])

    special_rules = {
        "绪论": f"""
【绪论特殊要求】
- 小标题格式：1.1、1.2、1.3
- 背景、问题、意义三段式
- 设计对象：{design_object}
""",
        "理论基础": f"""
【理论基础特殊要求】
- 小标题格式：2.1、2.2、2.3
- 围绕{design_type}专业理论展开
- 要有理论+理论如何指导本设计
""",
        "问题发现与分析": f"""
【问题发现与分析特殊要求】
- 小标题格式：3.1、3.2、3.3
- 围绕{design_object}的现状调研和分析
- 设计场景：{context}
- 涉及调研数据、结构对比、需求分层、方案比较时，使用<chart/>或<table/>标签
- 图表或表格必须写成独立的一整行，不得夹带解释文字，不得输出HTML源码片段
- 设计类正文只允许两类表格模板：对比型表格和评价/统计型表格，不要自行发明第三种结构
- 如果数据是功能对比、满意度、材料、维护模式、空间配置等描述性内容，优先使用<table/>；只有在分类清晰且数值完整时才使用<chart/>
- <chart/>的x必须是分类名称，y必须是纯数字序列，单位放到unit属性，不要写进y里
- 图表维度和分类必须覆盖正文全部对比项，不得遗漏
- 设计类以定性为主，表格可以放非数值内容（如功能对照、材料对比），不需要强求数字图表
- 设计类表格必须严格遵循列数：表头有几列，data 每行就必须有几列；如果是三列对比表，data 行就必须对应三组值，不得把后两组塞进第二列
- 本章原则上不输出<drawing/>标签，如确有必要也只保留1张分析图
""",
        "设计策略与方案": f"""
【设计策略与方案特殊要求】
- 小标题格式：4.1、4.2、4.3、4.4、4.5
- 只围绕关键设计方案输出图纸，不要对每个细节都配图
- 优先输出 3-5 张关键图：总图、平面图、立面图、节点图、效果图
- <drawing/>标签格式：<drawing id="编号" type="设计图类型" title="图纸标题" description="图纸描述"/>
- 设计图类型根据{design_type}专业特点选取
- 每一小节最多 1 张图，且不要重复同类图
- 全文同一图纸标题与描述只能出现一次；前文已经出现过的图纸，后文只用文字引用图号，不要再次输出相同的<drawing/>标签
- 如果需要插入图表或表格，仍然只能用标准自闭合标签，不得写成HTML源码或半截标签
- 设计类表格一律使用“行名 + 逐列值”的平铺格式，不要生成嵌套表、合并单元格或把多个案例值堆到第二列
""",
        "总结与反思": """
【总结与反思特殊要求】
- 小标题格式：5.1、5.2
- 总结主要成果，指出不足和改进方向
- 本章原则上不输出<drawing/>标签
""",
    }

    sp = special_rules.get(chapter_name, "")
    outline_text = f"\n大纲建议：\n{sub_outline}" if sub_outline else ""

    prompt = f"""你是一名{design_type}专业学术论文写作专家。写《{title}》第{chapter_num}章 {chapter_name}。
字数要求：{word_limit}字以内。
{FMT_RULE}
{sp}
{outline_text}

只输出正文内容，不要输出"第{chapter_num}章"字样。"""
    return prompt


def build_abstract_prompt(profile: dict) -> str:
    """设计类摘要提示词"""
    return f"""写设计类论文摘要，400字左右。
题目：{profile.get('title', '')}
设计对象：{profile.get('design_object', '')}
设计专业：{profile.get('design_type', '')}
核心问题：{profile.get('core_problems', [])}
设计策略：{profile.get('design_strategies', [])}
{FMT_RULE}
只输出摘要正文。摘要正文必须直接起写，不要输出“摘要：”或“摘要:”前缀。摘要中不得出现任何表格、图表、<table/>、<chart/>、<drawing/>、<tr>、<td>、<th>、<caption>等标签或HTML片段。"""


def build_keywords_prompt(profile: dict) -> str:
    """设计类关键词提示词"""
    return f"""生成3-5个设计类论文关键词，用分号隔开。
设计对象：{profile.get('design_object', '')}
设计专业：{profile.get('design_type', '')}
只输出关键词，如：设计关键词1；关键词2；关键词3"""


def build_references_prompt(profile: dict) -> str:
    """设计类参考文献提示词"""
    return f"""生成12-15条设计类论文规范参考文献，类型包括书籍[M]、期刊论文[J]、学位论文[D]。
主题：{profile.get('design_type', '')}、{profile.get('design_object', '')}。
每条占一行，格式如：[1] 作者. 书名[M]. 出版社, 年份."""


def build_drawing_analysis_prompt(full_text: str, profile: dict) -> str:
    """分析论文中图纸描述的提示词（用于补齐<drawing/>标签）"""
    return f"""你是一名设计论文图纸标注专家。分析以下论文全文，完成两项任务：

任务A：查找是否有总结性表述，如"本文共X张设计图：图1xxx、图2xxx..."等
任务B：查找各章节描述设计效果但缺少<drawing/>标签的段落

论文全文：
{full_text[:8000]}

设计专业：{profile.get('design_type', '')}
设计对象：{profile.get('design_object', '')}

硬约束：
1. 设计类总图纸数量建议控制在5-8张，不要为每个细节都补图
2. 重点补第3章和第4章的关键图，第1章、第2章、第5章原则上不补图
3. 若当前文本已经有足够的关键图，不要继续新增
4. 如果补图，只补总图、平面图、立面图、节点图、效果图这类关键图
5. 不要输出半成品 drawing 标签，不要输出属性不全或格式不闭合的标签
6. 同一图纸标题/描述不要重复输出；如果原文中已经有相同图纸，只保留一处，其余用“如图X所示”文字引用

输出严格JSON格式：
{{
    "has_summary": true/false,
    "summary_text": "如果有总结表述，原文",
    "existing_drawings": ["已存在的drawing的标题列表"],
    "missing_drawings": [
        {{"type": "效果图", "title": "建议标题", "description": "建议描述", "context": "对应段落前100字"}}
    ],
    "action": "修改建议: add_missing / keep_existing",
    "modified_text_snippets": [
        {{"original": "原文片段", "replacement": "替换为带标签的内容"}}
    ]
}}

只输出JSON。"""
