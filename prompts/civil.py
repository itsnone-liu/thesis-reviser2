# -*- coding: utf-8 -*-
"""
土木工程类论文提示词模板
按章节分类，重点解决：
1. 土木工程论文结构模板化
2. 建筑/结构/施工数据的文字与图表一致性
3. 图纸生成规则（建筑图、结构图、施工图、甘特图等）
"""
import sys, os, json
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from core import CIV_WORD_LIMITS, CIV_CHAPTERS

FMT_RULE = """【格式铁律 - 必须严格遵守】
1. 严禁使用任何markdown符号：# ** * ` > | --- |
2. 严禁使用加粗标记，所有文字统一纯文本
3. 段落之间用换行分隔，不要用空行
4. 正文中涉及的结构维度、对比项、分类或参数项，要保持完整，不得为了简化删减
5. <drawing/> 标签必须与正文分离，单独成行，不能混在句子中
6. <table/> 标签必须包含完整属性
"""

SCENE_LOCK_RULE = """【场景锁定 - 必须严格遵守】
1. 全文只能围绕当前论文题目中的同一个建筑工程对象展开，不得中途切换到其他建筑、结构或项目
2. 章节1~6必须前后呼应，所有建筑参数、结构尺寸、荷载数据、施工方案都要服务于同一项目
3. 不得把当前题目写成另一个土木工程课题
4. 不得同时混用两套不同建筑方案
5. 图纸标题、图纸描述、正文术语必须使用同一套对象命名，不能前后不一致
"""

DRAWING_RULE = """【图纸生成规则 - 土木工程专用】
1. 土木工程图纸使用 <drawing/> 标签（所有视觉图统一走drawing，不走chart）
2. 不同章节使用不同的图纸类型：
   - 第1-2章：建筑平面图、立面图、剖面图
   - 第3-4章：结构布置图、配筋图、节点详图、基础图
   - 第5章：施工进度计划用 <drawing type="gantt">、施工平面布置图、资源分配图等
   - type 字段要尽量具体，不要写成泛化的 drawing/结构图，优先写成与图纸内容一致的专用类型
3. 【description规则 — 最关键规则】drawing标签的description必须是直接可用的GPT生图提示词，规则如下：
   A. description必须包含本章正文中与该图相关的全部信息，不能概括、不能省略、不能改写
   B. description就是给GPT生图用的完整提示词文本，无需再用其他标签或结构
   C. 所有空间方位（东南/西北/南侧/北侧/左侧/右侧/中间等）、尺寸数值、材料名称、颜色、数量等，必须和正文完全一致
   D. 正确做法：把正文中描述该图的所有句子直接复制到description中，而不是自己重写一段摘要
   E. 错误做法（禁止）：用简短的摘要代替完整描述
   F. 一切让GPT画错图的问题根源都是description信息不完整——所以description必须包含足够让GPT画出正确图纸的完整文本，宁多勿少
4. 图纸标签title必须带序号（章节号+序号），格式严格如下：
   - 第1章图纸：title="图1-1 标准层建筑平面图"
   - 第2章图纸：title="图2-1 正立面图"、title="图2-2 侧立面图"
   - 第3章图纸：title="图3-1 标准层梁配筋图"
   - 第4章图纸：title="图4-1 各层水平地震作用分布图"
   - 第5章图纸：title="图5-1 施工进度横道图"
5. 【重要】正文中引用图纸时，直接用"图X-X所示"的格式，不要画蛇添足再加"图 X-X"前缀。正文中绝对不要出现"图1"、"图2"这种无章节号的格式
   【重要】drawing标签的id顺序必须和正文中引用的图纸顺序一致，不要交叉（正文先提到哪个图，id就排在前）
6. 正确示例：
   <drawing id="1" type="建筑平面图" title="图1-1 标准层建筑平面图" description="本教学楼建筑平面呈矩形，标准层平面尺寸为东西长54.0m，南北宽18.0m，内走廊宽2.4m。柱网7.2m×7.2m，共6跨7列。南侧布置6间普通教室每间72m²，北侧设教师办公室和辅助用房，东西两端各设一部封闭楼梯间，中部设一部电梯。外墙240mm加气混凝土砌块，内墙200mm。外窗采用断桥铝合金Low-E中空玻璃窗1.8m×2.4m。标准层层高3.6m。"/>
  这是错误示例（禁止）：
   <drawing id="1" type="建筑平面图" title="图1-1 标准层建筑平面图" description="标准层教室布置，柱网7.2m×7.2m，六层框架"/>

【甘特图/图表规则 - 使用 <drawing/> 标签】
7. 施工进度计划等时序内容，使用 <drawing type="gantt"> 标签：
   <drawing id="x" type="gantt" title="图5-1 施工进度横道图" description="本工程施工总工期365天。基础工程：施工工期60天，包括土方开挖、地基处理、基础混凝土浇筑；主体结构工程：施工工期120天（第61-180天），包括框架柱梁板钢筋模板混凝土，每层约20天，共6层；砌体工程：施工工期60天（第181-240天）；装饰装修工程：施工工期90天（第241-330天），包括内外墙抹灰、地面、门窗安装、涂料；竣工验收：30天（第331-365天）。"/>
8. 数据对比图（如资源分配、费用构成），使用 <drawing type="chart_bar"> 或 <drawing type="chart_pie"> 标签：
   <drawing id="x" type="chart_bar" title="图5-3 主要工种劳动力配置图" description="本工程高峰期总人数215人，各工种配置：钢筋工50人，木工40人，混凝土工25人，瓦工20人，抹灰工25人，油漆工15人，普工40人，其他辅助工种若干。以上数据以柱状图形式展示，横轴为工种名称，纵轴为人数。"/>

【参数表规则 - 必须使用 <table/> 标签】
9. 涉及设计参数汇总表（如表1-1 主要设计参数）必须使用 <table/> 标签渲染，不得用文本罗列
10. <table/> 格式：
   <table id="1" title="表1-1 主要设计参数汇总" headers="参数名称||参数值" data="总建筑面积(m²)||约8000;建筑层数||6层;建筑总高度(m)||22.8;标准层层高(m)||3.6;柱网尺寸(m)||7.2×7.2;地基承载力特征值(kPa)||180;基本风压(kN/m²)||0.40;基本雪压(kN/m²)||0.65;抗震设防烈度(度)||7;设计地震分组||第一组;场地类别||Ⅲ类;设计使用年限(年)||50;安全等级||二级;抗震等级||三级;结构形式||框架结构;基础形式||独立基础"/>
   【重要】列分隔符用"||"（双竖线），行分隔符用";"（分号）。数据中不要再用逗号","作为分隔符。
11. 表1-1必须包含：总建筑面积、建筑层数、建筑总高度、标准层层高、柱网尺寸、地基承载力、基本风压、基本雪压、抗震设防烈度、设计地震分组、场地类别、设计使用年限、安全等级、抗震等级、结构形式、基础形式
12. 表1-1放在第1章"工程概况"末尾，正文中说"主要设计参数见表1-1所示"
"""

CONSISTENCY_RULE = """【数据一致性铁律 - 必须严格遵守】
1. 全文所有图纸、表格、正文中的同一参数必须保持完全一致，禁止前后矛盾
2. 必须严格使用 project_spec 中给出的参数值，不得随意编造新数值：
   - 建筑规模：{building_scale}
   - 结构形式：{structure_type}
   - 层数/高度：{stories}
   - 抗震设防：{seismic}
   - 地基条件：{foundation}
   - 关键参数：{critical_params}
   - 工期：{duration}
3. 如果 project_spec 已给出参数，所有章节必须沿用，禁止改写
4. 如果 project_spec 未给出参数，必须在第1章首次出现时确定，并在后续所有章节、图纸、表格中保持一致
5. 每个 <drawing/> 标签的 description 必须包含至少2项 project_spec 中的关键参数
6. 同一参数在不同章节的名称、数值必须完全一致
7. 第5章（施工/造价/管理方案）的数据必须引用前文已确定的参数，不得重新发明
"""

CHART_RULE = """【图表输出规范 - 必须遵守】
1. 涉及数据对比、趋势分析、造价分析、进度计划的地方，必须使用<chart/>标签
2. 严禁用文字罗列数据
3. 土木工程常用图表类型：
   - bar: 柱状图（造价对比、材料用量对比）
   - line/trend: 折线图（造价趋势、沉降监测）
   - pie: 饼图（费用构成比、工期组成）
   - gantt: 甘特图（施工进度计划）← 施工组织设计必须使用！
   - comparison: 对比图（方案比选）
   - stacked: 堆叠图（成本构成）
4. 【甘特图专用格式】
   x="任务1,任务2,任务3,..." 
   y="start1,duration1;start2,duration2;start3,duration3,..."
   单位统一用"天"
5. 【施工组织设计特别要求】
   - 必须有甘特图展示进度计划
   - 必须有柱状图/饼图展示资源分配
6. y属性只能填数字，不能填文字标签
7. 多系列时每个系列的y值数量必须相等
"""

STRICT_OUTPUT_RULE = """【输出纪律 - 必须严格遵守】
1. 只输出论文正文，不要输出开场白、寒暄、角色说明、道谢、总结性口吻
2. 严禁出现“好的”“遵照您的要求”“作为土木工程专业论文写作专家”“我将为您撰写”等提示词复述
3. 不要复述题目，不要复述“严格遵循”“内容将严格遵循”等说明语
4. drawing 标签必须完整闭合，且单独成行，禁止输出 `<drawing />/`、`</draing>`、`<draing>`、`<drowing>` 之类的坏标签
5. 需要图纸时，把完整 drawing 标签放在该章节末尾，每个标签占一行
6. 需要表格时，直接输出完整 `<table/>` 标签，不要输出 HTML 片段
"""

def build_chapter_prompt(chapter_name: str, chapter_num: int, profile: dict,
                          chapter_outline: list, full_text: str = "") -> str:
    """构建土木工程单章的生成提示词"""
    title = profile.get("title", "")
    project_type = profile.get("project_type", "建筑与结构设计")
    object_name = profile.get("object_name", "")
    structure_type = profile.get("structure_type", "框架结构")
    
    # 章节专用prompt
    prompts = {
        "工程概况": f"""你是土木工程专业论文写作专家。请撰写毕业论文第{chapter_num}章：工程概况。

【论文题目】：{title}
【建筑类型】：{object_name}
【结构形式】：{structure_type}

【写作要点】：
1. 工程背景：项目名称、建设地点、使用功能、建设规模
2. 设计依据：采用的国家规范（GB50010-2010、GB50011-2010等）、设计使用年限
3. 地质与气象条件：地基承载力、基本风压、基本雪压、抗震设防烈度
4. 设计参数：建筑面积、层数、层高、总高度、结构形式、抗震等级
5. 如为施工组织/工程管理类论文，则写工程概况、建设条件、施工环境

【必须包含的数据】：
- 建筑面积（m²）、层数、建筑总高度（m）
- 层高（m）、柱网尺寸（m）
- 地基承载力特征值（kPa）
- 基本风压（kN/m²）、基本雪压（kN/m²）
- 抗震设防烈度（度）、设计地震分组、场地类别
- 设计使用年限（50年）、安全等级（二级）

【字数限制】：{CIV_WORD_LIMITS.get(chapter_name, 800)}字

        【输出要求】：
        - 纯文本，用数字和具体参数说话
        - 段落自然，不要列数字序号
        - 不输出<chart/>和<table/>标签
        - 文中必须使用"如图1-1所示""如表1-1所示"等引导语句（但不要实际生成标签）
        - 末尾输出本节对应的 <drawing/> 标签（如有必要）

{STRICT_OUTPUT_RULE}""",

        "建筑设计": f"""你是土木工程专业论文写作专家。请撰写毕业论文第{chapter_num}章：建筑设计。

【论文题目】：{title}
【建筑类型】：{object_name}

【写作要点】：
1. 建筑平面设计：功能分区、房间布局、交通组织、防火分区、疏散设计
2. 建筑立面设计：造型风格、外立面材料、幕墙/窗户设计
3. 建筑剖面设计：层高关系、竖向交通（楼梯电梯）
4. 建筑构造设计：屋面、墙体、地面、门窗

【必须包含的数据】：
- 标准层平面尺寸、房间功能及面积
- 楼梯宽度、踏步尺寸、电梯数量及载重
- 墙体材料及厚度
- 窗户尺寸及窗地比
- 防火分区面积、疏散距离、疏散宽度

【字数限制】：{CIV_WORD_LIMITS.get(chapter_name, 1000)}字

        【输出要求】：
- 正文结束后，输出1-2个<drawing/>标签
          <drawing id="1" type="建筑平面图" title="图2-1 标准层建筑平面图" description="..."/>
          <drawing id="2" type="建筑立面图" title="图2-2 正立面图" description="..."/>
        - 不输出<chart/>标签

{STRICT_OUTPUT_RULE}""",

        "结构设计": f"""你是土木工程专业论文写作专家。请撰写毕业论文第{chapter_num}章：结构设计。

【论文题目】：{title}
【建筑类型】：{object_name}
【结构形式】：{structure_type}

【写作要点】：
1. 结构选型与布置：结构体系选择、柱网布置、结构平面规则性
2. 荷载计算：恒载（梁板柱自重、填充墙、面层）、活载（楼面、屋面、雪荷载）、风荷载、地震作用
3. 内力分析：框架内力计算（竖向荷载分层法、水平荷载D值法）、内力组合
4. 构件设计：梁截面配筋、柱截面配筋、楼板配筋
5. 基础设计：基础选型、基础尺寸、配筋

【必须包含的结构参数】：
- 梁截面：主梁300×600mm/250×500mm，次梁200×400mm
- 柱截面：500×500mm/600×600mm
- 板厚：100mm/120mm
- 混凝土强度等级：C30/C35
- 钢筋：HRB400
- 基础类型及尺寸

【字数限制】：{CIV_WORD_LIMITS.get(chapter_name, 1500)}字

        【输出要求】：
        - 正文结束后，输出2-3个<drawing/>标签（结构布置图、配筋图、基础图）
        - 正文中数据对比处使用<chart/>或<table/>标签
        - <chart/>数据必须完整，不能简化删减

{STRICT_OUTPUT_RULE}""",

        "施工组织设计": f"""你是土木工程专业论文写作专家。请撰写毕业论文第{chapter_num}章：施工组织设计。

【论文题目】：{title}
【建筑类型】：{object_name}

【写作要点】：
1. 施工方案：基础施工、主体结构施工、砌体/装饰施工方案
2. 施工进度计划：划分施工段、确定工序、计算工期（必须标注总工期）
3. 施工平面布置：塔吊位置、材料堆场、施工道路、临设布置
4. 质量安全保证措施

【必须包含的数据】：
- 总工期（天/月）
- 各主要工序的持续时间
- 劳动力配置（各工种人数）
- 主要机械配置（塔吊型号数量等）
- 施工段划分

【字数限制】：{CIV_WORD_LIMITS.get(chapter_name, 1500)}字

        【输出要求 - 必须遵守】：
        1. 必须有 <drawing type="gantt"> 标签展示施工进度计划！格式：
           <drawing id="x" type="gantt" title="图5-1 施工进度横道图" description="基础工程:0-60天,主体结构:60-180天,砌体工程:180-240天,装饰装修:240-330天,竣工验收:330-360天"/>
        2. 必须有 <drawing type="chart_bar"> 或 <drawing type="chart_pie"> 展示资源/成本分配
        3. 正文结束后输出1-2个<drawing/>标签（施工平面布置图等，type 尽量写成施工平面布置图/劳动力配置图/资源分配图）

{STRICT_OUTPUT_RULE}""",

        "工程造价与成本控制": f"""你是土木工程专业论文写作专家。请撰写毕业论文第{chapter_num}章：工程造价与成本控制。

【论文题目】：{title}
【建筑类型】：{object_name}

【写作要点】：
1. 工程概况与造价构成：土建工程、安装工程、装饰工程费用占比
2. 造价影响因素分析：设计阶段、施工阶段、材料价格波动
3. 成本控制措施：招投标、限额设计、动态成本管理
4. 全过程造价管理方法

【必须包含的数据】：
- 总造价（万元）、单方造价（元/m²）
- 各分部工程造价及占比
- 人/材/机费用占比
- 造价对比分析（可引用的行业数据）

【字数限制】：{CIV_WORD_LIMITS.get(chapter_name, 1200)}字

        【输出要求】：
        - 必须有<drawing type="chart_pie">或<drawing type="chart_bar">展示造价构成
        - 必须有<table/>标签展示费用明细对比

{STRICT_OUTPUT_RULE}""",

        "质量管理与安全管理": f"""你是土木工程专业论文写作专家。请撰写毕业论文第{chapter_num}章：质量管理与安全管理。

【论文题目】：{title}
【建筑类型】：{object_name}

【写作要点】：
1. 质量管理体系：质量目标、管理体系、质量控制点
2. 各施工阶段质量控制：基础、主体、装饰
3. 安全管理体系：安全目标、危险源识别、安全措施
4. 应急预案

【字数限制】：{CIV_WORD_LIMITS.get(chapter_name, 1200)}字

        【输出要求】：
        - 可使用<table/>标签展示质量控制点/危险源清单

{STRICT_OUTPUT_RULE}""",

        "结论与展望": f"""你是土木工程专业论文写作专家。请撰写毕业论文第{chapter_num}章：结论与展望。

【论文题目】：{title}

【写作要点】：
1. 主要设计/研究成果总结
2. 存在的不足与改进方向
3. 展望

【字数限制】：{CIV_WORD_LIMITS.get(chapter_name, 500)}字

        【输出要求】：纯文本，不要图表标签

{STRICT_OUTPUT_RULE}""",
    }

    # 根据论文类型选择对应的提示词
    if chapter_name == "工程概况":
        return prompts["工程概况"]
    elif chapter_name == "建筑设计":
        return prompts["建筑设计"]
    elif chapter_name == "结构设计":
        return prompts["结构设计"]
    elif chapter_name == "施工组织设计":
        return prompts["施工组织设计"]
    elif chapter_name == "工程造价与成本控制":
        return prompts["工程造价与成本控制"]
    elif chapter_name == "质量管理与安全管理":
        return prompts["质量管理与安全管理"]
    elif chapter_name == "结论与展望":
        return prompts["结论与展望"]
    
    # 兜底
    dr = DRAWING_RULE if chapter_name in ["工程概况", "建筑设计", "结构设计", "结构计算", "施工组织设计"] else ""
    return f"""你是土木工程专业论文写作专家。请撰写毕业论文第{chapter_num}章：{chapter_name}。
【论文题目】：{title}
{dr}
【字数限制】：{CIV_WORD_LIMITS.get(chapter_name, 1000)}字
【输出要求】：纯文本段落

{STRICT_OUTPUT_RULE}"""


def build_abstract_prompt(profile: dict) -> str:
    """构建摘要提示词"""
    title = profile.get("title", "")
    prompt = f"""你是土木工程专业论文写作专家。请为以下论文撰写摘要（300字左右）。

【论文题目】：{title}

【摘要写作要求】：
1. 介绍工程概况（地点、规模、结构形式）
2. 概述主要设计/研究内容（建筑、结构、施工）
3. 简述主要成果和结论
4. 3-5个关键词（以"关键词："开头）

【输出要求】：纯文本，不要图表标签
5. 【重要】摘要正文内容不要以"摘要："开头，直接写正文内容
6. 正文写完后单独一行写关键词（格式：关键词：xxx；yyy；zzz）

{STRICT_OUTPUT_RULE}"""
    return prompt


def build_keywords_prompt(profile: dict) -> str:
    """构建关键词提示词（已包含在摘要中）"""
    return ""


def build_references_prompt(profile: dict) -> str:
    """构建参考文献提示词"""
    title = profile.get("title", "")
    return f"""你是土木工程专业论文写作专家。请为论文《{title}》列出10-12篇参考文献。

【要求】：
1. 必须包含以下国家标准和规范：
   - GB50010-2010《混凝土结构设计规范》
   - GB50011-2010《建筑抗震设计规范》
   - GB50009-2012《建筑结构荷载规范》
   - GB50007-2011《建筑地基基础设计规范》
2. 其余为近5年的学术论文、专业书籍
3. 参考文献格式：[1] 作者. 书名/论文名[D]. 出版社/期刊, 年份.
4. 按引用顺序编号

【输出要求】：纯文本，从参考文献标题开始

{STRICT_OUTPUT_RULE}"""


def build_drawing_analysis_prompt(full_text: str, profile: dict) -> str:
    """分析全文图纸标签，补齐缺失的drawing标签"""
    title = profile.get("title", "")
    return f"""你是土木工程图纸分析专家。请分析以下论文全文，检查图纸标签的完整性。

【论文题目】：{title}

【当前已有的图纸标签】：
请从文中提取所有<drawing/>标签。

【分析要求】：
1. 检查正文中是否有"如图X所示"但缺少对应<drawing/>标签的描述
2. 检查<drawing/>标签的description是否包含具体的建筑参数（尺寸、材料、配筋等）
3. 对于施工组织设计类论文，检查是否缺少甘特图<chart type="gantt">标签

【输出JSON格式】：
{{
    "existing_drawings": ["图1已有", "图2已有"],
    "missing_drawings": [{{"id": "3", "type": "结构图", "title": "...", "description": "..."}}],
    "action": "add_missing" 或 "keep_existing",
    "modified_text_snippets": []  // 如有替换，放替换片段
}}

请只输出JSON，不要其他文字。

{STRICT_OUTPUT_RULE}"""


def generate_project_spec(profile: dict, spec_text: str = "") -> str:
    """保存project_spec描述，用于一致性规则
    如果传入了spec_text（LLM生成的JSON字符串），直接保存使用。
    否则用简单规则生成fallback值。
    """
    import json
    if spec_text:
        try:
            # 尝试解析并返回格式化后的JSON
            spec = json.loads(spec_text)
            return json.dumps(spec, ensure_ascii=False, indent=2)
        except:
            pass
    
    title = profile.get("title", "")
    structure_type = profile.get("structure_type", "框架结构")
    
    spec = {
        "building_scale": "",
        "structure_type": structure_type,
        "stories": "",
        "height": "",
        "floor_height": "3.6",
        "grid_size": "7.2x7.2",
        "seismic": "7度设防",
        "design_life": "50年",
        "safety_level": "二级",
        "foundation": "独立基础，fak=180kPa",
        "wind_load": "0.40",
        "snow_load": "0.65",
        "concrete_grade": "C30",
        "steel_grade": "HRB400",
        "critical_params": [],
        "duration": "300",
    }

    if "六层" in title: spec["stories"] = "6层"
    elif "七层" in title: spec["stories"] = "7层"
    elif "十二层" in title: spec["stories"] = "12层"
    elif "五层" in title: spec["stories"] = "5层"
    elif "八层" in title: spec["stories"] = "8层"
    elif "九层" in title: spec["stories"] = "9层"
    elif "四层" in title: spec["stories"] = "4层"
    elif "高层" in title: spec["stories"] = "18层"

    if "剪力墙" in title: spec["structure_type"] = "剪力墙结构"
    elif "框架-剪力墙" in title: spec["structure_type"] = "框架-剪力墙结构"
    elif "钢结构" in title: spec["structure_type"] = "钢结构"

    return json.dumps(spec, ensure_ascii=False, indent=2)


def get_chapter_outline(profile: dict) -> list:
    """根据论文类型返回定制的章节大纲"""
    title = profile.get("title", "")
    
    # 判断论文类型
    if "施工组织设计" in title:
        return [
            ("工程概况", 1),
            ("施工方案", 2),
            ("施工进度计划", 3),
            ("施工平面布置", 4),
            ("质量安全保证措施", 5),
            ("结论与展望", 6),
        ]
    elif "造价" in title or "成本" in title:
        return [
            ("工程概况", 1),
            ("造价构成分析", 2),
            ("造价影响因素", 3),
            ("成本控制措施", 4),
            ("全过程造价管理", 5),
            ("结论与展望", 6),
        ]
    elif "质量管理" in title or "管理" in title:
        return [
            ("工程概况", 1),
            ("质量管理体系", 2),
            ("施工阶段质量控制", 3),
            ("安全管理措施", 4),
            ("质量安全保证体系", 5),
            ("结论与展望", 6),
        ]
    elif "基础" in title or "基坑" in title or "桩基" in title:
        return [
            ("工程概况", 1),
            ("工程地质条件", 2),
            ("基础方案比选", 3),
            ("基础设计计算", 4),
            ("施工方案", 5),
            ("结论与展望", 6),
        ]
    elif "道路" in title or "路基路面" in title:
        return [
            ("工程概况", 1),
            ("路线设计", 2),
            ("路基设计", 3),
            ("路面结构设计", 4),
            ("施工方案", 5),
            ("结论与展望", 6),
        ]
    elif "桥梁" in title or "简支梁桥" in title:
        return [
            ("工程概况", 1),
            ("桥型方案比选", 2),
            ("上部结构设计", 3),
            ("下部结构设计", 4),
            ("施工方案", 5),
            ("结论与展望", 6),
        ]
    elif "钢结构" in title or "鉴定" in title:
        return [
            ("工程概况", 1),
            ("结构检测与鉴定", 2),
            ("结构分析与计算", 3),
            ("加固设计方案", 4),
            ("施工工艺", 5),
            ("结论与展望", 6),
        ]
    else:
        # 标准建筑与结构设计
        return [
            ("工程概况", 1),
            ("建筑设计", 2),
            ("结构设计", 3),
            ("结构计算", 4),
            ("施工组织设计", 5),
            ("结论与展望", 6),
        ]
