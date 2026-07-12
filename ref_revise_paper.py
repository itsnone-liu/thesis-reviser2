# -*- coding: utf-8 -*-
"""
论文修改程序 - 解析与诊断模块（revise_paper.py 前半部分）
"""

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import re
import json
import os
import time
import shutil
import uuid
import argparse
import subprocess
from copy import deepcopy
from web import (
    call_llm,
    build_chart_rule,
    MG_WORD_LIMITS,
    MG_SOFT_MAX,
    MG_CHAPTERS,
    tasks_db,
    is_system_busy,
    system_lock,
    add_c,
    add_t,
    chart_to_bytes,
    extract_charts,
    extract_tables,
)


# ================== [ 字体设置辅助 ] ==================
def set_run_font(run, font_name, font_size, bold=False, color=None):
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(font_size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color


# ================== [ 封面信息提取 ] ==================
def extract_cover_info(docx_path: str) -> dict:
    """从原论文DOCX的前25个段落中提取封面信息"""
    info = {
        "title": "",
        "name": "",
        "student_id": "",
        "major": "",
        "level": "",
        "advisor": "",
        "year": "",
        "month": "",
        "day": "",
    }
    try:
        doc = Document(docx_path)
        patterns = {
            "title": re.compile(r'论\s*文\s*题\s*目[：:]\s*(.+)'),
            "name": re.compile(r'姓\s*名[：:]\s*(.+)'),
            "student_id": re.compile(r'学\s*号[：:]\s*(.+)'),
            "major": re.compile(r'专\s*业[：:]\s*(.+)'),
            "level": re.compile(r'层\s*次[：:]\s*(.+)'),
            "advisor": re.compile(r'指导教师[：:]\s*(.+)'),
            "date": re.compile(r'(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日'),
        }
        count = 0
        for para in doc.paragraphs:
            if count >= 25:
                break
            text = para.text.strip()
            if not text:
                count += 1
                continue
            for key, pat in patterns.items():
                if key == "date":
                    m = pat.search(text)
                    if m:
                        info["year"] = m.group(1)
                        info["month"] = m.group(2)
                        info["day"] = m.group(3)
                else:
                    m = pat.search(text)
                    if m:
                        info[key] = m.group(1).strip()
            count += 1
    except Exception as e:
        print(f"提取封面信息失败: {e}")
    return info


# ================== [ 封面生成 ] ==================
def build_cover(doc, cover_info: dict):
    """在DOCX文档开头生成标准封面页（匹配封面.docx格式）"""
    # 页面设置
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(3.175)
    section.right_margin = Cm(3.175)

    # 学校名：无锡太湖学院 + 自学考试（28pt加粗）
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p.add_run("无锡太湖学院")
    set_run_font(r1, "宋体", 28, bold=True)
    r2 = p.add_run("自学考试")
    set_run_font(r2, "宋体", 28, bold=True)

    # 空行
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 毕业论文（48pt加粗宋体，字间空格）
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("毕 业 论 文")
    set_run_font(r, "宋体", 48, bold=True)

    # 空行
    p = doc.add_paragraph()

    # 信息项（16pt加粗，左对齐，段后10pt，最小行高11pt）
    labels = [
        ("论文题目", "title"),
        ("姓    名", "name"),
        ("学    号", "student_id"),
        ("专    业", "major"),
        ("层    次", "level"),
        ("指导教师", "advisor"),
    ]
    for label, key in labels:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        text = f"{label}：{cover_info.get(key, '')}"
        r = p.add_run(text)
        set_run_font(r, "宋体", 16, bold=True)
        p.paragraph_format.space_after = Pt(10)
        pPr = p._element.get_or_add_pPr()
        spacing = OxmlElement('w:spacing')
        spacing.set(qn('w:line'), '220')
        spacing.set(qn('w:lineRule'), 'atLeast')
        spacing.set(qn('w:after'), '200')
        pPr.insert(0, spacing)

    # 3个空行
    for _ in range(3):
        p = doc.add_paragraph()

    # 日期（16pt加粗）
    year = cover_info.get("year", "")
    month = cover_info.get("month", "")
    day = cover_info.get("day", "")
    if year and month and day:
        date_text = f"{year}年{month}月{day}日"
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r = p.add_run(date_text)
        set_run_font(r, "宋体", 16, bold=True)

    # 封面末尾插入分页符
    doc.add_page_break()


# ================== [ 文本清理 ] ==================
def clean_text(text: str) -> str:
    """清理多余符号和空格"""
    if not text:
        return ""
    # 删除 markdown 符号
    text = re.sub(r'\*\*|\*|`|#|>', '', text)
    # 删除装饰符号
    text = re.sub(r'[★●■◇◆○◎]', '', text)
    # 全角括号转半角
    text = text.replace('［', '[').replace('］', ']').replace('（', '(').replace('）', ')')
    # 合并多余空格
    text = re.sub(r'\s+', ' ', text)
    # 去除行首行尾空格
    text = text.strip()
    return text


# ================== [ 文档解析 ] ==================
def extract_docx_text(path: str) -> str:
    """使用 python-docx 读取 .docx 文件的所有段落文本，保留换行，过滤空段落"""
    try:
        doc = Document(path)
        paragraphs = [p.text.strip() for p in doc.paragraphs if p.text and p.text.strip()]
        return "\n".join(paragraphs)
    except Exception as e:
        print(f"读取文档失败 {path}: {e}")
        return ""


# ================== [ 论文画像分析 ] ==================
def analyze_original_paper(text: str) -> dict:
    """调用 LLM 分析论文全文，提取研究画像"""
    prompt = f"""你是一名学术论文分析专家。请对以下论文全文进行深度分析，输出严格JSON格式（不要markdown代码块，不要任何其他文字）。

论文全文：
{text[:8000]}

请输出以下字段的JSON：
{{
  "title": "论文标题",
  "major": "专业",
  "company": "研究对象/公司",
  "industry": "所属行业",
  "core_problems": ["问题1", "问题2", "问题3"],
  "data_hints": ["数据线索1", "线索2", "线索3"],
  "outline": {{
    "引言": "本章核心内容摘要（100字内）",
    "国内外研究现状": "...",
    "现状分析": "...",
    "问题与原因分析": "...",
    "解决方案": "...",
    "结论": "..."
  }},
  "theories": ["理论1", "理论2"],
  "problem_analysis_summary": "第三章核心逻辑",
  "solution_summary": "第五章方案概要",
  "charts_mentioned": ["图表1描述"],
  "data_quality_note": "数据质量评价"
}}

只输出JSON，严禁输出其他内容。"""

    for attempt in range(3):  # 首次 + 最多2次重试
        response = call_llm(prompt, max_tokens=3000)
        if not response:
            time.sleep(1)
            continue
        try:
            json_match = re.search(r'\{.*\}', response, re.DOTALL)
            if json_match:
                result = json.loads(json_match.group())
                # 补齐缺失字段
                if "title" not in result:
                    result["title"] = text[:50].replace('\n', ' ').strip() if text else "未识别标题"
                if "major" not in result:
                    result["major"] = "工商管理"
                if "company" not in result:
                    result["company"] = "某企业"
                if "industry" not in result:
                    result["industry"] = "制造业"
                if "core_problems" not in result:
                    result["core_problems"] = []
                if "data_hints" not in result:
                    result["data_hints"] = []
                if "outline" not in result:
                    result["outline"] = {name: "" for name, _ in MG_CHAPTERS}
                else:
                    for name, _ in MG_CHAPTERS:
                        if name not in result["outline"]:
                            result["outline"][name] = ""
                if "theories" not in result:
                    result["theories"] = []
                if "problem_analysis_summary" not in result:
                    result["problem_analysis_summary"] = ""
                if "solution_summary" not in result:
                    result["solution_summary"] = ""
                if "charts_mentioned" not in result:
                    result["charts_mentioned"] = []
                if "data_quality_note" not in result:
                    result["data_quality_note"] = ""
                return result
        except Exception as e:
            print(f"解析论文分析结果失败（尝试{attempt+1}/3）：{e}")
            time.sleep(1)
            continue

    # 兜底默认
    fallback_title = text[:100].replace('\n', ' ').strip() if text else "未识别标题"
    return {
        "title": fallback_title,
        "major": "工商管理",
        "company": "某企业",
        "industry": "制造业",
        "core_problems": ["问题待识别", "问题待识别", "问题待识别"],
        "data_hints": ["数据线索待识别", "数据线索待识别", "数据线索待识别"],
        "outline": {name: "" for name, _ in MG_CHAPTERS},
        "theories": ["理论待识别"],
        "problem_analysis_summary": "待分析",
        "solution_summary": "待分析",
        "charts_mentioned": [],
        "data_quality_note": "LLM解析失败，使用默认画像"
    }


# ================== [ 深度诊断与重构 ] ==================
def diagnose_and_reconstruct(original_profile: dict) -> dict:
    """调用 LLM 基于论文画像进行深度诊断，输出修订计划"""
    prompt = f"""你是一名资深学术论文评审与修改专家。请基于以下论文画像进行深度诊断，并输出修订方案。

论文画像：
{json.dumps(original_profile, ensure_ascii=False, indent=2)}

请输出严格JSON格式（不要markdown代码块，不要任何其他文字）：
{{
  "diagnosis_report": {{
    "overall_score": "1-10分",
    "overall_comment": "总体评价（200字）",
    "issues": [
      {{
        "chapter": "涉及章节",
        "severity": "高/中/低",
        "issue_type": "理论错误/数据矛盾/逻辑混乱/结构失衡/表述问题/其他",
        "description": "问题描述",
        "fix_direction": "修正方向"
      }}
    ]
  }},
  "revised_profile": {{
    "title": "优化后的标题",
    "company": "公司",
    "industry": "行业",
    "core_problems": ["修正后问题1", "问题2", "问题3"],
    "data_hints": ["数据方向1", "方向2", "方向3"]
  }},
  "chapter_revisions": {{
    "引言": {{"issues": [], "fix_requirements": [], "key_points": []}},
    "国内外研究现状": {{"issues": [], "fix_requirements": [], "key_points": []}},
    "现状分析": {{"issues": [], "fix_requirements": ["必须含至少2个<chart/>和1个<table/>"], "key_points": []}},
    "问题与原因分析": {{"issues": [], "fix_requirements": ["必须含至少1个<chart/>和1个<table/>"], "key_points": []}},
    "解决方案": {{"issues": [], "fix_requirements": ["必须含至少1个<chart/>"], "key_points": []}},
    "结论": {{"issues": [], "fix_requirements": [], "key_points": []}}
  }},
  "theory_adjustment": "理论框架调整说明",
  "data_strategy": "数据/图表策略",
  "special_notes": "生成新论文时的特殊注意事项"
}}

注意：chapter_revisions 中的每个章节都要有 issues、fix_requirements、key_points 三个字段。如果某章节无问题，请留空列表[]。
只输出JSON，严禁输出其他内容。"""

    for attempt in range(3):  # 首次 + 最多2次重试
        response = call_llm(prompt, max_tokens=4000)
        if not response:
            time.sleep(1)
            continue
        try:
            json_match = re.search(r'\{.*\}', response, re.DOTALL)
            if json_match:
                result = json.loads(json_match.group())
                # 补齐缺失字段
                if "diagnosis_report" not in result:
                    result["diagnosis_report"] = {"overall_score": "5", "overall_comment": "默认评价", "issues": []}
                if "revised_profile" not in result:
                    result["revised_profile"] = {
                        "title": original_profile.get("title", ""),
                        "company": original_profile.get("company", ""),
                        "industry": original_profile.get("industry", ""),
                        "core_problems": original_profile.get("core_problems", []),
                        "data_hints": original_profile.get("data_hints", [])
                    }
                if "chapter_revisions" not in result:
                    result["chapter_revisions"] = {}
                for name, _ in MG_CHAPTERS:
                    if name not in result["chapter_revisions"]:
                        result["chapter_revisions"][name] = {"issues": [], "fix_requirements": [], "key_points": []}
                    else:
                        for key in ["issues", "fix_requirements", "key_points"]:
                            if key not in result["chapter_revisions"][name]:
                                result["chapter_revisions"][name][key] = []
                if "theory_adjustment" not in result:
                    result["theory_adjustment"] = ""
                if "data_strategy" not in result:
                    result["data_strategy"] = ""
                if "special_notes" not in result:
                    result["special_notes"] = ""
                return result
        except Exception as e:
            print(f"解析诊断结果失败（尝试{attempt+1}/3）：{e}")
            time.sleep(1)
            continue

    # 保守兜底
    return {
        "diagnosis_report": {
            "overall_score": "5",
            "overall_comment": "由于LLM解析失败，采用保守默认诊断。建议人工复核论文内容。",
            "issues": []
        },
        "revised_profile": {
            "title": original_profile.get("title", ""),
            "company": original_profile.get("company", ""),
            "industry": original_profile.get("industry", ""),
            "core_problems": original_profile.get("core_problems", []),
            "data_hints": original_profile.get("data_hints", [])
        },
        "chapter_revisions": {
            name: {"issues": [], "fix_requirements": [], "key_points": []}
            for name, _ in MG_CHAPTERS
        },
        "theory_adjustment": "",
        "data_strategy": "",
        "special_notes": "LLM诊断解析失败，请人工检查原始论文并制定修改策略。"
    }


# ================== [ 诊断报告持久化 ] ==================
def save_diagnosis_report(revision_plan: dict, output_path: str):
    """将 revision_plan['diagnosis_report'] 格式化为易读文本报告并写入文件"""
    diag = revision_plan.get("diagnosis_report", {})
    lines = []
    lines.append("=" * 50)
    lines.append("【论文诊断报告】")
    lines.append("=" * 50)
    lines.append("")
    lines.append(f"总体评分：{diag.get('overall_score', 'N/A')}")
    lines.append("")
    lines.append("【总体评价】")
    lines.append(diag.get("overall_comment", "暂无"))
    lines.append("")
    issues = diag.get("issues", [])
    if issues:
        lines.append("【问题清单】")
        for i, issue in enumerate(issues, 1):
            lines.append(f"  问题{i}：")
            lines.append(f"    涉及章节：{issue.get('chapter', 'N/A')}")
            lines.append(f"    严重程度：{issue.get('severity', 'N/A')}")
            lines.append(f"    问题类型：{issue.get('issue_type', 'N/A')}")
            lines.append(f"    问题描述：{issue.get('description', 'N/A')}")
            lines.append(f"    修正方向：{issue.get('fix_direction', 'N/A')}")
            lines.append("")
    else:
        lines.append("【问题清单】")
        lines.append("  未发现明显问题（或LLM未返回问题列表）")
        lines.append("")
    lines.append("=" * 50)

    dir_name = os.path.dirname(output_path)
    if dir_name:
        os.makedirs(dir_name, exist_ok=True)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write("\n".join(lines))


# ================== [ 画像融合 ] ==================
def merge_into_standard_profile(revised_profile: dict, major: str) -> dict:
    """将修订画像转换为兼容原 web.py 的标准 profile 格式"""
    return {
        "major": major,
        "title": revised_profile.get("title", ""),
        "company": revised_profile.get("company", ""),
        "industry": revised_profile.get("industry", ""),
        "core_problems": revised_profile.get("core_problems", []),
        "data_hints": revised_profile.get("data_hints", []),
    }


# ================== [ 重写章节生成 ] ==================
def revise_gen_chapter(name, num, profile, chapter_rev, update):
    """修改版章节生成：完全复用原 gen_chapter 逻辑，追加重写要求"""
    update(f"正在生成第{num}章 {name}...", int((num - 1) / 6 * 80))
    if name == "引言":
        sp = "\n- 禁止使用任何标题、编号、markdown。用2-3段写背景和意义。"
    elif name == "国内外研究现状":
        sp = f"\n- 禁止使用任何标题、编号、markdown。纯段落叙述，不谈{profile['company']}。"
    elif name == "现状分析":
        sp = "\n- 小标题：3.1、3.2、3.3。必须含至少2个<chart/>和1个<table/>。"
    elif name == "问题与原因分析":
        sp = "\n- 小标题：4.1、4.2、4.3。必须含至少1个<chart/>和1个<table/>。"
    elif name == "解决方案":
        sp = "\n- 小标题：5.1、5.2、5.3。必须含至少1个<chart/>"
    elif name == "结论":
        sp = "\n- 300-500字总结，无编号。"
    else:
        sp = ""
    cr = build_chart_rule() if name in ["现状分析", "问题与原因分析", "解决方案"] else ""
    hd = "禁止任何标题格式" if name in ["引言", "国内外研究现状", "结论"] else f"小标题：{num}.1 / {num}.2格式"

    revision_section = ""
    if chapter_rev:
        issues = "\n".join([f"- {i}" for i in chapter_rev.get("issues", [])])
        fixes = "\n".join([f"- {f}" for f in chapter_rev.get("fix_requirements", [])])
        keys = "\n".join([f"- {k}" for k in chapter_rev.get("key_points", [])])
        if issues or fixes or keys:
            revision_section = f"""
【重写要求 - 必须严格遵守】
本章在原论文中存在以下问题：
{issues}

修正方向：
{fixes}

必须包含的核心内容：
{keys}
"""

    prompt = f"""写《{profile['title']}》第{num}章 {name}。字数：{MG_WORD_LIMITS[name]}内。
格式铁律：1.{hd} 2.无大标题 3.无markdown符号 4.无加粗 5.纯文本正文{sp}
{cr} 对象：{profile['company']} 行业：{profile['industry']} 只输出正文。{revision_section}"""
    text = call_llm(prompt)
    return re.sub(r'^第[一二三四五六\d]+章.*?\n', '', text).strip()


# ================== [ 重写 TXT 管道 ] ==================
def run_revise_txt_pipeline(task_id: str, profile: dict, revision_plan: dict):
    """修改版论文生成管道：复用原 run_txt_pipeline 结构，融入 revision_plan"""
    folder = f"output/{task_id}"
    os.makedirs(folder, exist_ok=True)

    def update(msg, prog):
        tasks_db[task_id].update({"msg": msg, "progress": prog})
        print(f"[{prog}%] {msg}")

    try:
        update("正在生成摘要...", 5)
        special_notes = revision_plan.get("special_notes", "")
        theory_adjustment = revision_plan.get("theory_adjustment", "")
        abstract_prompt = f"写一段300字以内的论文摘要，只输出摘要正文，不要任何标题和说明文字。\n论文题目：{profile['title']}\n研究对象：{profile['company']}\n行业：{profile['industry']}"
        if special_notes or theory_adjustment:
            abstract_prompt += f"\n\n【特别说明】\n{special_notes}\n{theory_adjustment}"
        abstract = call_llm(abstract_prompt, 600)
        # 提取真正的摘要内容（去掉废话）
        abstract = re.sub(r'^.*?\n---.*?\n\*\*摘要\*\*', '', abstract, flags=re.DOTALL)
        abstract = re.sub(r'^.*这是一份.*?\n', '', abstract)
        abstract = re.sub(r'^摘\s*要[：:\s]*', '', abstract.strip())
        abstract = re.sub(r'^\*\*摘要\*\*', '', abstract.strip())
        abstract = abstract.strip()

        # 生成关键词
        keywords_prompt = f"根据以下论文信息生成3-5个关键词，用分号隔开。\n标题：{profile['title']}\n行业：{profile['industry']}\n公司：{profile['company']}\n只输出关键词，如：关键词1；关键词2；关键词3"
        keywords = call_llm(keywords_prompt, max_tokens=200).strip()
        # 清理关键词结果（去掉"关键词："等前缀，只取分号分隔的部分）
        keywords = re.sub(r'^关键词[：:\s]*', '', keywords, flags=re.IGNORECASE)
        keywords = re.sub(r'^\*\*关键词\*\*[：:\s]*', '', keywords)
        # 如果有关键词后面跟了其他说明，只取关键词部分
        if '。' in keywords:
            keywords = keywords.split('。')[0]
        keywords = keywords.strip()

        # 先生成所有章节内容
        chapter_revisions = revision_plan.get("chapter_revisions", {})
        chapters_content = {}
        for name, num in MG_CHAPTERS:
            chapter_rev = chapter_revisions.get(name)
            content = revise_gen_chapter(name, num, profile, chapter_rev, update)
            chapters_content[name] = content

        # 解析各章二级/三级标题用于目录
        toc_lines = []
        for name, num in MG_CHAPTERS:
            line_name = f"第{num}章 {name}"
            toc_lines.append(line_name)
            content = chapters_content.get(name, "")
            for line in content.split('\n'):
                m2 = re.match(r'^(\d+\.\d+)\s+(.+)', line.strip())
                if m2:
                    toc_lines.append(f"    {m2.group(1)} {m2.group(2)}")
                m3 = re.match(r'^(\d+\.\d+\.\d+)\s+(.+)', line.strip())
                if m3:
                    toc_lines.append(f"        {m3.group(1)} {m3.group(2)}")
        # 目录末尾加参考文献
        toc_lines.append("参考文献")

        # 合并论文（封面已有题目，正文不再重复）
        txt = "---PAGE_BREAK---\n摘要\n" + abstract.strip() + "\n\n关键词\n" + keywords + "\n\n"
        txt += "---PAGE_BREAK---\n目录\n" + "\n".join(toc_lines) + "\n\n"
        txt += "---PAGE_BREAK---\n"
        for name, num in MG_CHAPTERS:
            content = chapters_content.get(name, "")
            txt += f"第{num}章 {name}\n{content}\n\n"
            txt += "---PAGE_BREAK---\n"

        update("正在生成参考文献...", 85)
        refs_prompt = f"""生成12-15条规范参考文献，类型包括书籍[M]、期刊论文[J]、学位论文[D]。
主题涉及：{profile['title']}、{profile['industry']}相关理论。
每条占一行，格式如：[1] 作者. 书名[M]. 出版社, 年份.
"""
        refs_res = call_llm(refs_prompt, max_tokens=1500)
        refs = [r.strip() for r in refs_res.split('\n') if r.strip() and (r.strip()[0].isdigit() or r.strip().startswith('['))]
        refs = refs[:15]
        txt += "参考文献\n" + "\n".join(refs)

        txt_path = os.path.join(folder, "00_完整论文.txt")
        with open(txt_path, 'w', encoding='utf-8') as f:
            f.write(txt)
        update("TXT生成完毕！", 100)
        tasks_db[task_id].update({"status": "completed", "files": [txt_path]})
    except Exception as e:
        tasks_db[task_id].update({"status": "error", "msg": f"生成失败: {str(e)}"})
        print(f"生成失败: {e}")


# ================== [ DOCX 转换 ] ==================
def txt_to_docx_safe(txt_path, docx_path, update, cover_info=None):
    """完全仿照本地双击运行：从硬盘读TXT，写硬盘DOCX"""
    update("正在读取TXT...", 10)
    with open(txt_path, "r", encoding="utf-8") as f:
        full_text = f.read()
    update("正在提取图表数据...", 30)
    charts, tables = extract_charts(full_text), extract_tables(full_text)
    els = [(c["start_pos"], c["end_pos"], "chart", c) for c in charts] + [(t["start_pos"], t["end_pos"], "table", t) for t in tables]
    els.sort(key=lambda x: x[0])
    parts, le = [], 0
    for s, e, t, d in els:
        if s > le:
            parts.append(("text", full_text[le:s]))
        parts.append((t, d))
        le = e
    if le < len(full_text):
        parts.append(("text", full_text[le:]))
    update("正在生成Word排版...", 60)
    doc = Document()
    # 设置默认样式
    style = doc.styles['Normal']
    style.font.name = '宋体'
    style._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    style.font.size = Pt(12)
    if cover_info:
        build_cover(doc, cover_info)
    cn, tn = 0, 0
    is_first_content = True
    pt_last_heading = ""
    in_toc_section = False  # 标记是否在目录区域
    next_page_break = False  # 跨 part 的分页标记

    for pt, ct in parts:
        if pt == "chart":
            cn += 1
            try:
                add_c(doc, ct, chart_to_bytes(ct), cn)
            except Exception as e:
                print(f"生成图表失败: {e}")
        elif pt == "table":
            tn += 1
            try:
                add_t(doc, ct, tn)
            except Exception as e:
                print(f"生成表格失败: {e}")
        else:
            text = ct.strip()
            if not text:
                continue
            lines = text.split('\n')
            for raw_line in lines:
                line = clean_text(raw_line).strip()
                if not line:
                    continue

                # 处理显式分页标记
                if line == '---PAGE_BREAK---':
                    next_page_break = True
                    in_toc_section = False  # 分页标记也退出目录区域
                    continue

                # 如果上一段标记了分页，在当前段落加 pageBreakBefore
                if next_page_break:
                    need_page_break = True
                    next_page_break = False
                else:
                    need_page_break = False

                # 目录区域（"目录"之后的内容，直到遇到正文标题）
                if line == '目录':
                    is_level1 = True
                else:
                    is_level1 = bool(re.match(r'^第[一二三四五六\d]+章', line)) or bool(re.match(r'^摘\s*要|^关键词|^参考文献', line))

                is_level3 = bool(re.match(r'^\d+\.\d+\.\d+\s+', line))
                is_level2 = bool(re.match(r'^\d+\.\d+\s+', line)) and not is_level3
                is_toc_item = in_toc_section

                # 检测离开目录区域
                if in_toc_section and is_level1 and not re.match(r'^第[一二三四五六\d]+章', line):
                    in_toc_section = False
                # 进入目录区域
                if line == '目录':
                    in_toc_section = True

                if is_level1 and not is_toc_item:
                    # 去重：当前一级标题与上一个相同则跳过
                    if pt_last_heading:
                        stripped_last = re.sub(r'\s+', '', pt_last_heading.replace('第', '').replace('章', '')).strip()
                        stripped_cur = re.sub(r'\s+', '', line).strip()
                        if stripped_cur == stripped_last:
                            pt_last_heading = ""
                            continue
                    pt_last_heading = line

                    # 一级标题
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p.paragraph_format.space_before = Pt(12)
                    p.paragraph_format.space_after = Pt(12)
                    r = p.add_run(line)
                    set_run_font(r, "黑体", 16, bold=True)
                    is_first_content = False

                    # 需要在当前段落前加分页（来自 ---PAGE_BREAK--- 标记）
                    if need_page_break:
                        pPr = p._element.get_or_add_pPr()
                        from docx.oxml import OxmlElement
                        pb = OxmlElement('w:pageBreakBefore')
                        pPr.append(pb)
                    continue

                pt_last_heading = ""

                # 目录条目（"目录"~"参考文献"之间的行）—— 所有内容都用宋体14pt
                if is_toc_item:
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    r = p.add_run(line)
                    set_run_font(r, "宋体", 14)
                    p.paragraph_format.space_before = Pt(2)
                    p.paragraph_format.space_after = Pt(2)
                    is_first_content = False
                    continue

                if is_level3:
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    r = p.add_run(line)
                    set_run_font(r, "宋体", 12, bold=True)
                    is_first_content = False
                elif is_level2:
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    p.paragraph_format.space_before = Pt(6)
                    p.paragraph_format.space_after = Pt(6)
                    r = p.add_run(line)
                    set_run_font(r, "黑体", 14, bold=True)
                    is_first_content = False
                else:
                    p = doc.add_paragraph()
                    p.paragraph_format.first_line_indent = Cm(0.74)
                    p.paragraph_format.line_spacing = 1.5
                    r = p.add_run(line)
                    set_run_font(r, "宋体", 12)
                    is_first_content = False
    doc.save(docx_path)
    update("DOCX生成完毕！", 100)


def finalize_docx(docx_path, update):
    """分节符修正 + PDF 真实页码 + 目录更新 + 页脚"""
    update("正在处理分节符和页码...", 92)
    doc = Document(docx_path)
    paragraphs = doc.paragraphs
    body = doc.element.body

    toc_idx = -1
    for i, p in enumerate(paragraphs):
        if p.text.strip() == '目录':
            toc_idx = i
            break

    chapter1_idx = -1
    for i, p in enumerate(paragraphs):
        t = p.text.strip()
        if re.match(r'^第1章\s', t) and i > toc_idx + 3 and '\t' not in t:
            chapter1_idx = i
            break

    if chapter1_idx < 0:
        print("[跳过] 未找到第1章")
        return

    toc_end_idx = chapter1_idx - 1
    while toc_end_idx > toc_idx:
        if paragraphs[toc_end_idx].text.strip():
            break
        toc_end_idx -= 1

    toc_headings = []
    toc_range_start = toc_idx + 1 if toc_idx >= 0 else 0
    for i in range(toc_range_start, chapter1_idx):
        t = paragraphs[i].text.strip()
        if t:
            heading = t.split('\t')[0].strip() if '\t' in t else t
            toc_headings.append(heading)

    # 保存临时副本 → 转 PDF → 算页码
    tag = uuid.uuid4().hex[:8]
    temp_docx = docx_path.replace('.docx', f'_temp_{tag}.docx')
    doc.save(temp_docx)

    update("正在用 LibreOffice 渲染 PDF 计算页码...", 95)
    pdf_path = os.path.join(os.path.dirname(temp_docx) or '.', f'_temp_{tag}.pdf')
    subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf',
                    temp_docx, '--outdir', os.path.dirname(pdf_path)],
                   capture_output=True, text=True, timeout=120)
    expected = temp_docx.replace('.docx', '.pdf')
    if os.path.exists(expected):
        os.rename(expected, pdf_path)

    heading_abs_pages = {}
    if os.path.exists(pdf_path):
        result = subprocess.run(['pdfinfo', pdf_path], capture_output=True, text=True)
        num_pages = 0
        for line in result.stdout.split('\n'):
            if line.startswith('Pages'):
                num_pages = int(line.split(':')[1].strip())
                break

        # 找目录页并跳过
        toc_pdf_page = -1
        for pg in range(1, num_pages + 1):
            r = subprocess.run(['pdftotext', '-f', str(pg), '-l', str(pg), pdf_path, '-'],
                               capture_output=True, text=True)
            lines = [l.strip() for l in r.stdout.split('\n') if l.strip()]
            if any(l == '目录' for l in lines):
                toc_pdf_page = pg
                break

        start_pg = toc_pdf_page + 1 if toc_pdf_page >= 0 else 1
        for pg in range(start_pg, num_pages + 1):
            r = subprocess.run(['pdftotext', '-f', str(pg), '-l', str(pg), pdf_path, '-'],
                               capture_output=True, text=True)
            text_flat = re.sub(r'\s+', '', r.stdout)
            for h in toc_headings:
                if h in heading_abs_pages:
                    continue
                h_flat = re.sub(r'\s+', '', h)
                if h_flat in text_flat:
                    heading_abs_pages[h] = pg

        ch1_heading = next((h for h in toc_headings if h.startswith('第1章')), None)
        ch1_page = heading_abs_pages.get(ch1_heading, 1)
        offset = ch1_page - 1

        # 更新目录页码
        update("正在更新目录页码...", 97)
        for i in range(toc_range_start, chapter1_idx):
            p = paragraphs[i]
            t = p.text.strip()
            if not t:
                continue
            heading_text = t.split('\t')[0].strip() if '\t' in t else t
            if heading_text in heading_abs_pages:
                sec_page = heading_abs_pages[heading_text] - offset
                if sec_page >= 1:
                    p.clear()
                    tab_stops = p.paragraph_format.tab_stops
                    tab_stops.add_tab_stop(Cm(16), alignment=WD_TAB_ALIGNMENT.RIGHT,
                                           leader=WD_TAB_LEADER.DOTS)
                    r = p.add_run(heading_text)
                    set_run_font(r, "宋体", 14)
                    p.alignment = 0
                    p.add_run("\t")
                    rp = p.add_run(str(sec_page))
                    set_run_font(rp, "宋体", 14)

        if os.path.exists(pdf_path):
            os.remove(pdf_path)

    # 分节符
    last_sect_pr = None
    for child in list(body):
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'sectPr':
            last_sect_pr = child

    if last_sect_pr is not None:
        template = deepcopy(last_sect_pr)
        body.remove(last_sect_pr)

        toc_end_elem = paragraphs[toc_end_idx]._element
        toc_end_pPr = toc_end_elem.find(qn('w:pPr'))
        if toc_end_pPr is None:
            toc_end_pPr = OxmlElement('w:pPr')
            toc_end_elem.insert(0, toc_end_pPr)
        sPr0 = OxmlElement('w:sectPr')
        for c in template:
            tag = c.tag.split('}')[-1] if '}' in c.tag else c.tag
            if tag in ('pgSz', 'pgMar'):
                sPr0.append(deepcopy(c))
        type_elem = OxmlElement('w:type')
        type_elem.set(qn('w:val'), 'continuous')
        sPr0.append(type_elem)
        toc_end_pPr.append(sPr0)

        ts = OxmlElement('w:sectPr')
        for c in template:
            tag = c.tag.split('}')[-1] if '}' in c.tag else c.tag
            if tag in ('pgSz', 'pgMar'):
                ts.append(deepcopy(c))
        pgnt = OxmlElement('w:pgNumType')
        pgnt.set(qn('w:start'), '1')
        ts.append(pgnt)
        body.append(ts)

    # 移除第1章的 pageBreakBefore
    ch1_elem = paragraphs[chapter1_idx]._element
    ch1_pPr = ch1_elem.find(qn('w:pPr'))
    if ch1_pPr is not None:
        for pb in ch1_pPr.findall(qn('w:pageBreakBefore')):
            ch1_pPr.remove(pb)

    # 页脚
    try:
        sec0 = doc.sections[0]
        sec0.different_first_page_header_footer = True
        for fn in ['footer', 'even_page_footer', 'first_page_footer']:
            try:
                f = getattr(sec0, fn)
                f.is_linked_to_previous = False
                for pf in f.paragraphs:
                    pf.clear()
            except Exception:
                pass
    except Exception:
        pass

    try:
        if len(doc.sections) > 1:
            sec1 = doc.sections[1]
            footer = sec1.footer
            footer.is_linked_to_previous = False
            for pf in footer.paragraphs:
                pf.clear()
            pf = footer.paragraphs[0]
            pf.alignment = 1
            run = pf.add_run()
            fc1 = OxmlElement('w:fldChar')
            fc1.set(qn('w:fldCharType'), 'begin')
            run._element.append(fc1)
            it = OxmlElement('w:instrText')
            it.set(qn('xml:space'), 'preserve')
            it.text = ' PAGE '
            run._element.append(it)
            fc2 = OxmlElement('w:fldChar')
            fc2.set(qn('w:fldCharType'), 'end')
            run._element.append(fc2)
    except Exception as e:
        print(f"页脚处理异常: {e}")

    doc.save(docx_path)
    if os.path.exists(temp_docx):
        os.remove(temp_docx)
    update("最终排版完成", 100)


# ================== [ 主入口 ] ==================
def main():
    parser = argparse.ArgumentParser(description="论文修改程序 - 经管类论文深度重构")
    parser.add_argument("--input", "-i", required=True, help="输入的DOCX文件路径")
    parser.add_argument("--output", "-o", default="./output", help="输出目录（默认./output）")
    args = parser.parse_args()

    if not os.path.isfile(args.input):
        print(f"错误：输入文件不存在 {args.input}")
        return

    os.makedirs(args.output, exist_ok=True)

    print("=" * 50)
    print("【论文修改程序】启动")
    print("=" * 50)

    print("\n步骤1/4：提取DOCX文本...")
    original_text = extract_docx_text(args.input)
    if not original_text:
        print("错误：未能从DOCX中提取文本，程序终止。")
        return
    print(f"  提取成功，共 {len(original_text)} 字符")

    cover_info = extract_cover_info(args.input)

    print("\n步骤2/4：分析原文画像...")
    analysis_result = analyze_original_paper(original_text)
    print(f"  识别标题：{analysis_result.get('title', 'N/A')}")
    print(f"  识别专业：{analysis_result.get('major', 'N/A')}")
    print(f"  识别对象：{analysis_result.get('company', 'N/A')}")

    print("\n步骤3/4：诊断与重构...")
    revision_plan = diagnose_and_reconstruct(analysis_result)
    report_path = os.path.join(args.output, "诊断报告.txt")
    save_diagnosis_report(revision_plan, report_path)
    print(f"  诊断报告已保存：{report_path}")
    diag = revision_plan.get("diagnosis_report", {})
    print(f"  总体评分：{diag.get('overall_score', 'N/A')}")
    print(f"  共发现 {len(diag.get('issues', []))} 个问题")

    major = revision_plan.get("revised_profile", {}).get("major", analysis_result.get("major", "工商管理"))
    profile = merge_into_standard_profile(revision_plan.get("revised_profile", {}), major)

    task_id = str(uuid.uuid4())
    tasks_db[task_id] = {"status": "running", "progress": 0, "msg": "启动", "files": []}

    print(f"\n步骤4/4：生成修改版论文，任务ID：{task_id}")
    run_revise_txt_pipeline(task_id, profile, revision_plan)

    txt_path = os.path.join("output", task_id, "00_完整论文.txt")
    if not os.path.exists(txt_path):
        print("错误：TXT生成失败，请检查日志。")
        return

    docx_path = os.path.join("output", task_id, "论文_修改版.docx")

    def update_fn(msg, prog):
        print(f"[DOCX {prog}%] {msg}")

    txt_to_docx_safe(txt_path, docx_path, update_fn, cover_info)
    finalize_docx(docx_path, update_fn)

    final_docx = os.path.join(args.output, "论文_修改版.docx")
    shutil.move(docx_path, final_docx)
    print(f"\n{'=' * 50}")
    print(f"全部完成！")
    print(f"诊断报告：{report_path}")
    print(f"修改版论文：{final_docx}")
    print(f"{'=' * 50}")


if __name__ == "__main__":
    main()
