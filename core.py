# -*- coding: utf-8 -*-
"""
core.py — 论文系统基础层
=======================
LLM调用、DOCX排版工具、图表/图片渲染、标签解析
所有模块（profile/generator/renderer/reviser）都基于此层
"""
import os, re, io, json, time, uuid, threading, shutil, base64
import subprocess
from copy import deepcopy
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import requests
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from PIL import Image, ImageDraw, ImageFont
plt.rcParams["font.sans-serif"] = ["WenQuanYi Micro Hei", "SimHei"]
plt.rcParams["axes.unicode_minus"] = False

from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

try:
    from mechanical_cad import generate_engineering_png, freecad_available
except Exception:
    generate_engineering_png = None

# 确定性标签守卫（渲染前校验/修复/降级，无LLM）
try:
    from tagguard import (audit_and_repair, render_report_text,
                          check_numeric_consistency, consistency_summary,
                          normalize_chart_type)
except Exception:
    audit_and_repair = None
    render_report_text = None
    check_numeric_consistency = None
    consistency_summary = None

    def normalize_chart_type(t):
        return (t or "bar").strip().lower()

# ==================== 配置 ====================
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY", "")
DEEPSEEK_BASE_URL = os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com/v1")
DEEPSEEK_MODEL = os.getenv("DEEPSEEK_MODEL", "deepseek-chat")
IMAGE_API_KEY = os.getenv("IMAGE_API_KEY", "")
IMAGE_API_URL = os.getenv("IMAGE_API_URL", "https://aihubmix.com/v1/images/generations")
IMAGE_MODEL = "gpt-image-2"
COLORS = ["#5B9BD5", "#ED7D31", "#A5A5A5", "#FFC000", "#4472C4", "#70AD47", "#264478"]

# ==================== 章节配置 ====================
MG_WORD_LIMITS = {"引言": 1000, "国内外研究现状": 1500, "现状分析": 2500, "问题与原因分析": 3000, "解决方案": 3000, "结论": 500}
MG_SOFT_MAX = {"引言": 1200, "国内外研究现状": 1800, "现状分析": 3500, "问题与原因分析": 4500, "解决方案": 4000, "结论": 800}
MG_CHAPTERS = [("引言", 1), ("国内外研究现状", 2), ("现状分析", 3), ("问题与原因分析", 4), ("解决方案", 5), ("结论", 6)]

SJ_WORD_LIMITS = {"绪论": 800, "理论基础": 1000, "问题发现与分析": 1500, "设计策略与方案": 2500, "总结与反思": 800}
SJ_CHAPTERS = [("绪论", 1), ("理论基础", 2), ("问题发现与分析", 3), ("设计策略与方案", 4), ("总结与反思", 5)]

MC_WORD_LIMITS = {"概述": 900, "原理与方案分析": 1200, "总体设计": 1600, "关键部件设计": 2200, "计算与校核": 1600, "总结": 600}
MC_CHAPTERS = [("概述", 1), ("原理与方案分析", 2), ("总体设计", 3), ("关键部件设计", 4), ("计算与校核", 5), ("总结", 6)]

# 土木工程章节配置 - 标准六个章节
CIV_WORD_LIMITS = {"工程概况": 800, "建筑设计": 1000, "结构设计": 1500, "结构计算": 1200, "施工组织设计": 1500, "结论与展望": 500}
CIV_SOFT_MAX = {"工程概况": 1000, "建筑设计": 1200, "结构设计": 2000, "结构计算": 1800, "施工组织设计": 2000, "结论与展望": 800}
CIV_CHAPTERS = [("工程概况", 1), ("建筑设计", 2), ("结构设计", 3), ("结构计算", 4), ("施工组织设计", 5), ("结论与展望", 6)]

# 土木工程-扩展章节定义（用于施工组织/造价/管理等变体）
CIV2_WORD_LIMITS = {"工程概况": 800, "施工方案": 1000, "施工进度计划": 1000, "施工平面布置": 800, "质量安全保证措施": 1000, "结论与展望": 500}

# 论文类型映射
PAPER_TYPE_MAP = {
    "管理": MG_CHAPTERS,
    "设计": SJ_CHAPTERS,
    "机械": MC_CHAPTERS,
    "土木": CIV_CHAPTERS,
    "civil": CIV_CHAPTERS,
}
PAPER_LIMIT_MAP = {
    "管理": MG_WORD_LIMITS,
    "设计": SJ_WORD_LIMITS,
    "机械": MC_WORD_LIMITS,
    "土木": CIV_WORD_LIMITS,
    "civil": CIV_WORD_LIMITS,
}

# ==================== 单线程排队锁 ====================
tasks_db = {}
is_system_busy = False
system_lock = threading.Lock()


# ==================== LLM 调用 ====================
def call_llm(prompt: str, max_tokens: int = 4096, retry: int = 3) -> str:
    """调用 DeepSeek API，重试 retry 次
    - 默认附加 enable_thinking=false（推理模型如 deepseek-v4-flash 若开着思考，
      会先耗尽 max_tokens 导致正文为空）；设置 DEEPSEEK_EXTRA_JSON 可覆盖附加参数
    - 200但正文为空（思考耗尽tokens）→ 升高 max_tokens 重试"""
    if not DEEPSEEK_API_KEY:
        raise RuntimeError("DEEPSEEK_API_KEY is not set")
    headers = {"Authorization": f"Bearer {DEEPSEEK_API_KEY}", "Content-Type": "application/json"}
    try:
        extra_params = json.loads(os.getenv("DEEPSEEK_EXTRA_JSON", '{"enable_thinking": false}'))
    except Exception:
        extra_params = {}
    data = {
        "model": DEEPSEEK_MODEL,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.7,
        "max_tokens": max_tokens
    }
    if isinstance(extra_params, dict):
        data.update(extra_params)
    cur_max_tokens = max_tokens
    for attempt in range(retry):
        try:
            data["max_tokens"] = cur_max_tokens
            resp = requests.post(
                f"{DEEPSEEK_BASE_URL}/chat/completions",
                headers=headers, json=data, timeout=180
            )
            if resp.status_code == 200:
                payload = resp.json()
                choice = (payload.get("choices") or [{}])[0] or {}
                finish_reason = choice.get("finish_reason") or ""
                content = (choice.get("message") or {}).get("content") or ""
                content = content.strip()
                # length 表示服务端达到 token 上限；即使末尾有标点也可能是完整句中断，必须重试。
                if finish_reason == "length":
                    print(f"LLM输出因length截断(尝试{attempt+1}/{retry})，提高max_tokens重试")
                    cur_max_tokens = min(int(cur_max_tokens * 1.8) + 1000, 16000)
                    if attempt < retry - 1:
                        time.sleep(2)
                        continue
                    raise RuntimeError("LLM输出达到max_tokens仍被截断，阻断生成")
                if content:
                    return content
                # 正文为空：多为推理模型思考耗尽tokens → 升高max_tokens重试
                cur_max_tokens = min(int(cur_max_tokens * 1.8) + 1000, 16000)
                print(f"LLM返回空正文(尝试{attempt+1}/{retry})，升至max_tokens={cur_max_tokens}重试")
                if attempt < retry - 1:
                    time.sleep(2)
            elif resp.status_code == 400 and extra_params and "enable_thinking" in resp.text:
                # 该端点不支持 enable_thinking 参数 → 去掉附加参数重试
                for k in extra_params:
                    data.pop(k, None)
                extra_params = {}
                print(f"端点不支持enable_thinking，已移除附加参数(尝试{attempt+1}/{retry})")
            else:
                print(f"LLM调用失败(尝试{attempt+1}/{retry}): {resp.status_code} {resp.text[:200]}")
                if attempt < retry - 1:
                    time.sleep(2)
        except Exception as e:
            print(f"LLM异常(尝试{attempt+1}/{retry}): {e}")
            if attempt < retry - 1:
                time.sleep(2)
    print("LLM调用失败，返回空")
    return ""


# ==================== 文本清理 ====================
def clean_text(text: str, strip_markdown: bool = True) -> str:
    """清理文本，可选剥离markdown符号"""
    if not text:
        return ""
    if strip_markdown:
        # 剥离markdown标题符号
        text = re.sub(r'^#{1,6}\s*', '', text, flags=re.MULTILINE)
        # 剥离markdown加粗/斜体/行内代码
        text = re.sub(r'\*\*|\*|`|>', '', text)
        # 剥离装饰符号
        text = re.sub(r'[★●■◇◆○◎]', '', text)
        # 全角括号转半角
        text = text.replace('［', '[').replace('］', ']')
        text = text.replace('（', '(').replace('）', ')')
        # 剥离markdown表格行（兜底）
        text = re.sub(r'^\|.*\|.*\|$', '', text, flags=re.MULTILINE)
        text = re.sub(r'^\|[\s\-:]+\|[\s\-:]+\|', '', text, flags=re.MULTILINE)
        # 剥离误残留的孤立斜杠行
        text = re.sub(r'^\s*[\\/／]\s*$', '', text, flags=re.MULTILINE)
        # 合并多余空格
        text = re.sub(r' {2,}', ' ', text)
        text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()


def set_run_font(run, font_name: str, font_size: int, bold: bool = False, color=None):
    """设置 run 字体属性"""
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(font_size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color


def _init_s(doc):
    """初始化 docx section 默认样式"""
    s = doc.sections[0]
    s.page_width = Cm(21.0)
    s.page_height = Cm(29.7)
    s.top_margin = Cm(2.54)
    s.bottom_margin = Cm(2.54)
    s.left_margin = Cm(3.175)
    s.right_margin = Cm(3.175)
    style = doc.styles['Normal']
    style.font.name = '宋体'
    style._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    style.font.size = Pt(12)
    return doc


def _wt(doc, text, font='宋体', size=12, bold=False, alignment=None,
        first_line_indent=None, spacing=None, color=None) -> 'Paragraph':
    """简化版添加段落到 docx"""
    p = doc.add_paragraph()
    if alignment is not None:
        p.alignment = alignment
    if first_line_indent is not None:
        p.paragraph_format.first_line_indent = first_line_indent
    if spacing is not None:
        p.paragraph_format.line_spacing = spacing
    r = p.add_run(text)
    set_run_font(r, font, size, bold, color)
    return p


# ==================== 封面 ====================
def extract_cover_info(docx_path: str) -> dict:
    """从DOCX前25段提取封面信息"""
    info = {"title": "", "name": "", "student_id": "", "major": "",
            "level": "", "advisor": "", "year": "", "month": "", "day": ""}
    try:
        doc = Document(docx_path)
        for p in doc.paragraphs[:25]:
            t = p.text.strip()
            if not t:
                continue
            def _match(patterns):
                for pat in patterns:
                    m = re.search(pat, t)
                    if m:
                        return m
                return None
            m = _match([r'论\s*文\s*题\s*目[：:]\s*(.+)', r'题\s*目[：:]\s*(.+)'])
            if m: info["title"] = m.group(1).strip()
            m = _match([r'姓\s*名[：:]\s*(.+)', r'学生姓名[：:]\s*(.+)'])
            if m: info["name"] = m.group(1).strip()
            m = _match([r'学\s*号[：:]\s*(.+)', r'准考证号[：:]\s*(.+)', r'学籍号[：:]\s*(.+)'])
            if m: info["student_id"] = m.group(1).strip()
            m = _match([r'专\s*业[：:]\s*(.+)'])
            if m: info["major"] = m.group(1).strip()
            m = _match([r'层\s*次[：:]\s*(.+)', r'学\s*历[：:]\s*(.+)'])
            if m: info["level"] = m.group(1).strip()
            m = _match([r'指导教师[：:]\s*(.+)', r'指导老师[：:]\s*(.+)'])
            if m: info["advisor"] = m.group(1).strip()
            m = re.search(r'(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日', t)
            if m:
                info["year"], info["month"], info["day"] = m.group(1), m.group(2), m.group(3)
    except Exception as e:
        print(f"提取封面失败: {e}")
    return info


def build_cover(doc, cover_info: dict):
    """在DOCX文档开头生成标准封面页"""
    cover_info = cover_info or {}
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
    doc.add_paragraph().alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 毕业论文（48pt加粗宋体）
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("毕 业 论 文")
    set_run_font(r, "宋体", 48, bold=True)

    # 空行
    doc.add_paragraph()

    # 信息项
    labels = [
        ("论文题目", "title"), ("姓    名", "name"), ("学    号", "student_id"),
        ("专    业", "major"), ("层    次", "level"), ("指导教师", "advisor"),
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

    # 3空行
    for _ in range(3):
        doc.add_paragraph()

    # 日期
    y, m, d = cover_info.get("year"), cover_info.get("month"), cover_info.get("day")
    if y and m and d:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r = p.add_run(f"{y}年{m}月{d}日")
        set_run_font(r, "宋体", 16, bold=True)

    doc.add_page_break()


def normalize_cover_info(cover_info: Optional[dict] = None) -> dict:
    """补齐封面字段，允许空封面先输出。"""
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
    if cover_info:
        for k, v in cover_info.items():
            if k in info and v is not None:
                info[k] = str(v)
    return info


def extract_title_from_txt(text: str) -> str:
    """从TXT头部提取论文题目。"""
    if not text:
        return ""
    for line in text.splitlines()[:12]:
        t = line.strip()
        if not t:
            continue
        m = re.match(r'^(?:论文题目|题目)\s*[:：]\s*(.+)$', t)
        if m:
            return m.group(1).strip()
    return ""


def strip_title_from_txt(text: str) -> str:
    """移除TXT头部的题目行，保留正文/目录内容。"""
    if not text:
        return text
    lines = text.splitlines()
    out = []
    skipped = False
    i = 0
    while i < len(lines):
        line = lines[i]
        if not skipped and not line.strip():
            i += 1
            continue
        if not skipped and re.match(r'^(?:论文题目|题目)\s*[:：]\s*.+$', line.strip()):
            skipped = True
            i += 1
            while i < len(lines) and not lines[i].strip():
                i += 1
            continue
        out.extend(lines[i:])
        break
    return "\n".join(out).lstrip()


# ==================== 预处理：标签属性跨行/裸属性修复 ====================

def preprocess_tag_attrs(text):
    """
    预处理TXT文本，兼容GPT输出格式变体：
    1. 自闭合<tag.../>后的独立属性行 → 合并回标签内
    2. 无标签包裹的裸属性行（连续的key=value行）→ 包裹为<tag/>
    """
    lines = text.split('\n')
    result = []
    i = 0
    keys_table = ['header', 'rows', 'data', 'data_source']
    keys_chart = ['type', 'x', 'y', 'legend', 'unit', 'data_source']
    all_attr_keys = set(keys_table + keys_chart)
    
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()
        
        # 1. 自闭合标签
        m = re.match(r'^(<(?:table|chart)\s+[^>]*)/>\s*$', stripped)
        if m:
            tag_type = m.group(1)
            j, collected = _collect_next_attrs(lines, i+1, all_attr_keys)
            if collected:
                extra = ' '.join(f'{k}="{v}"' for k, v in collected.items())
                result.append(f'{tag_type} {extra}/>')
                i = j
                continue
            result.append(line)
            i += 1
            continue
        
        # 2. 裸属性行
        kv_start = re.match(r'^(\w+)=', stripped)
        if kv_start and kv_start.group(1) in all_attr_keys:
            first_key = kv_start.group(1)
            is_table = first_key in keys_table
            j, collected = _collect_next_attrs(lines, i, all_attr_keys)
            if collected:
                tag_type = 'table' if (is_table or ('header' in collected and 'rows' in collected)) else 'chart'
                extra = ' '.join(f'{k}="{v}"' for k, v in collected.items())
                result.append(f'<{tag_type} {extra}/>')
                i = j
                continue
        
        result.append(line)
        i += 1
    
    return '\n'.join(result)

def _collect_next_attrs(lines, start_idx, valid_keys):
    """从start_idx开始收集连续属性行，返回(结束索引, 属性字典)"""
    collected = {}
    current_key = None
    current_val_lines = []
    j = start_idx
    has_any = False
    
    while j < len(lines):
        stripped = lines[j].strip()
        if not stripped:
            break
        
        kv = re.match(r'^(\w+)=["\u201c\u201d]?(.*)', stripped)
        if kv and kv.group(1) in valid_keys:
            has_any = True
            if current_key and current_val_lines:
                collected[current_key] = ' '.join(current_val_lines)
            
            current_key = kv.group(1)
            after_eq = kv.group(2)
            
            closed = False
            for qc in ['"', '\u201d', '"', '\u201c']:
                idx = after_eq.rfind(qc)
                if idx >= 0:
                    val = after_eq[:idx]
                    collected[current_key] = val
                    current_key = None
                    current_val_lines = []
                    closed = True
                    break
            
            if not closed:
                clean = after_eq.rstrip('"').rstrip('\u201d').rstrip('"').rstrip('\u201c')
                if clean != after_eq:
                    collected[current_key] = clean
                    current_key = None
                    current_val_lines = []
                else:
                    current_val_lines = [after_eq]
            
            j += 1
        elif current_key:
            has_close = False
            for qc in ['"', '\u201d', '"', '\u201c']:
                idx = stripped.rfind(qc)
                if idx >= 0:
                    current_val_lines.append(stripped[:idx])
                    collected[current_key] = ' '.join(current_val_lines)
                    current_key = None
                    current_val_lines = []
                    has_close = True
                    break
            if not has_close:
                current_val_lines.append(stripped)
            j += 1
        else:
            break
    
    if current_key and current_val_lines:
        collected[current_key] = ' '.join(current_val_lines)
    
    return j, collected if has_any else {}


# ==================== 标签解析 ====================

# ==================== 健壮标签解析（三层：标准化→宽容解析→校验） ====================

def tolerant_extract_charts(text):
    """宽容解析<chart/>标签 — 支持属性跨行和中文引号"""
    result = []
    matches = list(re.finditer(r'<chart\b.*?/>', text, flags=re.DOTALL))
    if matches:
        for m in matches:
            tag = m.group()
            attrs = dict(re.findall(r'(\w+)=["\u201c\u201d\']([^"\u201c\u201d\']*)["\u201c\u201d\']', tag))
            if attrs:
                result.append({
                    "id": attrs.get("id", ""),
                    "title": attrs.get("title", ""),
                    "type": attrs.get("type", ""),
                    "x": attrs.get("x", ""),
                    "y": attrs.get("y", ""),
                    "legend": attrs.get("legend", ""),
                    "unit": attrs.get("unit", ""),
                    "data_source": attrs.get("data_source", attrs.get("datasource", "")),
                    "start_pos": m.start(), "end_pos": m.end()
                })
        return result

    for m in re.finditer(r'<chart\s+.*?(?:/>|>|(?=\n))', text, re.DOTALL):
        tag = m.group()
        start_pos = m.start()
        end_pos = m.end()
        attrs = dict(re.findall(r'(\w+)=["\u201c\u201d\']([^"\u201c\u201d\']*)["\u201c\u201d\']', tag))

        if tag.rstrip().endswith('/>'):
            remaining = text[end_pos:]
            next_lines = remaining.split('\n', 8)[:8]
            for line in next_lines:
                stripped = line.strip()
                kv_match = re.match(r'(\w+)=["\u201c\u201d]?(.*?)["\u201c\u201d]?\s*(?:/)?\s*$', stripped)
                if kv_match and kv_match.group(1) in ('type', 'x', 'y', 'legend', 'unit', 'data_source', 'id', 'title'):
                    key = kv_match.group(1)
                    val = kv_match.group(2).rstrip('/').rstrip()
                    if key not in attrs or not attrs[key]:
                        attrs[key] = val
                    end_pos = text.index(line, end_pos) + len(line)
                else:
                    break

        if attrs:
            result.append({
                "id": attrs.get("id", ""),
                "title": attrs.get("title", ""),
                "type": attrs.get("type", ""),
                "x": attrs.get("x", ""),
                "y": attrs.get("y", ""),
                "legend": attrs.get("legend", ""),
                "unit": attrs.get("unit", ""),
                "data_source": attrs.get("data_source", attrs.get("datasource", "")),
                "start_pos": start_pos, "end_pos": end_pos
            })
    return result


def tolerant_extract_tables(text):
    """宽容解析<table/>标签 — 不修改原文本，支持属性跨行"""
    result = []
    # 优先解析完整的自闭合 table 标签，支持属性跨行
    matches = list(re.finditer(r'<table\b.*?/>', text, flags=re.DOTALL))
    if matches:
        for m in matches:
            tag = m.group()
            attrs = dict(re.findall(r'(\w+)=["\u201c\u201d\']([^"\u201c\u201d\']*)["\u201c\u201d\']', tag))
            if 'header' not in attrs and 'headers' in attrs:
                attrs['header'] = attrs['headers']
            if attrs and ('title' in attrs or 'header' in attrs or 'rows' in attrs):
                result.append({
                    "id": attrs.get("id", ""),
                    "title": attrs.get("title", ""),
                    "header": attrs.get("header", ""),
                    "rows": attrs.get("rows", ""),
                    "data": attrs.get("data", ""),
                    "data_source": attrs.get("data_source", attrs.get("datasource", "")),
                    "start_pos": m.start(), "end_pos": m.end()
                })
        return result

    # 兼容旧版不完整输出：逐行扫描首行，再尝试吸收后续属性
    for m in re.finditer(r'<table\s+[^\n]*', text):
        tag = m.group()
        end_pos = m.end()
        start_pos = m.start()

        attrs = dict(re.findall(r'(\w+)=["\u201c\u201d\']([^"\u201c\u201d\']*)["\u201c\u201d\']', tag))
        if 'header' not in attrs and 'headers' in attrs:
            attrs['header'] = attrs['headers']

        remaining = text[end_pos:]
        next_lines = remaining.split('\n', 6)[:6]
        for line in next_lines:
            stripped = line.strip()
            kv_match = re.match(r'(\w+)=["\u201c\u201d]?(.*?)["\u201c\u201d]?\s*(?:/)?\s*$', stripped)
            if kv_match and kv_match.group(1) in ('header', 'rows', 'data', 'data_source', 'id', 'title'):
                key = kv_match.group(1)
                val = kv_match.group(2).rstrip('/').rstrip()
                if key not in attrs or not attrs[key]:
                    attrs[key] = val
                end_pos = text.index(line, end_pos) + len(line)
            else:
                break

        if attrs and ('title' in attrs or 'header' in attrs or 'rows' in attrs):
            result.append({
                "id": attrs.get("id", ""),
                "title": attrs.get("title", ""),
                "header": attrs.get("header", ""),
                "rows": attrs.get("rows", ""),
                "data": attrs.get("data", ""),
                "data_source": attrs.get("data_source", attrs.get("datasource", "")),
                "start_pos": start_pos, "end_pos": end_pos
            })
    return result


def _infer_drawing_type(text: str, fallback: str = "结构图") -> str:
    """根据标题/描述推断图纸类型。"""
    t = (text or "").replace(" ", "")
    rules = [
        ("运动过程图", ["运动过程", "动作过程", "动作时序", "流程", "工艺流程"]),
        ("装配示意图", ["装配示意", "装配关系", "装配图"]),
        ("受力分析图", ["受力分析", "受力", "力学分析", "载荷", "应力"]),
        ("原理图", ["原理图", "工作原理", "原理"]),
        ("总体布局图", ["总体布局", "布局图", "总布置", "总体方案"]),
        ("结构图", ["结构图", "结构示意", "结构分解", "结构"]),
        ("零件图", ["零件图", "零件结构", "零件"]),
        ("传动简图", ["传动简图", "传动链", "传动"]),
        ("流程图", ["流程图", "工艺流程", "流程"]),
        ("安装布局图", ["安装布局", "安装示意", "安装"]),
    ]
    for typ, kws in rules:
        if any(k in t for k in kws):
            return typ
    return fallback


def _is_placeholder_drawing_text(text: str) -> bool:
    """过滤明显的占位文本，避免把半成品 drawing 标签渲染进 DOCX。"""
    raw = (text or "").strip()
    if not raw:
        return True
    compact = re.sub(r'[\s\W_]+', '', raw, flags=re.UNICODE)
    if compact in {"", ">", "<", "/", "图", "图纸", "示意", "概览", "原图", "示意图", "结构图", "设计图", "工程图", "机械图"}:
        return True
    if raw in {">", "<", "/", "／", ">", ">>", "<<", "—", "-", "｜", "|"}:
        return True
    return len(compact) <= 1


def _prefer_structured_drawing_value(base_value: Any, structured_value: Any) -> Any:
    """当基础值明显是占位符时，用结构化值覆盖。"""
    if structured_value in (None, ""):
        return base_value
    if base_value in (None, ""):
        return structured_value
    if _is_placeholder_drawing_text(str(base_value)):
        return structured_value
    return base_value


def _canonicalize_drawing_line(line: str, seq: int = 0,
                               start_pos: int = 0, end_pos: int = 0) -> Optional[dict]:
    """
    解析单行 drawing 标签：
    - 支持严格格式
    - 支持 `<drawing/ 图1-1 xxx，yyy` 之类的半成品格式
    - 返回统一字典，供渲染阶段按顺序编号
    """
    raw = (line or "").strip()
    if "<drawing" not in raw.lower():
        return None

    strict = re.search(
        r'<drawing\s+id="([^"]+)"(?:\s+type="([^"]*)")?\s+title="([^"]+)"(?:\s+description="([^"]*)")?\s*/?>',
        raw, re.DOTALL
    )
    if strict:
        strict_title = strict.group(3).strip()
        strict_desc = (strict.group(4) or "").strip()
        if _is_placeholder_drawing_text(strict_title) and _is_placeholder_drawing_text(strict_desc):
            return None
        if _is_placeholder_drawing_text(strict_title) and strict_desc and not _is_placeholder_drawing_text(strict_desc):
            strict_title = strict_desc
        attrs = {
            "id": strict.group(1),
            "type": strict.group(2) or "",
            "title": strict_title,
            "description": strict_desc,
        }
        attrs["seq"] = seq or 0
        attrs["start_pos"] = start_pos
        attrs["end_pos"] = end_pos
        return attrs

    attrs = dict(re.findall(r'(\w+)=["\']([^"\']*)["\']', raw))
    tail = raw
    # 去掉前缀与多余闭合符
    tail = re.sub(r'(?is)^.*?<drawing\s*/*\s*', '', tail)
    tail = tail.replace('<drawing', '').strip()
    tail = re.sub(r'/\s*>?\s*$', '', tail).strip()
    tail = re.sub(r'^[图图]\s*[\d一二三四五六七八九十]+(?:[-.]\d+)?[：:、\s-]*', '', tail)
    tail = tail.strip(' /')
    if attrs:
        title = attrs.get("title", "").strip()
        desc = attrs.get("description", "").strip()
        typ = attrs.get("type", "").strip()
        if _is_placeholder_drawing_text(title) and _is_placeholder_drawing_text(desc):
            return None
        if _is_placeholder_drawing_text(title) and desc and not _is_placeholder_drawing_text(desc):
            title = desc
        if not title:
            # 只有 description 时，优先把 description 作为标题，而不是整段原始文本
            if desc:
                title = desc
            else:
                # 兜底：从原始文本提取一个较短标题
                parts = [p.strip() for p in re.split(r'[，,；;。]', tail) if p.strip()]
                title = parts[0] if parts else tail[:24]
        if not desc:
            parts = [p.strip() for p in re.split(r'[，,；;。]', tail) if p.strip()]
            desc = '；'.join(parts[1:]) if len(parts) > 1 else (parts[0] if parts else tail)
        if not typ:
            typ = _infer_drawing_type(f"{title} {desc}")
        return {
            "id": attrs.get("id", str(seq or 0)),
            "type": typ,
            "title": title,
            "description": desc,
            "seq": seq or 0,
            "start_pos": start_pos,
            "end_pos": end_pos,
        }

    if not tail:
        return None
    parts = [p.strip() for p in re.split(r'[，,；;。]', tail) if p.strip()]
    title = parts[0] if parts else tail[:24]
    desc = '；'.join(parts[1:]) if len(parts) > 1 else (parts[0] if parts else tail)
    if _is_placeholder_drawing_text(title) and _is_placeholder_drawing_text(desc):
        return None
    if len(title) < 3 and desc:
        title = desc[:18]
    typ = _infer_drawing_type(f"{title} {desc}")
    return {
        "id": str(seq or 0),
        "type": typ,
        "title": title,
        "description": desc,
        "seq": seq or 0,
        "start_pos": start_pos,
        "end_pos": end_pos,
    }


def _normalize_drawing_signature(drawing: dict) -> str:
    """为 drawing 生成内容签名，用于识别重复图纸标签。"""
    if not isinstance(drawing, dict):
        return ""
    pieces = []
    for key in ("type", "title", "description", "scene", "layout"):
        val = drawing.get(key)
        text = re.sub(r"\s+", "", str(val or ""))
        if text:
            pieces.append(text)
    return "|".join(pieces)


def tolerant_extract_drawings(text):
    """宽容解析<drawing/>标签 — 不修改原文本"""
    return extract_drawings_from_text(text)


def extract_charts(text: str) -> list:
    """解析 <chart/> 标签，返回列表 [{"id":..., "title":..., "type":..., "x":..., "y":..., "unit":..., "data_source":..., "start_pos":..., "end_pos":...}]"""
    charts = []
    for m in re.finditer(r'<chart\b.*?/>', text, flags=re.DOTALL | re.I):
        attrs = dict(re.findall(r'(\w+)\s*=\s*["\']([^"\']*)["\']', m.group()))
        if not attrs:
            continue
        if not all(attrs.get(k) for k in ("id", "title", "type", "x", "y")):
            continue
        charts.append({
            "id": attrs.get("id", ""),
            "title": attrs.get("title", ""),
            "type": attrs.get("type", ""),
            "x": attrs.get("x", ""),
            "y": attrs.get("y", ""),
            "legend": attrs.get("legend", ""),
            "unit": attrs.get("unit", ""),
            "data_source": attrs.get("data_source", ""),
            "start_pos": m.start(), "end_pos": m.end()
        })
    return charts


def extract_tables(text: str) -> list:
    """解析 <table/> 标签，返回列表（容忍LLM未闭合标签）"""
    tables = []
    for m in re.finditer(r'<table\b.*?/>', text, flags=re.DOTALL | re.I):
        attrs = dict(re.findall(r'(\w+)\s*=\s*["\']([^"\']*)["\']', m.group()))
        # 兼容 LLM 输出 header/headers 两种写法
        if 'header' not in attrs and 'headers' in attrs:
            attrs['header'] = attrs['headers']
        if not all(attrs.get(k) for k in ("id", "title", "header", "rows", "data")):
            continue
        tables.append({
            "id": attrs.get("id", ""),
            "title": attrs.get("title", ""),
            "header": attrs.get("header", ""),
            "rows": attrs.get("rows", ""),
            "data": attrs.get("data", ""),
            "data_source": attrs.get("data_source", "") or attrs.get("datasource", ""),
            "start_pos": m.start(), "end_pos": m.end()
        })
    return tables


def extract_drawings_from_text(text: str) -> list:
    """从文本中提取所有 <drawing/> 标签"""
    drawings = []
    seen_signatures = set()
    seq = 0

    # 优先处理完整的跨行 <drawing ... /> 标签
    matches = list(re.finditer(r'<drawing\b.*?/>', text, flags=re.DOTALL))
    if matches:
        for m in matches:
            raw = m.group()
            parsed = _canonicalize_drawing_line(raw, seq + 1, m.start(), m.end())
            if parsed:
                sig = _normalize_drawing_signature(parsed)
                if sig and sig in seen_signatures:
                    continue
                if sig:
                    seen_signatures.add(sig)
                seq += 1
                drawings.append(parsed)
        if drawings:
            return drawings

    pos = 0
    for line in text.splitlines(True):
        line_end = pos + len(line)
        if "<drawing" in line.lower():
            parsed = _canonicalize_drawing_line(line, seq + 1, pos, line_end)
            if parsed:
                sig = _normalize_drawing_signature(parsed)
                if sig and sig in seen_signatures:
                    pos = line_end
                    continue
                if sig:
                    seen_signatures.add(sig)
                seq += 1
                drawings.append(parsed)
        pos = line_end
    return drawings


_MECH_JSON_BLOCK_RE = re.compile(
    r'\[\[\s*MECH_JSON\s*\]\](.*?)\[\[\s*/\s*MECH_JSON\s*\]\]',
    re.DOTALL | re.IGNORECASE
)


def _parse_mech_json_payload(raw: str):
    """尽量把 MECH_JSON 文本解析成 Python 对象。"""
    if not raw:
        return None
    text = raw.strip()
    candidates = [text]
    if not (text.startswith("{") or text.startswith("[")):
        m = re.search(r'(\{.*\}|\[.*\])', text, re.DOTALL)
        if m:
            candidates.insert(0, m.group(1).strip())
    for candidate in candidates:
        try:
            return json.loads(candidate)
        except Exception:
            continue
    return None


def extract_mech_json_blocks(text: str) -> list:
    """提取 [[MECH_JSON]]...[[/MECH_JSON]] 结构化块。"""
    blocks = []
    if not text:
        return blocks
    for idx, m in enumerate(_MECH_JSON_BLOCK_RE.finditer(text), 1):
        payload = _parse_mech_json_payload(m.group(1))
        if payload is None:
            continue
        blocks.append({
            "index": idx,
            "start_pos": m.start(),
            "end_pos": m.end(),
            "raw": m.group(0),
            "payload": payload,
        })
    return blocks


def strip_mech_json_blocks(text: str) -> str:
    """移除所有 [[MECH_JSON]]...[[/MECH_JSON]] 块，避免泄漏到 DOCX。"""
    if not text:
        return ""
    stripped = _MECH_JSON_BLOCK_RE.sub("\n", text)
    stripped = re.sub(r'\n{3,}', '\n\n', stripped)
    return stripped


def _normalize_table_data(data_str: str) -> str:
    """Normalize table data: if data has commas but no pipes, convert commas to pipes."""
    if data_str and "|" not in data_str and "," in data_str:
        return data_str.replace(",", "|")
    return data_str


def _split_table_tokens(text: str) -> list[str]:
    """把设计类表格里的混合分隔符压平成 token 列表。"""
    if not text:
        return []
    tokens = [t.strip() for t in re.split(r"[;,，；|｜\n]+", text) if t.strip()]
    return tokens


def canonicalize_design_table_payload(header_str: str, rows_str: str, data_str: str) -> dict:
    """
    将设计类表格的 header/rows/data 规整成可渲染矩阵。
    返回:
      {
        "headers": [...],
        "rows": [[row_name, v1, v2, ...], ...],
      }
    """
    headers = [h.strip() for h in re.split(r"[;,，；|｜]+", header_str or "") if h.strip()]
    row_tokens = [r.strip() for r in re.split(r"[;,，；|｜]+", rows_str or "") if r.strip()]
    data_tokens = _split_table_tokens(data_str or "")

    if not headers:
        return {"headers": [], "rows": []}

    data_cols = max(len(headers) - 1, 1)
    matrix: list[list[str]] = []

    # 兼容 LLM 把整行数据直接塞进 rows 的情况：
    # - 必须至少形成两行完整矩阵，避免把“纯行名列表”误判为矩阵
    # - 优先只在 data 为空时启用，避免和标准 rows/data 结构冲突
    if (
        row_tokens
        and not data_tokens
        and len(row_tokens) >= len(headers) * 2
        and len(row_tokens) % len(headers) == 0
    ):
        for i in range(0, len(row_tokens), len(headers)):
            chunk = row_tokens[i:i + len(headers)]
            if len(chunk) < len(headers):
                chunk = chunk + [""] * (len(headers) - len(chunk))
            matrix.append(chunk[:len(headers)])
        return {"headers": headers, "rows": matrix}

    if row_tokens:
        idx = 0
        for row_name in row_tokens:
            if idx < len(data_tokens) and data_tokens[idx] == row_name:
                idx += 1
            values = data_tokens[idx:idx + data_cols]
            idx += len(values)
            row = [row_name] + values
            while len(row) < len(headers):
                row.append("")
            matrix.append(row[:len(headers)])

        # 如果还有剩余 token，优先补到最后一行，避免数据被截断
        if idx < len(data_tokens) and matrix:
            tail = data_tokens[idx:]
            last = matrix[-1]
            tail_pos = 1
            for val in tail:
                if tail_pos >= len(last):
                    break
                if not last[tail_pos]:
                    last[tail_pos] = val
                tail_pos += 1
    else:
        # 无 rows 时：含 | 分列符的数据按行结构解析（保留空首列——分组表组列可为空，
        # 詹娜表5/谢朋金表7案: "||经理"空首列被压平丢弃→定长分组错位一列）
        group_size = len(headers)
        if group_size <= 0:
            return {"headers": headers, "rows": []}
        raw_text = data_str or ""
        if re.search(r"[|｜]", raw_text):
            matrix = []
            for ln in re.split(r"[;；\n]+", raw_text):
                if not ln.strip():
                    continue
                cells = [c.strip() for c in re.split(r"[|｜]+", ln)]
                while len(cells) < group_size:
                    cells.append("")
                matrix.append(cells[:group_size])
            return {"headers": headers, "rows": matrix}
        for i in range(0, len(data_tokens), group_size):
            row = data_tokens[i:i + group_size]
            while len(row) < group_size:
                row.append("")
            matrix.append(row[:group_size])

    return {"headers": headers, "rows": matrix}


def _flatten_mech_json_entries(payload):
    """把 MECH_JSON payload 规范成若干字典条目。"""
    if payload is None:
        return []
    if isinstance(payload, list):
        items = payload
    elif isinstance(payload, dict):
        for key in ("drawings", "figures", "items", "blocks", "entries"):
            val = payload.get(key)
            if isinstance(val, list):
                items = val
                break
        else:
            items = [payload]
    else:
        return []

    normalized = []
    for item in items:
        if isinstance(item, dict):
            normalized.append(dict(item))
        elif isinstance(item, str):
            parsed = _parse_mech_json_payload(item)
            if isinstance(parsed, dict):
                normalized.append(parsed)
    return normalized


def _mech_json_entry_target(entry: dict):
    """从 MECH_JSON 条目里提取可匹配的 drawing 标识。"""
    for key in ("id", "seq", "drawing_id", "drawing_seq", "drawing", "target", "ref_id"):
        val = entry.get(key)
        if val not in (None, "", []):
            return str(val)
    return ""




def _extract_machine_spec_from_blocks(mech_blocks: list) -> dict:
    """从MECH_JSON块中提取全局machine_spec（取第一个非空的）。"""
    for block in mech_blocks:
        payload = block.get("payload") if isinstance(block, dict) else None
        if not isinstance(payload, dict):
            continue
        spec = payload.get("machine_spec")
        if isinstance(spec, dict) and spec:
            return dict(spec)
    return {}


def _stable_unique_texts(items, limit: int = 8) -> list[str]:
    """保序去重，空值过滤。"""
    seen = set()
    result = []
    for item in items or []:
        text = str(item).strip()
        if not text or text in seen:
            continue
        seen.add(text)
        result.append(text)
        if len(result) >= limit:
            break
    return result

def merge_mech_json_into_drawings(drawings: list, mech_blocks: list) -> list:
    """
    将 MECH_JSON 里的结构化字段合并进 drawing 列表。
    - 优先按 id/seq 直接匹配
    - 其次按顺序补位
    - 未匹配到 drawing 时，保留为新增图纸条目
    """
    if not drawings and not mech_blocks:
        return drawings
    if not mech_blocks:
        return drawings

    enriched = [dict(d) for d in drawings]
    by_id = {}
    by_seq = {}
    for idx, d in enumerate(enriched):
        if d.get("id") not in (None, ""):
            by_id[str(d.get("id"))] = idx
        if d.get("seq") not in (None, ""):
            by_seq[str(d.get("seq"))] = idx

    unmatched_indices = [i for i in range(len(enriched))]
    synthetic = []
    global_machine_spec = _extract_machine_spec_from_blocks(mech_blocks)
    global_annotations = []
    if isinstance(global_machine_spec, dict):
        for key in ("load", "precision", "stroke", "material", "working_condition"):
            item = global_machine_spec.get(key)
            if isinstance(item, dict):
                raw = item.get("raw") or item.get("value")
                if raw:
                    global_annotations.append(f"{key}：{raw}")
            elif item:
                global_annotations.append(f"{key}：{item}")
        critical = global_machine_spec.get("critical_params", [])
        if isinstance(critical, list):
            for item in critical:
                if isinstance(item, dict):
                    name = item.get("name") or item.get("key")
                    raw = item.get("raw") or item.get("value")
                    if name and raw and f"{name}：{raw}" not in global_annotations:
                        global_annotations.append(f"{name}：{raw}")

    for block in mech_blocks:
        for entry in _flatten_mech_json_entries(block.get("payload")):
            # 保留所有结构化字段，但避免覆盖基础三要素的兜底值
            entry = dict(entry)
            if not entry.get("template"):
                entry["template"] = _pick_mech_template(entry)
            # 把新机械 JSON 的参数类字段提前扁平化为可视化注释，避免信息被藏在 payload 里
            annotation_lines = []
            for key in (
                "parameters", "technical_params", "critical_params",
                "dimensions", "dimension_text", "forces", "force_text", "annotations",
                "annotation", "notes", "params", "label_text",
            ):
                annotation_lines.extend(_collect_annotation_lines(entry.get(key), key))
            existing = entry.get("numeric_annotations")
            merged_annotations = list(global_annotations)
            if isinstance(existing, list):
                merged_annotations.extend(str(x) for x in existing)
            merged_annotations.extend(annotation_lines)
            if merged_annotations:
                entry["numeric_annotations"] = _stable_unique_texts(merged_annotations, 8)
            target = _mech_json_entry_target(entry)
            idx = None
            if target in by_id:
                idx = by_id[target]
            elif target in by_seq:
                idx = by_seq[target]
            elif unmatched_indices:
                idx = unmatched_indices.pop(0)

            if idx is None:
                synthetic.append(entry)
                continue

            base = enriched[idx]
            for k, v in entry.items():
                if v in (None, ""):
                    continue
                if k in ("type", "title", "description"):
                    # 原始值缺失或明显是占位符时，采用结构化补充，避免脏标签落到 DOCX
                    base[k] = _prefer_structured_drawing_value(base.get(k), v)
                else:
                    base[k] = v
            if base.get("seq") not in (None, ""):
                by_seq[str(base.get("seq"))] = idx
            if base.get("id") not in (None, ""):
                by_id[str(base.get("id"))] = idx

    # 兜底：把无法匹配的结构化条目作为新图纸追加，保持机械图链路可见
    for entry in synthetic:
        if not entry.get("id"):
            entry["id"] = str(len(enriched) + len(synthetic))
        if not entry.get("seq"):
            entry["seq"] = len(enriched) + len(synthetic)
        if not entry.get("type"):
            entry["type"] = "结构图"
        if not entry.get("title"):
            entry["title"] = "机械示意图"
        if not entry.get("description"):
            entry["description"] = ""
        existing = entry.get("numeric_annotations", [])
        combined = [x for x in (global_annotations + existing) if x]
        entry["numeric_annotations"] = _stable_unique_texts(combined, 8)
        enriched.append(entry)

    return enriched


def extract_drawings_with_positions(text: str) -> list:
    """提取 <drawing/> 标签并保留位置信息"""
    return extract_drawings_from_text(text)


def extract_docx_text(path: str) -> str:
    """读取 DOCX 所有段落文本，换行分隔"""
    try:
        doc = Document(path)
        return "\n".join(p.text.strip() for p in doc.paragraphs if p.text and p.text.strip())
    except Exception as e:
        print(f"读取DOCX失败 {path}: {e}")
        return ""


# ==================== 图表渲染 ====================
def _render_table_format(ct, rows_data, col_labels):
    """渲染table_format类型为表格图"""
    title_text = ct.get("title", "")
    source = ct.get("data_source", "")
    import io
    n_rows = len(rows_data)
    n_cols = len(col_labels)
    fig, ax = plt.subplots(figsize=(max(6, n_cols * 2.5), max(3, n_rows * 0.8)))
    ax.axis('off')
    cell_text = []
    for row in rows_data:
        cells = [c.strip() for c in row.split(",")]
        # 补齐到列数
        while len(cells) < n_cols:
            cells.append("")
        cell_text.append(cells[:n_cols])
    table = ax.table(cellText=cell_text, colLabels=col_labels,
                    loc='center', cellLoc='center')
    table.auto_set_font_size(False)
    table.set_fontsize(9)
    table.scale(1, 1.5)
    if title_text:
        ax.set_title(title_text, fontsize=12, fontweight='bold', pad=15)
    if source:
        ax.text(0.5, -0.05, '数据来源：' + source, ha='center', va='top',
                transform=ax.transAxes, fontsize=8, color='gray')
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight')
    buf.seek(0)
    plt.close(fig)
    return buf.read()


def chart_to_bytes(ct: dict, colors=None) -> bytes:
    """依据 chart 标签信息生成 matplotlib 图表，返回 PNG bytes"""
    if colors is None:
        colors = COLORS
    titles = [t.strip() for t in ct.get("x", "").split(",") if t.strip()]
    chart_type = normalize_chart_type(ct.get("type", "bar") or "bar")
    legend_text = ct.get("legend", "") or ""
    legend_labels = [t.strip() for t in legend_text.split(",") if t.strip()]

    def _parse_series(series_text: str):
        vals = []
        for v_str in series_text.split(","):
            v_str = v_str.strip()
            if not v_str:
                continue
            num_match = re.search(r'-?\d+(?:\.\d+)?', v_str)
            if not num_match:
                continue  # 跳过非数字值，不放弃整图
            vals.append(float(num_match.group()))
        return vals if vals else None

    raw_y = ct.get("y", "")
    series_texts = [s.strip() for s in raw_y.split(";") if s.strip()]
    if not series_texts:
        return b""
    
    # table_format类型：直接渲染为表格图，跳过数值解析
    if chart_type == "table_format" and series_texts and titles:
        return _render_table_format(ct, series_texts, titles)
    
    multi_vals = []
    all_text_labels = True  # 是否所有y值都是纯文字（structure类型专用）
    for series_text in series_texts:
        parsed = _parse_series(series_text)
        if parsed is None:
            # 判断y值是否可转为纯文字展示（structure/scatter类型）
            if chart_type in ("structure", "scatter", "flow", "tree", "table_format"):
                all_text_labels = True
                multi_vals.append([])  # 空数值占位
            else:
                print(f'  ⚠️ 图表"{ct.get("title", "")}"的y值包含非数字或空值，无法渲染。')
                print(f'    y值: {ct.get("y", "")}')
                print(f'    【建议】在prompt中强化：每个系列必须输出完整可绘图的纯数字序列')
                return b""
        else:
            all_text_labels = False
            multi_vals.append(parsed)

    # 兼容 category-major：当 y 按分类分组、legend 按对比组命名时，自动转置为 series-major
    if legend_labels and len(multi_vals) == len(titles) and all(len(v) == len(legend_labels) for v in multi_vals):
        transposed = []
        for j in range(len(legend_labels)):
            transposed.append([multi_vals[i][j] for i in range(len(multi_vals))])
        multi_vals = transposed

    is_multi = len(multi_vals) > 1
    if not legend_labels and is_multi:
        legend_labels = [f"系列{i+1}" for i in range(len(multi_vals))]
    if legend_labels and len(legend_labels) != len(multi_vals):
        print(f'  ⚠️ 图表"{ct.get("title", "")}"的legend数量与系列数不一致，无法渲染。')
        print(f'    legend: {legend_text}')
        print(f'    y值: {ct.get("y", "")}')
        return b""
    if titles and len(titles) != len(multi_vals[0]):
        # 仅在单系列/多系列的每个系列都同长度时成立
        if any(len(v) != len(multi_vals[0]) for v in multi_vals):
            print(f'  ⚠️ 图表"{ct.get("title", "")}"的各系列数据点数不一致，无法渲染。')
            print(f'    x值: {ct.get("x", "")}')
            print(f'    y值: {ct.get("y", "")}')
            return b""
    if not titles:
        print(f'  ⚠️ 图表"{ct.get("title", "")}"缺少x轴分类，无法渲染。')
        return b""

    fig, ax = plt.subplots(figsize=(7, 4))
    title_text = ct.get("title", "")
    unit = ct.get("unit", "")
    source = ct.get("data_source", "")
    x_pos = np.arange(len(titles))
    # 计算最大值前过滤空列表
    non_empty = [series for series in multi_vals if series]
    max_val = max(max(series) for series in non_empty) if non_empty else 1.0

    def _bar_colors(idx):
        if len(colors) >= idx + 1:
            return colors[idx]
        return colors[idx % len(colors)]

    if chart_type in ("bar", "comparison"):
        if is_multi:
            n = len(multi_vals)
            width = 0.8 / n
            shift = (n - 1) / 2
            for i, vals in enumerate(multi_vals):
                offset = (i - shift) * width
                bars = ax.bar(
                    x_pos + offset, vals, width=width,
                    color=_bar_colors(i),
                    label=legend_labels[i]
                )
                for bar, v in zip(bars, vals):
                    ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + max_val * 0.01,
                            f"{v:.1f}", ha='center', va='bottom', fontsize=9)
        else:
            vals = multi_vals[0]
            bars = ax.bar(x_pos, vals, color=_bar_colors(0), width=0.6, label=legend_labels[0] if legend_labels else None)
            for bar, v in zip(bars, vals):
                ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + max_val * 0.01,
                        f"{v:.1f}", ha='center', va='bottom', fontsize=9)
    elif chart_type in ("line", "trend"):
        for i, vals in enumerate(multi_vals):
            marker = ['o', 's', '^', 'D', 'v', 'x'][i % 6]
            ax.plot(x_pos, vals, marker=marker, color=_bar_colors(i), linewidth=2, markersize=6,
                    label=legend_labels[i] if legend_labels else None)
            for j, v in enumerate(vals):
                ax.text(j, v + max_val * 0.02, f"{v:.1f}", ha='center', fontsize=9)
        ax.set_xticks(x_pos)
    elif chart_type == "stacked":
        bottom = np.zeros(len(titles))
        for i, vals in enumerate(multi_vals):
            ax.bar(x_pos, vals, bottom=bottom, color=_bar_colors(i), width=0.6,
                   label=legend_labels[i] if legend_labels else None)
            bottom = bottom + np.array(vals)
    elif chart_type == "pie":
        if is_multi:
            print(f'  ⚠️ 图表"{title_text}"为饼图时不支持多系列，无法渲染。')
            return b""
        vals = multi_vals[0]
        wedges, texts, autotexts = ax.pie(vals, labels=titles, autopct='%1.1f%%',
                                          colors=colors[:len(vals)] if len(colors) >= len(vals) else colors,
                                          startangle=90)
        for t in autotexts:
            t.set_fontsize(9)
    elif chart_type == "gantt":
        if not is_multi:
            print(f'  ⚠️ 甘特图"{title_text}"需要每个任务一组起止数据，当前数据不足。')
            return b""
        if len(multi_vals) != len(titles):
            print(f'  ⚠️ 甘特图"{title_text}"的任务数与x轴分类数不一致，无法渲染。')
            print(f'    x值: {ct.get("x", "")}')
            print(f'    y值: {ct.get("y", "")}')
            return b""
        y_pos = np.arange(len(titles))
        for i, vals in enumerate(multi_vals):
            if len(vals) != 2:
                print(f'  ⚠️ 甘特图"{title_text}"第{i+1}组数据不是start,duration两项，无法渲染。')
                return b""
            start, duration = vals
            ax.barh(y_pos[i], duration, left=start, height=0.5, color=_bar_colors(i))
        ax.set_yticks(y_pos)
        ax.set_yticklabels(titles, fontsize=9)
        ax.invert_yaxis()
    elif chart_type == "structure":
        raw_y = ct.get("y", "")
        y_labels_raw = [s.strip() for s in raw_y.split(";") if s.strip()]
        if all_text_labels and y_labels_raw:
            # 纯文字结构图：用文字标签直接显示
            n_groups = len(y_labels_raw)
            if len(titles) != n_groups:
                # 但可能x是一级分类，y是子项列表
                for i, (title, labels) in enumerate(zip(titles, y_labels_raw)):
                    ax.text(i, 0.5 - i*0.1, f"{title}\n{' / '.join(labels.split(','))}",
                           ha='center', va='center', fontsize=10,
                           bbox=dict(boxstyle='round,pad=0.5', facecolor='lightblue', alpha=0.7))
                ax.set_ylim(-0.5, len(titles))
                ax.set_yticks([])
                ax.set_xticks(range(len(titles)))
                ax.set_xticklabels(titles, fontsize=10)
            else:
                # 对称结构
                bar_width = 0.6
                x_pos = np.arange(len(titles))
                for i in range(len(titles)):
                    labels = y_labels_raw[i].split(",") if len(y_labels_raw) > i else []
                    label_text = "\n".join(labels) if labels else y_labels_raw[i]
                    ax.bar(x_pos[i], [1], width=bar_width, color=colors[i % len(colors)], alpha=0.7)
                    ax.text(x_pos[i], 0.5, label_text, ha='center', va='center', fontsize=9,
                           color='black', fontweight='bold')
                ax.set_ylim(0, 1.5)
                ax.set_yticks([])
        elif is_multi:
            bottom = np.zeros(len(titles))
            for i, vals in enumerate(multi_vals):
                ax.bar(x_pos, vals, bottom=bottom, color=_bar_colors(i), width=0.6,
                       label=legend_labels[i] if legend_labels else None)
                bottom = bottom + np.array(vals)
        else:
            vals = multi_vals[0]
            bars = ax.bar(x_pos, vals, color=_bar_colors(0), width=0.6, label=legend_labels[0] if legend_labels else None)
            for bar, v in zip(bars, vals):
                ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + max_val * 0.01,
                        f"{v:.1f}", ha='center', va='bottom', fontsize=9)
    elif chart_type == "table_format":
        # table_format：y值是纯文字，渲染为表格样式的柱状图
        raw_y = ct.get("y", "")
        rows_data = [s.strip() for s in raw_y.split(";") if s.strip()]
        if rows_data:
            n_rows = len(rows_data)
            n_cols = len(titles)
            fig, ax = plt.subplots(figsize=(max(6, n_cols * 2.5), max(3, n_rows * 1.2)))
            ax.axis('off')
            # 绘制表格
            col_labels = titles
            cell_text = []
            for row in rows_data:
                cells = [c.strip() for c in row.split(",")]
                cell_text.append(cells)
            table = ax.table(cellText=cell_text, colLabels=col_labels,
                            loc='center', cellLoc='center')
            table.auto_set_font_size(False)
            table.set_fontsize(9)
            table.scale(1, 1.5)
            if title_text:
                ax.set_title(title_text, fontsize=12, fontweight='bold', pad=15)
            # 输出到bytes
            buf = io.BytesIO()
            plt.savefig(buf, format='png', dpi=150, bbox_inches='tight')
            buf.seek(0)
            plt.close(fig)
            return buf.read()
        else:
            print(f'  ⚠️ 图表\"{title_text}\"的table_format数据为空，无法渲染。')
            return b""
    else:
        if any(len(v) != len(titles) for v in multi_vals):
            print(f'  ⚠️ 图表"{title_text}"的x轴与y轴长度不一致，无法渲染。')
            print(f'    x值: {ct.get("x", "")}')
            print(f'    y值: {ct.get("y", "")}')
            return b""
        # 未知类型兜底：按柱状图绘制，绝不输出"有标题没图形"的空图
        print(f'  ⚠️ 图表"{title_text}"类型"{ct.get("type", "")}"未识别，按柱状图绘制')
        for i, vals in enumerate(multi_vals):
            bars = ax.bar([p + i * 0.8 / max(len(multi_vals), 1) for p in x_pos], vals,
                          color=_bar_colors(i), width=0.6 / max(len(multi_vals), 1),
                          label=legend_labels[i] if legend_labels else None)
            for bar, v in zip(bars, vals):
                ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + max_val * 0.01,
                        f"{v:.1f}", ha='center', va='bottom', fontsize=9)
    ax.set_title(title_text, fontsize=14, fontweight='bold', pad=15)
    if chart_type != "pie":
        handles, labels = ax.get_legend_handles_labels()
        if any(lbl and not str(lbl).startswith("_") for lbl in labels):
            ax.legend(fontsize=9)
    if chart_type != "pie" and chart_type != "gantt":
        ax.set_xticks(x_pos)
        ax.set_xticklabels(titles, fontsize=9, rotation=15, ha='right')
        if unit:
            ax.set_ylabel(unit, fontsize=10)
    elif chart_type == "gantt" and unit:
        ax.set_xlabel(unit, fontsize=10)
    if source:
        ax.text(0.5, -0.12, f"数据来源：{source}", transform=ax.transAxes,
                ha='center', fontsize=8, style='italic', color='gray')
    plt.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=200, bbox_inches='tight')
    plt.close(fig)
    return buf.getvalue()


def _make_placeholder_image(title: str, note: str = "原图生成失败·占位待补",
                            width: int = 1200, height: int = 800) -> str:
    """生成占位图（PIL本地渲染，无网络依赖），返回临时文件路径。"""
    import tempfile
    img = Image.new("RGB", (width, height), color=(245, 245, 245))
    draw = ImageDraw.Draw(img)
    draw.rectangle([8, 8, width - 8, height - 8], outline=(180, 180, 180), width=4)
    try:
        font_big = _load_font(34, bold=True)
        font_small = _load_font(26)
    except Exception:
        font_big = font_small = None
    text = (title or "").strip() or "（未命名）"
    def _center(y, s, font, fill):
        if font:
            tw = draw.textlength(s, font=font)
            draw.text(((width - tw) / 2, y), s, font=font, fill=fill)
        else:
            draw.text((width / 3, y), s, fill=fill)
    _center(height * 0.40, text, font_big, (90, 90, 90))
    _center(height * 0.52, note, font_small, (150, 150, 150))
    fd, path = tempfile.mkstemp(suffix=".png", prefix="placeholder_")
    os.close(fd)
    img.save(path)
    return path


def _lookup_drawing_image(drawing: dict, drawing_images: dict) -> Optional[str]:
    """统一图片key查找：兼容 seq/id/int化 等多种约定，避免图被静默丢弃。"""
    if not drawing_images:
        return None
    raw_keys = []
    for k in (drawing.get("seq"), drawing.get("id")):
        if k is not None and str(k).strip():
            raw_keys.append(str(k))
    # 归一化映射：数字key转int字符串
    norm_map = {}
    for k, v in drawing_images.items():
        ks = str(k)
        norm_map[ks] = v
        try:
            norm_map[str(int(float(ks)))] = v
        except (ValueError, TypeError):
            pass
    for key in raw_keys:
        if key in norm_map and os.path.exists(str(norm_map[key])):
            return str(norm_map[key])
        try:
            if str(int(float(key))) in norm_map:
                p = str(norm_map[str(int(float(key)))])
                if os.path.exists(p):
                    return p
        except (ValueError, TypeError):
            pass
    return None


def add_c_placeholder(doc, ct: dict, cn: int, reason: str = ""):
    """图表渲染失败时的占位输出（保持图注结构完整，问题可见）"""
    title = ct.get("title", f"图{cn}")
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p_title.add_run(title)
    set_run_font(r, "宋体", 10, bold=True)
    note = "图表数据异常（渲染失败占位）" + (f"：{reason[:60]}" if reason else "")
    path = _make_placeholder_image(title, note, width=1000, height=500)
    try:
        doc.add_picture(path, width=Cm(14))
    except Exception as e:
        print(f"添加图表占位失败: {e}")
    finally:
        try:
            os.remove(path)
        except OSError:
            pass
    last_p = doc.add_paragraph()
    last_p.paragraph_format.space_after = Pt(6)


def add_c(doc, ct: dict, chart_bytes: bytes, cn: int):
    """向 docx 中添加图表（chart）"""
    title = ct.get("title", f"图{cn}")
    source = ct.get("data_source", "")
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p_title.add_run(title)  # 直接用LLM给的title原名，不再加"图{cn}"前缀
    set_run_font(r, "宋体", 10, bold=True)
    if chart_bytes:
        try:
            doc.add_picture(io.BytesIO(chart_bytes), width=Cm(14))
        except Exception as e:
            print(f"添加图片失败: {e}")
    if source:
        p_src = doc.add_paragraph()
        p_src.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p_src.add_run(f"数据来源：{source}")
        set_run_font(r, "宋体", 9)
    last_p = doc.add_paragraph()
    last_p.paragraph_format.space_after = Pt(6)


def _table_quality(ct: dict):
    """表格内容质量评估 → (行数, 空格率%)。用于回执判定"渲染成功但内容残缺"。"""
    try:
        shape = canonicalize_design_table_payload(
            ct.get("header", ""), ct.get("rows", ""),
            _normalize_table_data(ct.get("data", "")))
        rows = shape["rows"]
        if not rows:
            return 0, 100.0
        ncell = sum(len(r) for r in rows) or 1
        nempty = sum(1 for r in rows for c in r
                     if not str(c).strip() or str(c).strip() == "—")
        return len(rows), 100.0 * nempty / ncell
    except Exception:
        return 0, 100.0


def add_t(doc, ct: dict, tn: int):
    """向 docx 中添加表格（table）"""
    title = ct.get("title", f"表{tn}")
    header_str = ct.get("header", "")
    rows_str = ct.get("rows", "")
    data_str = _normalize_table_data(ct.get("data", ""))
    source = ct.get("data_source", "")
    table_shape = canonicalize_design_table_payload(header_str, rows_str, data_str)
    headers = table_shape["headers"]
    data_rows = table_shape["rows"]
    if not headers:
        raise ValueError(f"表{tn}缺少表头，禁止静默跳过")
    num_cols = max(len(headers), 1)
    table_data = data_rows
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    caption = title.strip()
    if not re.match(r'^表\d+(?:-\d+)?(?:\s|$)', caption):
        caption = f"表{tn} {caption}"
    r = p.add_run(caption)
    set_run_font(r, "宋体", 10, bold=True)
    t = doc.add_table(rows=1 + len(table_data), cols=num_cols)
    t.style = 'Table Grid'
    # 表头
    for j, h in enumerate(headers):
        cell = t.rows[0].cells[j]
        cell.text = h
        for para in cell.paragraphs:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in para.runs:
                set_run_font(run, "宋体", 9, bold=True)
    # 数据行
    for i, row_data in enumerate(table_data):
        for j, val in enumerate(row_data):
            cell = t.rows[i + 1].cells[j]
            cell.text = val
            for para in cell.paragraphs:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in para.runs:
                    set_run_font(run, "宋体", 9)
    # 0908加固: 分组表首列自动纵向合并 — 数据行第1列"组首行有值/后续为空"或"连续同值"(≥2行)时
    # vMerge并回填组值, 消除"通道类型/责任中心列大量空格"(詹娜表5/谢朋金表7案)
    try:
        if num_cols >= 2 and len(table_data) >= 2:
            col0 = [str(r[0]).strip() if r and len(r) > 0 else "" for r in table_data]
            # 空值继承上方组值
            filled, cur = [], ""
            for v in col0:
                if v:
                    cur = v
                filled.append(cur)
            from docx.oxml.ns import qn as _qn
            i = 0
            while i < len(filled):
                j = i
                while j + 1 < len(filled) and filled[j + 1] == filled[i]:
                    j += 1
                if filled[i]:
                    span = j - i + 1
                    top = t.rows[i + 1].cells[0]  # +1: 表头占docx第0行
                    if not col0[i]:
                        top.text = filled[i]  # 组首行为空时回填(极少见)
                    if span >= 2:
                        # 后续格保持空, 由vMerge承接(merge会拼接各格内容, 预填将致重复文本)
                        _mg = top.merge(t.rows[j + 1].cells[0])
                        _tcPr = _mg._tc.get_or_add_tcPr()
                        _va = _tcPr.find(_qn('w:vAlign'))
                        if _va is None:
                            _va = _tcPr.makeelement(_qn('w:vAlign'), {})
                            _tcPr.append(_va)
                        _va.set(_qn('w:val'), 'center')
                i = j + 1
    except Exception as _e:
        print(f"分组表合并(忽略): {_e}")
    if source:
        p_s = doc.add_paragraph()
        p_s.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p_s.add_run(f"数据来源：{source}")
        set_run_font(r, "宋体", 9)
    last_p = doc.add_paragraph()
    last_p.paragraph_format.space_after = Pt(6)


# ==================== 图片生成 ====================


def _build_engineering_image_prompt(drawing: dict) -> str:
    """为机械工程图构造受控的 GPT 图像生成 prompt，减少想象、锁定参数。"""
    title = drawing.get("title", "机械工程图")
    img_type = drawing.get("type", "结构图")
    description = drawing.get("description", "")

    # 收集所有参数
    params = []
    for key in ("technical_params", "critical_params", "parameters"):
        val = drawing.get(key)
        if isinstance(val, dict):
            for k, v in val.items():
                params.append(f"- {k}: {v}")
        elif isinstance(val, list):
            for item in val:
                if isinstance(item, dict):
                    name = item.get("name") or item.get("key")
                    raw = item.get("raw") or item.get("value")
                    if name and raw:
                        params.append(f"- {name}: {raw}")

    # 从 description 提取关键尺寸和规格
    desc = f"{title} {description}"
    patterns = [
        (r"(?:直径|φ|Φ|Ø)\s*(\d+(?:\.\d+)?(?:\s*mm)?)", "直径"),
        (r"(M\d+(?:x\d+(?:\.\d+)?)?(?:\s*\d+)?)", "螺栓"),
        (r"(\d+\s*mm\s*x\s*\d+\s*mm(?:\s*x\s*\d+\s*mm)?)", "尺寸"),
        (r"(HT\d+|45钢|40Cr|20Cr|T8A|Q235A|Cr12MoV)", "材料"),
        (r"(\d+(?:\.\d+)?\s*N(?:\s*|kN)?)", "力"),
    ]
    extracted = []
    for pat, label in patterns:
        for m in re.finditer(pat, desc, flags=re.IGNORECASE):
            item = m.group(0).strip()
            if item and item not in [e.split(":")[-1].strip() for e in extracted]:
                extracted.append(f"- {label}: {item}")
    extracted = extracted[:12]

    params_text = "\n".join(params[:8]) or "无具体参数"
    extracted_text = "\n".join(extracted) or "无提取参数"

    prompt = f"""You are a technical CAD illustrator. Create a SINGLE precise, professional engineering drawing (orthographic front view, CAD line art style) for the following mechanical part/fixture. Do NOT produce a 3D rendered image, perspective view, or artistic illustration.

Title: {title}
Type: {img_type}
Description: {description}

Known parameters from the design:
{params_text}

Extracted dimensions and specifications:
{extracted_text}

CRITICAL STYLE RULES:
- Orthographic 2D front view, NO perspective, NO 3D shading, NO artistic style
- Clean black line art on white background
- Draw a recognizable fixture/mold with: base plate, workpiece, locating pin, V-block, clamping plate, clamping bolt, milling cutter (for fixture) OR upper die plate, punch, lower die plate, die cavity, guide pillars (for mold)
- Section lines / hatching on cut metal surfaces
- Center lines (dash-dot), dimension lines with arrows and numbers
- Balloon part numbers (1, 2, 3...)
- Title block lower-right with title "{title}", scale "1:1", material "45 steel / HT200"
- All dimensions and part names MUST match the description above
- DO NOT invent parts or change dimensions
- NO shadows, gradients, colors, photorealistic effects
- Output 1024x1024 PNG, high contrast crisp lines
"""
    return prompt

def generate_single_image(drawing: dict, save_dir: str, max_retries: int = 3) -> Optional[str]:
    """调用AI图片API生成单张图纸图片，返回本地路径"""
    img_type = drawing.get("type", "设计图")
    title = drawing.get("title", "设计图")
    description = drawing.get("description", "")
    backend = classify_drawing_backend(drawing)
    # 所有类型（土木、机械diagram、设计effect）都走GPT图片API
    # 土木关键词→"diagram"分支，走工程制图风格
    if backend == "diagram":
        prompt_text = _build_design_diagram_prompt(drawing)
        print(f"  [GPT图示] 正在生成: {title}")
    else:
        prompt_text = _build_design_image_prompt(drawing)
        print(f"  [GPT效果图] 正在生成: {title}")

    img_path = _call_image_api(prompt_text, drawing, save_dir, max_retries)
    if img_path and _is_blank_image(img_path):
        try:
            os.remove(img_path)
        except Exception:
            pass
        retry_prompt = _strengthen_image_prompt(prompt_text, drawing)
        img_path = _call_image_api(retry_prompt, drawing, save_dir, max_retries)
        if img_path and _is_blank_image(img_path):
            try:
                os.remove(img_path)
            except Exception:
                pass
            img_path = None
    return img_path


def _build_design_image_prompt(drawing: dict) -> str:
    """为设计类效果图构造更强约束的图片 prompt。"""
    title = drawing.get("title", "设计图")
    img_type = drawing.get("type", "效果图")
    description = drawing.get("description", "")
    scene = drawing.get("scene", "")
    layout = drawing.get("layout", "")
    extra = " ".join([str(scene), str(layout)]).strip()
    prompt = f"""Create a complete, detailed architectural/interior design illustration.
Subject: {title}
Type: {img_type}
Description: {description}
Context: {extra}

Requirements:
- The image must not be blank or mostly white.
- Fill the canvas with a clear composition, with foreground, middle ground, and background.
- Show the main spatial elements, furniture, walls, floor, lighting, materials, and atmosphere.
- Use a realistic but clean design presentation style.
- If it is a plan/elevation/section/nodes/analysis image, draw the corresponding architectural line drawing instead of a photo-like scene.
- No empty corners, no tiny isolated strokes, no placeholder marks.
- 1024x1024.
"""
    return prompt


def _build_design_diagram_prompt(drawing: dict) -> str:
    """为设计/土木/机械类平面/立面/节点/分析图构造更强约束的 GPT prompt。
    
    设计原则：GPT足够强大，不需要结构化分类。直接把drawing的完整文本喂给它，
    加上统一的风格约束和内容忠实度约束即可。
    """
    title = drawing.get("title", "设计图")
    img_type = drawing.get("type", "图示")
    description = drawing.get("description", "")

    return f"""Create a professional civil/mechanical engineering drawing in CAD line-art style, with white background and black/dark gray linework.

Title: {title}
Type: {img_type}
Description: {description}

CRITICAL REQUIREMENTS:
1. CONTENT FIDELITY — Every detail in the Description above is the ground truth and must be reproduced EXACTLY:
   - All spatial/positional relationships (方位: 东南/西北/南侧/北侧/左侧/右侧/中间 etc.) must be exactly as written
   - All colors, materials, dimensions must match the Description exactly
   - All numerical values (楼高,层高,柱距,配筋 etc.) must be displayed as labels/annotations on the drawing
   - Do NOT change, omit, or invent any detail from the Description

2. PROFESSIONAL QUALITY — Fill at least 70% of the canvas with dense, meaningful content:
   - Use clean engineering linework with proper line weights (thick for outlines, thin for annotations)
   - Include dimension chains, grid lines, labels, callouts, annotations, legend, title block
   - No blank areas, no isolated tiny strokes, no placeholder marks

3. STYLE — Draw in standard engineering drafting convention:
   - Plans: show walls, columns, room labels, grid axes, dimension chains, north arrow, scale
   - Elevations/sections: show floor levels, openings, structural elements, height annotations
   - Details: show enlarged nodes, reinforcement bars, material layers, callout bubbles
   - Diagrams: show clear composition with labeled elements, arrows, and legends

4. Output 1024x1024 PNG with substantial content filling the frame.
"""


def _strengthen_image_prompt(prompt_text: str, drawing: dict) -> str:
    """在首次结果过白时，追加更强约束后重试。"""
    title = drawing.get("title", "设计图")
    return prompt_text + f"""

HARD FAIL CONDITIONS:
- If the image is mostly blank or mostly white, redraw it with substantially more content.
- The title is: {title}
- Put large enough objects or line elements in the center and corners so the image does not look empty.
"""
def _call_image_api(prompt_text: str, drawing: dict, save_dir: str, max_retries: int = 6) -> Optional[str]:
    if not IMAGE_API_KEY:
        raise RuntimeError("IMAGE_API_KEY is not set")
    headers = {
        "Authorization": f"Bearer {IMAGE_API_KEY}",
        "Content-Type": "application/json"
    }
    data = {
        "model": IMAGE_MODEL,
        "prompt": prompt_text,
        "n": 1,
        "size": "1024x1024"
    }
    for attempt in range(max_retries):
        try:
            resp = requests.post(IMAGE_API_URL, headers=headers, json=data, timeout=300)
            if resp.status_code == 200:
                result = resp.json()
                img_data = result.get("data", [{}])[0]
                b64_data = img_data.get("b64_json") or img_data.get("url")
                if b64_data:
                    if b64_data.startswith("http"):
                        img_resp = requests.get(b64_data, timeout=60)
                        img_bytes = img_resp.content
                    else:
                        img_bytes = base64.b64decode(b64_data)
                    key = drawing.get("seq") or drawing.get("id") or uuid.uuid4().hex[:8]
                    filename = f"drawing_{key}.png"
                    path = os.path.join(save_dir, filename)
                    with open(path, "wb") as f:
                        f.write(img_bytes)
                    if _is_blank_image(path):
                        try:
                            os.remove(path)
                        except Exception:
                            pass
                        print(f"  图片过于空白，重试(尝试{attempt+1})")
                        time.sleep(2)
                        continue
                    print(f"  图片生成成功: {path}")
                    return path
            else:
                body = ""
                try:
                    body = resp.text[:120]
                except Exception:
                    pass
                print(f"  图片API失败(尝试{attempt+1}): {resp.status_code} {body}")
                if resp.status_code == 403:
                    # 403=平台滚动限流/配额窗口: 立即重试必然连败,指数退避等窗口
                    wait = min(30 * (attempt + 1), 120)
                    print(f"    限流退避{wait}秒...")
                    time.sleep(wait)
                elif resp.status_code == 429:
                    time.sleep(15 * (attempt + 1))
                else:
                    time.sleep(5)
        except Exception as e:
            print(f"  图片生成异常(尝试{attempt+1}): {e}")
            time.sleep(2)
    return None

    prompt_text = f"{title}，{description}，{img_type}，高清设计效果图"
    return _call_image_api(prompt_text, drawing, save_dir, max_retries)


def classify_drawing_backend(drawing: dict) -> str:
    """按图纸类型分类：effect=真实效果图走GPT；diagram=平面/立面/节点/分析类走本地图示。"""
    text = " ".join([
        str(drawing.get("type", "")),
        str(drawing.get("title", "")),
        str(drawing.get("description", "")),
        str(drawing.get("template", "")),
        str(drawing.get("scene", "")),
        str(drawing.get("layout", "")),
    ]).lower()
    effect_keywords = ["效果图", "外观", "渲染", "场景图", "展示图", "概念图", "空间效果"]
    diagram_keywords = [
        "平面图", "总平面图", "立面图", "节点图", "剖面图", "分析图", "功能分区",
        "构造", "流线", "布置图", "关系图", "示意图", "详图", "节点详图", "轴测图"
    ]
    # 土木/建筑类图纸关键词 —— 优先匹配，不走机械CAD
    civil_keywords = [
        "标准层", "楼梯", "柱网", "框架", "结构平面", "梁配筋", "板配筋",
        "柱配筋", "基础平面", "基础详图", "建筑平面", "建筑立面", "建筑剖面",
        "正立面", "侧立面", "背立面", "施工平面", "施工进度", "施工部署",
        "资源分配", "劳动力", "施工横道", "横道图",
    ]
    mech_signals = [
        "mechanical", "机械", "机构", "零件", "部件", "装配", "受力", "载荷", "应力",
        "行程", "参数", "尺寸", "标注", "公差", "motion", "force", "assembly", "parameter", "structure"
    ]
    # 土木/建筑优先 —— 归入diagram分支（GPT工程制图生图）
    if any(k in text for k in civil_keywords):
        return "diagram"
    if any(k in text for k in diagram_keywords):
        return "diagram"
    if any(k in text for k in effect_keywords):
        return "effect"
    if any(k in text for k in mech_signals) or any(k in str(drawing.get(k, "")).lower() for k in ("structured", "meta", "payload", "parameters", "technical_params", "critical_params", "numeric_annotations")):
        return "diagram"
    if any(k in text for k in ["gantt", "甘特", "时序", "流程", "结构", "原理", "装配", "受力", "布局"]):
        return "diagram"
    return "effect"


def _load_font(size: int, bold: bool = False):
    candidates = [
        "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
        "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
    ]
    for path in candidates:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size=size)
            except Exception:
                continue
    return ImageFont.load_default()


def _mech_text_blob(drawing: dict) -> str:
    """把 drawing 及其结构化字段压成一个便于分类的文本块。"""
    parts = [
        str(drawing.get("type", "")),
        str(drawing.get("title", "")),
        str(drawing.get("description", "")),
        str(drawing.get("template", "")),
        str(drawing.get("kind", "")),
        str(drawing.get("scene", "")),
        str(drawing.get("layout", "")),
    ]
    for key in (
        "structured", "meta", "payload", "annotation", "annotations",
        "annotations_text", "steps", "step_text", "parts", "part_text",
        "dimensions", "dimension_text", "forces", "force_text", "labels",
        "label_text", "params", "parameters", "flow_steps", "motion_steps",
        "nodes", "links", "arrows", "notes", "numeric_annotations",
    ):
        val = drawing.get(key)
        if val not in (None, "", [], {}):
            parts.append(json.dumps(val, ensure_ascii=False) if isinstance(val, (dict, list)) else str(val))
    return " ".join(parts)


def _iter_text_tokens(value):
    if value is None:
        return
    if isinstance(value, dict):
        for k, v in value.items():
            yield str(k)
            yield from _iter_text_tokens(v)
    elif isinstance(value, (list, tuple, set)):
        for item in value:
            yield from _iter_text_tokens(item)
    else:
        text = str(value)
        for token in re.split(r"[，,；;。、\n\r\t]+", text):
            token = token.strip()
            if token:
                yield token


def _collect_annotation_lines(value, key: str = "") -> list:
    """
    递归提取适合直接画在图上的注释行。
    兼容新机械 JSON 中的 parameters / dimensions / forces / critical_params 等字段。
    """
    if value is None or value == "" or value == [] or value == {}:
        return []

    lines = []
    key_norm = str(key or "").strip()

    if isinstance(value, list):
        for item in value:
            lines.extend(_collect_annotation_lines(item, key_norm))
        return lines

    if isinstance(value, dict):
        if key_norm in {"payload", "structured", "meta", "data", "info", "machine_spec"}:
            lines = []
            for sub_key, sub_val in value.items():
                if sub_key in {"id", "seq", "type", "title", "description", "template", "kind", "scene", "layout"}:
                    continue
                lines.extend(_collect_annotation_lines(sub_val, str(sub_key)))
            return lines

        # 机械 master-spec 的结构化参数
        if {"name", "value"} & set(value.keys()):
            name = str(value.get("name") or key_norm or "").strip()
            raw = value.get("raw")
            val = value.get("value")
            unit = str(value.get("unit") or "").strip()
            if raw not in (None, ""):
                text = str(raw).strip()
            else:
                text = str(val).strip()
                if unit and unit not in text:
                    text = f"{text}{unit}"
            if name and text:
                lines.append(f"{name}：{text}")
            elif text:
                lines.append(text)
            return lines

        # 常见机械字段：尽量把 key/value 显式化
        priority_keys = {
            "load": "载荷",
            "precision": "精度",
            "stroke": "行程",
            "material": "材料",
            "working_condition": "工况",
            "technical_params": "",
            "critical_params": "",
            "dimensions": "尺寸",
            "dimension": "尺寸",
            "forces": "受力",
            "force": "受力",
            "annotations": "",
            "annotation": "",
            "numeric_annotations": "",
        }
        display_key = priority_keys.get(key_norm, key_norm)
        for sub_key, sub_val in value.items():
            if sub_key in {"id", "seq", "type", "title", "description", "template", "kind", "scene", "layout"}:
                continue
            if isinstance(sub_val, (dict, list)):
                lines.extend(_collect_annotation_lines(sub_val, str(sub_key)))
                continue
            sub_text = str(sub_val).strip()
            if not sub_text:
                continue
            if key_norm in {"technical_params", "critical_params", "dimensions", "forces"} or display_key:
                name = str(sub_key).strip()
                if name and sub_text:
                    if name in {"value", "raw"} and display_key:
                        continue
                    if display_key:
                        # technical_params: 载荷=500N -> 载荷：500N
                        if key_norm == "technical_params":
                            lines.append(f"{name}：{sub_text}")
                        elif key_norm == "critical_params":
                            lines.append(f"{sub_text}")
                        else:
                            lines.append(f"{display_key}{name}：{sub_text}")
                    else:
                        lines.append(f"{name}：{sub_text}")
        return lines

    text = str(value).strip()
    if text:
        if key_norm:
            lines.append(f"{key_norm}：{text}")
        else:
            lines.append(text)
    return lines


def _collect_structured_text(drawing: dict) -> str:
    tokens = []
    for key in (
        "structured", "meta", "payload", "annotations", "steps", "parts",
        "dimensions", "forces", "labels", "params", "parameters",
        "flow_steps", "motion_steps", "nodes", "links", "arrows", "notes",
        "numeric_annotations", "critical_params", "technical_params",
    ):
        val = drawing.get(key)
        if val not in (None, "", [], {}):
            tokens.extend(list(_iter_text_tokens(val)))
            tokens.extend(_collect_annotation_lines(val, key))
    return " ".join(tokens)


def _extract_numbers_from_text(text: str) -> list:
    """提取文本中的数字/带单位数字作为注释候选。"""
    if not text:
        return []
    found = []
    patterns = [
        r'[φΦØ]?\s*\d+(?:\.\d+)?\s*(?:mm|cm|m|μm|°|N|kN|MPa|rpm|r/min|s|min|%)?',
        r'[A-Za-zα-ωΑ-Ω][A-Za-z0-9_]{0,8}\s*[:=]\s*-?\d+(?:\.\d+)?\s*(?:mm|cm|m|μm|°|N|kN|MPa|rpm|r/min|s|min|%)?',
    ]
    for pat in patterns:
        for m in re.finditer(pat, text, flags=re.IGNORECASE):
            item = m.group(0).strip()
            if item and item not in found:
                found.append(item)
    if not found:
        for m in re.finditer(r'-?\d+(?:\.\d+)?(?:\s*(?:mm|cm|m|μm|°|N|kN|MPa|rpm|r/min|s|min|%))?', text, flags=re.IGNORECASE):
            item = m.group(0).strip()
            if item and item not in found:
                found.append(item)
    return found[:8]


def _split_phrases(text: str, limit: int = 5) -> list:
    phrases = [s.strip(" ，,；;。:\n\t") for s in re.split(r"[，,；;。:\n\t、]+", text or "") if s.strip(" ，,；;。:\n\t")]
    # 去重保序
    out = []
    for p in phrases:
        if p not in out:
            out.append(p)
    return out[:limit]


def _pick_mech_template(drawing: dict) -> str:
    explicit = [
        drawing.get("template"),
        drawing.get("diagram_type"),
        drawing.get("figure_type"),
        drawing.get("kind"),
        drawing.get("role"),
        drawing.get("category"),
        drawing.get("template_name"),
    ]
    for cand in explicit:
        text = re.sub(r"\s+", "", str(cand or ""))
        if not text:
            continue
        if any(k in text for k in ("flow", "流程", "工艺", "时序", "过程")):
            return "flow"
        if any(k in text for k in ("force", "受力", "载荷", "应力", "扭矩", "压力", "load", "stress", "moment")):
            return "force"
        if any(k in text for k in ("assembly", "装配", "拆装", "爆炸", "分解", "安装", "组合", "explode")):
            return "assembly"
        if any(k in text for k in ("motion", "运动", "位移", "轨迹", "行程", "转动", "旋转", "kinematics", "velocity")):
            return "motion"
        if any(k in text for k in ("parameter", "参数", "标注", "公差", "注释", "dimension", "annotation", "tolerance")):
            return "parameter"
        if any(k in text for k in ("structure", "结构", "布局", "原理", "示意")):
            return "structure"

    text = re.sub(r"\s+", "", _mech_text_blob(drawing))
    if any(k in text for k in ("流程", "工艺", "步骤", "过程", "时序", "flow", "sequence")):
        return "flow"
    if any(k in text for k in ("受力", "载荷", "应力", "扭矩", "压力", "force", "load", "stress", "moment")):
        return "force"
    if any(k in text for k in ("装配", "拆装", "爆炸", "分解", "安装", "组合", "assembly", "explode")):
        return "assembly"
    if any(k in text for k in ("运动", "位移", "轨迹", "行程", "转动", "旋转", "motion", "kinematics", "velocity")):
        return "motion"
    if any(k in text for k in ("尺寸", "参数", "标注", "公差", "注释", "dimension", "annotation", "tolerance")):
        return "parameter"
    return "structure"


def _draw_arrow(draw: ImageDraw.ImageDraw, start, end, color, width=5, head=16):
    draw.line((*start, *end), fill=color, width=width)
    dx = end[0] - start[0]
    dy = end[1] - start[1]
    length = max((dx * dx + dy * dy) ** 0.5, 1)
    ux, uy = dx / length, dy / length
    px, py = -uy, ux
    tip = end
    p1 = (tip[0] - ux * head + px * head * 0.6, tip[1] - uy * head + py * head * 0.6)
    p2 = (tip[0] - ux * head - px * head * 0.6, tip[1] - uy * head - py * head * 0.6)
    draw.polygon([tip, p1, p2], fill=color)


def _fit_text(draw, text, font, max_width):
    words = list(text)
    lines, buf = [], ""
    for ch in words:
        test = buf + ch
        if draw.textbbox((0, 0), test, font=font)[2] <= max_width or not buf:
            buf = test
        else:
            lines.append(buf)
            buf = ch
    if buf:
        lines.append(buf)
    return "\n".join(lines)


def _is_blank_image(img_path: str) -> bool:
    """粗略判断生成图是否近乎空白，避免把白图写入成果目录。"""
    try:
        img = Image.open(img_path).convert("RGB")
    except Exception:
        return False
    px = list(img.getdata())
    total = len(px) or 1
    white = sum(1 for r, g, b in px if r >= 248 and g >= 248 and b >= 248)
    white_ratio = white / total
    # 既要接近纯白，又要缺乏明显图形变化
    if white_ratio >= 0.96:
        return True
    if white_ratio >= 0.92:
        # 对于高白底图，再看一下色彩波动是否极低
        samples = px[:: max(total // 5000, 1)]
        if not samples:
            return True
        channel_spread = 0.0
        for ch in range(3):
            vals = [p[ch] for p in samples]
            mean = sum(vals) / len(vals)
            var = sum((v - mean) ** 2 for v in vals) / len(vals)
            channel_spread += var ** 0.5
        return channel_spread < 25
    return False


def generate_diagram_image(drawing: dict, save_dir: str) -> Optional[str]:
    """生成轻量机械/设计结构图、流程图、原理图的本地图片后端。"""
    os.makedirs(save_dir, exist_ok=True)
    img_type = drawing.get("type", "结构图")
    title = drawing.get("title", "结构示意图")
    description = drawing.get("description", "")
    # 判断是否为土木/建筑类图纸
    arch_text = re.sub(r'\s+', '', " ".join([
        str(drawing.get("type", "")),
        str(drawing.get("title", "")),
        str(drawing.get("description", "")),
    ]))
    is_civil = any(k in arch_text for k in (
        "平面图", "立面图", "剖面图", "标准层", "柱网", "框架",
        "配筋图", "基础图", "节点图", "构造节点", "节点详图", "详图",
        "施工平面", "施工进度", "横道图", "正立面", "侧立面",
        "建筑平面", "建筑立面", "建筑剖面", "结构平面",
    ))
    # 非土木/建筑类的diagram才走机械CAD（如机械零件、机构运动等）
    if not is_civil:
        try:
            from mechanical_cad import generate_mechanical_cad_image
        except Exception as e:
            generate_mechanical_cad_image = None
            print(f"机械CAD后端导入失败，回退到模板后端: {e}")
        if generate_mechanical_cad_image is not None:
            try:
                mech_path = generate_mechanical_cad_image(drawing, save_dir)
                if mech_path and os.path.exists(mech_path):
                    return mech_path
            except Exception as e:
                print(f"机械CAD后端生成失败，回退到模板后端: {e}")
    width, height = 1400, 900
    img = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(img)
    title_font = _load_font(38, bold=True)
    box_font = _load_font(28)
    small_font = _load_font(22)
    mini_font = _load_font(18)
    typ_norm = re.sub(r'\s+', '', _mech_text_blob(drawing))
    template = _pick_mech_template(drawing)
    arch_text = re.sub(r'\s+', '', " ".join([
        str(drawing.get("type", "")),
        str(drawing.get("title", "")),
        str(drawing.get("description", "")),
    ]))
    if any(k in arch_text for k in ("平面图", "总平面图", "功能分区", "布局图", "布置图")):
        template = "plan"
    elif any(k in arch_text for k in ("立面图", "剖面图")):
        template = "elevation"
    elif any(k in arch_text for k in ("节点图", "构造节点", "节点详图", "详图")):
        template = "detail"
    elif any(k in arch_text for k in ("分析图", "流线图", "关系图")):
        template = "analysis"
    struct_text = _collect_structured_text(drawing)
    annotation_lines = []
    for key in (
        "numeric_annotations", "parameters", "technical_params", "critical_params",
        "dimensions", "dimension_text", "forces", "force_text", "annotations",
        "annotation", "notes", "params", "label_text", "payload", "structured",
        "meta", "data", "info",
    ):
        annotation_lines.extend(_collect_annotation_lines(drawing.get(key), key))
    # 保留顺序去重
    annotation_lines = [x for i, x in enumerate(annotation_lines) if x and x not in annotation_lines[:i]]
    numeric_annotations = drawing.get("numeric_annotations")
    if not numeric_annotations:
        numeric_annotations = _extract_numbers_from_text(" ".join([description, struct_text, " ".join(annotation_lines)]))
    if not isinstance(numeric_annotations, list):
        numeric_annotations = [str(numeric_annotations)]
    numeric_annotations = [str(x) for x in numeric_annotations if str(x).strip()][:8]
    if annotation_lines:
        for line in annotation_lines:
            if line not in numeric_annotations:
                numeric_annotations.append(line)
    numeric_annotations = numeric_annotations[:8]
    items = _split_phrases(struct_text or description, limit=6)
    if not items:
        items = ["输入", "处理", "传递", "输出"]

    # 标题
    header_fill = {
        "flow": (236, 244, 255),
        "force": (252, 240, 240),
        "assembly": (242, 247, 240),
        "motion": (243, 241, 253),
        "parameter": (250, 246, 236),
        "structure": (245, 247, 250),
    }.get(template, (245, 247, 250))
    draw.rectangle((60, 40, width - 60, 150), outline=(60, 60, 60), width=3, fill=header_fill)
    draw.text((90, 70), title, fill=(20, 20, 20), font=title_font)
    draw.text((90, 190), f"类型：{img_type}  |  模板：{template}", fill=(80, 80, 80), font=small_font)
    if numeric_annotations:
        draw.text((90, 220), "注释：" + "；".join(numeric_annotations[:4]), fill=(95, 95, 95), font=mini_font)

    if template == "flow":
        steps = items[:6] or ["输入", "处理", "判断", "执行", "输出"]
        x0, y0, w, h, gap = 120, 300, 220, 92, 46
        for i, step in enumerate(steps):
            x = x0 + i * (w + gap)
            y = y0 + (i % 2) * 22
            draw.rounded_rectangle((x, y, x + w, y + h), radius=16, outline=(43, 104, 168), width=4, fill=(236, 244, 255))
            draw.text((x + 18, y + 30), f"{i+1}. {_fit_text(draw, step[:12], box_font, 180)}", fill=(20, 20, 20), font=box_font)
            if i < len(steps) - 1:
                _draw_arrow(draw, (x + w, y + h // 2), (x + w + gap - 8, y + h // 2), (43, 104, 168), width=4)
        draw.text((90, 760), "流程/步骤图", fill=(90, 90, 90), font=small_font)

    elif template == "plan":
        # 平面/总平面/功能分区图：用更完整的空间分块和流线表达，避免输出近白图
        outer = (120, 260, 1280, 710)
        draw.rounded_rectangle(outer, radius=18, outline=(54, 76, 108), width=4, fill=(248, 250, 253))
        # 入口与边界
        draw.rectangle((110, 470, 145, 550), outline=(54, 76, 108), width=3, fill=(236, 244, 255))
        draw.text((100, 440), "入口", fill=(54, 76, 108), font=mini_font)
        # 三分区布局
        zones = [
            ((180, 300, 520, 520), "核心区", (236, 244, 255)),
            ((540, 300, 920, 520), "辅助区", (242, 247, 250)),
            ((940, 300, 1220, 520), "过渡区", (245, 247, 241)),
        ]
        for rect, label, fill in zones:
            draw.rounded_rectangle(rect, radius=14, outline=(54, 76, 108), width=3, fill=fill)
            draw.text((rect[0] + 20, rect[1] + 18), label, fill=(20, 20, 20), font=box_font)
        # 家具/功能块
        blocks = [
            (210, 350, 290, 410), (320, 350, 400, 410), (570, 350, 650, 410),
            (680, 350, 760, 410), (970, 350, 1050, 410), (1080, 350, 1160, 410)
        ]
        for rect in blocks:
            draw.rectangle(rect, outline=(96, 120, 160), width=2, fill=(255, 255, 255))
        # 流线
        _draw_arrow(draw, (145, 510), (180, 410), (43, 104, 168), width=4)
        _draw_arrow(draw, (520, 410), (540, 410), (43, 104, 168), width=4)
        _draw_arrow(draw, (920, 410), (940, 410), (43, 104, 168), width=4)
        draw.text((90, 760), "平面/功能分区图", fill=(90, 90, 90), font=small_font)

    elif template == "elevation":
        # 立面/剖面图：用分层界面、开窗、界面材料条带做出完整建筑立面感
        ground = 680
        draw.line((120, ground, 1280, ground), fill=(50, 50, 50), width=4)
        façade = (180, 260, 1220, 680)
        draw.rounded_rectangle(façade, radius=10, outline=(74, 84, 100), width=4, fill=(248, 249, 252))
        # 材料带
        bands = [
            (200, 280, 1200, 360, (236, 244, 255), "上部檐口/屋面"),
            (200, 360, 1200, 520, (245, 247, 241), "中部界面/开窗"),
            (200, 520, 1200, 660, (242, 247, 252), "底部基座/收边"),
        ]
        for x1, y1, x2, y2, fill, label in bands:
            draw.rectangle((x1, y1, x2, y2), outline=(74, 84, 100), width=2, fill=fill)
            draw.text((x1 + 18, y1 + 10), label, fill=(20, 20, 20), font=mini_font)
        # 窗洞/门洞
        openings = [(280, 390, 430, 500), (500, 390, 660, 500), (720, 390, 860, 500), (940, 390, 1120, 500)]
        for rect in openings:
            draw.rectangle(rect, outline=(54, 76, 108), width=3, fill=(255, 255, 255))
        # 尺寸标记与节点说明
        _draw_arrow(draw, (260, 240), (260, 280), (54, 76, 108), width=3)
        _draw_arrow(draw, (1140, 240), (1140, 280), (54, 76, 108), width=3)
        draw.text((90, 760), "立面/剖面图", fill=(90, 90, 90), font=small_font)

    elif template == "detail":
        # 节点/构造详图：做成放大节点 + 标注框，避免只剩几条线
        main = (300, 290, 980, 620)
        draw.rounded_rectangle(main, radius=18, outline=(170, 128, 46), width=4, fill=(252, 247, 236))
        core = (520, 370, 760, 540)
        draw.rectangle(core, outline=(130, 96, 32), width=4, fill=(255, 253, 245))
        draw.text((575, 438), "节点核心", fill=(20, 20, 20), font=box_font)
        # 叠层材料
        layers = [
            (540, 330, 740, 365, "上层收口"),
            (540, 540, 740, 575, "下层收边"),
        ]
        for x1, y1, x2, y2, label in layers:
            draw.rectangle((x1, y1, x2, y2), outline=(170, 128, 46), width=2, fill=(255, 244, 215))
            draw.text((x1 + 14, y1 + 6), label, fill=(120, 86, 32), font=mini_font)
        callouts = [
            ((520, 370), (380, 320), "1"),
            ((760, 370), (900, 320), "2"),
            ((760, 540), (900, 600), "3"),
            ((520, 540), (380, 600), "4"),
        ]
        for s, e, lab in callouts:
            _draw_arrow(draw, s, e, (170, 128, 46), width=4, head=12)
            draw.ellipse((e[0]-22, e[1]-22, e[0]+22, e[1]+22), outline=(170, 128, 46), width=3, fill=(255, 244, 215))
            draw.text((e[0]-7, e[1]-11), lab, fill=(120, 86, 32), font=mini_font)
        draw.text((90, 760), "节点/构造详图", fill=(90, 90, 90), font=small_font)

    elif template == "analysis":
        panels = [
            ((170, 300, 470, 560), "现状区", (236, 244, 255)),
            ((565, 300, 865, 560), "问题区", (252, 240, 240)),
            ((960, 300, 1260, 560), "优化区", (242, 247, 241)),
        ]
        for rect, label, fill in panels:
            draw.rounded_rectangle(rect, radius=16, outline=(54, 76, 108), width=4, fill=fill)
            draw.text((rect[0] + 20, rect[1] + 18), label, fill=(20, 20, 20), font=box_font)
        _draw_arrow(draw, (470, 430), (565, 430), (198, 72, 72), width=5)
        _draw_arrow(draw, (865, 430), (960, 430), (78, 126, 72), width=5)
        for x in (215, 620, 1015):
            draw.rectangle((x, 360, x + 140, 450), outline=(100, 100, 100), width=2, fill=(255, 255, 255))
        draw.text((90, 760), "分析/关系图", fill=(90, 90, 90), font=small_font)

    elif template == "force":
        body = (510, 340, 890, 560)
        draw.rounded_rectangle(body, radius=26, outline=(80, 80, 80), width=4, fill=(250, 248, 248))
        draw.text((620, 430), "受力对象", fill=(20, 20, 20), font=box_font)
        force_names = ["F1", "F2", "F3", "F4"]
        if numeric_annotations:
            force_names = numeric_annotations[:4] + force_names
        arrows = [
            ((250, 420), (510, 420), force_names[0]),
            ((1150, 420), (890, 420), force_names[1]),
            ((700, 760), (700, 560), force_names[2] if len(force_names) > 2 else "F3"),
            ((700, 190), (700, 340), force_names[3] if len(force_names) > 3 else "F4"),
        ]
        for s, e, lab in arrows:
            _draw_arrow(draw, s, e, (198, 72, 72), width=6)
            tx, ty = ((s[0] + e[0]) // 2, (s[1] + e[1]) // 2)
            draw.rounded_rectangle((tx - 64, ty - 24, tx + 64, ty + 24), radius=12, fill=(255, 236, 236), outline=(198, 72, 72), width=2)
            draw.text((tx - 54, ty - 12), lab[:14], fill=(130, 35, 35), font=mini_font)
        draw.text((90, 760), "受力/载荷图", fill=(90, 90, 90), font=small_font)

    elif template == "assembly":
        center = (700, 470)
        parts = items[:5] or ["零件A", "零件B", "零件C", "连接件", "总成"]
        positions = [
            (180, 260), (1020, 250), (150, 610), (1030, 600), (560, 690)
        ]
        draw.ellipse((center[0]-128, center[1]-128, center[0]+128, center[1]+128), outline=(78, 126, 72), width=5, fill=(237, 247, 235))
        draw.text((center[0]-54, center[1]-18), "总成", fill=(20, 20, 20), font=box_font)
        for i, part in enumerate(parts):
            x, y = positions[i % len(positions)]
            rect = (x, y, x + 220, y + 88)
            draw.rounded_rectangle(rect, radius=16, outline=(78, 126, 72), width=4, fill=(245, 252, 243))
            draw.text((x + 18, y + 28), part[:12], fill=(22, 22, 22), font=box_font)
            _draw_arrow(draw, (x + 110, y + 44), center, (78, 126, 72), width=4, head=14)
            draw.line((x + 110, y + 44, center[0], center[1]), fill=(140, 140, 140), width=2)
        draw.text((90, 760), "装配/爆炸图", fill=(90, 90, 90), font=small_font)

    elif template == "motion":
        path_pts = [(170, 640), (360, 570), (550, 500), (760, 420), (980, 350), (1190, 280)]
        for i in range(len(path_pts) - 1):
            _draw_arrow(draw, path_pts[i], path_pts[i + 1], (112, 92, 196), width=4, head=14)
        for idx, pt in enumerate(path_pts[::2]):
            x, y = pt
            draw.ellipse((x - 24, y - 24, x + 24, y + 24), outline=(112, 92, 196), width=4, fill=(239, 236, 252))
            draw.text((x - 10, y - 10), str(idx + 1), fill=(70, 50, 140), font=mini_font)
        draw.rounded_rectangle((520, 260, 860, 390), radius=16, outline=(112, 92, 196), width=4, fill=(243, 241, 253))
        draw.text((555, 305), "运动主体", fill=(20, 20, 20), font=box_font)
        if items:
            draw.text((90, 760), "；".join(items[:4]), fill=(90, 90, 90), font=small_font)
        else:
            draw.text((90, 760), "运动/位移示意图", fill=(90, 90, 90), font=small_font)

    elif template == "parameter":
        body = (300, 280, 1020, 620)
        draw.rounded_rectangle(body, radius=20, outline=(170, 128, 46), width=4, fill=(252, 247, 236))
        draw.rectangle((450, 360, 760, 520), outline=(170, 128, 46), width=4, fill=(255, 253, 245))
        draw.text((510, 415), "主体零件", fill=(20, 20, 20), font=box_font)
        dims = numeric_annotations[:4] or ["L", "W", "H", "φd"]
        callouts = [
            ((450, 360), (350, 300), dims[0]),
            ((760, 360), (950, 300), dims[1] if len(dims) > 1 else "W"),
            ((760, 520), (960, 600), dims[2] if len(dims) > 2 else "H"),
            ((450, 520), (350, 620), dims[3] if len(dims) > 3 else "φd"),
        ]
        for s, e, lab in callouts:
            _draw_arrow(draw, s, e, (170, 128, 46), width=4, head=12)
            cx, cy = (e[0], e[1])
            box = (cx - 86, cy - 20, cx + 86, cy + 20)
            draw.rounded_rectangle(box, radius=10, fill=(255, 244, 215), outline=(170, 128, 46), width=2)
            draw.text((cx - 74, cy - 11), lab[:18], fill=(120, 86, 32), font=mini_font)
        draw.text((90, 760), "参数/尺寸标注图", fill=(90, 90, 90), font=small_font)

    else:  # structure
        # 结构模板：三层剖面/模块堆叠布局，避免再出现“两个方块+圆圈”的通用样式
        top_label = items[0] if len(items) > 0 else "上层模块"
        mid_label = items[1] if len(items) > 1 else "核心机构"
        bot_label = items[2] if len(items) > 2 else "底座/支撑"
        if "定位" in typ_norm:
            top_label, mid_label, bot_label = "定位基准", "定位元件", "工件/夹具体"
        elif "装配" in typ_norm:
            top_label, mid_label, bot_label = "部件A", "部件B", "总成/装配体"
        elif "布局" in typ_norm:
            top_label, mid_label, bot_label = "上游模块", "核心模块", "下游模块"

        layer_boxes = [
            ((390, 255, 1010, 345), top_label, (242, 247, 250)),
            ((330, 390, 1070, 510), mid_label, (235, 243, 252)),
            ((390, 555, 1010, 645), bot_label, (242, 247, 250)),
        ]
        for rect, label, fill in layer_boxes:
            draw.rounded_rectangle(rect, radius=16, outline=(72, 84, 100), width=4, fill=fill)
            draw.text((rect[0] + 28, rect[1] + 28), label, fill=(20, 20, 20), font=box_font)

        # 左侧支撑/输入，右侧输出/接口，用线连接三层结构
        side_blocks = [
            ((110, 388, 280, 472), "输入/驱动"),
            ((1120, 388, 1290, 472), "输出/接口"),
        ]
        for rect, label in side_blocks:
            draw.rounded_rectangle(rect, radius=14, outline=(43, 104, 168), width=4, fill=(236, 244, 255))
            draw.text((rect[0] + 18, rect[1] + 26), label, fill=(35, 72, 120), font=mini_font)

        _draw_arrow(draw, (280, 430), (330, 430), (43, 104, 168), width=4)
        _draw_arrow(draw, (1070, 430), (1120, 430), (43, 104, 168), width=4)
        _draw_arrow(draw, (700, 345), (700, 390), (43, 104, 168), width=4)
        _draw_arrow(draw, (700, 510), (700, 555), (43, 104, 168), width=4)

        # 结构说明与参数显式标注
        if numeric_annotations:
            for i, ann in enumerate(numeric_annotations[:4]):
                y = 700 + i * 34
                draw.rounded_rectangle((90, y, 380, y + 28), radius=8, fill=(236, 244, 255), outline=(43, 104, 168), width=2)
                draw.text((104, y + 4), f"参数 {i+1}：{ann[:24]}", fill=(35, 72, 120), font=mini_font)
        draw.text((90, 760), "结构/剖面示意图", fill=(90, 90, 90), font=small_font)

    key = drawing.get("seq") or drawing.get("id") or uuid.uuid4().hex[:8]
    filename = f"drawing_{key}.png"
    out_path = os.path.join(save_dir, filename)
    img.save(out_path)
    return out_path


def generate_all_images(drawings: list, save_dir: str, max_workers: int = 2) -> dict:
    """并发生成所有图纸图片，返回 {seq: local_path}。
    注意：只给没有seq的drawing分配序号——补图场景传入的是缺失子集，
    若无条件重编号1..N会导致文件名与原seq错位、覆盖已有图。"""
    result = {}
    os.makedirs(save_dir, exist_ok=True)
    used = {d.get("seq") for d in drawings if d.get("seq")}
    nxt = 1
    for d in drawings:
        if not d.get("seq"):
            while nxt in used:
                nxt += 1
            d["seq"] = nxt
            used.add(nxt)
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        fut_map = {executor.submit(generate_single_image, d, save_dir): str(d["seq"]) for d in drawings}
        for fut in as_completed(fut_map):
            did = str(fut_map[fut])
            path = fut.result()
            if path:
                result[did] = path
    print(f"图纸生成完成: {len(result)}/{len(drawings)} 成功")
    return result


def add_drawing_image(doc, drawing: dict, img_path: str, dn: int, skip_title: bool = False):
    """向docx中添加设计图纸图片。
    skip_title: 正文紧邻处已有同款图注行时跳过插入title(防图注×2, 0908修复)"""
    title = drawing.get("title", f"设计图{dn}")
    desc = drawing.get("description", "")
    if not skip_title:
        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p_title.add_run(title)  # 直接用LLM给的title原名，不再加"图{dn}"前缀
        set_run_font(r, "宋体", 10, bold=True)
    try:
        doc.add_picture(img_path, width=Cm(14))
    except Exception as e:
        print(f"添加设计图失败: {e}")
    if desc:
        p_desc = doc.add_paragraph()
        p_desc.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p_desc.add_run(desc)
        set_run_font(r, "宋体", 9)
    last_p = doc.add_paragraph()
    last_p.paragraph_format.space_after = Pt(6)


# ==================== DOCX 最终排版 ====================
def finalize_docx(docx_path: str, update=None):
    """分节符修正 + PDF真实页码 + 目录更新 + 页脚"""
    if update is None:
        def update(msg, prog): print(f"[{prog}%] {msg}")
    update("正在处理分节符和页码...", 92)
    doc = Document(docx_path)
    paragraphs = doc.paragraphs
    body = doc.element.body

    # 找目录和第一章位置
    toc_idx = -1
    for i, p in enumerate(paragraphs):
        if p.text.strip() == '目录':
            toc_idx = i
            break
    chapter1_idx = -1
    for i, p in enumerate(paragraphs):
        t = p.text.strip()
        # 正文第1章：排除TOC中的（后面带数字页码）和标题中的（居中对齐）
        if re.match(r'^第1章\s', t) and i > toc_idx + 3:
            # 检查：如果不包含数字（页码未回填），或者是正文的第一章（前面有空白段落组）
            has_digit_suffix = bool(re.search(r'\d+$', t.replace(' ', '')))
            if not has_digit_suffix:
                chapter1_idx = i
                break
            # 有数字后缀，可能还是TOC中的。检查前面段落是否是"参考文献 "
            for j in range(max(0, i-5), i):
                if '参考文献' in paragraphs[j].text.strip():
                    # TOC中的，跳过
                    break
            else:
                chapter1_idx = i
                break
    if chapter1_idx < 0:
        raise RuntimeError("finalize_docx失败：未找到正文第1章，禁止继续交付")

    def _apply_toc_line(p, heading_text: str, sec_page: int):
        p.clear()
        pPr = p._element.get_or_add_pPr()
        p_format = p.paragraph_format
        try:
            p_format.tab_stops.clear_all()
        except Exception:
            pass
        try:
            p_format.tab_stops.add_tab_stop(Cm(16), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        except Exception:
            pass
        tabs_elem = OxmlElement('w:tabs')
        tab_elem = OxmlElement('w:tab')
        tab_elem.set(qn('w:val'), 'right')
        tab_elem.set(qn('w:leader'), 'dot')
        tab_elem.set(qn('w:pos'), str(int(Cm(16).twips)))
        tabs_elem.append(tab_elem)
        for old_tabs in pPr.findall(qn('w:tabs')):
            pPr.remove(old_tabs)
        pPr.append(tabs_elem)
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(heading_text)
        set_run_font(r, "宋体", 14)
        p.add_run("\t")
        rp = p.add_run(str(sec_page))
        set_run_font(rp, "宋体", 14)

    toc_end_idx = chapter1_idx - 1
    while toc_end_idx > toc_idx:
        if paragraphs[toc_end_idx].text.strip():
            break
        toc_end_idx -= 1

    # 收集目录条目
    toc_headings = []
    toc_range_start = toc_idx + 1 if toc_idx >= 0 else 0
    for i in range(toc_range_start, chapter1_idx):
        t = paragraphs[i].text.strip()
        if t:
            heading = t.split('\t')[0].strip() if '\t' in t else t
            toc_headings.append(heading)

    # 生成临时PDF计算页码
    tag = uuid.uuid4().hex[:8]
    temp_docx = docx_path.replace('.docx', f'_temp_{tag}.docx')
    doc.save(temp_docx)
    update("正在用LibreOffice渲染PDF计算页码...", 95)
    pdf_path = os.path.join(os.path.dirname(temp_docx) or '.', f'_temp_{tag}.pdf')
    subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf',
                    temp_docx, '--outdir', os.path.dirname(pdf_path)],
                   capture_output=True, text=True, timeout=300)
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
        # 找目录页
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

        update("正在更新目录页码...", 97)
        # 对所有TOC条目，不做PDF匹配，直接使用PDF页码计算
        # 方案：用pdftotext扫描PDF获取每个标题的页码
        last_sec_page = 1  # 从第1页开始
        for i in range(toc_range_start, chapter1_idx):
            try:
                p = paragraphs[i]
                t = p.text.strip()
                if not t:
                    continue
                heading_text = t.split('\t')[0].strip() if '\t' in t else t
                # PDF精确匹配
                if heading_text in heading_abs_pages:
                    sec_page = heading_abs_pages[heading_text] - offset
                    last_sec_page = sec_page
                else:
                    # 回退：用上一个已知页码
                    sec_page = last_sec_page
                    last_sec_page = sec_page + 0  # 同一页（TOC子项通常在同一页或下一页）
                if sec_page is not None and sec_page >= 1:
                    _apply_toc_line(p, heading_text, sec_page)
            except Exception as e:
                print("  页码更新异常 [%d]: %s" % (i, str(e)))
                continue

        # 二次兜底：如果还有未回填页码的TOC条目，沿用上一个有效页码补齐
        last_known_page = None
        for i in range(toc_range_start, chapter1_idx):
            try:
                p = paragraphs[i]
                t = p.text.strip()
                if not t:
                    continue
                if '\t' in t:
                    parts = t.rsplit('\t', 1)
                    if len(parts) == 2 and parts[1].strip().isdigit():
                        last_known_page = int(parts[1].strip())
                    continue
                if last_known_page is None:
                    continue
                _apply_toc_line(p, t, last_known_page)
            except Exception as e:
                print("  页码二次回填异常 [%d]: %s" % (i, str(e)))
                continue
        if os.path.exists(pdf_path):
            os.remove(pdf_path)

    # 分节符处理
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

    # 移除第1章的pageBreakBefore
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


# ==================== 诊断报告 ====================
def save_diagnosis_report(revision_plan: dict, output_path: str):
    """保存诊断报告到文本文件"""
    diag = revision_plan.get("diagnosis_report", {})
    lines = [
        "=" * 50, "【论文诊断报告】", "=" * 50, "",
        f"总体评分：{diag.get('overall_score', 'N/A')}", "",
        "【总体评价】", diag.get("overall_comment", "暂无"), ""
    ]
    issues = diag.get("issues", [])
    if issues:
        lines.append("【问题清单】")
        for i, issue in enumerate(issues, 1):
            lines.extend([
                f"  问题{i}：",
                f"    涉及章节：{issue.get('chapter', 'N/A')}",
                f"    严重程度：{issue.get('severity', 'N/A')}",
                f"    问题类型：{issue.get('issue_type', 'N/A')}",
                f"    问题描述：{issue.get('description', 'N/A')}",
                f"    修正方向：{issue.get('fix_direction', 'N/A')}", ""
            ])
    else:
        lines.extend(["【问题清单】", "  未发现明显问题（或LLM未返回问题列表）", ""])
    lines.append("=" * 50)
    dir_name = os.path.dirname(output_path)
    if dir_name:
        os.makedirs(dir_name, exist_ok=True)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write("\n".join(lines))


# ==================== 目录排序渲染 ====================
def _render_toc(doc, entries, need_page_break, page_break_applied):
    """按章节顺序排序并渲染目录条目"""
    if not entries:
        return
    
    # 构建章节序号→名称映射
    chapters = {}
    subs_by_chapter = {}
    others = []
    
    chapter_order = []
    chapter_num = 0
    
    for typ, text in entries:
        if typ == 'chapter':
            chapter_num += 1
            chapters[chapter_num] = text
            chapter_order.append((chapter_num, 'chapter', text))
        elif typ == 'subsection':
            # 提取章节号 2.1 → 2
            m = re.match(r'^(\d+)\.', text)
            if m:
                ch = int(m.group(1))
                if ch not in subs_by_chapter:
                    subs_by_chapter[ch] = []
                subs_by_chapter[ch].append(text)
            else:
                others.append((0, 'subsection', text))
        else:
            others.append((99, typ, text))
    
    # 按章节顺序输出
    sorted_entries = []
    for cn in range(1, chapter_num + 1):
        if cn in chapters:
            sorted_entries.append(('chapter', chapters[cn]))
        if cn in subs_by_chapter:
            for s in subs_by_chapter[cn]:
                sorted_entries.append(('subsection', s))
    
    # 最后加参考文献/摘要/关键词
    for _, typ, text in others:
        sorted_entries.append((typ, text))
    
    # 先输出"目录"标题
    from docx.shared import Pt
    from docx.shared import Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    # 目录必须单独起页，避免和封面或前一部分挤在同一页
    if any(p.text.strip() for p in doc.paragraphs):
        doc.add_page_break()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("目录")
    set_run_font(r, "黑体", 18, bold=True)
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(12)
    
    for typ, text in sorted_entries:
        p = doc.add_paragraph()
        if need_page_break and not page_break_applied:
            from docx.oxml import OxmlElement
            pPr = p._element.get_or_add_pPr()
            pb = OxmlElement('w:pageBreakBefore')
            pPr.append(pb)
            page_break_applied = True
            need_page_break = False
        r = p.add_run(text)
        if typ == 'chapter':
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            set_run_font(r, "宋体", 14)
        elif typ == 'subsection':
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.left_indent = Cm(0.74)
            set_run_font(r, "宋体", 14)
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            set_run_font(r, "宋体", 14)

# ==================== TXT→DOCX 通用排版 ====================
def txt_to_docx_safe(txt_path: str, docx_path: str, update=None,
                     cover_info: Optional[dict] = None,
                     drawing_images: Optional[dict] = None):
    """通用 TXT → DOCX 排版（支持图表+图纸+封面+分节）"""
    if update is None:
        def update(msg, prog): print(f"[{prog}%] {msg}")
    if drawing_images is None:
        drawing_images = {}

    update("正在读取TXT...", 10)
    with open(txt_path, "r", encoding="utf-8") as f:
        full_text = f.read()
    full_text = strip_mech_json_blocks(full_text)
    # 0908修复: 剥离[FIGURES]自报清单块(审计对账用, 不入正文——此前被当普通文本
    # 渲染导致图注×2, 何晋案: 10种图/表注重复)
    # 0908一般性加固: 剥离前先对账(自报清单 vs 实际标签), 缺标签图进receipt报警(全专业)
    figures_missing_tags = []
    _fig_blocks = re.findall(r'\[FIGURES\]\s*(.*?)\[/FIGURES\]', full_text, re.S)
    if _fig_blocks:
        _declared = set()
        for _blk in _fig_blocks:
            for _ln in _blk.split("\n"):
                _m = re.match(r"^\s*([图表]\s*\d{1,2}(?:-\d{1,2})?)\s+\S", _ln)
                if _m:
                    _declared.add(re.sub(r"\s+", "", _m.group(1)))
        _actual = set()
        for _m in re.finditer(r'<(?:drawing|table|chart)\b[^>]*?title="([图表]\s*\d{1,2}(?:-\d{1,2})?)', full_text):
            _actual.add(re.sub(r"\s+", "", _m.group(1)))
        figures_missing_tags = sorted(_declared - _actual)
        if figures_missing_tags:
            print(f"⚠ [自报对账] 自报但无标签: {figures_missing_tags} — 该图/表将缺失")
    full_text = re.sub(r'\[FIGURES\]\s*.*?\[/FIGURES\]\s*', '', full_text, flags=re.S)
    txt_title = extract_title_from_txt(full_text)
    if txt_title:
        full_text = strip_title_from_txt(full_text)
    # 预处理：吸收跨行属性、包裹裸属性行（兼容GPT输出格式变体）
    full_text = preprocess_tag_attrs(full_text)
    # 标签守卫：确定性校验/修复/降级（错拼、闭合、数量对齐、rows污染等）
    guard_report = None
    if audit_and_repair is not None:
        try:
            full_text, guard_report = audit_and_repair(full_text)
            summary = render_report_text(guard_report) if render_report_text else ""
            if summary:
                print(f"  [{summary}]")
                update("标签守卫: " + summary, 35)
        except Exception as e:
            raise RuntimeError(f"标签守卫失败，禁止继续渲染: {e}") from e
    else:
        raise RuntimeError("标签守卫不可用，禁止继续渲染")
    # 一致性守卫：跨章数值冲突检测与修正（标题事实优先/多数值投票）
    consistency_report = None
    if check_numeric_consistency is not None:
        try:
            full_text, consistency_report = check_numeric_consistency(
                full_text, title=txt_title or "")
            csum = consistency_summary(consistency_report) if consistency_summary else ""
            if csum:
                print(f"  [{csum}]")
                update("一致性守卫: " + csum, 36)
        except Exception as e:
            raise RuntimeError(f"一致性守卫失败，禁止继续渲染: {e}") from e
    else:
        raise RuntimeError("一致性守卫不可用，禁止继续渲染")

    update("正在提取标签数据...", 30)
    # 使用宽容解析器（支持多种标签格式变体）
    charts = tolerant_extract_charts(full_text)
    tables = tolerant_extract_tables(full_text)
    drawings = tolerant_extract_drawings(full_text)

    els = (
        [(c["start_pos"], c["end_pos"], "chart", c) for c in charts]
        + [(t["start_pos"], t["end_pos"], "table", t) for t in tables]
        + [(d["start_pos"], d["end_pos"], "drawing", d) for d in drawings]
    )
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
    _init_s(doc)

    cover = normalize_cover_info(cover_info)
    if txt_title and not cover.get("title"):
        cover["title"] = txt_title
    if cover_info is not None:
        build_cover(doc, cover)  # 允许空字段先出封面
    else:
        build_cover(doc, cover)  # 默认空封面，后续可补录

    cn, tn, dn = 0, 0, 0
    _prev_text_tail = []  # 0908修复: 上一个text part的末尾行(供drawing图注去重判断)
    next_page_break = False
    need_page_break = False
    page_break_applied = False
    in_toc_section = False
    pt_last_heading = ""
    seen_toc_refs = False  # 是否已经在TOC中看到"参考文献"（用于无PAGE_BREAK过渡检测）

    receipt = {
        "schema_version": 2,
        "render_status": "running",
        "charts": {"total": 0, "ok": 0, "placeholder": 0, "failed": []},
        "tables": {"total": 0, "ok": 0, "fallback": 0, "failed": [], "manual_review": []},
        "drawings": {"total": 0, "ok": 0, "placeholder": 0, "missing_image": []},
        "postflight": {"docx_valid": False, "mismatches": []},
        "figures_missing_tags": figures_missing_tags,  # 0908: 自报清单有但无标签的图/表
        "guard": {"status": "ok" if isinstance(guard_report, dict) else "failed", "errors": []},
        "consistency": {"status": "ok" if isinstance(consistency_report, dict) else "failed", "errors": []},
    }
    if isinstance(guard_report, dict):
        receipt["guard"] = {
            "status": "ok",
            "scanned": dict(guard_report.get("scanned", {})),
            "fixed": dict(guard_report.get("fixed", {})),
            "errors": list(guard_report.get("errors", [])),
        }
    if isinstance(consistency_report, dict):
        receipt["consistency"] = {
            "status": "ok",
            "params_scanned": consistency_report.get("params_scanned", 0),
            "fixed_count": consistency_report.get("fixed_count", 0),
            "conflicts": consistency_report.get("conflicts", []),
            "errors": list(consistency_report.get("errors", [])),
        }

    for pt, ct in parts:
        if pt == "chart":
            cn += 1
            receipt["charts"]["total"] += 1
            try:
                cb = chart_to_bytes(ct)
                if not cb:
                    raise ValueError("图表渲染返回空(gantt/数据格式不符等)")
                add_c(doc, ct, cb, cn)
                receipt["charts"]["ok"] += 1
            except Exception as e:
                print(f"生成图表失败: {e}")
                receipt["charts"]["failed"].append(
                    {"title": ct.get("title", ""), "error": str(e)[:150]})
                try:
                    add_c_placeholder(doc, ct, cn, str(e))
                    receipt["charts"]["placeholder"] += 1
                except Exception as e2:
                    print(f"图表占位也失败: {e2}")
        elif pt == "table":
            tn += 1
            receipt["tables"]["total"] += 1
            try:
                nrows, empty_ratio = _table_quality(ct)
                if nrows == 0 or empty_ratio > 50:
                    raise ValueError(f"内容残缺({nrows}行,空格率{empty_ratio:.0f}%)")
                add_t(doc, ct, tn)
                receipt["tables"]["ok"] += 1
            except Exception as e:
                print(f"生成表格失败: {e}")
                receipt["tables"]["failed"].append(
                    {"title": ct.get("title", ""), "error": str(e)[:150]})
                try:
                    # 兜底：以原始token强行铺一张表，保结构可见
                    fallback_ct = dict(ct)
                    fallback_ct["header"] = ct.get("header", "") or "项目,内容"
                    if not ct.get("data") and ct.get("rows"):
                        fallback_ct["data"] = ct.get("rows", "")
                        fallback_ct["rows"] = ""
                    fr, fe = _table_quality(fallback_ct)
                    if fr == 0 or fe > 50:
                        # 结构化兜底也残缺 → 单列清单保住全部内容
                        toks = [t for t in re.split(
                            r"[;,，；|｜\n]+",
                            str(ct.get("rows", "")) + ";" + str(ct.get("data", "")))
                            if t.strip()]
                        fallback_ct = dict(ct)
                        fallback_ct["header"] = "内容"
                        fallback_ct["rows"] = ""
                        fallback_ct["data"] = ";".join(toks)
                    add_t(doc, fallback_ct, tn)
                    receipt["tables"]["fallback"] += 1
                    receipt["tables"].setdefault("manual_review", []).append(
                        {"title": ct.get("title", ""), "reason": "fallback_used"})
                except Exception as e2:
                    print(f"表格兜底也失败: {e2}")
        elif pt == "drawing":
            dn += 1
            receipt["drawings"]["total"] += 1
            # 0908修复: 正文紧邻处(前一个text part末尾3行)已有同款图注 → 不再插title(防图注×2)
            _skip_title = False
            if _prev_text_tail:
                _tit = re.sub(r"\s+", "", str(ct.get("title", "")))
                for _ln in _prev_text_tail[-3:]:
                    _lns = re.sub(r"\s+", "", _ln)
                    if _tit and (_tit in _lns or
                                 re.match(r"^图\d{1,2}(-\d{1,2})?\b", _lns) and _lns[:6] == _tit[:6]):
                        _skip_title = True
                        break
            _prev_text_tail = []
            try:
                img_path = _lookup_drawing_image(ct, drawing_images)
                if img_path:
                    add_drawing_image(doc, ct, img_path, dn, skip_title=_skip_title)
                    receipt["drawings"]["ok"] += 1
                else:
                    img_key = str(ct.get("seq") or ct.get("id") or "")
                    print(f"图纸图片缺失 key={img_key}, 使用占位图")
                    receipt["drawings"]["missing_image"].append(
                        {"title": ct.get("title", ""), "key": img_key})
                    ph = _make_placeholder_image(ct.get("title", f"设计图{dn}"))
                    try:
                        add_drawing_image(doc, ct, ph, dn, skip_title=_skip_title)
                        receipt["drawings"]["placeholder"] += 1
                    finally:
                        try:
                            os.remove(ph)
                        except OSError:
                            pass
            except Exception as e:
                print(f"插入图纸失败: {e}")
                receipt["drawings"]["missing_image"].append(
                    {"title": ct.get("title", ""), "error": str(e)[:150]})
        else:
            for raw_line in ct.split('\n'):
                line = clean_text(raw_line).strip()
                # 【修复】跳过因标签位置偏移残留的片段（如 "<", "<c", "<ta" 等）
                if line in ('/', '／'):
                    continue
                if line and line.startswith('<') and not line.startswith('---PAGE') and line != '目录':
                    if len(line) < 10 and not re.match(r'^<\w+\s', line) and not re.match(r'^</?\w+>', line):
                        continue
                if not line:
                    continue
                # 0908修复: 记录文本行供紧随drawing的图注去重
                _prev_text_tail.append(line)
                if len(_prev_text_tail) > 8:
                    _prev_text_tail.pop(0)
                # 分页标记
                if line == '---PAGE_BREAK---':
                    # 如果在目录收集中，先渲染目录
                    if in_toc_section:
                        in_toc_section = False
                        _render_toc(doc, _toc_entries, need_page_break, page_break_applied)
                        pt_last_heading = ""  # 重置，避免正文一级标题被去重
                    next_page_break = True
                    in_toc_section = False
                    continue  # NOT seen: seen_toc_refs只在看到"参考文献"时设置

                if next_page_break:
                    need_page_break = True
                    next_page_break = False
                else:
                    need_page_break = False

                is_toc_header = line == '目录'
                is_level1 = bool(re.match(r'^第[一二三四五六七八九十\d]+章\s+\S', line)) or bool(
                    re.match(r'^摘\s*要$|^关键词$|^参考文献$', line))
                is_level2 = bool(re.match(r'^\d+\.\d+\s+', line)) and not bool(re.match(r'^\d+\.\d+\.\d+\s+', line))
                is_level3 = bool(re.match(r'^\d+\.\d+\.\d+\s+', line))

                # 【修复】检测 'X. 同标题' — 紧跟在第X章标题后的带点重复
                is_dotted_repeat = False
                if pt_last_heading and re.match(r'^\d+\.\s+', line):
                    # pt_last_heading 是 '第X章 YYYY'，提取 YYYY 部分
                    title_part = re.sub(r'^第[一二三四五六七八九十\d]+章\s+', '', pt_last_heading).strip()
                    dotted_text = re.sub(r'^\d+\.\s+', '', line).strip()
                    if title_part and dotted_text and title_part[:4] in dotted_text or dotted_text in title_part:
                        is_dotted_repeat = True

                # 进入/退出目录区域
                if line == '目录':
                    in_toc_section = True
                    _toc_entries = []  # 收集目录条目
                    continue

                # 目录模式：收集所有条目，稍后排序渲染
                if in_toc_section:
                    if line.startswith('第') and '章' in line:
                        _toc_entries.append(('chapter', line))
                    elif re.match(r'^\d+\.\d+(\.\d+)?\s+', line):
                        # 二级(N.1)与三级(N.1.1)目录条目都收集; 三级行此前掉入
                        # else分支→目录收集中断+该行被丢弃(每章首条三级必丢)
                        _toc_entries.append(('subsection', line))
                    elif re.match(r'^参考文献\s*$', line):
                        _toc_entries.append(('ref', line))
                    elif re.match(r'^摘\s*要$', line):
                        _toc_entries.append(('abstract', line))
                    elif re.match(r'^关键词$', line):
                        _toc_entries.append(('keyword', line))
                    elif line == '---PAGE_BREAK---':
                        # 遇到分页符时，结束目录收集并渲染
                        in_toc_section = False
                        # 排序并渲染目录
                        _render_toc(doc, _toc_entries, need_page_break, page_break_applied)
                        next_page_break = True
                        continue
                    else:
                        # 非目录内容→结束目录收集
                        in_toc_section = False
                        _render_toc(doc, _toc_entries, need_page_break, page_break_applied)
                        pt_last_heading = ""  # 重置，避免正文一级标题被去重
                        # 此条内容作为正文段落处理，不continue
                    continue

                if is_level1 and not in_toc_section:
                    # 一级标题去重
                    if pt_last_heading:
                        sl = re.sub(r'\s+', '', pt_last_heading.replace('第', '').replace('章', '')).strip()
                        sc = re.sub(r'\s+', '', line.replace('第', '').replace('章', '')).strip()
                        if sc == sl:
                            pt_last_heading = ""
                            continue
                    pt_last_heading = line
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p.paragraph_format.space_before = Pt(12)
                    p.paragraph_format.space_after = Pt(12)
                    r = p.add_run(line)
                    set_run_font(r, "黑体", 16, bold=True)
                    if need_page_break:
                        pPr = p._element.get_or_add_pPr()
                        pb = OxmlElement('w:pageBreakBefore')
                        pPr.append(pb)
                    continue

                # 【修复】跳过 "X. 同标题" — 紧跟在第X章标题后的带点重复行
                if is_dotted_repeat:
                    continue

                pt_last_heading = ""

                if is_toc_header and not in_toc_section:
                    p = doc.add_paragraph()
                    # 分页在当前段落
                    if need_page_break and not page_break_applied:
                        pPr = p._element.get_or_add_pPr()
                        pb = OxmlElement('w:pageBreakBefore')
                        pPr.append(pb)
                        page_break_applied = True
                        need_page_break = False
                    # 检测TOC结束标记
                    if in_toc_section and re.match(r'^参考文献\s*$', line.strip()):
                        seen_toc_refs = True
                    # 目录居中显示
                    if line == '目录':
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        r = p.add_run(line)
                        set_run_font(r, "黑体", 18, bold=True)
                        p.paragraph_format.space_before = Pt(12)
                        p.paragraph_format.space_after = Pt(12)
                    # TOC条目（包含所有章标题+子标题，不跳过）
                    elif in_toc_section:
                        if seen_toc_refs and line.startswith('第'):  # 已过参考文献出现"第X章"→进入正文
                            in_toc_section = False
                            pt_last_heading = ""
                        else:
                            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                            r = p.add_run(line)
                            set_run_font(r, "宋体", 14)
                            p.paragraph_format.space_before = Pt(2)
                            p.paragraph_format.space_after = Pt(2)
                    continue

                if is_level3:
                    p = doc.add_paragraph()
                    r = p.add_run(line)
                    set_run_font(r, "宋体", 12, bold=True)
                elif is_level2:
                    p = doc.add_paragraph()
                    r = p.add_run(line)
                    set_run_font(r, "黑体", 14, bold=True)
                    p.paragraph_format.space_before = Pt(6)
                    p.paragraph_format.space_after = Pt(6)
                else:
                    p = doc.add_paragraph()
                    # 分页在当前段落（不是上一段）
                    if need_page_break and not page_break_applied:
                        pPr = p._element.get_or_add_pPr()
                        pb = OxmlElement('w:pageBreakBefore')
                        pPr.append(pb)
                        page_break_applied = True
                        need_page_break = False
                    r = p.add_run(line)
                    set_run_font(r, "宋体", 12)
                    p.paragraph_format.first_line_indent = Cm(0.74)
                    p.paragraph_format.line_spacing = 1.5

    doc.save(docx_path)
    update("DOCX生成完毕！", 98)
    # 最终排版（分节符+页码+页脚）- 放到最后
    update("正在处理最终排版（分节符/页码/页脚）...", 95)
    finalize_docx(docx_path, update=lambda m, p: update(m, 95 + int(p * 0.05)))
    # finalize 后立即重新打开，确保输出仍是可读的有效DOCX；不能只相信 save() 返回。
    try:
        _postflight = Document(docx_path)
        _postflight.element.body
        receipt["postflight"]["docx_valid"] = True
        receipt["postflight"]["tables"] = len(_postflight.tables)
        receipt["postflight"]["drawings"] = len(_postflight.element.body.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing'))
    except Exception as exc:
        receipt["postflight"]["mismatches"].append(str(exc))
        raise RuntimeError(f"DOCX后置验收失败，禁止交付: {exc}") from exc

    # ===== 渲染回执：写JSON旁车文件 + 控制台摘要（失败可见，不再静默） =====
    receipt["file"] = docx_path
    try:
        report_path = docx_path + ".report.json"
        with open(report_path, "w", encoding="utf-8") as f:
            json.dump(receipt, f, ensure_ascii=False, indent=1)
        c_, t_, d_ = receipt["charts"], receipt["tables"], receipt["drawings"]
        flags = []
        if c_["failed"]:
            flags.append(f"图表失败{len(c_['failed'])}")
        if t_["failed"]:
            flags.append(f"表格降级{len(t_['failed'])}")  # 质量门控/异常→兜底渲染
        if d_["missing_image"]:
            flags.append(f"图纸缺图{len(d_['missing_image'])}")
        cons = receipt.get("consistency")
        if cons and cons.get("conflicts"):
            unfixed = sum(1 for c in cons["conflicts"] if not c.get("fixed"))
            if unfixed:
                flags.append(f"数值冲突未修{unfixed}")
        if c_["placeholder"] or d_["placeholder"] or t_.get("fallback") or flags:
            receipt["render_status"] = "degraded"
            receipt["verdict"] = "manual_review"
        else:
            receipt["render_status"] = "ok"
            receipt["verdict"] = "pass"
        status = "⚠️ " + " / ".join(flags) if flags else "✅ 全部渲染成功"
        print(f"  [渲染回执] chart {c_['ok']}/{c_['total']} | table {t_['ok']}/{t_['total']}"
              f" | drawing {d_['ok']}/{d_['total']} → {status}")
        if cons and cons.get("conflicts"):
            print(f"  [一致性] 扫描{cons['params_scanned']}项, 冲突{len(cons['conflicts'])}, "
                  f"自动修正{cons['fixed_count']}")
        print(f"  [渲染回执] 已写入 {report_path}")
    except Exception as e:
        print(f"渲染回执写入失败: {e}")
    return receipt
    update("DOCX生成完毕！", 100)
