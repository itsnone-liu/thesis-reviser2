#!/usr/bin/env python3
"""
环艺设计类论文批量修改脚本
处理35篇低/无等级论文，针对审核表逐一修复

主要修改：
1. 删掉摘要正文开头多余的"摘要："
2. 从TXT中提取表格数据，重新生成docx表格
3. 其他问题（层次、学号、引号等）

用法：
  python3 fix_hanyi.py [论文准考证号...]
  不指定参数则处理全部35篇
"""

import os
import re
import sys
import json
import shutil
from pathlib import Path
from docx import Document
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ==================== 路径配置 ====================
HUANYI_DIR = Path("/root/project/workspace/论文资料/整理/环艺")
MODIFIED_DIR = Path("/root/project/workspace/论文资料/整理/设计类_修改版")
OUTPUT_DIR = Path("/root/project/workspace/论文资料/整理/环艺_修改版")
AUDIT_FILE = Path("/root/project/workspace/论文资料/整理/设计类审核表.xlsx")

# ==================== 字体工具 ====================
def set_run_font(run, font_name="宋体", size=12, bold=False, color=None):
    """设置run的字体属性"""
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color


def copy_paragraph_format(src_para, dst_para):
    """复制段落格式"""
    dst_para.paragraph_format.alignment = src_para.paragraph_format.alignment
    dst_para.paragraph_format.space_before = src_para.paragraph_format.space_before
    dst_para.paragraph_format.space_after = src_para.paragraph_format.space_after
    dst_para.paragraph_format.first_line_indent = src_para.paragraph_format.first_line_indent
    dst_para.paragraph_format.line_spacing = src_para.paragraph_format.line_spacing

def copy_run_format(original_para, new_para, text):
    """复制段落格式并保留字体属性"""
    if original_para.runs:
        # 取第一个run的格式
        src_run = original_para.runs[0]
        run = new_para.add_run(text)
        run.font.name = src_run.font.name or '宋体'
        try:
            run._element.rPr.rFonts.set(qn('w:eastAsia'), 
                src_run._element.rPr.rFonts.get(qn('w:eastAsia')) if src_run._element.rPr is not None and src_run._element.rPr.rFonts is not None else '宋体')
        except:
            pass
        run.font.size = src_run.font.size or Pt(12)
        run.font.bold = src_run.font.bold
    else:
        run = new_para.add_run(text)
        run.font.name = '宋体'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
        run.font.size = Pt(12)

# ==================== 表格解析 ====================
def parse_html_table_to_rows(tag_content):
    """
    从HTML风格的<table>标签内容中提取表格数据。
    支持几种格式：
    1. <table/><tr><td>...</td></tr>
    2. <table><caption>标题</caption><tr><th>...</th></tr>
    3. <table\n<tr<td...格式（无缩进）
    4. 表格数据格式（标题/列/行）
    
    返回 (caption, headers, rows_data)
    caption: 标题文字
    headers: 列表头
    rows_data: [[行数据], ...]
    """
    caption = ""
    headers = []
    rows_data = []
    
    # 提取caption
    cap_m = re.search(r'<caption[^>]*>([^<]*)</caption', tag_content)
    if cap_m:
        caption = cap_m.group(1).strip()
    
    # 提取所有行（包括内容从<table>开始的）
    # 统一处理：先移除换行符，再提取所有tr块
    flat = tag_content.replace('\n', ' ').replace('\r', '')
    
    # 找所有<tr>...</tr>块
    tr_blocks = re.findall(r'<tr\b[^>]*>(.*?)</tr\s*>', flat, re.DOTALL)
    
    if not tr_blocks:
        # 尝试匹配没有闭合标签的格式
        # 例如：<tr<td...格式
        tr_blocks = re.findall(r'<tr\b[^>]*>(.*?)(?:</tr|$)', flat, re.DOTALL)
    
    if not tr_blocks:
        return caption, headers, rows_data
    
    for tr_idx, tr_content in enumerate(tr_blocks):
        # 提取th或td
        cells = re.findall(r'<(?:th|td)\b[^>]*>(.*?)</(?:th|td)\s*>', tr_content, re.DOTALL)
        if not cells:
            # 尝试没有闭合标签
            cells = re.findall(r'<(?:th|td)\b[^>]*>(.*?)(?:</(?:th|td)|$)', tr_content, re.DOTALL)
        
        cell_texts = [c.strip() for c in cells]
        
        if not cell_texts:
            continue
        
        # 第一行如果有th，是表头
        if tr_idx == 0 and any('<th' in tr_content for _ in ['']):
            headers = cell_texts
        elif headers and len(cell_texts) == len(headers):
            rows_data.append(cell_texts)
        else:
            # 没有th的情况下，第一行当表头
            if tr_idx == 0 and not headers:
                headers = cell_texts
            else:
                rows_data.append(cell_texts)
    
    # 处理特殊的表格数据格式（如功能分区面积分配表）
    if not headers and not rows_data:
        rows_data = _extract_table_data_format(flat)
    
    return caption, headers, rows_data


def _extract_table_data_format(text):
    """处理特殊格式如：标题：xxx 列：xxx 行1：xxx"""
    # 找"行\d："格式的数据
    rows = []
    # 按行分割
    for line in text.split('\n'):
        line = line.strip()
        m = re.match(r'行\d[：:](.*)', line)
        if m:
            cells = [c.strip() for c in m.group(1).split('|')]
            rows.append(cells)
    return rows


def add_table_to_doc(doc, caption, headers, rows_data, table_idx, original_para=None):
    """
    在doc中插入一个格式化表格。
    返回最后插入的段落
    """
    if not headers or len(headers) == 0:
        return None
    
    num_cols = len(headers)
    
    # 添加表标题（如果原始段落位置附近有）
    if original_para:
        # 在原始段落之前插入
        title_para = original_para.insert_paragraph_before(f"表{table_idx} {caption}" if caption else f"表{table_idx}")
    else:
        title_para = doc.add_paragraph(f"表{table_idx} {caption}" if caption else f"表{table_idx}")
    
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in title_para.runs:
        set_run_font(run, "宋体", 10, bold=True)
    
    # 创建表格
    num_rows = 1 + len(rows_data)
    t = doc.add_table(rows=num_rows, cols=num_cols)
    t.style = 'Table Grid'
    
    # 设置表格自动适应宽度
    try:
        tbl = t._tbl
        tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement('w:tblPr')
        tblW = OxmlElement('w:tblW')
        tblW.set(qn('w:w'), '5000')
        tblW.set(qn('w:type'), 'pct')
        tblPr.append(tblW)
    except:
        pass
    
    # 表头
    for j, h in enumerate(headers):
        cell = t.rows[0].cells[j]
        cell.text = h
        for para in cell.paragraphs:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in para.runs:
                set_run_font(run, "宋体", 9, bold=True)
    
    # 数据行
    for i, row_data in enumerate(rows_data):
        for j, val in enumerate(row_data):
            if j < num_cols:
                cell = t.rows[i + 1].cells[j]
                cell.text = val
                for para in cell.paragraphs:
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in para.runs:
                        set_run_font(run, "宋体", 9)
    
    # 表格后的空行
    last_p = doc.add_paragraph()
    last_p.paragraph_format.space_after = Pt(6)
    
    return last_p


# ==================== 在DOCX中查找表格位置 ====================
def find_table_in_docx(doc, chapter_num, section_num=None):
    """
    在DOCX中查找指定位置的表格。
    返回 (paragraph_index, 表格段落对象)
    注意：python-docx的表格是document级别存储的，需要根据文本位置推断。
    """
    # 找章节位置
    chapter_markers = []
    for i, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        # 匹配章标题 如 "第3章" 或 "3.2"
        if re.match(rf'^第{chapter_num}章', text) or re.match(rf'^{chapter_num}\.', text[:3]):
            chapter_markers.append(i)
    
    return chapter_markers


# ==================== 修复功能 ====================

def fix_abstract_extra_tag(doc, backup_doc=None):
    """
    修复1：删除摘要正文开头的"摘要："
    关键：只删除"摘要："这两个字及其后的冒号，保持后续文本的字体格式完全不变
    """
    changes = 0
    for i, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        # 找到"摘要："开头（不是单独成段的"摘要"）
        if not text.startswith('摘要：'):
            continue
        
        # 确认这不是一个单独的摘要标题
        prev_text = doc.paragraphs[i-1].text.strip() if i > 0 else ""
        if prev_text == '摘要' or prev_text == '摘  要':
            pass  # 这是摘要正文段
        elif text == '摘要：':
            continue  # 只有"摘要："没有内容，跳过
        
        # 需要删除"摘要："前缀，但保持格式不变
        # 先计算"摘要："的长度
        prefix = ''
        for tag in ['摘要：', '摘要:']:
            if text.startswith(tag):
                prefix = tag
                break
        
        if not prefix:
            continue
        
        new_text = text[len(prefix):].strip()
        
        # 修改段落中所有run的内容，移除开头的"摘要："
        full_text_before = ''.join(r.text for r in p.runs)
        
        if full_text_before.startswith(prefix):
            # 需要在run级别操作
            remaining = len(prefix)
            for run in p.runs:
                if not run.text:
                    continue
                if remaining <= 0:
                    break
                if len(run.text) <= remaining:
                    remaining -= len(run.text)
                    run.text = ''
                else:
                    run.text = run.text[remaining:]
                    remaining = 0
        
        changes += 1
    
    return changes


def fix_table_from_txt(doc, exam_id, chapter_section, docx_path=None):
    """
    修复2：从TXT提取表格数据重建DOCX表格。
    在原始DOCX中找到指定章节（如"3.2"或"1.4"）位置，
    用TXT中的表格数据替换该位置的1列格式错误表格。
    """
    # 找到对应的TXT文件
    # 搜索所有paper.txt
    txt_paths = list(MODIFIED_DIR.glob(f"*{exam_id}*/paper.txt"))
    if not txt_paths:
        # 从考试号前几位模糊匹配
        short_id = exam_id[-6:]
        txt_paths = list(MODIFIED_DIR.glob(f"*{short_id}*/paper.txt"))
    if not txt_paths:
        # 尝试从准考证号精确匹配目录
        for d in MODIFIED_DIR.iterdir():
            if d.is_dir() and exam_id in d.name:
                txt_path = d / "paper.txt"
                if txt_path.exists():
                    txt_paths = [txt_path]
                    break
    
    if not txt_paths:
        return f"找不到TXT文件 (exam_id={exam_id})"
    
    txt_path = txt_paths[0]
    
    # 读取TXT
    with open(txt_path, 'r') as f:
        txt_content = f.read()
    
    # 在TXT中提取所有表格
    all_tables = _extract_all_tables_from_txt(txt_content)
    
    if not all_tables:
        return "TXT中无表格数据"
    
    # 找到DOCX中目标章节位置
    chapter_num = chapter_section.split('.')[0]
    section_num = chapter_section.split('.')[1] if '.' in chapter_section else None
    
    # 在DOCX中找到该区域的第一个表格（1列格式的）
    doc_tables_to_replace = []
    for ti, table in enumerate(doc.tables):
        if len(table.columns) == 1 and len(table.rows) >= 3:
            doc_tables_to_replace.append(ti)
    
    if not doc_tables_to_replace:
        return f"章节{chapter_section}附近找不到1列表格"
    
    # 用TXT中的表格数据替换
    table_idx = 0
    replaced = 0
    for ti in doc_tables_to_replace:
        if table_idx < len(all_tables):
            caption, headers, rows_data = all_tables[table_idx]
            if not headers:
                table_idx += 1
                continue
            
            table = doc.tables[ti]
            
            # 清空并重建表格
            _rebuild_table(table, caption, headers, rows_data)
            replaced += 1
            table_idx += 1
    
    return f"替换了{replaced}个表格"


def _extract_all_tables_from_txt(txt_content):
    """从TXT中提取所有表格数据"""
    tables = []
    
    # 找所有<table>...</table>块
    # 注意：有些标签没有闭合的>，用<table或</table做边界
    raw_tables = re.split(r'(?=<table)', txt_content)
    
    for chunk in raw_tables:
        if not chunk.startswith('<table'):
            continue
        # 找到结束位置
        end_m = re.search(r'</?table[^>]*>', chunk[6:]) if len(chunk) > 6 else None
        if end_m:
            tag_content = chunk[:end_m.end() + 6]
        else:
            # 尝试用---PAGE_BREAK---或下一章标题做边界
            next_break = re.search(r'(?:---PAGE_BREAK|---|第\d+章)', chunk[6:])
            if next_break:
                tag_content = chunk[:next_break.start() + 6]
            else:
                tag_content = chunk
        
        caption, headers, rows_data = parse_html_table_to_rows(tag_content)
        if headers or rows_data:
            tables.append((caption, headers, rows_data))
    
    # 如果上面没找到，尝试直接用正则
    if not tables:
        for m in re.finditer(r'<table\b[^>]*>(.*?)</table\s*>', txt_content, re.DOTALL):
            tag_content = m.group()
            caption, headers, rows_data = parse_html_table_to_rows(tag_content)
            if headers or rows_data:
                tables.append((caption, headers, rows_data))
    
    # 再尝试不带>的结束标签
    if not tables:
        for m in re.finditer(r'<table\b[^>]*>(.*?)(?:</table\s*$|</table\s*\n)', txt_content, re.DOTALL):
            tag_content = m.group()
            caption, headers, rows_data = parse_html_table_to_rows(tag_content)
            if headers or rows_data:
                tables.append((caption, headers, rows_data))
    
    return tables


def _rebuild_table(table, caption, headers, rows_data):
    """重建一个已有表格的内容"""
    num_cols = len(headers)
    num_rows = 1 + len(rows_data)
    
    # 如果表格行列不够，增加行列
    while len(table.rows) < num_rows:
        table.add_row()
    
    while len(table.columns) < num_cols:
        # 无法直接添加列，需要通过底层操作
        pass
    
    # 设置表格属性
    try:
        tbl = table._tbl
        tblPr = tbl.tblPr
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr')
            tbl.insert(0, tblPr)
        tblW = tblPr.find(qn('w:tblW'))
        if tblW is None:
            tblW = OxmlElement('w:tblW')
            tblPr.append(tblW)
        tblW.set(qn('w:w'), '5000')
        tblW.set(qn('w:type'), 'pct')
    except:
        pass
    
    # 填充表头
    for j in range(min(num_cols, len(table.columns))):
        if j < len(headers):
            cell = table.rows[0].cells[j]
            cell.text = headers[j]
            for para in cell.paragraphs:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in para.runs:
                    set_run_font(run, "宋体", 9, bold=True)
    
    # 填充数据行
    for i, row_data in enumerate(rows_data):
        for j in range(min(num_cols, len(table.columns))):
            if j < len(row_data):
                cell = table.rows[i + 1].cells[j]
                cell.text = row_data[j]
                for para in cell.paragraphs:
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in para.runs:
                        set_run_font(run, "宋体", 9)


def fix_cover_info(doc, exam_id):
    """
    修复封面信息（学号、层次、专业等）
    """
    changes = []
    
    for i, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        
        # 修复学号
        if '学    号' in text or '学号' in text:
            old_text = ''.join(r.text for r in p.runs)
            # 替换为正确的学号
            new_text = re.sub(r'学\s*号[：:]\s*\S*', f'学    号：{exam_id}', old_text)
            if new_text != old_text:
                # 只改第一个run
                for run in p.runs:
                    if '学' in run.text:
                        run.text = new_text
                        break
                changes.append(f"学号 → {exam_id}")
        
        # 修复层次
        if '层    次' in text or '层次' in text:
            for run in p.runs:
                if '专升本' not in run.text and ('层' in run.text or '次' in run.text):
                    # 替换层次
                    if '本科' in run.text:
                        run.text = run.text.replace('本科', '专升本')
                        changes.append("层次: 本科→专升本")
                    elif '专科' in run.text:
                        run.text = run.text.replace('专科', '专升本')
                        changes.append("层次: 专科→专升本")
                    break
        
        # 修复专业
        if '专    业' in text or '专业' in text:
            if '环境设计' not in text:
                for run in p.runs:
                    if '环境' not in run.text:
                        run.text = re.sub(r'专业[：:]\s*\S+', '专业：环境设计', run.text)
                        changes.append("专业→环境设计")
                        break
    
    return changes


def fix_quotes(doc, text_corrections):
    """
    修复引号等问题
    text_corrections: [(old_text, new_text), ...]
    """
    changes = 0
    for p in doc.paragraphs:
        for old, new in text_corrections:
            if old in p.text:
                for run in p.runs:
                    if old in run.text:
                        run.text = run.text.replace(old, new)
                        changes += 1
    return changes


# ==================== 主处理流程 ====================

def process_paper(exam_id, problems, output_dir):
    """
    处理单篇论文
    """
    # 找到DOCX文件
    docx_name = None
    for f in os.listdir(HUANYI_DIR):
        if exam_id in f and f.endswith('.docx'):
            docx_name = f
            break
    
    if not docx_name:
        return {'status': 'skipped', 'reason': f'找不到文件 (exam_id={exam_id})'}
    
    src_path = HUANYI_DIR / docx_name
    dst_path = output_dir / docx_name
    
    # 读取DOCX
    doc = Document(str(src_path))
    
    report = {'status': 'processed', 'fixes': []}
    
    # 根据问题类型逐一修复
    problems_list = problems.split('\n') if problems else []
    
    for prob in problems_list:
        prob = prob.strip()
        if not prob:
            continue
        
        # 1. 删"摘要："
        if '摘要部分' in prob and ('多了' in prob or '摘要：' in prob):
            changes = fix_abstract_extra_tag(doc)
            report['fixes'].append(f"删摘要前缀: {changes}处")
        
        # 2. 修复表格
        table_match = re.search(r'"([\d.]+)"中的表格格式不对', prob)
        if not table_match:
            table_match = re.search(r'"([\d.]+)"中的表格内容不对', prob)
        if table_match:
            section = table_match.group(1)
            result = fix_table_from_txt(doc, exam_id, section)
            report['fixes'].append(f"修复表格({section}): {result}")
        
        # 3. 修复学号
        if '学号' in prob:
            m = re.search(r'应为["‘\u201c]?(\d+)["’\u201d]?', prob)
            correct_id = m.group(1) if m else exam_id
            changes = fix_cover_info(doc, correct_id)
            if changes:
                report['fixes'].append(f"修复封面: {', '.join(changes)}")
        
        # 4. 修复层次
        if '层次' in prob and '专升本' in prob:
            changes = fix_cover_info(doc, exam_id)
            if changes:
                report['fixes'].append(f"修复层次: {', '.join(changes)}")
        
        # 5. 修复引号
        quote_fixes = []
        for m in re.finditer(r"['\u2018](\S+)['\u2019]", prob):
            word = m.group(1)
            quote_fixes.append((f"'{word}'", f"「{word}」"))
            quote_fixes.append((f"'{word}'", f"「{word}」"))
        
        if '双引号' in prob or '引号' in prob:
            # 尝试自动匹配
            pass
        
        if quote_fixes:
            changes = fix_quotes(doc, quote_fixes)
            if changes:
                report['fixes'].append(f"引号修复: {changes}处")
        
        # 6. 正文格式不对
        if '正文格式不对' in prob or '正文格式不正确' in prob:
            report['fixes'].append("正文格式: 需人工确认")
    
    # 保存
    output_dir.mkdir(parents=True, exist_ok=True)
    doc.save(str(dst_path))
    
    report['output'] = str(dst_path)
    return report


def main():
    import openpyxl
    
    # 读取审核表
    wb = openpyxl.load_workbook(str(AUDIT_FILE))
    ws = wb['Sheet1']
    
    # 筛选低/无等级
    papers = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        name, exam_id_raw, filename, problems, level = row
        exam_id = str(exam_id_raw).strip() if exam_id_raw else ''
        if len(exam_id) < 12:
            exam_id = exam_id.zfill(12)
        
        level_str = str(level).strip() if level else 'None'
        if level_str not in ('低', 'None', 'None'):
            continue
        
        papers.append({
            'name': name,
            'exam_id': exam_id,
            'problems': problems or '',
            'level': level_str
        })
    
    print(f"需要处理的论文: {len(papers)}篇\n")
    
    # 如果指定了准考证号，只处理指定
    if len(sys.argv) > 1:
        target_ids = sys.argv[1:]
        papers = [p for p in papers if any(tid in p['exam_id'] for tid in target_ids)]
        print(f"筛选后: {len(papers)}篇\n")
    
    results = []
    
    for i, paper in enumerate(papers):
        name = paper['name']
        exam_id = paper['exam_id']
        problems = paper['problems']
        level = paper['level']
        
        print(f"[{i+1}/{len(papers)}] {name} ({exam_id}) [{level}]")
        print(f"  问题: {problems[:100]}...")
        
        result = process_paper(exam_id, problems, OUTPUT_DIR)
        
        if result['status'] == 'skipped':
            print(f"  ⏭ 跳过: {result['reason']}")
        else:
            print(f"  ✅ 处理完成")
            for fix in result['fixes']:
                print(f"    - {fix}")
            print(f"  📄 保存: {result['output']}")
        
        results.append(result)
        print()
    
    # 汇总
    success = sum(1 for r in results if r['status'] == 'processed')
    skipped = sum(1 for r in results if r['status'] == 'skipped')
    print(f"\n{'='*50}")
    print(f"完成！成功 {success}, 跳过 {skipped}")
    
    # 输出问题较多需要人工确认的
    manual = [r for r in results if any('人工确认' in f for f in r.get('fixes', []))]
    if manual:
        print(f"\n⚠️ {len(manual)}篇需要人工确认正文格式问题")


if __name__ == '__main__':
    main()
