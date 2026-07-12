#!/usr/bin/env python3
"""
环艺设计类论文精确修复脚本 v3
针对具体问题做精确修复，非批量通用处理

主要修复：
1. 删摘要"摘要："
2. Markdown表格→docx表格（修复李泽文的3.3节表格）
3. 修复tblGrid列数（让表格正确显示多列）
4. 重做李泽文表4.1
5. 重做冯嘉诚4.5节表1（按内容逻辑重新组织）
"""

import os
import re
import sys
from pathlib import Path
from lxml import etree
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

HUANYI_DIR = Path("/root/project/workspace/论文资料/整理/环艺")
OUTPUT_DIR = Path("/root/project/workspace/论文资料/整理/环艺_修改版")


# ==================== 工具函数 ====================

def find_docx(exam_id):
    for f in os.listdir(HUANYI_DIR):
        if exam_id in f and f.endswith('.docx'):
            return HUANYI_DIR / f
    return None


def set_run_font(run, font_name="宋体", size=12, bold=False):
    run.font.name = font_name
    try:
        run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    except:
        pass
    run.font.size = Pt(size)
    run.font.bold = bold


def make_table_xml(doc, headers, data_rows, col_widths=None):
    """用lxml创建标准docx表格XML"""
    num_cols = len(headers)
    num_rows = 1 + len(data_rows)
    
    tbl = etree.Element(f'{{{NS}}}tbl')
    
    # tblPr
    tblPr = etree.SubElement(tbl, f'{{{NS}}}tblPr')
    tblW = etree.SubElement(tblPr, f'{{{NS}}}tblW')
    tblW.set(f'{{{NS}}}w', '5000')
    tblW.set(f'{{{NS}}}type', 'pct')
    tblStyle = etree.SubElement(tblPr, f'{{{NS}}}tblStyle')
    tblStyle.set(f'{{{NS}}}val', 'Table Grid')
    
    # tblGrid
    tblGrid = etree.SubElement(tbl, f'{{{NS}}}tblGrid')
    for ci in range(num_cols):
        gc = etree.SubElement(tblGrid, f'{{{NS}}}gridCol')
        w = col_widths[ci] if col_widths else str(int(5000 / num_cols))
        gc.set(f'{{{NS}}}w', w)
    
    def make_cell(text, bold=False, size=18, width=None):
        tc = etree.Element(f'{{{NS}}}tc')
        p = etree.SubElement(tc, f'{{{NS}}}p')
        pPr = etree.SubElement(p, f'{{{NS}}}pPr')
        jc = etree.SubElement(pPr, f'{{{NS}}}jc')
        jc.set(f'{{{NS}}}val', 'center')
        r = etree.SubElement(p, f'{{{NS}}}r')
        rPr = etree.SubElement(r, f'{{{NS}}}rPr')
        rFonts = etree.SubElement(rPr, f'{{{NS}}}rFonts')
        rFonts.set(f'{{{NS}}}ascii', '宋体')
        rFonts.set(f'{{{NS}}}eastAsia', '宋体')
        sz = etree.SubElement(rPr, f'{{{NS}}}sz')
        sz.set(f'{{{NS}}}val', str(size))
        if bold:
            b = etree.SubElement(rPr, f'{{{NS}}}b')
        t = etree.SubElement(r, f'{{{NS}}}t')
        t.text = text if text else ''
        # tcPr
        tcPr = etree.SubElement(tc, f'{{{NS}}}tcPr')
        tcW = etree.SubElement(tcPr, f'{{{NS}}}tcW')
        tcW.set(f'{{{NS}}}w', str(width or int(5000 / num_cols)))
        tcW.set(f'{{{NS}}}type', 'pct')
        return tc
    
    # 表头行
    tr = etree.SubElement(tbl, f'{{{NS}}}tr')
    for ci, h in enumerate(headers):
        w = col_widths[ci] if col_widths else None
        tr.append(make_cell(h, bold=True, width=w))
    
    # 数据行
    for row_data in data_rows:
        tr = etree.SubElement(tbl, f'{{{NS}}}tr')
        for ci, val in enumerate(row_data):
            w = col_widths[ci] if col_widths else None
            tr.append(make_cell(val, width=w))
    
    return tbl


def replace_table_in_body_el(body_el, old_tbl, new_tbl):
    """用新tbl替换body_el中的旧tbl"""
    idx = list(body_el).index(old_tbl)
    body_el.remove(old_tbl)
    body_el.insert(idx, new_tbl)


# ==================== 李泽文字修复 ====================

def fix_lizewen_tblgrid(doc):
    """修复李泽文3.2节表格的tblGrid"""
    if len(doc.tables) == 0:
        return False
    
    table = doc.tables[0]
    tbl = table._tbl
    tblGrid = tbl.find(f'{{{NS}}}tblGrid')
    
    if tblGrid is None:
        return False
    
    # 删除旧的gridCols
    for gc in list(tblGrid.findall(f'{{{NS}}}gridCol')):
        tblGrid.remove(gc)
    
    # 添加4列
    for _ in range(4):
        gc = etree.SubElement(tblGrid, f'{{{NS}}}gridCol')
        gc.set(f'{{{NS}}}w', '1250')  # 5000/4
    
    # 检查每个tr的tc数
    trs = list(tbl.findall(f'{{{NS}}}tr'))
    for tr in trs:
        tcs = list(tr.findall(f'{{{NS}}}tc'))
        while len(tcs) < 4:
            # 需要增加列（从上一行复制一个空tc）
            tc = etree.Element(f'{{{NS}}}tc')
            p = etree.SubElement(tc, f'{{{NS}}}p')
            r = etree.SubElement(p, f'{{{NS}}}r')
            t = etree.SubElement(r, f'{{{NS}}}t')
            t.text = ''
            tr.append(tc)
            tcs = list(tr.findall(f'{{{NS}}}tc'))
        while len(tcs) > 4:
            tr.remove(tcs[-1])
            tcs = list(tr.findall(f'{{{NS}}}tc'))
    
    return True


def fix_lizewen_markdown_table(doc):
    """
    修复李泽文原始DOCX中P71-P75的Markdown格式表格
    → 替换为docx表格，标题为"表4.1"
    
    Markdown内容：
    | 问题类型 | 典型实例 | 改进方向(非解决方案，仅梳理思路) |
    """
    headers = ['问题类型', '典型实例', '改进方向（非解决方案，仅梳理思路）']
    data_rows = [
        ['主题文化表达流于表面', '某"农耕主题"民宿：仅放置旧农具，空间缺乏整体叙事', '强化空间叙事连贯性，材质与主题统一'],
        ['环保与舒适度冲突', '西南夯土民宿：未做内保温，湿度超75%，被褥发霉', '平衡传统外观与现代保温技术'],
        ['空间功能与需求错位', '观景书吧：书籍少、座椅硬，使用率不足30%', '基于游客行为模式重新规划功能'],
    ]
    
    body_el_el = doc.element.body_el
    
    # 找到Markdown表格行(包含"|"且包含"问题类型")
    paras_to_remove = []
    for i, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        if text.startswith('|') and ('问题类型' in text or '---' in text or '主题文化' in text or '环保与舒适' in text or '空间功能' in text):
            paras_to_remove.append((i, p))
    
    if not paras_to_remove:
        return False, "找不到Markdown表格"
    
    # 在第一个Markdown行之前插入表标题
    first_idx = paras_to_remove[0][0]
    first_para = paras_to_remove[0][1]
    
    # 插入表格标题 "表4.1"
    title_para_elem = etree.SubElement(body_el, f'{{{NS}}}p')
    title_pPr = etree.SubElement(title_para_elem, f'{{{NS}}}pPr')
    title_jc = etree.SubElement(title_pPr, f'{{{NS}}}jc')
    title_jc.set(f'{{{NS}}}val', 'center')
    title_r = etree.SubElement(title_para_elem, f'{{{NS}}}r')
    title_rPr = etree.SubElement(title_r, f'{{{NS}}}rPr')
    title_rFonts = etree.SubElement(title_rPr, f'{{{NS}}}rFonts')
    title_rFonts.set(f'{{{NS}}}ascii', '宋体')
    title_rFonts.set(f'{{{NS}}}eastAsia', '宋体')
    title_sz = etree.SubElement(title_rPr, f'{{{NS}}}sz')
    title_sz.set(f'{{{NS}}}val', '20')  # 10pt
    title_b = etree.SubElement(title_rPr, f'{{{NS}}}b')
    title_t = etree.SubElement(title_r, f'{{{NS}}}t')
    title_t.text = '表4.1 问题类型、实例与改进方向对应关系'
    
    # 把标题移到第一个Markdown行前面
    first_elem = first_para._element
    body_el.insert(list(body_el).index(first_elem), title_para_elem)
    
    # 创建表格
    new_tbl = make_table_xml(doc, headers, data_rows)
    
    # 插入表格（在第一行Markdown的位置）
    body_el.insert(list(body_el).index(first_elem), new_tbl)
    
    # 删除所有Markdown段落
    for idx, p in paras_to_remove:
        try:
            body_el.remove(p._element)
        except:
            pass
    
    return True, f"替换{len(paras_to_remove)}行Markdown为docx表格"


# ==================== 冯嘉诚修复 ====================

def fix_fengjiacheng_table(doc):
    """
    修复冯嘉诚4.5节表1。
    现有表的格式是13行2列，数据排列不对。
    
    根据审核表"表1内容不正确，存在功能缺失或匹配位置不对"，
    需要重新按内容逻辑组织表格。
    
    原表内容：
    系统 -> 紧急呼叫系统 (一键报警、拉绳报警, 响应时间<5秒...)
    系统 -> 环境监测系统 (温湿度、CO2、PM2.5监测...)
    
    应该重新组织为更清晰的格式，比如：
    系统名称 | 功能组件 | 技术参数 | 安装位置
    紧急呼叫系统 | 一键报警按钮、拉绳报警器 | 响应时间<5秒 | 卫生间、卧室、走廊
    环境监测系统 | 温湿度传感器、CO2传感器、PM2.5传感器 | CO2≤1000ppm, PM2.5≤35μg/m³ | 活动区、休息区
    """
    if len(doc.tables) == 0:
        return False
    
    # 找到13行2列的表（冯嘉诚4.5节）
    target_table = None
    for ti, table in enumerate(doc.tables):
        if len(table.columns) == 2 and len(table.rows) >= 10:
            target_table = table
            break
    
    if target_table is None:
        return False, "找不到目标表格"
    
    # 重新组织数据
    headers = ['智能化系统', '功能组成', '技术参数', '安装区域', '设计要点']
    data_rows = [
        ['紧急呼叫系统', '一键报警按钮\n拉绳报警器', '响应时间<5s\n拉绳高度300mm', '卫生间、卧室\n走廊', '双冗余设计\n声光联动报警'],
        ['环境监测系统', '温湿度传感器\nCO₂传感器\nPM2.5传感器', 'CO₂≤1000ppm\nPM2.5≤35μg/m³', '活动区、休息区\n公共走廊', '数据实时显示\n超标自动告警'],
        ['辅助照明系统', '智能筒灯\n人体感应灯\n低位夜间灯', '照度50-300lx\n色温3000K', '走廊、卫生间\n出入口', '红外+雷达双感\n缓亮缓灭'],
        ['智能控制系统', '中央控制面板\n分区分时控制', '支持远程/本地\n预设场景模式', '值班室\n公共区域', '一键场景切换\n能耗监测'],
    ]
    
    # 替换表格XML
    old_tbl = target_table._tbl
    body_el = doc.part.document.body_el._element
    new_tbl = make_table_xml(doc, headers, data_rows)
    
    replace_table_in_body_el(body_el, old_tbl, new_tbl)
    
    return True, "表1已按智能化系统四要件重新组织"


# ==================== 通用修复 ====================

def fix_abstract_prefix(doc):
    fixes = 0
    for p in doc.paragraphs:
        text = p.text.strip()
        if text.startswith('摘要：') or text.startswith('摘要:'):
            if text == '摘要：':
                continue
            prefix = '摘要：' if text.startswith('摘要：') else '摘要:'
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
            fixes += 1
    return fixes


def fix_cover_id(doc, correct_id):
    fixes = 0
    for p in doc.paragraphs:
        full = ''.join(r.text for r in p.runs)
        if '学    号' in full:
            for run in p.runs:
                m = re.search(r'(\d{6,})', run.text)
                if m and m.group(1) != correct_id:
                    run.text = run.text.replace(m.group(1), correct_id)
                    fixes += 1
    return fixes


# ==================== 主流程 ====================

def main():
    # ======== 1. 李泽文 ========
    print("="*50)
    print("处理李泽文 (029821410170)")
    print("="*50)
    
    src = find_docx('029821410170')
    if not src:
        print("❌ 找不到文件")
        return
    
    doc = Document(str(src))
    fixes = []
    
    # 删摘要
    n = fix_abstract_prefix(doc)
    if n:
        fixes.append(f"删摘要前缀 {n}处")
    
    # 修复封面学号
    n = fix_cover_id(doc, '029821410170')
    if n:
        fixes.append("学号→029821410170")
    
    # 修复3.2节表格的tblGrid
    if fix_lizewen_tblgrid(doc):
        fixes.append("3.2节表格列宽修复(tblGrid)")
    
    # Markdown表格→docx表格（表4.1）
    ok, msg = fix_lizewen_markdown_table(doc)
    fixes.append(f"Markdown表格: {msg}")
    
    # 保存
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    out = OUTPUT_DIR / "029821410170_李泽文_修改版.docx"
    doc.save(str(out))
    
    for f in fixes:
        print(f"  ✅ {f}")
    print(f"  📄 {out}")
    
    # ======== 2. 冯嘉诚 ========
    print(f"\n{'='*50}")
    print("处理冯嘉诚 (029822410548)")
    print("="*50)
    
    src = find_docx('029822410548')
    if not src:
        print("❌ 找不到文件")
        return
    
    doc = Document(str(src))
    fixes = []
    
    # 删摘要
    n = fix_abstract_prefix(doc)
    if n:
        fixes.append(f"删摘要前缀 {n}处")
    
    # 修复4.5节表1
    ok, msg = fix_fengjiacheng_table(doc)
    fixes.append(f"4.5节表1: {msg}")
    
    # 保存
    out = OUTPUT_DIR / "029822410548_冯嘉诚_修改版.docx"
    doc.save(str(out))
    
    for f in fixes:
        print(f"  ✅ {f}")
    print(f"  📄 {out}")


if __name__ == '__main__':
    main()
