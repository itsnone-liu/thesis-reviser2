#!/usr/bin/env python3
"""全面检查所有35篇论文的表格序号。"""
import os, re
from docx import Document
from docx.oxml.ns import qn

MODIFIED_DIR = "/root/project/workspace/论文资料/整理/环艺_修改版"
files = sorted([f for f in os.listdir(MODIFIED_DIR) if f.endswith('.docx')])

all_good = True

for fname in files:
    path = os.path.join(MODIFIED_DIR, fname)
    doc = Document(path)
    tables = doc.tables
    name = fname.split('_')[1]
    
    issues = []
    
    for t_idx, table in enumerate(tables):
        tbl_element = table._tbl
        
        # 找前面最近的标题
        sib = tbl_element.getprevious()
        immediate_before = None
        while sib is not None:
            if sib.tag == qn('w:p'):
                texts = [r.text or '' for r in sib.findall('.//' + qn('w:t'))]
                ft = ''.join(texts).strip()
                if ft:
                    immediate_before = ft
                    break
            sib = sib.getprevious()
        
        # 检查后面是否有重复标题
        sib = tbl_element.getnext()
        repeat_title = None
        while sib is not None:
            if sib.tag == qn('w:p'):
                texts = [r.text or '' for r in sib.findall('.//' + qn('w:t'))]
                ft = ''.join(texts).strip()
                if ft and re.search(r'^表\s*\d+\s', ft[:30]):
                    repeat_title = ft[:60]
                    break
                if ft:  # 有文字但不是表标题，停止
                    break
            sib = sib.getnext()
        
        has_title = immediate_before and re.search(r'^表\s*\d+\s', immediate_before[:30])
        
        if not has_title:
            issues.append(f"表格#{t_idx+1}: ⚠️ 无标题")
        elif repeat_title:
            issues.append(f"表格#{t_idx+1}: ⚠️ 后面有重复标题 '{repeat_title}'")
        
        # 多表格检查序号连续性
        if len(tables) > 1 and has_title:
            num_match = re.search(r'^表\s*(\d+)', immediate_before[:30])
            if num_match:
                num = int(num_match.group(1))
                if num != t_idx + 1:
                    issues.append(f"表格#{t_idx+1}: ⚠️ 序号={num}，期望={t_idx+1}")
    
    # 单表格的特殊检查 - 确保标题是表1
    if len(tables) == 1:
        tbl_element = tables[0]._tbl
        sib = tbl_element.getprevious()
        immediate_before = None
        while sib is not None:
            if sib.tag == qn('w:p'):
                texts = [r.text or '' for r in sib.findall('.//' + qn('w:t'))]
                ft = ''.join(texts).strip()
                if ft:
                    immediate_before = ft
                    break
            sib = sib.getprevious()
        
        if immediate_before:
            m = re.search(r'^表\s*(\d+)', immediate_before[:30])
            if m and int(m.group(1)) != 1:
                issues.append(f"单表格: ⚠️ 序号不是表1，而是表{m.group(1)}")
    
    status = "✅" if not issues else "⚠️"
    if issues:
        all_good = False
        print(f"{status} {name} ({len(tables)}表):")
        for iss in issues:
            print(f"   {iss}")
    # else: print(f"{status} {name} ({len(tables)}表)")

if all_good:
    print(f"✅ 全部35篇论文的表格序号和标题完全正确！")
else:
    print(f"\n⚠️ 仍有问题需要修复")
