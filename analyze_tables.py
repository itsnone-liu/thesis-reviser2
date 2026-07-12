#!/usr/bin/env python3
"""分析无标题表格的上下文，确定合适的标题。"""
import os, re, json
from docx import Document
from docx.oxml.ns import qn

MODIFIED_DIR = "/root/project/workspace/论文资料/整理/环艺_修改版"

# 需要检查的无标题表格
targets = [
    ("029822410501_贾燕霞_修改版.docx", 0),   # 第1个表格
    ("029823410134_高立凯_修改版.docx", 0),
    ("029823410119_何素珍_修改版.docx", 0),
    ("029823410164_臧斯恒_修改版.docx", 0),
    ("029823410164_臧斯恒_修改版.docx", 1),
    ("029823410190_许秀敏_修改版.docx", 0),
    ("029823410338_毕守荣_修改版.docx", 0),
    ("029823410272_王琛_修改版.docx", 1),
    ("029823410558_张颖_修改版.docx", 0),
    ("029823410558_张颖_修改版.docx", 2),
]

for fname, t_idx in targets:
    path = os.path.join(MODIFIED_DIR, fname)
    doc = Document(path)
    tables = doc.tables
    
    if t_idx >= len(tables):
        print(f"\n⚠️ {fname}: 表格#{t_idx+1}不存在（共{len(tables)}个）")
        continue
    
    table = tables[t_idx]
    
    # 找表格前后段落的文本
    tbl_element = table._tbl
    body = tbl_element.getparent()
    
    prev_paras = []
    sib = tbl_element.getprevious()
    while sib is not None and len(prev_paras) < 3:
        if sib.tag == qn('w:p'):
            texts = [r.text or '' for r in sib.findall('.//' + qn('w:t'))]
            prev_paras.append(''.join(texts).strip())
        sib = sib.getprevious()
    
    next_paras = []
    sib = tbl_element.getnext()
    while sib is not None and len(next_paras) < 3:
        if sib.tag == qn('w:p'):
            texts = [r.text or '' for r in sib.findall('.//' + qn('w:t'))]
            next_paras.append(''.join(texts).strip())
        sib = sib.getnext()
    
    # 表格内容样本
    rows_data = []
    for r_idx, row in enumerate(table.rows):
        if r_idx > 5: break
        cells = [cell.text.strip()[:30] for cell in row.cells]
        rows_data.append(cells)
    
    name = fname.split('_')[1] if '_' in fname else fname[:15]
    print(f"\n{'='*60}")
    print(f"📄 {name} — 表格#{t_idx+1} ({len(table.rows)}行×{len(table.columns)}列)")
    print(f"{'='*60}")
    print(f"【前文段落】:")
    for p in reversed(prev_paras):
        print(f"  {p[:100]}")
    print(f"【表格内容】:")
    for r in rows_data:
        print(f"  {r}")
    print(f"【后文段落】:")
    for p in next_paras:
        print(f"  {p[:100]}")
