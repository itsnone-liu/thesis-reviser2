#!/usr/bin/env python3
"""验证修复后的表格标题情况。"""
import os, re
from docx import Document
from docx.oxml.ns import qn

MODIFIED_DIR = "/root/project/workspace/论文资料/整理/环艺_修改版"

targets = [
    "029821410170_李泽文_修改版.docx",
    "029822410501_贾燕霞_修改版.docx",
    "029823410134_高立凯_修改版.docx",
    "029823410119_何素珍_修改版.docx",
    "029823410164_臧斯恒_修改版.docx",
    "029823410190_许秀敏_修改版.docx",
    "029823410338_毕守荣_修改版.docx",
    "029823410272_王琛_修改版.docx",
    "029823410558_张颖_修改版.docx",
]

for fname in targets:
    path = os.path.join(MODIFIED_DIR, fname)
    doc = Document(path)
    tables = doc.tables
    name = fname.split('_')[1]
    
    print(f"\n{'='*60}")
    print(f"📄 {name} — 共 {len(tables)} 个表格")
    print(f"{'='*60}")
    
    for t_idx, table in enumerate(tables):
        tbl_element = table._tbl
        parent = tbl_element.getparent()
        
        # 检查前后的标题
        titles_found = []
        sib = tbl_element.getprevious()
        while sib is not None:
            if sib.tag == qn('w:p'):
                texts = [r.text or '' for r in sib.findall('.//' + qn('w:t'))]
                full_text = ''.join(texts).strip()
                if re.search(r'表\s*\d', full_text[:30]):
                    titles_found.append(('前', full_text[:80]))
            sib = sib.getprevious()
        
        sib = tbl_element.getnext()
        while sib is not None:
            if sib.tag == qn('w:p'):
                texts = [r.text or '' for r in sib.findall('.//' + qn('w:t'))]
                full_text = ''.join(texts).strip()
                if re.search(r'表\s*\d', full_text[:30]):
                    titles_found.append(('后', full_text[:80]))
            sib = sib.getnext()
        
        rows_info = f"{len(table.rows)}行×{len(table.columns)}列"
        print(f"  表格#{t_idx+1} ({rows_info})")
        if titles_found:
            for pos, t in titles_found:
                print(f"    [{pos}] {t}")
        else:
            print(f"    ⚠️ 无标题")
