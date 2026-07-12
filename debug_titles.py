#!/usr/bin/env python3
"""打印每个表格前后的段落（含旧标题），以便精确匹配。"""
import os, re
from docx import Document
from docx.oxml.ns import qn

MODIFIED_DIR = "/root/project/workspace/论文资料/整理/环艺_修改版"

files_to_fix = [
    "029822410501_贾燕霞_修改版.docx",
    "029823410134_高立凯_修改版.docx",
    "029823410119_何素珍_修改版.docx",
    "029823410164_臧斯恒_修改版.docx",
    "029823410190_许秀敏_修改版.docx",
    "029823410338_毕守荣_修改版.docx",
    "029823410272_王琛_修改版.docx",
    "029823410558_张颖_修改版.docx",
    "029821410170_李泽文_修改版.docx",
]

for fname in files_to_fix:
    path = os.path.join(MODIFIED_DIR, fname)
    doc = Document(path)
    tables = doc.tables
    name = fname.split('_')[1]
    
    print(f"\n{'='*60}")
    print(f"📄 {name} — {len(tables)} 个表格")
    print(f"{'='*60}")
    
    for t_idx, table in enumerate(tables):
        tbl_element = table._tbl
        parent = tbl_element.getparent()
        idx = list(parent).index(tbl_element)
        
        print(f"\n  表格#{t_idx+1} (idx={idx}):")
        
        # 前3个兄弟
        sib = tbl_element.getprevious()
        count = 0
        while sib is not None and count < 3:
            if sib.tag == qn('w:p'):
                texts = [r.text or '' for r in sib.findall('.//' + qn('w:t'))]
                ft = ''.join(texts).strip()
                if ft:
                    print(f"    ← [{ft[:100]}]")
            count += 1
            sib = sib.getprevious()
        
        # 后3个兄弟
        sib = tbl_element.getnext()
        count = 0
        while sib is not None and count < 3:
            if sib.tag == qn('w:p'):
                texts = [r.text or '' for r in sib.findall('.//' + qn('w:t'))]
                ft = ''.join(texts).strip()
                if ft:
                    print(f"    → [{ft[:100]}]")
            count += 1
            sib = sib.getnext()
