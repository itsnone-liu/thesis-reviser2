#!/usr/bin/env python3
"""精确验证：每个表格前后最近的带"表"的段落。"""
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

all_ok = True

for fname in targets:
    path = os.path.join(MODIFIED_DIR, fname)
    doc = Document(path)
    tables = doc.tables
    name = fname.split('_')[1]
    
    print(f"\n📄 {name} — {len(tables)} 个表格")
    
    for t_idx, table in enumerate(tables):
        tbl_element = table._tbl
        parent = tbl_element.getparent()
        
        # 仅检查紧邻的前一个段落
        prev_sib = tbl_element.getprevious()
        immediate_before = None
        while prev_sib is not None:
            if prev_sib.tag == qn('w:p'):
                texts = [r.text or '' for r in prev_sib.findall('.//' + qn('w:t'))]
                ft = ''.join(texts).strip()
                if ft:
                    immediate_before = ft
                    break
            prev_sib = prev_sib.getprevious()
        
        # 仅检查紧邻的后一个段落
        next_sib = tbl_element.getnext()
        immediate_after = None
        while next_sib is not None:
            if next_sib.tag == qn('w:p'):
                texts = [r.text or '' for r in next_sib.findall('.//' + qn('w:t'))]
                ft = ''.join(texts).strip()
                if ft:
                    immediate_after = ft
                    break
            next_sib = next_sib.getnext()
        
        # 检查：前面是否有表N标题
        has_title_before = bool(immediate_before and re.search(r'^表\s*\d+\s', immediate_before[:30]))
        has_title_after = bool(immediate_after and re.search(r'表\s*\d', immediate_after[:30]))
        
        rows_cols = f"{len(table.rows)}行×{len(table.columns)}列"
        status = "✅" if has_title_before else "⚠️"
        
        print(f"  {status} 表格#{t_idx+1} ({rows_cols})")
        if has_title_before:
            print(f"     ← [{immediate_before[:60]}]")
        else:
            print(f"     ← [无表标题: {immediate_before[:50] if immediate_before else '空'}]")
        if has_title_after:
            print(f"     → ⚠️ [后面有表标题: {immediate_after[:60]}]")
            all_ok = False
        
        # 检查表格后面有没有重复表标题
        sib = tbl_element.getnext()
        count = 0
        while sib is not None and count < 2:
            if sib.tag == qn('w:p'):
                texts = [r.text or '' for r in sib.findall('.//' + qn('w:t'))]
                ft = ''.join(texts).strip()
                if ft and re.search(r'^表\s*\d+\s', ft[:30]):
                    print(f"     → ⚠️ 有重复标题: [{ft[:60]}]")
                    all_ok = False
            count += 1
            sib = sib.getnext()

if all_ok:
    print(f"\n✅ 全部检查通过！所有表格都有正确标题，无重复。")
else:
    print(f"\n⚠️ 部分表格仍有问题，见上。")
