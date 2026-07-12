#!/usr/bin/env python3
"""清理重复标题和旧标题（表格后的标题）。"""
import os, re
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

MODIFIED_DIR = "/root/project/workspace/论文资料/整理/环艺_修改版"

files_to_fix = [
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

# 表格后应删除的旧标题文本（关键词）
# 这些是原始文档中表格后面的文字标题（无表序号或旧序号）
OLD_TITLES_BEFORE = [
    "表4.1 问题类型、实例与改进方向对应关系",
    "表3.1 设计策略与问题对应表",
    "理论-设计要素映射表",
    "理论与设计关联对照表",
    "问题-策略映射表",
    "设计转译模型表",
    "问题-需求对照表",
    "测点照度实测数据",
    "设计成效对照表",
    "案例生态主题表现对比表",
    "国内外案例对比分析表",
    "对比表格如下",
    "问题-策略映射表(设计阶段使用)",
    "设计转译模型可用以下表格概括",
    "对比分析如下表",
    "问题-需求对照表",
    "对照关系如下",
    "对照表如下",
    "对照表格如下",
    "映射关系如下",
    "如下表所示",
    "如下表",
    "如下图所示",
    "如下图",
    "映射表",
]

def is_header_or_title_para(text):
    """检查段落是否可能是标题/序号段落（要保留的）。"""
    text = text.strip()
    if not text:
        return False
    # 带"表N"的标题 — 已修复过的要保留
    if re.search(r'表\s*\d+\s', text):
        return True
    # 可能的大标题
    if re.match(r'^第\d+章\s', text):
        return True
    if re.match(r'^[一二三四五六七八九十、]', text):
        return True
    return False

for fname in files_to_fix:
    path = os.path.join(MODIFIED_DIR, fname)
    doc = Document(path)
    tables = doc.tables
    name = fname.split('_')[1]
    
    # 直接操作XML：对每个表格，删除后面的旧标题
    modified = False
    
    for t_idx, table in enumerate(tables):
        tbl_element = table._tbl
        parent = tbl_element.getparent()
        idx_in_parent = list(parent).index(tbl_element)
        
        # 检查表格后面有没有旧标题需要删除
        # 收集后面紧跟的段落
        sib = tbl_element.getnext()
        paras_to_delete = []
        while sib is not None:
            if sib.tag == qn('w:p'):
                texts = [r.text or '' for r in sib.findall('.//' + qn('w:t'))]
                full_text = ''.join(texts).strip()
                
                # 如果是旧标题（无表序号但匹配关键词，或者有旧序号）
                is_old = False
                for old in OLD_TITLES_BEFORE:
                    if old in full_text:
                        is_old = True
                        break
                
                # 检查是否是"表4.1""表3.1"这种旧序号（新序号应该是表1、表2、表3...）
                if re.match(r'表\s*[\d.]+\s', full_text) and not re.match(r'表\s*\d\s', full_text):
                    # 非连续编号（含小数点），如"表4.1""表3.1"
                    is_old = True
                
                # 只有"表"无序号且匹配关键词的
                if re.match(r'表\s*[^\d]', full_text[:10]) and len(full_text) < 100:
                    # 去掉"表"本身
                    rest = re.sub(r'^表\s*', '', full_text)
                    for old in OLD_TITLES_BEFORE:
                        if old[:10] in full_text or full_text[:15] in old:
                            is_old = True
                            break
                
                if is_old:
                    paras_to_delete.append(sib)
                    sib = sib.getnext()
                else:
                    break
            else:
                # 遇到下一个非段落元素（比如另一个表格）就停止
                break
        
        # 删除这些段落
        for p in paras_to_delete:
            parent.remove(p)
            modified = True
        
        if paras_to_delete:
            print(f"  ✅ {name}: 表格#{t_idx+1} → 删除了 {len(paras_to_delete)} 个旧标题段落")
        
        # 再检查表格前面有没有重复的提示文本（不带序号的）
        # 比如"理论-设计要素映射表"（无表X字样）
        sib = tbl_element.getprevious()
        paras_before = []
        while sib is not None:
            if sib.tag == qn('w:p'):
                texts = [r.text or '' for r in sib.findall('.//' + qn('w:t'))]
                full_text = ''.join(texts).strip()
                
                # 如果是表格前的提示文字（无表序号但内容与表格标题重复）
                # 比如"对比分析如下表"、"对照表格如下"
                is_hint = False
                for hint in ["对比分析如下表", "对照表格如下", "对比表格如下", "对照关系如下", "对照表如下", "如下表所示", "如下表"]:
                    if hint in full_text:
                        is_hint = True
                        break
                
                if is_hint:
                    paras_before.append(sib)
                    sib = sib.getprevious()
                else:
                    break
            else:
                break
        
        for p in paras_before:
            parent.remove(p)
            modified = True
        
        if paras_before:
            print(f"  ✅ {name}: 表格#{t_idx+1} → 删除了 {len(paras_before)} 个提示文字段落")
    
    if modified:
        doc.save(path)
        print(f"  💾 {name}: 已保存")

print(f"\n✅ 所有清理完成")
