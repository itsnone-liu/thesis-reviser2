#!/usr/bin/env python3
"""
从备份DOCX复制封面到修复后的DOCX。
封面定义为：从文档开头到"摘要"或"目录"之间的所有内容（含分页符、分节符）。
使用python-docx操作，严格按段落复制（保留格式但部分域代码可能丢失）。
"""

import sys
import os
from pathlib import Path
import shutil
from docx import Document
from docx.oxml.ns import qn
import re

BACKUP_DIR = Path("/root/project/workspace/论文资料/整理/论文终版_备份_修复前")
FINAL_DIR = Path("/root/project/workspace/论文资料/整理/论文终版")

def get_covers_from_docx(docx_path):
    """
    提取封面段落：从文档开头到"摘要"或"目录"所在的段落之前。
    返回段落列表(paragraph objects)。
    """
    doc = Document(str(docx_path))
    cover_paras = []
    found_end = False
    
    for p in doc.paragraphs:
        text = p.text.strip()
        # 遇到摘要或目录（一级标题标记）就停止
        if text in ("摘要", "摘  要", "目录", "目  录", "Abstract"):
            # 包含标题标记本身也属于封面？不包含，只包含之前的内容
            found_end = True
            break
        cover_paras.append(p)
    
    # 如果没有找到摘要/目录，说明可能目录在第一页
    # 保守策略：只取前3页的内容（约50-80段）
    if not found_end and len(cover_paras) > 0:
        cover_paras = cover_paras[:80]
    
    return cover_paras

def get_cover_element_tree(cover_paras):
    """
    获取封面段落的XML元素树（用于替换）。
    返回最早的p元素之前的兄弟节点+所有p元素。
    """
    if not cover_paras:
        return []
    
    # 获取第一个段落之前的所有兄弟节点（包括sectPr分节符等）
    first_p_elem = cover_paras[0]._element
    parent = first_p_elem.getparent()
    
    # 找到第一个p元素的索引
    idx = list(parent).index(first_p_elem)
    
    # 获取封面范围：从body开始到第一个p之前的内容（分节符、段落属性等）
    # 以及所有cover段落
    preceding = list(parent)[:idx]  # 封面之前可能有的内容
    cover_elems = list(parent)[idx:idx + len(cover_paras)]
    
    return preceding, cover_elems

def backup_current_file(docx_path):
    """备份当前文件（如果还没备份过封面恢复的版本）"""
    bak_path = docx_path.with_suffix('.docx.covers_bak')
    if not bak_path.exists():
        shutil.copy2(str(docx_path), str(bak_path))
        return bak_path
    return bak_path

def get_body_xml_structure(doc):
    """获取文档的body元素结构"""
    body = doc.part.document.body
    return body

def restore_covers_from_lxml(backup_docx, target_docx):
    """
    使用lxml直接替换body子元素来实现封面替换。
    更精确，保留封面格式。
    """
    from lxml import etree
    
    # 解析两个文档
    backup_doc = Document(str(backup_docx))
    target_doc = Document(str(target_docx))
    
    backup_body = backup_doc.part.document.body
    target_body = target_doc.part.document.body
    
    backup_children = list(backup_body._element)
    target_children = list(target_body._element)
    
    # 在备份文档中找到封面结束位置（"摘要"或"目录"）
    backup_cover_end = None
    for i, child in enumerate(backup_children):
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'p':
            text = ''.join(node.text or '' for node in child.iter() if node.tag.endswith('}t'))
            text = text.strip()
            if text in ("摘要", "摘  要", "目录", "目  录", "Abstract"):
                backup_cover_end = i
                break
    
    if backup_cover_end is None:
        # 找不到，保守取前80个元素
        backup_cover_end = min(80, len(backup_children))
    
    # 在目标文档中找到相同的结束位置
    target_cover_end = None
    for i, child in enumerate(target_children):
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'p':
            text = ''.join(node.text or '' for node in child.iter() if node.tag.endswith('}t'))
            text = text.strip()
            if text in ("摘要", "摘  要", "目录", "目  录", "Abstract"):
                target_cover_end = i
                break
    
    if target_cover_end is None:
        target_cover_end = backup_cover_end
    
    # 获取封面元素（备份文档的前N个元素）
    cover_elements = backup_children[:backup_cover_end]
    
    # 从备份中复制sectPr（分节符属性）—— 可能在第一个p之前，也可能在body末尾
    # sectPr在封面和正文之间通常用于分节，需要保留目标文档的sectPr
    # 获取目标文档中封面后的第一个sectPr
    target_sectPr = None
    for child in target_children[target_cover_end:]:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'sectPr':
            target_sectPr = child
            break
    
    # 从备份中获取封面的sectPr（如果有）
    backup_cover_sectPr = None
    for child in cover_elements:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'sectPr':
            backup_cover_sectPr = child
            break
    
    # 移除目标文档中的封面元素
    for child in target_children[:target_cover_end]:
        target_body._element.remove(child)
    
    # 在目标文档的body开头插入封面元素（逆序插入以保持顺序）
    for elem in reversed(cover_elements):
        target_body._element.insert(0, elem)
    
    # 保存
    target_doc.save(str(target_docx))
    return backup_cover_end, target_cover_end


def restore_covers_simple(backup_docx, target_docx):
    """
    简单方法：直接把备份DOCX的封面段落内容逐段复制到目标DOCX的封面位置。
    更安全但可能丢失复杂格式（表格、文本框等）。
    """
    backup_doc = Document(str(backup_docx))
    target_doc = Document(str(target_docx))
    
    # 找到备份和目标的封面边界
    backup_cover_end = None
    for i, p in enumerate(backup_doc.paragraphs):
        text = p.text.strip()
        if text in ("摘要", "摘  要", "目录", "目  录", "Abstract"):
            backup_cover_end = i
            break
    if backup_cover_end is None:
        backup_cover_end = min(80, len(backup_doc.paragraphs))
    
    target_cover_end = None
    for i, p in enumerate(target_doc.paragraphs):
        text = p.text.strip()
        if text in ("摘要", "摘  要", "目录", "目  录", "Abstract"):
            target_cover_end = i
            break
    if target_cover_end is None:
        target_cover_end = backup_cover_end
    
    # 确保边界有效
    if target_cover_end > len(target_doc.paragraphs):
        target_cover_end = len(target_doc.paragraphs)
    
    print(f"  备份封面段落: {backup_cover_end}段, 目标封面段落: {target_cover_end}段")
    
    # 提取备份封面的段落元素XML
    backup_cover_paras = list(backup_doc.paragraphs)[:backup_cover_end]
    
    # 获取目标文档中要替换的段落
    target_cover_paras = list(target_doc.paragraphs)[:target_cover_end]
    
    # 获取body元素
    target_body = target_doc.part.document.body._element
    
    # 获取目标文档中所有p元素的索引
    all_p_elements = target_body.findall(qn('w:p'))
    
    # 获取要替换的段落元素
    target_cover_elements = list(all_p_elements)[:target_cover_end]
    
    # 获取备份封面的段落元素
    backup_body = backup_doc.part.document.body._element
    backup_all_p = backup_body.findall(qn('w:p'))
    backup_cover_elements = list(backup_all_p)[:backup_cover_end]
    
    # 备份当前文件
    backup_current_file(target_docx)
    
    # 替换：移除目标封面元素，插入备份封面元素
    for elem in target_cover_elements:
        target_body.remove(elem)
    
    # 在第一个sectPr之前或body开头插入备份封面元素
    first_sectPr = target_body.find(qn('w:sectPr'))
    insert_idx = 0
    
    for elem in reversed(backup_cover_elements):
        # deep copy
        new_elem = elem.__deepcopy__(True)
        if first_sectPr is not None:
            target_body.insert(list(target_body).index(first_sectPr), new_elem)
        else:
            target_body.append(new_elem)
    
    target_doc.save(str(target_docx))
    return backup_cover_end, target_cover_end


def get_extracted_cover_range(docx_path):
    """用lxml直接提取body子元素来识别封面范围"""
    from lxml import etree
    doc = Document(str(docx_path))
    body = doc.part.document.body._element
    children = list(body)
    
    cover_end = None
    for i, child in enumerate(children):
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'p':
            # 检查是否是摘要/目录段落
            texts = [node.text or '' for node in child.iter() 
                    if node.tag.endswith('}t')]
            text = ''.join(texts).strip()
            if text in ("摘要", "摘  要", "目录", "目  录", "Abstract"):
                cover_end = i
                break
    
    if cover_end is None:
        cover_end = min(80, len(children))
    
    return cover_end


def restore_covers_lxml_direct(backup_docx, target_docx):
    """
    终极方法：直接用lxml操作XML树。
    从备份文档body复制前N个子元素，替换目标文档body的前N个子元素。
    """
    from lxml import etree
    import copy
    
    # 读取两个docx的document.xml
    import zipfile
    import tempfile
    
    def read_document_xml(docx_path):
        with zipfile.ZipFile(str(docx_path), 'r') as z:
            return etree.fromstring(z.read('word/document.xml'))
    
    def write_document_xml(docx_path, xml_tree):
        import tempfile
        tmp_path = str(docx_path) + '.tmp'
        shutil.copy2(str(docx_path), tmp_path)
        with zipfile.ZipFile(tmp_path, 'r') as z:
            entries = z.namelist()
            data = {name: z.read(name) for name in entries}
        
        data['word/document.xml'] = etree.tostring(xml_tree, xml_declaration=True, encoding='UTF-8', standalone=True)
        
        with zipfile.ZipFile(str(docx_path), 'w', zipfile.ZIP_DEFLATED) as z:
            for name, content in data.items():
                z.writestr(name, content)
        
        os.remove(tmp_path)
        return True
    
    backup_xml = read_document_xml(backup_docx)
    target_xml = read_document_xml(target_docx)
    
    # 获取body元素
    nsmap = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    backup_body = backup_xml.find('.//w:body', nsmap)
    target_body = target_xml.find('.//w:body', nsmap)
    
    backup_children = list(backup_body)
    target_children = list(target_body)
    
    # 找到备份的封面结束位置
    backup_cover_end = None
    for i, child in enumerate(backup_children):
        tag = child.tag.split('}')[-1]
        if tag == 'p':
            texts = [node.text or '' for node in child.iter() 
                    if node.tag.endswith('}t')]
            text = ''.join(texts).strip()
            if text in ("摘要", "摘  要", "目录", "目  录", "Abstract"):
                backup_cover_end = i
                break
    
    if backup_cover_end is None:
        backup_cover_end = min(80, len(backup_children))
    
    # 找到目标的封面结束位置
    target_cover_end = None
    for i, child in enumerate(target_children):
        tag = child.tag.split('}')[-1]
        if tag == 'p':
            texts = [node.text or '' for node in child.iter() 
                    if node.tag.endswith('}t')]
            text = ''.join(texts).strip()
            if text in ("摘要", "摘  要", "目录", "目  录", "Abstract"):
                target_cover_end = i
                break
    
    if target_cover_end is None:
        target_cover_end = min(80, len(target_children))
    
    print(f"  备份封面: {backup_cover_end}元素 → 目标封面: {target_cover_end}元素")
    
    # 获取备份的封面元素
    backup_cover_elems = list(backup_children)[:backup_cover_end]
    
    # 获取目标的封面元素
    target_cover_elems = list(target_children)[:target_cover_end]
    
    # 移除目标封面元素
    for elem in target_cover_elems:
        target_body.remove(elem)
    
    # 插入备份封面元素（逆序保持顺序）
    for elem in reversed(backup_cover_elems):
        new_elem = copy.deepcopy(elem)
        target_body.insert(0, new_elem)
    
    write_document_xml(target_docx, target_xml)
    return backup_cover_end, target_cover_end


def main():
    # 获取已修改的论文清单（备份目录中有对应文件）
    updated_papers = []
    
    for backup_year_dir in sorted(BACKUP_DIR.iterdir()):
        if not backup_year_dir.is_dir() or not backup_year_dir.name.isdigit():
            continue
        year = backup_year_dir.name
        
        for backup_major_dir in sorted(backup_year_dir.iterdir()):
            if not backup_major_dir.is_dir():
                continue
            major = backup_major_dir.name
            
            for backup_paper_dir in sorted(backup_major_dir.iterdir()):
                if not backup_paper_dir.is_dir():
                    continue
                
                # 对应的终版目录
                final_paper_dir = FINAL_DIR / year / major / backup_paper_dir.name
                if not final_paper_dir.exists():
                    continue
                
                # 找到DOCX文件
                backup_docxs = list(backup_paper_dir.glob('*.docx'))
                final_docxs = list(final_paper_dir.glob('*.docx'))
                
                if not backup_docxs or not final_docxs:
                    continue
                
                backup_docx = backup_docxs[0]
                final_docx = final_docxs[0]
                
                # 检查是否需要恢复封面（备份和修复版内容不同？直接比较大小）
                if backup_docx.stat().st_size == final_docx.stat().st_size:
                    continue  # 没修改过的跳过
                
                updated_papers.append({
                    'name': backup_paper_dir.name,
                    'backup': backup_docx,
                    'final': final_docx,
                    'year': year,
                    'major': major,
                })
    
    print(f"发现 {len(updated_papers)} 篇修改过的论文需要恢复封面")
    
    success = []
    failed = []
    
    for i, paper in enumerate(updated_papers):
        name = paper['name']
        print(f"\n[{i+1}/{len(updated_papers)}] {name}...", end=' ', flush=True)
        
        try:
            backup_cover_end, target_cover_end = restore_covers_lxml_direct(
                paper['backup'], paper['final']
            )
            print(f"✓ 封面已恢复 ({backup_cover_end}个元素)", end=' ')
            
            # 验证：打开目标文件检查封面是否恢复
            verify_doc = Document(str(paper['final']))
            first_texts = [p.text.strip() for p in verify_doc.paragraphs[:10] if p.text.strip()]
            if first_texts:
                print(f"[前段: '{first_texts[0][:20]}...']", end=' ')
            
            success.append(name)
            print()
            
        except Exception as e:
            print(f"✗ 失败: {e}")
            failed.append({'name': name, 'error': str(e)})
    
    print(f"\n\n{'='*50}")
    print(f"完成！成功 {len(success)}, 失败 {len(failed)}")
    
    if failed:
        print(f"\n失败列表:")
        for f in failed:
            print(f"  ✗ {f['name']}: {f['error']}")

if __name__ == '__main__':
    main()
