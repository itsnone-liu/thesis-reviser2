#!/usr/bin/env python3
"""
从备份目录恢复论文封面到论文终版目录。
修复：恢复封面完整内容（不仅仅是首段，而是整个封面区域的完整内容）。
"""

import sys
import os
from pathlib import Path
import shutil
import re
from lxml import etree
import copy

BACKUP_DIR = Path("/root/project/workspace/论文资料/整理/论文终版_备份_修复前")
FINAL_DIR = Path("/root/project/workspace/论文资料/整理/论文终版")

def find_final_docx(backup_name):
    """根据备份文件名找到对应的终版DOCX"""
    stem = backup_name.stem
    m = re.match(r'^(\d+)_(.+)_(\d+)$', stem)
    if not m:
        return None
    year = m.group(1)
    name = m.group(2)
    exam_id = m.group(3)
    paper_filename = f"{name}_{exam_id}.docx"
    for major in ['经管', '机械', '土木']:
        docx_path = FINAL_DIR / year / major / paper_filename
        if docx_path.exists():
            return docx_path
    return None

def read_document_xml(docx_path):
    import zipfile
    with zipfile.ZipFile(str(docx_path), 'r') as z:
        return etree.fromstring(z.read('word/document.xml'))

def write_document_xml(docx_path, xml_tree):
    import zipfile
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

def find_cover_end(children):
    """找到封面结束位置（以摘要/目录段落为准）"""
    for i, child in enumerate(children):
        tag = child.tag.split('}')[-1]
        if tag == 'p':
            texts = [node.text or '' for node in child.iter() 
                    if node.tag.endswith('}t')]
            text = ''.join(texts).strip()
            if text in ("摘要", "摘  要", "目录", "目  录", "Abstract", "ABSTRACT"):
                return i
    return min(80, len(children))

def is_cover_empty(xml_tree):
    """检查封面是否为空（论文题目、姓名、学号、专业这些字段是否有值）"""
    ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    body = xml_tree.find(f'.//{{{ns}}}body')
    if body is None:
        return True
    cover_end = find_cover_end(list(body))
    
    empty_data_fields = True
    for child in list(body)[:cover_end]:
        tag = child.tag.split('}')[-1]
        if tag == 'p':
            texts = [node.text or '' for node in child.iter() 
                    if node.tag.endswith('}t')]
            text = ''.join(texts).strip()
            # 检查信息字段
            if any(kw in text for kw in ['论文题目', '姓    名', '学    号', '专    业', '准考证号']):
                # 提取冒号后的内容
                parts = text.split('：')
                if len(parts) >= 2 and parts[1].strip():
                    empty_data_fields = False
    return empty_data_fields

def restore_cover(backup_docx, final_docx, verbose=True):
    """核心：用备份的封面替换终版的封面"""
    import zipfile
    
    # 先把目标文件备份
    bak_path = final_docx.with_suffix('.docx.covers_bak')
    if not bak_path.exists():
        shutil.copy2(str(final_docx), str(bak_path))
    
    # 读取两个文件的document.xml
    with zipfile.ZipFile(str(final_docx), 'r') as z:
        final_zip_data = {name: z.read(name) for name in z.namelist()}
    
    with zipfile.ZipFile(str(backup_docx), 'r') as z:
        backup_doc_xml = etree.fromstring(z.read('word/document.xml'))
    
    # 解析终版的document.xml
    final_doc_xml = etree.fromstring(final_zip_data['word/document.xml'])
    
    ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    backup_body = backup_doc_xml.find(f'.//{{{ns}}}body')
    final_body = final_doc_xml.find(f'.//{{{ns}}}body')
    
    if backup_body is None or final_body is None:
        raise ValueError("找不到body元素")
    
    backup_children = list(backup_body)
    final_children = list(final_body)
    
    # 找封面结束位置
    backup_cover_end = find_cover_end(backup_children)
    final_cover_end = find_cover_end(final_children)
    
    if verbose:
        print(f"  备份封面: {backup_cover_end}个元素, 终版封面: {final_cover_end}个元素", end=' ')
    
    # 取各自身份的封面长度（不使用相同长度，因为格式可能不同）
    backup_cover = list(backup_children)[:backup_cover_end]
    remaining = list(final_children)[final_cover_end:]
    
    # 完全重建final body：备份封面 + 终版剩余
    for child in list(final_body):
        final_body.remove(child)
    
    for elem in backup_cover:
        final_body.append(copy.deepcopy(elem))
    
    for elem in remaining:
        final_body.append(copy.deepcopy(elem))
    
    # 写回document.xml
    final_zip_data['word/document.xml'] = etree.tostring(final_doc_xml, xml_declaration=True, encoding='UTF-8', standalone=True)
    
    # 重建DOCX文件
    tmp_path = str(final_docx) + '.tmp'
    shutil.copy2(str(final_docx), tmp_path)
    with zipfile.ZipFile(tmp_path, 'r') as z:
        orig_names = z.namelist()
        orig_data = {name: z.read(name) for name in orig_names}
    
    # 用新读的完整数据
    with zipfile.ZipFile(str(final_docx), 'w', zipfile.ZIP_DEFLATED) as z:
        for name in orig_names:
            if name == 'word/document.xml':
                z.writestr(name, final_zip_data['word/document.xml'])
            else:
                z.writestr(name, orig_data[name])
    
    os.remove(tmp_path)
    return backup_cover_end, final_cover_end


def get_cover_details(xml_tree):
    """获取封面中的关键信息字段"""
    ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    body = xml_tree.find(f'.//{{{ns}}}body')
    if body is None:
        return []
    cover_end = find_cover_end(list(body))
    
    details = []
    for child in list(body)[:cover_end]:
        tag = child.tag.split('}')[-1]
        if tag == 'p':
            texts = [node.text or '' for node in child.iter() 
                    if node.tag.endswith('}t')]
            text = ''.join(texts).strip()
            if text and any(kw in text for kw in ['论文题目', '姓', '学', '专', '准考证']):
                details.append(text)
    return details


def main():
    backup_files = sorted(BACKUP_DIR.glob('*.docx'))
    print(f"备份目录中有 {len(backup_files)} 个DOCX文件\n")
    
    papers = []
    for bf in backup_files:
        final_docx = find_final_docx(bf)
        if final_docx is None:
            continue
        papers.append((bf, final_docx, bf.stem))
    
    print(f"匹配到 {len(papers)} 篇论文\n")
    
    # 检查封面状态
    print("封面状态检查:")
    need_restore = []
    skip = []
    for backup_docx, final_docx, name in papers:
        backup_xml = read_document_xml(backup_docx)
        final_xml = read_document_xml(final_docx)
        
        backup_details = get_cover_details(backup_xml)
        final_details = get_cover_details(final_xml)
        empty = is_cover_empty(final_xml)
        
        status = "封面字段为空!" if empty else "有内容"
        print(f"  {name}: [{status}]", end='')
        if backup_details:
            print(f" 备份: {backup_details[0][:30]}", end='')
        if final_details:
            print(f" 当前: {final_details[0][:30]}", end='')
        print()
        
        if empty:
            need_restore.append((backup_docx, final_docx, name))
        else:
            skip.append(name)
    
    print(f"\n需恢复: {len(need_restore)}, 跳过(已有内容): {len(skip)}")
    
    if not need_restore:
        print("无需恢复操作")
        return
    
    print(f"\n{'='*50}")
    print("开始恢复封面...\n")
    
    success = []
    failed = []
    
    for i, (backup_docx, final_docx, name) in enumerate(need_restore):
        print(f"[{i+1}/{len(need_restore)}] {name}...", end=' ', flush=True)
        
        try:
            bc, fc = restore_cover(backup_docx, final_docx, verbose=True)
            
            # 验证
            verify_xml = read_document_xml(final_docx)
            details = get_cover_details(verify_xml)
            print(f"✓ → {details[0][:35] if details else '(无信息)'}")
            
            success.append(name)
        except Exception as e:
            print(f"✗ 失败: {e}")
            failed.append({'name': name, 'error': str(e)})
    
    print(f"\n{'='*50}")
    print(f"完成！成功 {len(success)}, 失败 {len(failed)}")
    if failed:
        print(f"失败: {[f['name'] for f in failed]}")

if __name__ == '__main__':
    main()
