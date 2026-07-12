#!/usr/bin/env python3
"""
从备份目录恢复论文封面到论文终版目录。
备份目录：26_王奕文_029823431266.docx（年份前缀+文件名）
终版目录：{年份}/{专业}/{姓名_准考证号}.docx（扁平结构，无子目录）
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
    stem = backup_name.stem  # 去掉.docx
    m = re.match(r'^(\d+)_(.+)_(\d+)$', stem)
    if not m:
        return None
    
    year = m.group(1)
    name = m.group(2)
    exam_id = m.group(3)
    
    # 终版是扁平结构：{年份}/{专业}/{姓名_准考证号}.docx
    paper_filename = f"{name}_{exam_id}.docx"
    for major in ['经管', '机械', '土木']:
        docx_path = FINAL_DIR / year / major / paper_filename
        if docx_path.exists():
            return docx_path
    
    return None

def read_document_xml(docx_path):
    """从docx中读取document.xml"""
    import zipfile
    with zipfile.ZipFile(str(docx_path), 'r') as z:
        return etree.fromstring(z.read('word/document.xml'))

def write_document_xml(docx_path, xml_tree):
    """将修改后的document.xml写回docx"""
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
    """找到封面结束位置（摘要/目录段落之前的索引）"""
    for i, child in enumerate(children):
        tag = child.tag.split('}')[-1]
        if tag == 'p':
            texts = [node.text or '' for node in child.iter() 
                    if node.tag.endswith('}t')]
            text = ''.join(texts).strip()
            if text in ("摘要", "摘  要", "目录", "目  录", "Abstract", "ABSTRACT"):
                return i
    
    # 没找到，保守取前80个元素（含sectPr等）
    return min(80, len(children))

def has_cover_content(docx_path):
    """检查DOCX是否有封面内容（非空首段）"""
    doc_xml = read_document_xml(docx_path)
    ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    body = doc_xml.find(f'.//{{{ns}}}body')
    if body is None:
        return False
    
    for child in list(body)[:5]:
        tag = child.tag.split('}')[-1]
        if tag == 'p':
            texts = [node.text or '' for node in child.iter() 
                    if node.tag.endswith('}t')]
            text = ''.join(texts).strip()
            if text:
                return True
    return False

def get_cover_first_text(docx_path, max_paras=5):
    """获取封面首段文本"""
    doc_xml = read_document_xml(docx_path)
    ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    body = doc_xml.find(f'.//{{{ns}}}body')
    if body is None:
        return []
    
    texts = []
    for child in list(body)[:max_paras]:
        tag = child.tag.split('}')[-1]
        if tag == 'p':
            t = ''.join(node.text or '' for node in child.iter() 
                       if node.tag.endswith('}t'))
            texts.append(t.strip())
    return texts

def restore_cover(backup_docx, final_docx, verbose=True):
    """从备份DOCX恢复封面到目标DOCX"""
    
    # 先备份目标文件
    bak_path = final_docx.with_suffix('.docx.covers_bak')
    if not bak_path.exists():
        shutil.copy2(str(final_docx), str(bak_path))
    
    backup_xml = read_document_xml(backup_docx)
    final_xml = read_document_xml(final_docx)
    
    ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    
    backup_body = backup_xml.find(f'.//{{{ns}}}body')
    final_body = final_xml.find(f'.//{{{ns}}}body')
    
    if backup_body is None or final_body is None:
        raise ValueError("找不到body元素")
    
    backup_children = list(backup_body)
    final_children = list(final_body)
    
    # 找封面结束位置
    backup_cover_end = find_cover_end(backup_children)
    final_cover_end = find_cover_end(final_children)
    
    if verbose:
        print(f"  备份封面: {backup_cover_end}元素, 终版封面: {final_cover_end}元素", end=' ')
    
    # 如果封面长度差异太大，取较小的
    cover_len = min(backup_cover_end, final_cover_end)
    if abs(backup_cover_end - final_cover_end) > 20:
        print(f"(差异{abs(backup_cover_end-final_cover_end)}，取{cover_len})", end=' ')
    
    # 保留目标文档中封面范围之后的元素
    remaining = list(final_children)[cover_len:]
    
    # 备份封面元素
    backup_cover = list(backup_children)[:cover_len]
    
    # 清除目标body的所有子元素
    for child in list(final_body):
        final_body.remove(child)
    
    # 插入备份封面元素
    for elem in backup_cover:
        final_body.append(copy.deepcopy(elem))
    
    # 插入剩余元素
    for elem in remaining:
        final_body.append(copy.deepcopy(elem))
    
    # 保存
    write_document_xml(final_docx, final_xml)
    
    return cover_len


def main():
    backup_files = sorted(BACKUP_DIR.glob('*.docx'))
    print(f"备份目录中有 {len(backup_files)} 个DOCX文件\n")
    
    # 先列出所有备份的匹配情况
    papers = []
    unmatched = []
    for bf in backup_files:
        final_docx = find_final_docx(bf)
        if final_docx is None:
            unmatched.append(bf.name)
        else:
            papers.append((bf, final_docx, bf.stem))
    
    if unmatched:
        print(f"无法匹配 {len(unmatched)} 个文件:")
        for u in unmatched:
            print(f"  ! {u}")
        print()
    
    print(f"匹配到 {len(papers)} 篇论文\n")
    
    # 先检查当前封面状态
    print("检查当前封面状态:")
    for backup_docx, final_docx, name in papers:
        backup_texts = get_cover_first_text(backup_docx, 3)
        final_texts = get_cover_first_text(final_docx, 3)
        has_final = has_cover_content(final_docx)
        
        print(f"  {name}: ", end='')
        if not has_final:
            print(f"[封面为空] 备份首段: '{backup_texts[0][:30] if backup_texts else '(无)'}' → 需要恢复")
        else:
            if backup_texts and final_texts:
                same = backup_texts[0][:20] == final_texts[0][:20]
                if same:
                    print(f"[已有封面] '{final_texts[0][:30]}' → 跳过")
                else:
                    print(f"[封面不同] 当前: '{final_texts[0][:25]}' 备份: '{backup_texts[0][:25]}' → 恢复")
            else:
                print(f"[不确定] 备份: {backup_texts[:2]}, 当前: {final_texts[:2]}")
    
    print(f"\n{'='*50}")
    
    # 只恢复封面为空的或不同的
    success = []
    failed = []
    skipped = []
    
    for i, (backup_docx, final_docx, name) in enumerate(papers):
        # 检查是否需要恢复
        has_final = has_cover_content(final_docx)
        backup_texts = get_cover_first_text(backup_docx, 3)
        final_texts = get_cover_first_text(final_docx, 3)
        
        if has_final:
            if backup_texts and final_texts and backup_texts[0][:20] == final_texts[0][:20]:
                print(f"[{i+1}/{len(papers)}] {name} 封面已相同，跳过")
                skipped.append(name)
                continue
        
        print(f"[{i+1}/{len(papers)}] {name} 恢复封面...", end=' ', flush=True)
        
        try:
            cover_len = restore_cover(backup_docx, final_docx, verbose=True)
            
            # 验证
            verify_texts = get_cover_first_text(final_docx, 3)
            print(f"✓ → '{verify_texts[0][:30] if verify_texts else '(无)'}'")
            
            success.append(name)
            
        except Exception as e:
            print(f"✗ 失败: {e}")
            failed.append({'name': name, 'error': str(e)})
    
    print(f"\n{'='*50}")
    print(f"完成！成功 {len(success)}, 跳过 {len(skipped)}, 失败 {len(failed)}")
    if failed:
        print(f"失败: {[f['name'] for f in failed]}")

if __name__ == '__main__':
    main()
