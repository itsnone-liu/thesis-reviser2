#!/usr/bin/env python3
"""
从备份目录恢复论文封面到论文终版目录。
备份目录结构：26_王奕文_029823431266.docx（年份前缀+文件名）
终版目录结构：{年份}/{专业}/{姓名_准考证号}/{姓名_准考证号}.docx
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
    """根据备份文件名（如 26_王奕文_029823431266.docx）找到对应的终版DOCX"""
    # 解析备份文件名
    stem = backup_name.stem  # 去掉.docx
    # 格式：年份_姓名_准考证号
    m = re.match(r'^(\d+)_(.+)_(\d+)$', stem)
    if not m:
        return None
    
    year = m.group(1)
    name = m.group(2)
    exam_id = m.group(3)
    
    # 在终版目录中搜索
    paper_dir_name = f"{name}_{exam_id}"
    for major in ['经管', '机械', '土木']:
        paper_dir = FINAL_DIR / year / major / paper_dir_name
        if paper_dir.exists():
            docx = paper_dir / f"{paper_dir_name}.docx"
            if docx.exists():
                return docx
    
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
    
    # 没找到，保守取前80个元素
    return min(80, len(children))

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
        print(f"  备份封面: {backup_cover_end}个元素, 终版封面: {final_cover_end}个元素", end=' ')
    
    # 如果封面长度差异太大，取较小的
    if abs(backup_cover_end - final_cover_end) > 20:
        cover_len = min(backup_cover_end, final_cover_end)
        print(f"(差异大,取最小值{cover_len})", end=' ')
    else:
        cover_len = backup_cover_end
    
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
    print(f"备份目录中有 {len(backup_files)} 个DOCX文件")
    
    papers = []
    for bf in backup_files:
        final_docx = find_final_docx(bf)
        if final_docx is None:
            print(f"! 无法匹配终版: {bf.name}")
            continue
        papers.append((bf, final_docx, bf.stem))
    
    print(f"匹配到 {len(papers)} 篇论文需要恢复封面\n")
    
    success = []
    failed = []
    skipped = []
    
    for i, (backup_docx, final_docx, name) in enumerate(papers):
        print(f"[{i+1}/{len(papers)}] {name}...", end=' ', flush=True)
        
        try:
            cover_len = restore_cover(backup_docx, final_docx, verbose=True)
            
            # 验证
            verify_doc = etree.parse(str(final_docx))
            ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
            body = verify_doc.find(f'.//{{{ns}}}body')
            first_paras = []
            if body is not None:
                for child in list(body)[:5]:
                    tag = child.tag.split('}')[-1]
                    if tag == 'p':
                        texts = [node.text or '' for node in child.iter() 
                                if node.tag.endswith('}t')]
                        text = ''.join(texts).strip()
                        if text:
                            first_paras.append(text)
            
            print(f"✓ ({cover_len}元素)", end='')
            if first_paras:
                print(f" 首段: '{first_paras[0][:30]}'", end='')
            print()
            
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
