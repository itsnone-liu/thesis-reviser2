#!/usr/bin/env python3
"""
batch_revise_papers.py — 批量修改经管类论文
从源目录读取 DOCX，调用 reviser 修改管道，输出到目标目录。
分批处理，每批完成后抽样检查。
"""
import sys, os, re, json, time, shutil, random, subprocess
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from core import (
    extract_docx_text, extract_cover_info, call_llm,
    normalize_cover_info, clean_text,
)
from pygments import highlight
from reviser import _analyze_original_paper, _diagnose_and_reconstruct, _generate_revised
from renderer import render as render_it

SRC_DIR = "/root/project/workspace/论文资料/整理/经管类"
OUT_DIR = f"{SRC_DIR}/已处理"
COVER_MD = f"{OUT_DIR}/封面缺失汇总.md"

BATCH_SIZE = 20

def extract_title_from_docx(path):
    """从DOCX提取论文题目"""
    from zipfile import ZipFile
    import re
    try:
        with ZipFile(path) as z:
            xml = z.read('word/document.xml').decode('utf-8')
            paras = re.findall(r'<w:t[^>]*>([^<]+)', xml)
            for p in paras:
                if len(p) > 5 and ("研究" in p or "分析" in p or "问题" in p or "优化" in p or "对策" in p):
                    return p.strip()
    except:
        pass
    return ""

def fix_cover_info(cover, profile, name, student_id):
    """补齐封面信息"""
    if not cover:
        cover = {}
    cover['name'] = name
    cover['student_id'] = student_id
    if not cover.get('title') and profile.get('title'):
        cover['title'] = profile['title']
    if not cover.get('major'):
        cover['major'] = profile.get('major', '财务管理')
    if not cover.get('advisor'):
        cover['advisor'] = '王艳'
    if not cover.get('year') or not cover.get('month') or not cover.get('day'):
        cover['year'] = cover.get('year', '2025')
        cover['month'] = cover.get('month', '5')
        cover['day'] = cover.get('day', '20')
    for k in ('title','name','student_id','major','level','advisor','year','month','day'):
        if k not in cover:
            cover[k] = ''
    return cover

def process_one(fname, out_dir, missing_covers):
    """处理一篇论文"""
    path = os.path.join(SRC_DIR, fname)
    parts = fname.replace('.docx','').split('-')
    name = parts[0] if len(parts) >= 1 else ''
    student_id = parts[1] if len(parts) >= 2 else ''
    major = parts[2] if len(parts) >= 3 else ''
    student_dir = os.path.join(out_dir, f"{name}_{student_id}")
    os.makedirs(student_dir, exist_ok=True)
    txt_out = os.path.join(student_dir, "paper.txt")
    docx_out = os.path.join(student_dir, "paper.docx")
    profile_json = os.path.join(student_dir, "profile.json")
    cover_json = os.path.join(student_dir, "cover.json")

    # 1. 提取原论文信息
    try:
        text = extract_docx_text(path)
    except:
        text = ""
    if not text:
        print(f"  ⚠️ 无法读取论文: {fname}，跳过")
        return False

    cover = extract_cover_info(path)
    cover['name'] = name
    cover['student_id'] = student_id
    cover['major'] = major

    # 2. 分析画像
    print(f"  [画像分析] 分析原论文...")
    analysis = _analyze_original_paper(text[:6000], "管理")
    analysis['title'] = extract_title_from_docx(path) or analysis.get('title', '')
    analysis['company'] = analysis.get('company', name)
    analysis['major'] = major

    # 3. 记录封面缺失
    missing_fields = []
    for k in ['title','name','student_id','major','level','advisor','year','month','day']:
        v = cover.get(k, '')
        if not v:
            missing_fields.append(k)
    if missing_fields:
        missing_covers.append({
            'name': name,
            'student_id': student_id,
            'file': fname,
            'missing': ', '.join(missing_fields),
            'cover_info': cover
        })

    # 4. 写profile
    with open(profile_json, 'w', encoding='utf-8') as f:
        json.dump(analysis, f, ensure_ascii=False, indent=2)

    # 5. 诊断并生成修改版
    def update(msg, prog):
        print(f"    [{prog}%] {msg}", flush=True)
    print(f"  [诊断] 正在诊断...")
    revision_plan = _diagnose_and_reconstruct(analysis, "管理")
    # 保存profile里原有的题目和company
    rp = revision_plan.get("revised_profile", {})
    rp['title'] = analysis.get('title', '')
    rp['company'] = analysis.get('company', '')
    rp['major'] = major
    # 如果诊断没有给出公司名，保留原有profile
    if not rp.get('company'):
        rp['company'] = analysis.get('company', name)

    print(f"  [生成] 正在生成修改版文本...")
    try:
        txt = _generate_revised(analysis, revision_plan, "管理", update)
    except Exception as e:
        print(f"  ❌ 生成失败: {e}")
        return False

    with open(txt_out, 'w', encoding='utf-8') as f:
        f.write(txt)

    # 6. 渲染DOCX
    print(f"  [渲染] 正在渲染DOCX...")
    cover_fixed = fix_cover_info(cover, analysis, name, student_id)
    with open(cover_json, 'w', encoding='utf-8') as f:
        json.dump(cover_fixed, f, ensure_ascii=False, indent=2)

    try:
        render_it(txt_out, docx_out, "管理", cover_fixed, None, update)
        print(f"  ✅ 完成: {name}_{student_id}")
        return True
    except Exception as e:
        print(f"  ❌ 渲染失败: {e}")
        return False

def spot_check(out_dir, batch_num):
    """对批次做抽检"""
    dirs = sorted(os.listdir(out_dir))
    if not dirs:
        return
    # 抽检这一批中的1-2篇
    batch_start = (batch_num - 1) * BATCH_SIZE
    batch_end = min(batch_num * BATCH_SIZE, len(dirs))
    candidates = dirs[batch_start:batch_end]
    if not candidates:
        return
    samples = random.sample(candidates, min(2, len(candidates)))
    for d in samples:
        txt_path = os.path.join(out_dir, d, "paper.txt")
        if not os.path.isfile(txt_path):
            print(f"  ⚠️ 抽检 {d}: paper.txt 不存在")
            continue
        with open(txt_path, 'r', encoding='utf-8') as f:
            content = f.read()
        charts = content.count('<chart ')
        tables = content.count('<table ')
        has_unclosed = bool(re.search(r'<(chart|table)[^>]*[^/]>', content))
        ids = re.findall(r'id="(\d+)"', content)
        dup_ids = sorted({x for x in ids if ids.count(x) > 1})
        print(f"  抽检 {d}: charts={charts}, tables={tables}, 未闭合={has_unclosed}, 重复id={dup_ids}")

def main():
    files = sorted([f for f in os.listdir(SRC_DIR) if f.endswith('.docx')])
    # 跳过已处理的
    done = set()
    if os.path.isdir(OUT_DIR):
        for d in os.listdir(OUT_DIR):
            if os.path.isfile(os.path.join(OUT_DIR, d, 'paper.docx')):
                done.add(d)
    
    # 过滤出待处理的
    todo = []
    for fname in files:
        parts = fname.replace('.docx','').split('-')
        name = parts[0] if len(parts) >= 1 else ''
        student_id = parts[1] if len(parts) >= 2 else ''
        key = f"{name}_{student_id}"
        if key in done:
            continue
        todo.append(fname)
    
    if not todo:
        print("所有论文已处理完成。")
        return

    print(f"待处理: {len(todo)} 篇 (共 {len(files)} 篇)")
    missing_covers = []
    batches = [todo[i:i+BATCH_SIZE] for i in range(0, len(todo), BATCH_SIZE)]
    
    for batch_num, batch in enumerate(batches, 1):
        print(f"\n=== 第 {batch_num} 批 ({len(batch)} 篇) ===")
        for i, fname in enumerate(batch, 1):
            print(f"[{i}/{len(batch)}] {fname}")
            success = process_one(fname, OUT_DIR, missing_covers)
            if not success:
                print(f"  ⚠️ 跳过 {fname}")

        # 抽检
        print(f"\n--- 第 {batch_num} 批抽检 ---")
        spot_check(OUT_DIR, batch_num)

    # 写封面缺失汇总
    if missing_covers:
        md_lines = ["# 封面信息缺失汇总", "", f"> 共 {len(missing_covers)} 篇论文封面信息不完整", ""]
        for item in missing_covers:
            md_lines.append(f"## {item['name']} ({item['student_id']})")
            md_lines.append(f"- 文件名: {item['file']}")
            md_lines.append(f"- 缺失字段: {item['missing']}")
            md_lines.append(f"- 已补齐封面: {json.dumps(item['cover_info'], ensure_ascii=False)}")
            md_lines.append("")
        with open(COVER_MD, 'w', encoding='utf-8') as f:
            f.write("\n".join(md_lines))
        print(f"\n封面缺失汇总已保存: {COVER_MD}")

    print(f"\n✅ 全部处理完成。输出目录: {OUT_DIR}")

if __name__ == "__main__":
    main()
