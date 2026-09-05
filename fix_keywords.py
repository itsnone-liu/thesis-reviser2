# -*- coding: utf-8 -*-
"""关键词修订: ①内容全部改为名册LWGJC(分隔符统一为；, 词形保持名册原文)
②格式统一为单行"关键词：X"(50篇新版样式, 12pt不加粗); 删独立标签行; 修周昊双标签。
备份: 论文终版/_关键词备份/; 明细: 论文终版/关键词修订记录.csv"""
import os, re, csv, shutil
from docx import Document
from docx.shared import Pt
import pandas as pd

df = pd.read_excel('论文终版/25年12月(157人)+26年6月(116人)(1).xls', header=0, dtype=str)
ros = {str(r['XH']).strip(): str(r['LWGJC']).strip() for _, r in df.iterrows()}

def kw_fmt(raw):
    """名册关键词→论文书写格式: 剥"关键词："残留前缀, 去空白, 分隔符统一全角；, 词形原样保留。"""
    s = str(raw).strip().replace('\u3000', ' ')
    s = re.sub(r'^\s*关键词\s*[：:]\s*', '', s)
    parts = [p.strip() for p in re.split(r'[;；,，、]+', s) if p.strip()]
    return '；'.join(parts)

def rm_para(p):
    p._element.getparent().remove(p._element)

BAK = '论文终版/_关键词备份'
os.makedirs(BAK, exist_ok=True)
files = []
for cohort in ('25', '26'):
    for tp in os.listdir(f'论文终版/{cohort}'):
        p = f'论文终版/{cohort}/{tp}'
        if os.path.isdir(p):
            files += [f'{p}/{fn}' for fn in os.listdir(p) if fn.endswith('.docx') and not fn.startswith('~')]

log = []
for path in files:
    fn = os.path.basename(path)
    sid0 = re.sub(r'\D', '', fn.rsplit('_', 1)[-1].replace('.docx', ''))
    sid = sid0 if sid0 in ros else next((k for k in ros if k.lstrip('0') == sid0.lstrip('0')), None)
    if not sid:
        log.append([fn, '跳过:名册无此学号', '']); continue
    target = kw_fmt(ros[sid])
    bak = f'{BAK}/{fn}'
    if not os.path.exists(bak):          # 幂等重跑: 不覆盖首轮备份
        shutil.copy2(path, bak)
    d = Document(path)
    ops = []
    idx = None
    for i, p in enumerate(d.paragraphs[:30]):
        if re.search(r'关键词', p.text):
            idx = i; break
    if idx is None:
        log.append([fn, '未找到关键词段落', '']); continue
    head_p = d.paragraphs[idx]
    nxt = None
    if idx + 1 < len(d.paragraphs) and d.paragraphs[idx+1].text.strip():
        nxt = d.paragraphs[idx+1]
    if re.match(r'^\s*关键词\s*[：:]\s*\S', head_p.text):
        # 行内式(含周昊双标签的内容行)
        line = head_p
        if nxt is not None and re.match(r'^\s*关键词\s*[：:]\s*\S', nxt.text):
            pass
    elif re.match(r'^\s*关键词\s*$', head_p.text) and nxt is not None:
        # 独立标签行+内容行: 删标签行, 用内容行承载
        rm_para(head_p)
        line = nxt
    elif re.match(r'^\s*关键词\s*$', head_p.text):
        log.append([fn, '标签行后无内容行, 未处理', '']); continue
    else:
        log.append([fn, f'非常规关键词段: {head_p.text[:20]!r}', '']); continue
    # 周昊双标签: head_p已删, 但若行内式且下一行也是关键词行(反向)不处理——上面独立式已覆盖
    old = line.text.strip()
    if line.runs:
        line.runs[0].text = f'关键词：{target}'
        for r in line.runs[1:]:
            r.text = ''
        r0 = line.runs[0]
        if r0.bold or (r0.font.size and r0.font.size.pt > 13):
            r0.bold = None
            r0.font.size = Pt(12)
    else:
        line.add_run(f'关键词：{target}')
    if old != f'关键词：{target}':
        ops.append(f'内容[{old[:24]}→{target[:24]}]')
    else:
        ops.append('格式核验(内容已一致)')
    d.save(path)
    log.append([fn, '; '.join(ops) if ops else '无变化', target[:50]])

with open('论文终版/关键词修订记录.csv', 'w', newline='', encoding='utf-8-sig') as f:
    w = csv.writer(f)
    w.writerow(['文件', '修订操作', '名册关键词(写入值)'])
    w.writerows(log)
n = sum(1 for _, o, _ in log if '内容[' in o)
print(f"处理{len(log)}篇: 内容修订{n}篇(其余仅格式统一/已一致), 备份_关键词备份/, 记录→关键词修订记录.csv")
