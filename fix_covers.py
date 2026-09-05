# -*- coding: utf-8 -*-
"""封面批量修订: 题目/专业/导师按名册, 日期统一 25届→2025年11月 / 26届→2026年5月。
保留label格式; 值写入冒号后的run(保下划线等格式); 缺日期则在指导教师后插入。"""
import os, re, csv, json, shutil
from docx import Document
import pandas as pd

df = pd.read_excel('论文终版/25年12月(157人)+26年6月(116人)(1).xls', header=0, dtype=str)
ros = {}
for _, r in df.iterrows():
    sid = str(r['XH']).strip()
    ros[sid] = {'sid': sid, 'major': str(r['ZYMC']).strip(), 'tutor': str(r['DSXM']).strip(),
                'title': str(r['LWTM']).strip(),
                'cohort': '25' if str(r['HXWRQ']).startswith('2025') else '26'}
BAK = '论文终版/_封面备份'
os.makedirs(BAK, exist_ok=True)
UNI_DATE = {'25': '2025年11月', '26': '2026年5月'}

def norm_t(s):
    s = re.sub(r'\s+', '', str(s).replace('\u3000', ''))
    s = re.sub(r'[-－–—―‑‒−﹘﹣ー]+', '—', s)
    return s.replace('：', ':').replace('，', ',').replace('、', ',').replace('（', '(').replace('）', ')').replace(';', ',').replace('；', ',')
sq = lambda s: re.sub(r'\s+', '', str(s))

def set_field(p, label_re, new_value):
    """把段落中 label冒号后的值替换为new_value, 保留label所在run格式。
    值写入冒号后的第一个run(通常带下划线等值格式); 无多run则整体重写。"""
    full = p.text
    m = label_re.search(full)
    if not m: return False
    label = full[:m.end(1)]          # 含冒号
    if len(p.runs) >= 2:
        # 找到冒号结束位置所在的run索引
        acc = ""
        cut = None
        for i, r in enumerate(p.runs):
            acc += r.text
            if len(acc) >= m.end(1):
                cut = i; break
        if cut is None: return False
        p.runs[0].text = label if cut == 0 else label
        # cut run: 若label恰在此run结束, 值写入下一run; 否则截断该run
        if len("".join(r.text for r in p.runs[:cut])) < m.end(1):
            pass
        # 简化策略: label全部并入runs[0], runs[1]写值, 其余清空
        p.runs[0].text = label
        p.runs[1].text = new_value
        for r in p.runs[2:]: r.text = ""
    elif p.runs:
        p.runs[0].text = label + new_value
    else:
        p.add_run(label + new_value)
    return True

DATE_ANY = re.compile(r'^\s*[0-9０-９]{4}\s*年\s*[0-9０-９]{1,2}\s*(月\s*[0-9０-９]{0,2}\s*日?)?\s*$')

files = []
for cohort in ('25', '26'):
    for tp in os.listdir(f'论文终版/{cohort}'):
        p = f'论文终版/{cohort}/{tp}'
        if os.path.isdir(p):
            files += [(cohort, f'{p}/{fn}') for fn in os.listdir(p) if fn.endswith('.docx') and not fn.startswith('~')]

log = []
for cohort, path in files:
    fn = os.path.basename(path)
    fsid = re.sub(r'\D', '', fn.rsplit('_', 1)[-1].replace('.docx', ''))
    ro = ros.get(fsid) or next((v for k, v in ros.items() if k.lstrip('0') == fsid.lstrip('0')), None)
    if ro is None:
        log.append([fn, '跳过: 名册无此学号']); continue
    shutil.copy2(path, f'{BAK}/{fn}')
    d = Document(path)
    paras = d.paragraphs[:40]
    ops = []
    # 题目/专业/导师
    for label_re, key, nm in ((re.compile(r'(论文题目[：:])'), 'title', '题目'),
                              (re.compile(r'(专\s*业[：:])'), 'major', '专业'),
                              (re.compile(r'(指导教师[：:])'), 'tutor', '导师')):
        for p in paras:
            if label_re.search(p.text):
                old = label_re.search(p.text)
                old_v = p.text[old.end(1):].strip()
                same = norm_t(old_v) == norm_t(ro[key]) if key == 'title' else sq(old_v) == sq(ro[key])
                if not same:
                    set_field(p, label_re, ro[key])
                    ops.append(f'{nm}[{old_v[:16]}→{ro[key][:16]}]')
                break
    # 日期: 找到日期段落则统一改写; 没有则在指导教师段后插入
    want_date = UNI_DATE[ro['cohort']]
    date_p = None
    for p in paras:
        if DATE_ANY.match(p.text) and re.search(r'年', p.text):
            date_p = p; break
    if date_p is not None:
        old_d = date_p.text.strip()
        if old_d != want_date:
            set_field(date_p, re.compile(r'(\s*)'), want_date) if False else None
            if date_p.runs:
                date_p.runs[0].text = want_date
                for r in date_p.runs[1:]: r.text = ""
            else:
                date_p.add_run(want_date)
            ops.append(f'日期[{old_d}→{want_date}]')
    else:
        # 插入: 在指导教师段之后
        for i, p in enumerate(paras):
            if re.search(r'指导教师[：:]', p.text):
                newp = p.insert_paragraph_before('')
                # insert_paragraph_before插在p之前 → 需要插在p之后: 换用下一元素法
                newp._element.getparent().remove(newp._element)
                nxt = p._p.getnext()
                from docx.text.paragraph import Paragraph
                from docx.oxml import OxmlElement
                new_el = OxmlElement('w:p')
                if nxt is not None:
                    nxt.addprevious(new_el)
                else:
                    p._p.getparent().append(new_el)
                np = Paragraph(new_el, p._parent)
                np.style = p.style
                np.add_run(want_date)
                ops.append(f'日期[缺失→插入{want_date}]')
                break
        else:
            ops.append('日期缺失且无指导教师段, 未插入')
    d.save(path)
    log.append([fn, '; '.join(ops) if ops else '无改动(已一致)'])

with open('论文终版/封面修订记录.csv', 'w', newline='', encoding='utf-8-sig') as f:
    w = csv.writer(f)
    w.writerow(['文件', '修订操作'])
    w.writerows(log)
n_change = sum(1 for _, o in log if o and '无改动' not in o and '跳过' not in o)
print(f"处理{len(log)}篇: 有修订{n_change}篇, 明细见 封面修订记录.csv (备份: _封面备份/)")
