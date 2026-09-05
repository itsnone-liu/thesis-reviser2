# -*- coding: utf-8 -*-
"""26张占比问题表批量修复: A表头改名(数据正确) B改数 C破表重建。
备份到 论文终版/_修改备份2/。"""
import os, re, shutil
from docx import Document

ROOT = '论文终版'
BAK2 = f'{ROOT}/_修改备份2'
os.makedirs(BAK2, exist_ok=True)

def set_cell(cell, text):
    """保格式改单元格文本"""
    p = cell.paragraphs[0]
    if p.runs:
        p.runs[0].text = text
        for r in p.runs[1:]: r.text = ""
    else:
        p.add_run(text)
    for extra in cell.paragraphs[1:]:
        extra._element.getparent().remove(extra._element)

def get_tables(path):
    d = Document(path)
    return d, d.tables

def find_table(tables, key_in_header, first_cell=None):
    for t in tables:
        hdr = [c.text.strip() for c in t.rows[0].cells] if t.rows else []
        if any(key_in_header in h for h in hdr) and (first_cell is None or (hdr and hdr[0] == first_cell)):
            return t
    return None

def rename_headers(t, mapping):
    n = 0
    for c in t.rows[0].cells:
        h = c.text.strip()
        if h in mapping:
            set_cell(c, mapping[h]); n += 1
    return n

def edit_cell(t, r, c, text):
    set_cell(t.rows[r].cells[c], text)

log_all = []

def fix(path, fn_name):
    if not os.path.exists(f'{BAK2}/{os.path.basename(path)}'):
        shutil.copy2(path, f'{BAK2}/{os.path.basename(path)}')
    d, tables = get_tables(path)
    ops = fn_name(tables)
    d.save(path)
    log_all.append((os.path.basename(path), ops))
    print(f"✓ {os.path.basename(path)}: {ops}")

# ── A. 表头改名 ──
def a_刘晓玲(ts):
    t = find_table(ts, '转移支付占比')
    return [f"表头{rename_headers(t, {'转移支付占比(%)': '转移支付依赖度(%)'})}处"] if t else ["未找到"]
def a_宋妞妞(ts):
    t = find_table(ts, '适用人群占比')
    return [f"表头{rename_headers(t, {'适用人群占比(%)': '适用人群覆盖率(%)'})}处"] if t else ["未找到"]
def a_徐烨(ts):
    t = find_table(ts, '主动承担额外任务占比')
    return [f"表头{rename_headers(t, {'主动承担额外任务占比(%)': '额外任务承担率(%)'})}处"] if t else ["未找到"]
def a_徐琛雁(ts):
    t = find_table(ts, '提及人数', first_cell='离职原因类别')
    return [f"表头{rename_headers(t, {'占比': '提及率(%)'})}处"] if t else ["未找到"]
def a_魏添添(ts):
    t = find_table(ts, '提及人数', first_cell='离职原因')
    return [f"表头{rename_headers(t, {'占比': '提及率(%)'})}处"] if t else ["未找到"]
def a_方舒_93(ts):
    t = find_table(ts, '涉及记录数')
    return [f"表头{rename_headers(t, {'占比(%)': '占受检记录比例(%)'})}处"] if t else ["未找到"]
def a_崔宇蓉(ts):
    t = find_table(ts, '希望加强培训')
    return [f"表头{rename_headers(t, {'希望加强培训的员工占比(百分比)': '希望加强培训的员工选择率(%)'})}处"] if t else ["未找到"]
def a_刘灿青(ts):
    t = find_table(ts, '占比范围')
    return [f"表头{rename_headers(t, {'占比范围': '等级分布区间'})}处"] if t else ["未找到"]
def a_王友航(ts):
    t = find_table(ts, '达标率')
    return [f"表头{rename_headers(t, {'达标率(≥4分占比)': '达标率(≥4分比例)'})}处"] if t else ["未找到"]
def a_侍继典65(ts):
    t = find_table(ts, '定量指标数量')
    return [f"表头{rename_headers(t, {'定量指标占比': '定量指标比例'})}处"] if t else ["未找到"]
def a_侍继典100(ts):
    t = find_table(ts, '部门数量')
    n = rename_headers(t, {'定量指标占比(%)': '定量指标比例(%)', '定性指标占比(%)': '定性指标比例(%)'})
    return [f"表头{n}处"] if t else ["未找到"]
def a_黄雪梅(ts):
    t = find_table(ts, '制度设计绩效工资')
    n = rename_headers(t, {'制度设计绩效工资占比': '制度设计绩效工资比例', '实际平均绩效工资占比': '实际平均绩效工资比例'})
    return [f"表头{n}处"] if t else ["未找到"]
def a_豆菲(ts):
    t = find_table(ts, '当前占比')
    n = rename_headers(t, {'当前占比(%)': '当前比例(%)', '目标占比(%)': '目标比例(%)'})
    return [f"表头{n}处"] if t else ["未找到"]
def a_陈新80(ts):
    t = find_table(ts, '招聘环节性别筛选')
    n = rename_headers(t, {'女性管理岗占比': '女性管理岗比例', '女性核心技术岗占比': '女性核心技术岗比例'})
    return [f"表头{n}处"] if t else ["未找到"]
def a_李昕玥(ts):
    t = find_table(ts, '三年合计占比')
    return [f"表头{rename_headers(t, {'三年合计占比': '三年合计(万元)'})}处"] if t else ["未找到"]

# ── B. 改数(叙述定性: 白奶主导55%, 功能性/奶酪上调) ──
def b_陆骏杰(ts):
    t = find_table(ts, '优化前收入占比')
    if not t: return ["未找到"]
    ops = []
    for ri in range(1, len(t.rows)):
        name = t.rows[ri].cells[0].text.strip()
        if name == '奶酪及其他':
            edit_cell(t, ri, 1, '5.0')   # 优化前 20→5 (55+25+15+5=100)
            edit_cell(t, ri, 2, '10.0')  # 优化后 30→10
            ops.append("奶酪及其他 优化前20→5 优化后30→10")
        if name == '功能性乳制品':
            edit_cell(t, ri, 2, '20.0')  # 优化后 25→20 (40+30+20+10=100)
            ops.append("功能性乳制品 优化后25→20")
    return ops

# ── C. 破表重建 ──
def c_方舒55(ts):
    t = find_table(ts, '指标类别')
    if not t: return ["未找到"]
    rename_headers(t, {'占比': '比例(%)'})
    names = {1: '员工总数', 2: '管理人员', 3: '专业技术人员', 4: '项目施工人员',
             5: '辅助及后勤人员', 6: '本科及以上学历人员', 7: '35岁以下青年员工'}
    for r, nm in names.items():
        edit_cell(t, r, 0, nm)
    return ["行名重建7行+表头占比→比例(%) (岗位36+27+30+7=100, 另学历32%/年龄24%两指标行)"]

def c_朱香妃69(ts):
    t = find_table(ts, '绩效评估覆盖率')
    if not t: return ["未找到"]
    edit_cell(t, 2, 0, '2022')
    edit_cell(t, 3, 0, '2023')
    return ["年份列重建(2021/2022/2023, 原95/42串行)"]

def c_张扬56(ts):
    t = find_table(ts, '外籍/海归占比')
    if not t: return ["未找到"]
    rename_headers(t, {'外籍/海归占比': '外籍/海归占比(%)', '本土培养占比': '本土培养占比(%)'})
    edit_cell(t, 1, 0, '总部副总裁及以上')
    edit_cell(t, 2, 0, '区域总经理')
    edit_cell(t, 3, 0, '门店店长')
    for r, v in ((1, ('65.0', '35.0')), (2, ('30.0', '70.0')), (3, ('2.0', '98.0'))):
        edit_cell(t, r, 1, v[0]); edit_cell(t, r, 2, v[1])
    return ["管理层级拆行(高管/区域总经理/门店店长) + 小数比例转百分点(0.65→65.0)"]

def c_许明50(ts):
    t = find_table(ts, '认为通道单一', first_cell='指标')
    if not t: return ["未找到"]
    rename_headers(t, {'占比(%)': '比例(%)'})
    n = 0
    for ri in range(1, len(t.rows)):
        txt = t.rows[ri].cells[0].text
        if ';' in txt or '；' in txt:
            clean = re.split(r"[;；]", txt)[-1].strip()
            edit_cell(t, ri, 0, clean); n += 1
    return [f"行名去污染{n}行+表头→比例(%)"]

def c_许明62(ts):
    t = find_table(ts, '下班后常收到', first_cell='指标')
    if not t: return ["未找到"]
    rename_headers(t, {'占比(%)': '认同率(%)'})
    n = 0
    for ri in range(1, len(t.rows)):
        txt = t.rows[ri].cells[0].text
        if ';' in txt or '；' in txt:
            clean = re.split(r"[;；]", txt)[-1].strip()
            edit_cell(t, ri, 0, clean); n += 1
    return [f"行名去污染{n}行+表头→认同率(%)"]

JOBS = [
 ('25/经管/刘晓玲_029822430013.docx', a_刘晓玲),
 ('25/经管/宋妞妞_029823430767.docx', a_宋妞妞),
 ('25/经管/徐烨_029823430224.docx', a_徐烨),
 ('25/经管/徐琛雁_029823430961.docx', a_徐琛雁),
 ('25/经管/方舒_029823430683.docx', a_方舒_93),
 ('25/经管/方舒_029823430683.docx', c_方舒55),
 ('25/经管/朱香妃_029822430815.docx', c_朱香妃69),
 ('25/经管/李昕玥_029823430513.docx', a_李昕玥),
 ('25/经管/陆骏杰_029823431249.docx', b_陆骏杰),
 ('25/经管/陈新_029823431409.docx', a_陈新80),
 ('25/经管/魏添添_029823430240.docx', a_魏添添),
 ('26/经管/许明_029822430926.docx', c_许明50),
 ('26/经管/许明_029822430926.docx', c_许明62),
 ('26/经管/侍继典_029822430978.docx', a_侍继典65),
 ('26/经管/侍继典_029822430978.docx', a_侍继典100),
 ('26/经管/刘灿青_029823430328.docx', a_刘灿青),
 ('26/经管/崔宇蓉_029823430177.docx', a_崔宇蓉),
 ('26/经管/张扬_029823431494.docx', c_张扬56),
 ('26/经管/李好_029823430288.docx', None),   # 审计器豁免, 无需改
 ('26/经管/柳晴蓉_029823431251.docx', None),
 ('26/经管/胡吉阳_029822430169.docx', None),
 ('26/经管/郁洪斌_029823430461.docx', None),
 ('26/经管/王友航_029823431049.docx', a_王友航),
 ('26/经管/豆菲_029823431261.docx', a_豆菲),
 ('26/经管/黄雪梅_029822430246.docx', a_黄雪梅),
]
done = set()
for rel, fn in JOBS:
    if fn is None: continue
    key = (rel, fn.__name__)
    if key in done: continue
    done.add(key)
    fix(f'{ROOT}/{rel}', fn)
print(f"\n完成 {len(log_all)} 项修改, 备份在 {BAK2}/")
