#!/usr/bin/env python3
"""
扫描论文终版中的每篇过长标题论文的封面，读取摘要第一句，
根据摘要内容拟定新标题，再修改封面和更新状态表。

处理流程：
1. 找到论文文件
2. 读取封面/摘要，分析论文主题
3. 拟定标题（XXX问题研究-以XXX为例格式）
4. 修改封面
5. 更新状态表
"""
import os, re, sys
from pathlib import Path
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import openpyxl

BASE = "/root/project/workspace/论文资料/整理"
XLSX = f"{BASE}/学位申请合并_含论文状态.xlsx"

# === 待处理列表 ===
# 从上一轮扫描结果得来
CANDIDATES = [
    # (姓名, 准考证号, 专业)
    ("顾安奇", "29820430393", "旅游管理"),
    ("王刚", "29823410665", "环境设计"),
    ("唐心怡", "29822440013", "财务管理"),
    ("柳萌", "29822440011", "财务管理"),
    ("巩校鎛", "29822440007", "财务管理"),
    ("刘傲", "29822440003", "财务管理"),
    ("施雅琪", "29822440012", "财务管理"),
    ("周子杰", "29822440016", "财务管理"),
    ("唐卓妍", "29822440002", "财务管理"),
    ("卞妤柯", "29822440017", "财务管理"),
    ("方慧", "29822440005", "财务管理"),
    ("周童", "29823430793", "人力资源管理"),
    ("李妍妍", "29822430644", "金融学"),
    ("侍继典", "29822430978", "人力资源管理"),
    ("崔宇蓉", "29823430177", "人力资源管理"),
]

# 搜索结果
def find_paper_file(name, sid):
    """在论文终版中查找文件。"""
    for batch in ["25", "26"]:
        for major in ["经管", "机械", "设计"]:
            d = os.path.join(BASE, "论文终版", batch, major)
            if not os.path.exists(d):
                continue
            for f in os.listdir(d):
                if sid in f and f.endswith('.docx'):
                    return os.path.join(d, f)
    return None

def read_abstract(doc):
    """从封面和摘要读取信息。"""
    title_raw = ""
    abstract = ""
    keywords = ""
    
    for i, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        if '论文题目' in text and not title_raw:
            title_raw = text
        elif text == '摘要' and i + 1 < len(doc.paragraphs):
            abstract = doc.paragraphs[i+1].text.strip()
        elif text.startswith('关键词') or text.startswith('Keywords'):
            keywords = text
    
    return title_raw, abstract, keywords

def generate_title(name, abstract, major):
    """根据摘要内容拟定标题。
    格式：XXX问题研究-以XXX为例
    """
    if not abstract:
        return None
    
    # 分析摘要中的关键信息
    # 找研究对象/公司名称
    company_patterns = [
        r'以(.+?)(?:为(?:案例|研究对象|例|样本|背景))',
        r'选取(.+?)(?:为(?:案例|研究对象|例))',
        r'以\s*(.+?)\s*为\s*(?:案例|研究对象|例)',
        r'(?:针对|聚焦|基于)\s*(.+?)(?:的|,)',
    ]
    
    company = None
    for pat in company_patterns:
        m = re.search(pat, abstract[:200])
        if m:
            company = m.group(1).strip()
            # 清理
            company = re.sub(r'[，,。]', '', company)[:30]
            break
    
    # 分析主题关键词
    theme_keywords = []
    theme_patterns = [
        r'(财务风险|融资风险|投资风险|营运资金管理|内部控制|成本管理|现金流)',
        r'(绩效管理|培训体系|招聘问题|薪酬管理|职业发展|职业生涯管理|人力资源管理)',
        r'(营销策略|品牌传播|客户关系|市场策略)',
        r'(风险管理|信用风险|信贷风险)',
    ]
    
    for pat in theme_patterns:
        m = re.search(pat, abstract[:300])
        if m:
            theme_keywords.append(m.group(1))
    
    # 专业对应的默认主题
    major_themes = {
        "财务管理": "财务风险控制",
        "人力资源管理": "人力资源管理优化",
        "金融学": "风险管理",
        "旅游管理": "旅游管理",
        "环境设计": "环境设计",
    }
    
    # 构建标题
    theme = theme_keywords[0] if theme_keywords else major_themes.get(major, "管理")
    
    if company and len(company) >= 3:
        title = f"{theme}问题研究-以{company}为例"
    else:
        title = f"{theme}问题研究"
    
    return title

def update_cover(docx_path, new_title):
    """修改封面论文题目。"""
    doc = Document(docx_path)
    modified = False
    
    for p in doc.paragraphs:
        text = p.text.strip()
        if '论文题目' in text:
            for r in p.runs:
                if '论文题目' in r.text:
                    r.text = f"论文题目：{new_title}"
                    modified = True
                else:
                    r.text = ''
            break
    
    if modified:
        doc.save(docx_path)
    return modified

# === 主流程 ===
results = []

for name, sid, major in CANDIDATES:
    print(f"\n{'='*60}")
    print(f"📄 {name}({major}) 准考证={sid}")
    
    filepath = find_paper_file(name, sid)
    if not filepath:
        print(f"  ⚠ 未找到论文文件")
        results.append((name, sid, None, "文件不存在"))
        continue
    
    print(f"  文件: {filepath}")
    
    doc = Document(filepath)
    title_raw, abstract, keywords = read_abstract(doc)
    
    # 生成新标题
    new_title = generate_title(name, abstract, major)
    if not new_title:
        print(f"  ⚠ 无法生成标题")
        results.append((name, sid, None, "无法生成标题"))
        continue
    
    print(f"  摘要首句: {abstract[:100]}...")
    print(f"  新标题: {new_title}")
    
    # 修改封面
    success = update_cover(filepath, new_title)
    if success:
        print(f"  ✅ 封面已修改")
    else:
        print(f"  ⚠ 封面修改失败")
    
    results.append((name, sid, new_title, "成功" if success else "封面修改失败"))

# === 更新状态表 ===
print(f"\n{'='*60}")
print(f"📊 更新状态表...")

wb = openpyxl.load_workbook(XLSX)
ws = wb['学位申请汇总']

updated = 0
for name, sid, new_title, status in results:
    if new_title is None:
        continue
    
    for r in range(2, ws.max_row + 1):
        sid_raw = ws.cell(r, 2).value
        if sid_raw and str(int(sid_raw)) == sid:
            ws.cell(r, 6).value = new_title
            ws.cell(r, 8).value = '无缺失'
            updated += 1
            print(f"  ✅ {name}: 论文名称→{new_title}")
            break

wb.save(XLSX)
print(f"\n✅ 共更新 {updated} 篇论文的状态表")
print(f"  其中成功修改封面: {sum(1 for _,_,t,s in results if t and s=='成功')}")
print(f"  跳过/失败的: {sum(1 for _,_,t,s in results if not t or s!='成功')}")
