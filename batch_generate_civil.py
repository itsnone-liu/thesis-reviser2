#!/usr/bin/env python3
"""
batch_generate_civil.py — 批量生成土木工程论文

用法：
  python3 batch_generate_civil.py [--single 姓名]

流程：
  1. 从学位表读取土木学生名单
  2. 根据预设题目列表生成每篇论文
  3. 调用LLM逐章生成 + 渲染DOCX + 生图
  4. 输出到 论文资料/整理/论文终版/{25|26}/土木/
"""
import os, sys, json, re, time, shutil, zipfile, uuid
from pathlib import Path
import openpyxl
from xml.etree import ElementTree as ET

# 加入项目路径
PROJECT = str(Path(__file__).resolve().parent)
sys.path.insert(0, PROJECT)

from core import (
    call_llm, clean_text,
    CIV_WORD_LIMITS, CIV_CHAPTERS, PAPER_TYPE_MAP,
    build_cover, normalize_cover_info,
    add_t,
    tolerant_extract_tables,
    extract_drawings_from_text, merge_mech_json_into_drawings,
    generate_single_image, classify_drawing_backend,
    txt_to_docx_safe,
    _init_s, _wt,
)
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from prompts.civil import (
    build_chapter_prompt, build_abstract_prompt,
    build_references_prompt, build_drawing_analysis_prompt,
    get_chapter_outline, generate_project_spec,
    CONSISTENCY_RULE, DRAWING_RULE, FMT_RULE, SCENE_LOCK_RULE
)

# ===== 配置 =====
BASE = "/root/project/workspace/论文资料/整理"
XLSX = f"{BASE}/学位申请合并.xlsx"
OUTPUT = f"{BASE}/论文终版"

# 土木工程23个题目（已和老师确认）
CIVIL_TITLES = [
    ("严涛", "29820430414", "南京市鼓楼区紫峰中学六层框架教学楼建筑与结构设计", "2025年12月"),
    ("何晋", "29822431015", "苏州市工业园区星海七层框架办公楼建筑与结构设计", "2026年6月"),
    ("冯国栋", "29823430874", "杭州市西湖区绿城桂花园十二层剪力墙住宅楼建筑与结构设计", "2026年6月"),
    ("冯然", "29822430107", "无锡市新吴区科技创业园五层框架综合楼建筑与结构设计", "2025年12月"),
    ("冯磊", "29822430252", "宁波市海曙区南苑八层框架-剪力墙结构酒店建筑与结构设计", "2025年12月"),
    ("刘德荣", "29822430813", "合肥市包河区第一人民医院六层框架门诊楼建筑与结构设计", "2025年12月"),
    ("刘飞", "29821430999", "武汉市洪山区银泰九层框架商业综合体建筑与结构设计", "2026年6月"),
    ("周亮", "29823430350", "杭州市滨江区钱江湾四层框架幼儿园建筑与结构设计", "2025年12月"),
    ("周振宇", "29822430523", "苏州市吴中区新城金郡高层住宅小区施工组织设计", "2026年6月"),
    ("席勤峰", "29822430087", "常州市新北区创意产业园框架结构办公楼施工组织设计及进度控制", "2026年6月"),
    ("张德轩", "29821431063", "南京市江宁区同仁医院综合楼施工组织设计与质量管理", "2025年12月"),
    ("张超", "29823431459", "无锡市滨湖区太湖高级中学教学楼施工组织设计与安全控制", "2025年12月"),
    ("曹龙华", "29823430221", "杭州市拱墅区香积寺路框架结构住宅楼工程造价控制研究", "2025年12月"),
    ("杨晓", "29822430133", "苏州市姑苏区观前街商业广场高层建筑项目全过程造价管理研究", "2026年6月"),
    ("杨轩", "29823430891", "宁波市鄞州区南部商务区工程建设项目质量管理体系研究", "2026年6月"),
    ("柳骁蕾", "29823430287", "合肥市高新区创新产业园装配式建筑项目施工管理优化研究", "2025年12月"),
    ("王尧", "29823431344", "杭州市萧山区钱江世纪城高层建筑基础选型与设计", "2025年12月"),
    ("王楠", "29822430386", "南京市河西新城软土地区深基坑支护方案设计", "2025年12月"),
    ("管伟琪", "29823430248", "苏州市工业园区生物医药产业园工程桩基础设计与施工方案研究", "2026年6月"),
    ("袁园", "29821430002", "宁波市北仑区泰山路城市道路工程路基路面设计", "2025年12月"),
    ("邵喆", "29820430090", "常州市武进区大运河跨河简支梁桥上部结构设计", "2025年12月"),
    ("金凯", "29823431399", "上海市浦东新区世博展厅钢结构建筑设计与稳定性分析", "2025年12月"),
    ("陈顾磊", "29823431467", "南京市秦淮区老门东历史街区既有建筑结构安全性鉴定与加固设计", "2025年12月"),
]

BATCH_MAP = {"2025年12月": "25", "2026年6月": "26"}

# 土木工程指导老师（从之前汇总表来的）
CIVIL_TEACHERS = ["偶丹萍", "陈海军", "赵巧明", "庄红军", "冯林", "王永健", "孙乃立", "廖祥兵"]

# 论文类型判断
def classify_paper(title):
    if "施工组织设计" in title:
        return "施工组织"
    elif "造价" in title or "成本" in title:
        return "造价"
    elif "质量管理" in title or "管理" in title:
        return "管理"
    elif "基础" in title or "基坑" in title or "桩基" in title:
        return "基础"
    elif "道路" in title or "路基" in title:
        return "道路"
    elif "桥梁" in title or "简支梁" in title:
        return "桥梁"
    elif "钢结构" in title or "鉴定" in title or "加固" in title:
        return "鉴定加固"
    else:
        return "建筑结构"


# ===== 主生成函数 =====
def generate_paper(name, sid, title, batch, output_dir, teacher=""):
    """生成一篇土木论文"""
    print(f"\n{'='*60}")
    print(f"生成: {name} | {sid}")
    print(f"题目: {title}")
    print(f"{'='*60}")
    
    paper_type = classify_paper(title)
    chapters = get_chapter_outline({"title": title})
    
    # 构建profile
    profile = {
        "title": title,
        "name": name,
        "student_id": sid,
        "major": "土木工程",
        "level": "专升本",
        "advisor": teacher,  # 传入教师，封面会显示
        "year": "2025" if batch == "2025年12月" else "2026",
        "month": "12" if batch == "2025年12月" else "5",
        "day": "1",
        "project_type": paper_type,
        "object_name": title,
        "structure_type": "框架结构",
        "paper_type": "土木",
    }
    
    if "剪力墙" in title: profile["structure_type"] = "剪力墙结构"
    elif "框架-剪力墙" in title: profile["structure_type"] = "框架-剪力墙结构"
    elif "钢结构" in title: profile["structure_type"] = "钢结构"
    
    # 生成project_spec用于数据一致性 - 先调用LLM生成完整参数
    print(f"\n  [工程参数] 正在生成建筑技术参数...")
    spec_prompt = f"""你是一位土木工程总工。请根据以下论文题目，给出该工程唯一确定的核心设计参数。

论文题目：{title}
结构形式：{profile.get("structure_type", "框架结构")}

请以JSON格式返回，只输出JSON，不要任何其他文字：
{{
    "building_scale": "总建筑面积（m²）",
    "structure_type": "{profile.get("structure_type", "框架结构")}",
    "stories": "层数（如"6层"）",
    "height": "建筑总高度（m）",
    "floor_height": "标准层层高（m）",
    "grid_size": "主要柱网尺寸（m×m）",
    "seismic": "抗震设防烈度和措施（如"7度设防，三级抗震"）",
    "design_life": "设计使用年限",
    "safety_level": "安全等级",
    "foundation": "基础形式和地基承载力",
    "wind_load": "基本风压（kN/m²）",
    "snow_load": "基本雪压（kN/m²）",
    "concrete_grade": "混凝土强度等级",
    "steel_grade": "钢筋等级",
    "critical_params": ["关键参数列表，如建筑面积、层高、总工期等"],
    "duration": "施工总工期（天）"
}}

注意：参数必须符合工程实际，数据在合理范围内。对于无法从标题确定的参数，按常规中等规模同类工程给出合理值。"""
    spec_raw = call_llm(spec_prompt, max_tokens=1000)
    if spec_raw:
        spec_raw = clean_text(spec_raw)
        # 尝试提取JSON
        import json as _json
        try:
            _json.loads(spec_raw)  # 验证JSON
            project_spec = generate_project_spec(profile, spec_raw)
        except:
            import re as _re
            m = _re.search(r'```(?:json)?\n(.*?)\n```', spec_raw, _re.DOTALL)
            if m:
                try:
                    _json.loads(m.group(1))
                    project_spec = generate_project_spec(profile, m.group(1))
                except:
                    project_spec = generate_project_spec(profile)
            else:
                project_spec = generate_project_spec(profile)
    else:
        project_spec = generate_project_spec(profile)
    print(f"  → 工程参数已生成")
    
    # ===== 摘要 =====
    print(f"\n  [摘要] 正在生成...")
    abstract_prompt = build_abstract_prompt(profile)
    abstract_text = call_llm(abstract_prompt, max_tokens=1000)
    abstract_text = clean_text(abstract_text) if abstract_text else ""
    
    if not abstract_text:
        print("  ⚠ 摘要生成为空，跳过本篇")
        return False
    
    # 去除摘要正文中可能多余的"摘要："前缀
    abstract_text = re.sub(r'^摘要[：:]\s*', '', abstract_text.strip())
    
    # 提取关键词
    keywords = ""
    if "关键词" in abstract_text:
        parts = abstract_text.rsplit("关键词", 1)
        abstract_text = parts[0].strip()
        keywords = "关键词" + parts[1] if len(parts) > 1 else ""
    
    # ===== 逐章生成 =====
    all_chapters = []
    all_tables = []
    full_text = ""

    for ch_name, ch_num in chapters:
        print(f"\n  [第{ch_num}章] {ch_name} 正在生成...")
        prompt = build_chapter_prompt(ch_name, ch_num, profile, chapters, full_text)
        
        # 注入project_spec一致性约束
        if project_spec:
            try:
                spec = json.loads(project_spec)
            except:
                spec = {}
            extra_rules = CONSISTENCY_RULE
            extra_rules = extra_rules.replace("{building_scale}", spec.get("building_scale", ""))
            extra_rules = extra_rules.replace("{structure_type}", spec.get("structure_type", ""))
            extra_rules = extra_rules.replace("{stories}", spec.get("stories", ""))
            extra_rules = extra_rules.replace("{seismic}", spec.get("seismic", ""))
            extra_rules = extra_rules.replace("{foundation}", spec.get("foundation", ""))
            extra_rules = extra_rules.replace("{critical_params}", json.dumps(spec.get("critical_params", []), ensure_ascii=False))
            extra_rules = extra_rules.replace("{duration}", str(spec.get("duration", "")))
            
            # 第1章也注入project_spec约束，让LLM从源头统一参数
            prompt = prompt + "\n\n" + extra_rules
        
        # 土木：所有图走drawing标签，不追加CHART_RULE
        if ch_name in ["工程概况", "建筑设计", "结构设计", "结构计算", "施工组织设计", "施工方案", "施工进度计划"]:
            prompt = prompt + "\n\n" + DRAWING_RULE
        
        prompt = prompt + "\n\n" + FMT_RULE
        
        chapter_text = call_llm(prompt, max_tokens=4096)
        chapter_text = clean_text(chapter_text) if chapter_text else ""
        
        if chapter_text:
            # 1) 去除LLM可能自带的重复章节标题（如"第X章 xxx"或仅"xxx"出现在第一行）
            lines = chapter_text.strip().split('\n')
            cleaned_lines = []
            skip_first_line = False
            if lines:
                first = lines[0].strip()
                # 情况A：LLM输出了"第X章 xxx"
                if re.search(r'^第[一二三四五六七八九十\d]章', first):
                    skip_first_line = True
                # 情况B：LLM输出了纯章节名（和ch_name相同或包含ch_name）
                elif first == ch_name or ch_name in first:
                    # 但要去掉像"第3章 结构设计"这样的误伤
                    pass  # 暂时保留，后面再处理
            
            for i, line in enumerate(lines):
                if i == 0 and skip_first_line:
                    continue
                # 修复不完整的drawing标签
                line = re.sub(r'<drawing\s*/\s*$', '<drawing/>', line)
                line = re.sub(r'<drawing[^>]*?/\s*$', lambda m: m.group(0).rstrip('/ ') + '/>', line)
                cleaned_lines.append(line)
            chapter_text = '\n'.join(cleaned_lines)
            
            # 2) 在正文中如果出现"第X章 xxx"后面紧跟"xxx"的行，删除重复的纯标题行
            chapter_text = re.sub(
                rf'第{ch_num}章\s*{re.escape(ch_name)}\s*\n\s*{re.escape(ch_name)}\s*',
                f'第{ch_num}章 {ch_name}\n',
                chapter_text
            )
            
            # 3) 格式化为章节（不再添加"第X章"前缀，因为LLM可能已经输出）
            formatted = f"第{ch_num}章 {ch_name}\n\n{chapter_text}\n\n"
            all_chapters.append(formatted)
            full_text += formatted
            
            # 收集表格标签
            tables = tolerant_extract_tables(chapter_text)
            all_tables.extend(tables)
            
            print(f"    → {len(chapter_text)}字, 图表: 0个(GPT生图), 表格: {len(tables)}个")
        else:
            print(f"    ⚠ 章节生成为空")
    
    # ===== 参考文献 =====
    print(f"\n  [参考文献] 正在生成...")
    ref_prompt = build_references_prompt(profile)
    ref_text = call_llm(ref_prompt, max_tokens=2000)
    ref_text = clean_text(ref_text) if ref_text else ""
    
    # ===== 组装完整文本（原代码期望格式：摘要→PAGE_BREAK→目录→PAGE_BREAK→正文→参考文献）=====
    combined_chapters = "".join(all_chapters)
    
    # 修复drawing标签问题
    # 1) 修复跨行未闭合的drawing标签（如 <drawing ...\n... 没有/>结尾的）
    combined_chapters = re.sub(
        r'(<drawing\s[^>]*?)(?:\n[^<]*?)(/?>)?',
        lambda m: m.group(1) + '/>' if not m.group(2) else m.group(0),
        combined_chapters
    )
    # 2) 对没有type属性的drawing标签，根据title推断type
    def fix_drawing_type(m):
        tag = m.group(0)
        if 'type=' not in tag:
            title_m = re.search(r'title="([^"]*)"', tag)
            if title_m:
                title = title_m.group(1)
                if '平面' in title or '立面' in title: new_type = '建筑图'
                elif '结构' in title: new_type = '结构图'
                elif '基础' in title: new_type = '基础图'
                elif '配筋' in title: new_type = '配筋图'
                elif '施工' in title: new_type = '施工图'
                else: new_type = '工程图'
                tag = tag.replace('<drawing', f'<drawing type="{new_type}"')
        return tag
    combined_chapters = re.sub(r'<drawing\s[^>]*?/?>', fix_drawing_type, combined_chapters)
    # 3) 删除裸标签 <drawing/>（无任何属性）
    combined_chapters = re.sub(r'<drawing\s*/>\s*\n?', '', combined_chapters)
    
    # 收集所有章节标题和二级标题用于目录
    toc_entries = []
    for ch_name, ch_num in chapters:
        toc_entries.append(f"第{ch_num}章 {ch_name}")
    # 从已生成的全文提取二级标题
    for m in re.finditer(r'^(\d+\.\d+)\s+(.+?)$', combined_chapters, re.MULTILINE):
        toc_entries.append(f"    {m.group(1)} {m.group(2).strip()}")
    
    toc_text = "\n".join(toc_entries) + "\n参考文献"
    
    full_paper = f"摘要\n\n{abstract_text}\n\n"
    if keywords:
        full_paper += f"{keywords}\n\n"
    full_paper += f"---PAGE_BREAK---\n目录\n{toc_text}\n---PAGE_BREAK---\n"
    full_paper += combined_chapters
    full_paper += f"\n参考文献\n{ref_text}\n"
    
    # ===== 后处理清理 =====
    # 1. 去除"图 description=..."这种错误格式
    full_paper = re.sub(r'(?m)^\s*图\d*\s*description\s*=\s*"[^"]*"\s*$', '', full_paper)
    # 2. 修复"图X 图X-X"重复前缀 → "图X-X"（LLM有时在正文中也写了）
    full_paper = re.sub(r'(?m)^(图\d+)\s+(图[\d-]+\s)', r'\2', full_paper)
    # 也修复drawing标签内title的"图1 图2-1"问题
    full_paper = re.sub(r'(title=")图\d+\s+(图[\d-]+\s)', r'\1\2', full_paper)
    # 3. 去除多余的"图X"（单独成行且后面没跟具体内容的）
    full_paper = re.sub(r'(?m)^图\d+\s*$', '', full_paper)
    # 4. 确保每个drawing标签的title带"图X-X"格式（去掉"图X 图X-X"中的多余部分）
    full_paper = re.sub(
        r'(<drawing[^>]*title=")(?:图\d+\s*)?(图[\d-]+\s[^"]*)(")',
        r'\1\2\3', full_paper
    )
    
    # 5. 将在LLM输出中可能残留的 <chart> 标签转换为 <drawing> 标签
    #    （所有图统一走drawing，chart被废弃）
    full_paper = re.sub(
        r'<chart\s+([^>]*?)/>',
        lambda m: '<drawing ' + m.group(1) + ' type="chart_auto"/>',
        full_paper
    )
    
    # ===== 生成图片（GPT生图API）=====
    print(f"\n  [图片] 正在用GPT生图...")
    img_dir = os.path.join(output_dir, "images2")
    os.makedirs(img_dir, exist_ok=True)
    from core import tolerant_extract_drawings, generate_single_image
    from concurrent.futures import ThreadPoolExecutor, as_completed
    drawings_list = tolerant_extract_drawings(full_paper)
    drawing_images = {}
    if drawings_list:
        total = len(drawings_list)
        print(f"  [图片] 共 {total} 张待生成，并发数=3")
        # 为每张图分配唯一seq（先全部预分配，避免并发竞争）
        for d in drawings_list:
            raw_seq = str(d.get("seq", d.get("id", str(uuid.uuid4().hex[:8]))))
            d["_seq"] = raw_seq
        # 用线程池并发生图，控制并发数=3
        start_time = time.time()
        completed = 0
        failed = 0
        with ThreadPoolExecutor(max_workers=3) as pool:
            fut_map = {pool.submit(generate_single_image, d, img_dir): d for d in drawings_list}
            for fut in as_completed(fut_map):
                d = fut_map[fut]
                seq = d["_seq"]
                result_path = fut.result()
                if result_path and os.path.exists(result_path):
                    key = os.path.basename(result_path).replace(".png", "").split("_")[-1]
                    if key.isdigit():
                        drawing_images[str(int(key))] = result_path
                    completed += 1
                    elapsed = time.time() - start_time
                    avg = elapsed / completed if completed > 0 else 0
                    eta = (total - completed) * avg / 3  # 除并发数以估算剩余时间
                    title = d.get("title", "?")
                    print(f"    ✓ [{completed}/{total}] {os.path.basename(result_path)} ({title}) | 已用{elapsed:.0f}s 预估剩余{eta:.0f}s")
                else:
                    failed += 1
                    print(f"    ✗ [{completed+failed}/{total}] seq={seq} 生成失败")
        total_time = time.time() - start_time
        print(f"  [图片] 完成: {completed}张成功, {failed}张失败, 耗时{total_time:.0f}s ({total_time/60:.1f}min)")
    else:
        print(f"  [图片] 无需生成")
    
    # ===== 渲染DOCX =====
    print(f"\n  [DOCX] 正在渲染...")
    try:
        docx_path = render_to_docx(profile, full_paper, all_tables, output_dir, name, sid)
        print(f"  ✅ DOCX生成成功: {docx_path}")
        
        # 保存TXT
        txt_path = os.path.join(output_dir, f"{name}_{sid}.txt")
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write(full_paper)
        print(f"  ✅ TXT保存成功: {txt_path}")
        
        return True
    except Exception as e:
        print(f"  ❌ DOCX渲染失败: {e}")
        import traceback
        traceback.print_exc()
        return False


def render_to_docx(profile, full_text, tables, output_dir, name, sid):
    """渲染DOCX — 完全复用原代码的txt_to_docx_safe管道"""
    # 1. 保存TXT
    txt_path = os.path.join(output_dir, f"{name}_{sid}.txt")
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write(full_text)
    
    # 2. 加载已有图片（已经由generate_paper中的GPT生图API预生成）
    drawing_images = {}
    img_dir = os.path.join(output_dir, "images2")
    if os.path.isdir(img_dir):
        for fname in sorted(os.listdir(img_dir)):
            if fname.startswith("drawing_") and fname.endswith(".png"):
                key = fname.replace("drawing_", "").replace(".png", "")
                key = str(int(key))
                drawing_images[key] = os.path.join(img_dir, fname)
        print(f"  [图片] 已加载 {len(drawing_images)} 张图")
    
    # 3. 构建封面信息
    cover_info = {
        "title": profile.get("title", ""),
        "name": profile.get("name", ""),
        "student_id": profile.get("student_id", ""),
        "major": profile.get("major", "土木工程"),
        "level": profile.get("level", "专升本"),
        "advisor": profile.get("advisor", ""),
        "year": profile.get("year", "2025"),
        "month": profile.get("month", "12"),
        "day": profile.get("day", "1"),
    }
    
    # 4. 用原管道渲染
    docx_path = os.path.join(output_dir, f"{name}_{sid}.docx")
    txt_to_docx_safe(txt_path, docx_path, cover_info=cover_info, drawing_images=drawing_images)
    
    return docx_path


def set_run_font(run, font_name='宋体', size=12, bold=False, color=None):
    """设置run字体"""
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color


# ===== 主程序 =====
def main():
    import random
    random.seed(42)
    
    # 只处理指定的一篇
    single_mode = None
    if "--single" in sys.argv:
        idx = sys.argv.index("--single")
        if idx + 1 < len(sys.argv):
            single_mode = sys.argv[idx + 1]
    
    # 过滤要处理的
    to_process = CIVIL_TITLES
    if single_mode:
        to_process = [t for t in CIVIL_TITLES if t[0] == single_mode]
        if not to_process:
            print(f"未找到学生: {single_mode}")
            return
    
    print(f"待处理论文: {len(to_process)} 篇\n")
    
    # 教师分配
    random.shuffle(CIVIL_TEACHERS)
    
    results = {"success": 0, "fail": 0}
    for i, (name, sid, title, batch) in enumerate(to_process):
        teacher = CIVIL_TEACHERS[i % len(CIVIL_TEACHERS)]
        batch_dir = BATCH_MAP.get(batch, "25")
        target_dir = os.path.join(OUTPUT, batch_dir, "土木")
        os.makedirs(target_dir, exist_ok=True)
        
        output_dir = os.path.join("/tmp/civil_papers", f"{name}_{sid}")
        os.makedirs(output_dir, exist_ok=True)
        
        ok = generate_paper(name, sid, title, batch, output_dir, teacher)
        
        if ok:
            # 复制成品到目标目录
            src = os.path.join(output_dir, f"{name}_{sid}.docx")
            dst = os.path.join(target_dir, f"{name}_{sid}.docx")
            if os.path.exists(src):
                shutil.copy2(src, dst)
                print(f"  → 已复制到: {dst}")
            results["success"] += 1
        else:
            results["fail"] += 1
    
    print(f"\n{'='*60}")
    print(f"批量生成完成！")
    print(f"  成功: {results['success']} 篇")
    print(f"  失败: {results['fail']} 篇")
    print(f"输出目录: {OUTPUT}/{{25|26}}/土木/")


if __name__ == "__main__":
    main()
