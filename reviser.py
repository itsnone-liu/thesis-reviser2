# -*- coding: utf-8 -*-
"""
reviser.py — 论文修改管道
========================
读取已有论文DOCX → 分析画像 → LLM诊断 → 生成修改版

使用：
  python reviser.py --input 原论文.docx --output ./revised/ --type 管理
  python reviser.py --input 原论文.docx --output ./revised/ --type 设计
"""
import sys, os, re, json, time, shutil, uuid
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import argparse

from core import (
    call_llm, extract_docx_text, extract_cover_info,
    save_diagnosis_report, MG_CHAPTERS, SJ_CHAPTERS, MC_CHAPTERS,
)
from generator import generate


# ==================== 论文画像分析 ====================
def _analyze_original_paper(text: str, paper_type: str = "管理") -> dict:
    """调用LLM分析原文，提取研究画像"""
    if paper_type in ("管理", "经管", "manage", "mg"):
        chapters = MG_CHAPTERS
        chapter_names = [n for n, _ in chapters]
        fields = """
  "title": "论文标题",
  "major": "专业",
  "company": "研究对象/公司",
  "industry": "所属行业",
  "core_problems": ["问题1", "问题2", "问题3"],
  "data_hints": ["数据线索1", "线索2", "线索3"],
  "outline": {
    "引言": "...",
    "国内外研究现状": "...",
    "现状分析": "...",
    "问题与原因分析": "...",
    "解决方案": "...",
    "结论": "..."
  },
  "theories": ["理论1", "理论2"],
  "charts_mentioned": ["图表1描述"],
  "data_quality_note": "数据质量评价"
"""
    elif paper_type in ("机械", "mechanical", "mech", "mj"):
        chapter_names = [n for n, _ in MC_CHAPTERS]
        fields = """
  "title": "论文标题",
  "major": "机械专业",
  "company": "设计对象/装置",
  "mech_object": "机械设计对象",
  "mech_type": "夹具/传动/送料/搬运/升降/结构优化",
  "working_condition": "工况",
  "technical_params": {"载荷": "数值+单位", "行程": "数值+单位", "精度": "数值+单位"},
  "materials": "材料/热处理",
  "core_problems": ["问题1", "问题2", "问题3"],
  "data_hints": ["参数1", "参数2", "参数3"],
  "calc_items": ["校核项1", "校核项2", "校核项3"],
  "outline": {
    "概述": "...",
    "原理与方案分析": "...",
    "总体设计": "...",
    "关键部件设计": "...",
    "计算与校核": "...",
    "总结": "..."
  },
  "theories": ["理论1", "理论2"],
  "data_quality_note": "参数/图纸质量评价"
"""
    else:
        chapter_names = ["绪论", "理论基础", "问题发现与分析", "设计策略与方案", "总结与反思"]
        fields = """
  "title": "论文标题",
  "major": "设计专业",
  "design_type": "设计专业类型",
  "design_object": "设计对象",
  "context": "使用场景/背景",
  "target_users": ["用户群体1", "用户群体2"],
  "core_problems": ["问题1", "问题2", "问题3"],
  "design_strategies": ["策略1", "策略2", "策略3"],
  "design_elements": ["要素1", "要素2", "要素3"],
  "outline": {
    "绪论": "...",
    "理论基础": "...",
    "问题发现与分析": "...",
    "设计策略与方案": "...",
    "总结与反思": "..."
  },
  "theories": ["理论1", "理论2"],
  "data_quality_note": "数据/图纸质量评价"
"""

    prompt = f"""你是一名学术论文分析专家。请对以下论文全文进行深度分析，输出严格JSON格式（不要markdown代码块，不要任何其他文字）。

论文全文：
{text[:8000]}

请输出以下字段的JSON：
{fields}

只输出JSON，严禁输出其他内容。"""

    for attempt in range(3):
        response = call_llm(prompt, max_tokens=3000)
        if not response:
            time.sleep(1)
            continue
        try:
            json_match = re.search(r'\{.*\}', response, re.DOTALL)
            if json_match:
                result = json.loads(json_match.group())
                # 补齐缺失字段
                if "title" not in result:
                    result["title"] = text[:50].replace('\n', ' ').strip() if text else "未识别"
                if "major" not in result:
                    if paper_type in ("管理", "经管"):
                        result["major"] = "工商管理"
                    elif paper_type in ("机械", "mechanical", "mech", "mj"):
                        result["major"] = "机械设计"
                    else:
                        result["major"] = "设计学"
                for name in chapter_names:
                    if "outline" not in result:
                        result["outline"] = {}
                    if name not in result.get("outline", {}):
                        result["outline"][name] = ""
                return result
        except Exception as e:
            print(f"分析失败（尝试{attempt+1}）：{e}")
            time.sleep(1)
            continue
    # 兜底
    return {"title": text[:50].replace('\n', ' ').strip() or "未识别",
            "major": "工商管理" if paper_type in ("管理", "经管") else ("机械设计" if paper_type in ("机械", "mechanical", "mech", "mj") else "设计学"),
            "outline": {n: "" for n in chapter_names}}


# ==================== 深度诊断 ====================
def _diagnose_and_reconstruct(profile: dict, paper_type: str = "管理") -> dict:
    """调用LLM诊断，输出修订计划"""
    if paper_type in ("管理", "经管", "manage", "mg"):
        chapters = MG_CHAPTERS
        chapter_names = [n for n, _ in chapters]
    elif paper_type in ("机械", "mechanical", "mech", "mj"):
        chapter_names = [n for n, _ in MC_CHAPTERS]
    else:
        chapter_names = ["绪论", "理论基础", "问题发现与分析", "设计策略与方案", "总结与反思"]

    chapter_revisions_json = json.dumps(
        {n: {"issues": [], "fix_requirements": [], "key_points": []} for n in chapter_names},
        ensure_ascii=False, indent=2
    )

    prompt = f"""你是一名资深学术论文评审与修改专家。请基于以下论文画像进行深度诊断，并输出修订方案。

论文画像：
{json.dumps(profile, ensure_ascii=False, indent=2)}

请输出严格JSON格式（不要markdown代码块，不要任何其他文字）：
{{
  "diagnosis_report": {{
    "overall_score": "1-10分",
    "overall_comment": "总体评价（200字）",
    "issues": [
      {{
        "chapter": "涉及章节",
        "severity": "高/中/低",
        "issue_type": "理论错误/数据矛盾/逻辑混乱/结构失衡/表述问题/其他",
        "description": "问题描述",
        "fix_direction": "修正方向"
      }}
    ]
  }},
  "revised_profile": {{
    "title": "优化后的标题",
    "company": "公司",
    "industry": "行业",
    "core_problems": ["修正后问题1", "问题2", "问题3"],
    "data_hints": ["数据方向1", "方向2", "方向3"]
  }},
  "chapter_revisions": {chapter_revisions_json},
  "theory_adjustment": "理论框架调整说明",
  "data_strategy": "数据/图表策略",
  "special_notes": "生成新论文时的特殊注意事项"
}}

注意：chapter_revisions 中的每个章节都要有 issues、fix_requirements、key_points 三个字段。
如果某章节无问题，请留空列表[]。
只输出JSON，严禁输出其他内容。"""

    for attempt in range(3):
        response = call_llm(prompt, max_tokens=4000)
        if not response:
            time.sleep(1)
            continue
        try:
            json_match = re.search(r'\{.*\}', response, re.DOTALL)
            if json_match:
                result = json.loads(json_match.group())
                if "diagnosis_report" not in result:
                    result["diagnosis_report"] = {"overall_score": "5", "overall_comment": "默认评价", "issues": []}
                if "chapter_revisions" not in result:
                    result["chapter_revisions"] = {}
                for n in chapter_names:
                    if n not in result.get("chapter_revisions", {}):
                        result["chapter_revisions"][n] = {"issues": [], "fix_requirements": [], "key_points": []}
                    else:
                        for k in ["issues", "fix_requirements", "key_points"]:
                            if k not in result["chapter_revisions"][n]:
                                result["chapter_revisions"][n][k] = []
                return result
        except Exception as e:
            print(f"诊断失败（尝试{attempt+1}）：{e}")
            time.sleep(1)
            continue
    return {
        "diagnosis_report": {"overall_score": "5", "overall_comment": "默认诊断", "issues": []},
        "chapter_revisions": {n: {"issues": [], "fix_requirements": [], "key_points": []} for n in chapter_names}
    }


def _generate_revised(profile: dict, revision_plan: dict, paper_type: str, update=None) -> str:
    """基于修订计划生成修改版论文文本"""
    if update is None:
        def update(msg, prog): print(f"[{prog}%] {msg}")

    # 获取修订后的画像
    rp = revision_plan.get("revised_profile", profile)
    if paper_type in ("管理", "经管", "manage", "mg"):
        rp["major"] = rp.get("major", profile.get("major", "工商管理"))
        rp["title"] = rp.get("title", profile.get("title", ""))
        rp["company"] = rp.get("company", profile.get("company", ""))
        rp["industry"] = rp.get("industry", profile.get("industry", ""))
    elif paper_type in ("机械", "mechanical", "mech", "mj"):
        rp["major"] = rp.get("major", profile.get("major", "机械设计"))
        rp["title"] = rp.get("title", profile.get("title", ""))
        rp["company"] = rp.get("company", profile.get("company", ""))
        rp["mech_object"] = rp.get("mech_object", profile.get("mech_object", profile.get("company", "")))
        rp["mech_type"] = rp.get("mech_type", profile.get("mech_type", ""))
        rp["working_condition"] = rp.get("working_condition", profile.get("working_condition", ""))
    else:
        rp["design_object"] = rp.get("design_object", profile.get("design_object", ""))
        rp["design_type"] = rp.get("design_type", profile.get("design_type", ""))

    # 直接调用 generator.py 生成
    update("正在调用生成管道重写论文...", 20)
    txt = generate(rp, paper_type, update)
    return txt


def revise(input_docx: str, output_dir: str, paper_type: str = "管理", update=None):
    """
    论文修改主流程
    - input_docx: 输入DOCX路径
    - output_dir: 输出目录
    - paper_type: "管理" / "设计" / "机械"
    """
    if update is None:
        def update(msg, prog): print(f"[{prog}%] {msg}")

    os.makedirs(output_dir, exist_ok=True)

    update("步骤1/4：提取DOCX文本...", 5)
    original_text = extract_docx_text(input_docx)
    if not original_text:
        update("错误：无法读取DOCX文本", 0)
        return
    print(f"  提取成功，共 {len(original_text)} 字符")

    cover_info = extract_cover_info(input_docx)
    print(f"  封面信息：{cover_info.get('title', '')} / {cover_info.get('name', '')}")

    update("步骤2/4：分析原文画像...", 15)
    analysis = _analyze_original_paper(original_text, paper_type)
    print(f"  标题：{analysis.get('title', 'N/A')}")

    update("步骤3/4：诊断与重构...", 30)
    revision_plan = _diagnose_and_reconstruct(analysis, paper_type)
    report_path = os.path.join(output_dir, "诊断报告.txt")
    save_diagnosis_report(revision_plan, report_path)
    print(f"  诊断报告已保存：{report_path}")
    diag = revision_plan.get("diagnosis_report", {})
    print(f"  评分：{diag.get('overall_score', 'N/A')} / 问题：{len(diag.get('issues', []))}")

    update("步骤4/4：生成修改版论文...", 50)
    txt = _generate_revised(analysis, revision_plan, paper_type, update)

    task_id = str(uuid.uuid4())[:8]
    txt_path = os.path.join(output_dir, f"00_完整论文_{task_id}.txt")
    with open(txt_path, 'w', encoding='utf-8') as f:
        f.write(txt)
    print(f"  TXT已保存：{txt_path}")

    docx_path = os.path.join(output_dir, f"论文_修改版_{task_id}.docx")
    # 调用 renderer
    update("正在渲染DOCX...", 85)
    from renderer import render
    render(txt_path, docx_path, paper_type, cover_info, None, update)

    # 复制最终版本
    final_docx = os.path.join(output_dir, "论文_修改版.docx")
    shutil.copy(docx_path, final_docx)
    print(f"\n  ✅ 修改完成！")
    print(f"  诊断报告：{report_path}")
    print(f"  最终论文：{final_docx}")


def main():
    parser = argparse.ArgumentParser(description="论文修改程序")
    parser.add_argument("--input", "-i", required=True, help="输入的DOCX文件路径")
    parser.add_argument("--output", "-o", default="./output", help="输出目录（默认./output）")
    parser.add_argument("--type", "-t", default="管理", choices=["管理", "设计", "机械"],
                        help="论文类型")
    args = parser.parse_args()

    if not os.path.isfile(args.input):
        print(f"错误：输入文件不存在 {args.input}")
        return

    def update(msg, prog):
        print(f"[{prog}%] {msg}")

    revise(args.input, args.output, args.type, update)


if __name__ == "__main__":
    main()
