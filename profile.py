# -*- coding: utf-8 -*-
"""
profile.py — 独立画像生成模块
============================
输入：专业类型 + 题目 + 对象
输出：结构化画像JSON
可嵌入使用，也可独立CLI调用

支持类型：管理、设计、机械、土木。

使用：
  python profile.py --type 管理 --title "xxx" --object "K公司" -o profile.json
  python profile.py --type 设计 --title "xxx" --object "新中式女装" --context "日常穿着" -o profile.json
  python profile.py --type 机械 --title "xxx" --object "夹具/装置名称" --context "工况" -o profile.json
"""
import sys, os, json, re, argparse
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from core import call_llm
from mechanical_spec import (
    attach_mechanical_master_spec,
    build_mechanical_master_spec,
    lock_mechanical_master_spec,
)


def generate_profile(profile_type: str, title: str, target: str,
                     major: str = "", context: str = "") -> dict:
    """
    生成研究画像（统一入口）
    profile_type: "管理" 或 "设计" 或 "机械"
    """
    profile_type = profile_type.strip()
    if profile_type in ("管理", "经管", "manage", "mg"):
        return _gen_manage_profile(title, target, major)
    elif profile_type in ("设计", "design", "sj"):
        return _gen_design_profile(major or "设计学", title, target, context)
    elif profile_type in ("机械", "mechanical", "mech", "mj"):
        return _gen_mechanical_profile(title, target, major or "机械设计", context)
    elif profile_type in ("土木", "civil", "cw"):
        return _gen_civil_profile(title, target, major or "土木工程", context)
    else:
        print(f"不支持的专业类型: {profile_type}，可选：管理、设计、机械、土木")
        return {}


def get_mechanical_machine_spec(profile: dict):
    """
    稳定的机械 master-spec 读取入口。
    返回只读 MappingProxyType，可供下游直接锁定使用。
    """
    return lock_mechanical_master_spec(profile or {})


# 兼容性别名：方便外部代码按不同命名习惯调用
get_machine_spec = get_mechanical_machine_spec
build_machine_spec = build_mechanical_master_spec
lock_machine_spec = lock_mechanical_master_spec


def _gen_manage_profile(title: str, company: str, major: str = "") -> dict:
    """生成经管类画像"""
    prompt = f"""你是经管类论文研究设计专家。
【专业】：{major or "工商管理"}
【论文题目】：{title}
【研究对象】：{company}
请分析对象，生成严格JSON格式的研究画像：
{{
  "company": "{company}",
  "org_type": "类型",
  "industry": "行业",
  "core_problems": ["问题1","问题2","问题3"],
  "data_hints": ["数据1","数据2","数据3","数据4"]
}}
只输出JSON。"""
    response = call_llm(prompt, max_tokens=1200)
    try:
        json_match = re.search(r'\{.*\}', response, re.DOTALL)
        if json_match:
            profile = json.loads(json_match.group())
            profile["major"] = major or "工商管理"
            profile["title"] = title
            return profile
    except:
        pass
    return {
        "company": company, "org_type": "企业", "industry": "制造业",
        "core_problems": ["成本控制不足", "资金周转慢", "管理效率低"],
        "data_hints": ["近三年财务数据", "行业平均数据", "运营指标"],
        "major": major or "工商管理", "title": title
    }


def _gen_design_profile(design_type: str, title: str, target_name: str,
                        context: str = "", force_variation: bool = False) -> dict:
    """生成设计类画像"""
    variation_hint = "\n必须生成完全不同的设计问题和策略。" if force_variation else ""
    prompt = f"""你是设计类论文研究设计专家。
【设计专业】：{design_type}
【论文题目】：{title}
【设计对象】：{target_name}
{variation_hint}
请分析设计对象，生成严格JSON格式的研究画像：
{{
    "design_object": "{target_name}",
    "design_type": "{design_type}",
    "object_type": "对象类型（根据专业判断）",
    "context": "设计背景/使用场景",
    "target_users": ["用户群体1", "用户群体2"],
    "core_problems": ["问题1", "问题2", "问题3"],
    "design_strategies": ["策略1", "策略2", "策略3"],
    "design_elements": ["设计要素1", "设计要素2", "设计要素3"]
}}
只输出JSON，不要其他内容。"""
    response = call_llm(prompt, max_tokens=1200)
    try:
        json_match = re.search(r'\{.*\}', response, re.DOTALL)
        if json_match:
            profile = json.loads(json_match.group())
            profile["design_type"] = design_type
            profile["title"] = title
            if context:
                profile["context"] = context
            return profile
    except:
        pass
    return {
        "design_object": target_name, "design_type": design_type,
        "object_type": "设计对象", "context": context or "使用场景",
        "target_users": ["目标用户"],
        "core_problems": ["问题1", "问题2", "问题3"],
        "design_strategies": ["策略一", "策略二", "策略三"],
        "design_elements": ["要素1", "要素2", "要素3"],
        "title": title
    }


def _gen_mechanical_profile(title: str, target_name: str,
                            major: str = "机械设计", context: str = "") -> dict:
    """生成机械设计类画像"""
    prompt = f"""你是机械设计类论文研究设计专家。
【专业】：{major}
【论文题目】：{title}
【设计对象】：{target_name}
【应用场景】：{context or "机械设计继续教育论文"}
请分析对象，生成严格JSON格式的研究画像：
{{
  "mech_object": "{target_name}",
  "mech_type": "机械类型（必须从：夹具、传动、送料、搬运、升降、结构优化 中选择）",
  "working_condition": "工作条件/工况",
  "technical_params": {{"载荷": "数值+单位", "行程": "数值+单位", "精度": "数值+单位"}},
  "materials": "主要材料/热处理",
  "core_problems": ["问题1", "问题2", "问题3"],
  "calc_items": ["校核项1", "校核项2", "校核项3"],
  "data_hints": ["参数1", "参数2", "参数3", "参数4"]
}}
只输出JSON，不要其他内容。"""
    response = call_llm(prompt, max_tokens=1200)
    try:
        json_match = re.search(r'\{.*\}', response, re.DOTALL)
        if json_match:
            profile = json.loads(json_match.group())
            profile["major"] = major
            profile["title"] = title
            profile["mech_object"] = profile.get("mech_object", target_name)
            profile["company"] = target_name
            profile["industry"] = "机械设计"
            if context:
                profile["context"] = context
            return attach_mechanical_master_spec(profile)
    except:
        pass
    fallback = {
        "mech_object": target_name,
        "mech_type": "夹具",
        "working_condition": context or "机械设计工况",
        "technical_params": {"载荷": "500N", "行程": "300mm", "精度": "0.1mm"},
        "materials": "45钢",
        "core_problems": ["结构不稳定", "效率不足", "精度不足"],
        "calc_items": ["受力校核", "强度校核", "稳定性校核"],
        "data_hints": ["设计参数", "工况参数", "校核参数"],
        "company": target_name,
        "industry": "机械设计",
        "major": major,
        "title": title
    }
    return attach_mechanical_master_spec(fallback)


def _gen_civil_profile(title: str, target_name: str,
                       major: str = "土木工程", context: str = "") -> dict:
    """生成土木工程类画像"""
    structure_type = "框架结构"
    if "剪力墙" in title:
        structure_type = "剪力墙结构"
    elif "框架-剪力墙" in title:
        structure_type = "框架-剪力墙结构"
    elif "钢结构" in title:
        structure_type = "钢结构"

    return {
        "title": title,
        "object_name": target_name or title,
        "project_type": "建筑与结构设计",
        "structure_type": structure_type,
        "major": major,
        "context": context,
        "paper_type": "土木",
    }


def main():
    parser = argparse.ArgumentParser(description="论文画像生成模块")
    parser.add_argument("--type", "-t", required=True, choices=["管理", "经管", "manage", "mg", "设计", "design", "sj", "机械", "mechanical", "mech", "mj", "土木", "civil", "cw"],
                        help="专业类型：管理/设计/机械/土木")
    parser.add_argument("--title", required=True, help="论文题目")
    parser.add_argument("--object", "-o", required=True, help="研究对象/设计对象")
    parser.add_argument("--major", "-m", default="", help="专业名称（经管类）")
    parser.add_argument("--context", "-c", default="", help="使用场景（设计类）")
    parser.add_argument("--output", "-f", default="", help="输出JSON文件路径（默认stdout）")
    args = parser.parse_args()

    profile = generate_profile(args.type, args.title, args.object, args.major, args.context)
    if not profile:
        sys.exit(1)

    output = json.dumps(profile, ensure_ascii=False, indent=2)
    if args.output:
        with open(args.output, 'w', encoding='utf-8') as f:
            f.write(output)
        print(f"画像已保存: {args.output}")
    else:
        print(output)


if __name__ == "__main__":
    main()
