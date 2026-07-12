# -*- coding: utf-8 -*-
"""
generator.py — 论文生成管道（纯文本阶段）
======================================
调用 LLM 生成含标签的纯文本论文，供 renderer.py 渲染。
生成 = LLM 输出纯文本（含 <chart/> <table/> <drawing/> 标签）

使用：
  profile = {...}
  txt = generate(profile, '管理')
  txt = generate(profile, '设计')
"""
import sys, os, re, json, time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from core import (
    call_llm, clean_text,
    MG_CHAPTERS, SJ_CHAPTERS, CIV_WORD_LIMITS, MG_WORD_LIMITS, SJ_WORD_LIMITS, MC_WORD_LIMITS,
    extract_drawings_from_text, _canonicalize_drawing_line,
    canonicalize_design_table_payload,
)
from prompts.manage import (
    build_chapter_prompt as mg_chapter_prompt,
    build_abstract_prompt as mg_abstract_prompt,
    build_keywords_prompt as mg_kw_prompt,
    build_references_prompt as mg_ref_prompt,
)
from prompts.design import (
    build_chapter_prompt as dj_chapter_prompt,
    build_abstract_prompt as dj_abstract_prompt,
    build_keywords_prompt as dj_kw_prompt,
    build_references_prompt as dj_ref_prompt,
    build_drawing_analysis_prompt as dj_drawing_analysis_prompt,
)
from prompts.mechanical import (
    build_outline as mc_build_outline,
    build_machine_spec_prompt as mc_machine_spec_prompt,
    build_chapter_prompt as mc_chapter_prompt,
    build_abstract_prompt as mc_abstract_prompt,
    build_keywords_prompt as mc_kw_prompt,
    build_references_prompt as mc_ref_prompt,
    build_drawing_analysis_prompt as mc_drawing_analysis_prompt,
    MECH_TYPE_MAP as MC_MECH_TYPE_MAP,
)
from prompts.civil import (
    build_chapter_prompt as cv_chapter_prompt,
    build_abstract_prompt as cv_abstract_prompt,
    build_references_prompt as cv_ref_prompt,
    get_chapter_outline as cv_get_chapter_outline,
)


_VALID_TABLE_TAG_RE = re.compile(
    r'^\s*<table\b'
    r'(?=[^>]*\bid="[^"]+")'
    r'(?=[^>]*\btitle="[^"]+")'
    r'(?=[^>]*\bheader="[^"]+")'
    r'(?=[^>]*\brows="[^"]+")'
    r'(?=[^>]*\bdata="[^"]+")'
    r'[^>]*\s*/>\s*$',
    re.I,
)


def _escape_attr_text(value: str) -> str:
    """把属性值压成单行，避免标签属性跨行或引号冲突。"""
    if value is None:
        return ""
    text = str(value).replace('"', "'")
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _normalize_drawing_signature_from_parts(drawing_type: str, title: str, description: str) -> str:
    """生成 drawing 内容签名，用于去重。"""
    parts = []
    for value in (drawing_type, title, description):
        text = re.sub(r"\s+", "", str(value or ""))
        if text:
            parts.append(text)
    return "|".join(parts)


def _drawing_signature_from_text(text: str) -> str:
    """从一行或一段文本中提取 drawing 内容签名。"""
    if not text or "<drawing" not in text.lower():
        return ""
    parsed = _canonicalize_drawing_line(text, 1, 0, len(text))
    if not parsed:
        m = re.search(r'<drawing\b[^>]*>', text, re.I | re.S)
        if not m:
            return ""
        parsed = _canonicalize_drawing_line(m.group(0), 1, 0, len(m.group(0)))
    if not parsed:
        return ""
    return _normalize_drawing_signature_from_parts(
        parsed.get("type", ""),
        parsed.get("title", ""),
        parsed.get("description", ""),
    )


def _strip_design_abstract_prefix(abstract: str) -> str:
    """删除摘要开头的“摘要：/摘要:”前缀。"""
    text = (abstract or "").lstrip()
    text = re.sub(r'^\s*摘要\s*[:：]\s*', '', text)
    text = re.sub(r'^\s*摘要\s*\n\s*', '', text)
    # 摘要中如果混入图表/表格标签或HTML表格碎片，直接剔除
    text = re.sub(r'(?is)<(?:table|chart|tr|td|th|caption)\b.*?(?:/>|>|$)', '', text)
    text = re.sub(r'(?is)</(?:table|chart|tr|td|th|caption)\s*>', '', text)
    text = re.sub(r'(?is)<[^>\n]+>', '', text)
    cleaned_lines = []
    for line in text.splitlines():
        s = line.strip()
        if not s:
            continue
        if s in {"|", "---", "—"}:
            continue
        cleaned_lines.append(line)
    text = "\n".join(cleaned_lines)
    return text.strip()


def _looks_like_numeric_cell(value: str) -> bool:
    text = (value or "").strip()
    if not text:
        return False
    return bool(re.fullmatch(r'[\d.]+%?', text) or re.fullmatch(r'[\d.]+[一二三四五六七八九十]?', text))


def _repair_html_like_table_block(block_lines: list[str], table_id: int, chapter_num: int | None = None, table_index: int | None = None) -> str | None:
    """把 LLM 输出的 HTML 风格表格片段收敛成标准 <table/> 标签。"""
    raw = "\n".join(block_lines)
    if "<tr" not in raw.lower() and "<td" not in raw.lower() and "<th" not in raw.lower():
        return None

    title = ""
    title_match = re.search(r'<caption\s*([^<\n\r]*)', raw, re.I)
    if title_match:
        title = title_match.group(1)
    title = re.sub(r'</caption.*$', '', title, flags=re.I).strip()
    title = title.strip(' ：:、,.')

    headers: list[str] = []
    row_names: list[str] = []
    data_rows: list[str] = []
    current_group = ""
    group_left = 0

    row_blocks = re.findall(r'<tr\b.*?</tr\s*>?', raw, re.I | re.S)
    if not row_blocks:
        row_blocks = [line for line in block_lines if "<tr" in line.lower()]

    for row_block in row_blocks:
        low = row_block.lower()
        cells = re.findall(r'<(?:th|td)(?:\s+[^<>]*?)?(?:>|)([^<]*?)</(?:th|td)\s*>?', row_block, re.I | re.S)
        cells = [re.sub(r'rowspan\s*=\s*"?\d+"?', '', c, flags=re.I).strip() for c in cells]
        cells = [c.strip(' ：:、,.') for c in cells if c is not None and str(c).strip()]
        if not cells:
            continue
        if not headers and "<th" in low:
            headers = cells
            continue

        span_match = re.search(r'rowspan\s*=\s*"?(\d+)"?', row_block, re.I)
        if span_match:
            current_group = cells[0]
            group_left = max(int(span_match.group(1)), 1) - 1
            row_names.append(current_group)
            data_rows.append("|".join(cells[1:]))
            continue

        if current_group and group_left > 0:
            row_names.append(current_group)
            data_rows.append("|".join(cells))
            group_left -= 1
            continue

        if len(cells) >= 2 and not _looks_like_numeric_cell(cells[0]):
            current_group = cells[0]
            row_names.append(current_group)
            data_rows.append("|".join(cells[1:]))
            continue

        row_names.append(current_group or cells[0])
        data_rows.append("|".join(cells[1:] if len(cells) > 1 else cells))

    if not headers or not row_names or not data_rows:
        return None

    if headers:
        first_header = headers[0]
        if any(k in first_header for k in ("满意", "评分", "评价", "频次", "占比", "满意度")):
            headers[0] = "评价维度"

    header_text = ",".join(_escape_attr_text(v) for v in headers if _escape_attr_text(v))
    rows_text = ",".join(_escape_attr_text(v) for v in row_names if _escape_attr_text(v))
    data_text = ";".join(_escape_attr_text(v) for v in data_rows if _escape_attr_text(v))
    if not header_text or not rows_text or not data_text:
        return None

    title = _infer_civil_table_title(chapter_num, title, header_text, raw, table_index)

    parts = [
        f'id="{table_id}"',
        f'title="{_escape_attr_text(title or f"表{table_id}")}"',
        f'header="{header_text}"',
        f'rows="{rows_text}"',
        f'data="{data_text}"',
    ]
    return "<table " + " ".join(parts) + "/>"


def _convert_chart_attrs_to_table(attrs: dict, table_id: int) -> str | None:
    """把 chart 属性转换成可渲染的 table 标签。"""
    if not attrs:
        return None
    data = (attrs.get("data") or "").strip()
    x_values = [x.strip() for x in (attrs.get("x") or "").split(",") if x.strip()]
    title = _escape_attr_text(attrs.get("title") or f"表{table_id}")

    row_names: list[str] = []
    data_rows: list[str] = []
    header: list[str] = []

    # 形如：A,42,8.2,3.1,0; B,38,5.6,2.8,0
    if ";" in data and all("," in seg for seg in [s.strip() for s in data.split(";") if s.strip()]):
        segments = [s.strip() for s in data.split(";") if s.strip()]
        for seg in segments:
            cells = [c.strip() for c in seg.split(",") if c.strip()]
            if len(cells) < 2:
                continue
            row_names.append(cells[0])
            data_rows.append("|".join(cells[1:]))
        if x_values:
            header = ["对比项"] + x_values[1:]
        else:
            max_cols = max((len(row.split("|")) for row in data_rows), default=0)
            header = ["对比项"] + [f"指标{i}" for i in range(1, max_cols + 1)]

    # 形如：早间8-10点:12人次,上午10-12点:45人次
    elif ":" in data:
        pairs = [p.strip() for p in re.split(r"[;,]", data) if p.strip()]
        for pair in pairs:
            if ":" not in pair:
                continue
            label, value = pair.split(":", 1)
            num = re.search(r"[-+]?\d+(?:\.\d+)?", value)
            if not num:
                continue
            row_names.append(label.strip())
            data_rows.append(num.group(0))
        header = ["统计维度", "数值"]

    # 回退到已有 chart->table 逻辑
    else:
        fallback = _chart_to_table_attrs(attrs)
        if fallback:
            return _canonical_table_line(fallback)
        return None

    if not row_names or not data_rows or not header:
        return None

    header_text = ",".join(_escape_attr_text(v) for v in header if _escape_attr_text(v))
    rows_text = ",".join(_escape_attr_text(v) for v in row_names if _escape_attr_text(v))
    data_text = ";".join(_escape_attr_text(v) for v in data_rows if _escape_attr_text(v))
    if not header_text or not rows_text or not data_text:
        return None

    parts = [
        f'id="{table_id}"',
        f'title="{title}"',
        f'header="{header_text}"',
        f'rows="{rows_text}"',
        f'data="{data_text}"',
    ]
    data_source = _escape_attr_text(attrs.get("data_source") or attrs.get("datasource") or "")
    if data_source:
        parts.append(f'data_source="{data_source}"')
    return "<table " + " ".join(parts) + "/>"


def _repair_malformed_design_table_blocks(full_text: str) -> str:
    """把设计类正文中失控的 HTML 风格表格块收敛成标准 table 标签。"""
    if not full_text:
        return full_text

    lines = full_text.splitlines()
    out: list[str] = []
    i = 0
    table_seq = 0

    while i < len(lines):
        line = lines[i]
        low = line.lower().strip()

        if "<chart" in low:
            chart_attrs = _extract_tag_attrs(line, "chart")
            if chart_attrs:
                converted = _convert_chart_attrs_to_table(chart_attrs, table_seq + 1)
                if converted:
                    out.append(converted)
                    table_seq += 1
                    i += 1
                    continue

        if low.startswith("<table"):
            # 单行表格标签，先尝试直接规范化；若失败再按 HTML 风格表格块处理
            table_attrs = _extract_tag_attrs(line, "table")
            if table_attrs:
                repaired = _repair_table_attrs(table_attrs)
                canonical = _canonical_table_line(repaired)
                if canonical:
                    out.append(canonical)
                    table_seq += 1
                    i += 1
                    continue

        if low.startswith("<table") and not _VALID_TABLE_TAG_RE.match(line):
            block = [line]
            j = i + 1
            while j < len(lines):
                block.append(lines[j])
                if "</table" in lines[j].lower():
                    j += 1
                    break
                j += 1
            repaired = _repair_html_like_table_block(block, table_seq + 1)
            if repaired:
                out.append(repaired)
                table_seq += 1
                i = j
                continue
        out.append(line)
        i += 1

    return "\n".join(out)


# ==================== 设计类动态大纲 ====================
def generate_outline(profile: dict) -> dict:
    """为设计类论文动态生成大纲"""
    design_type = profile.get('design_type', '')
    design_object = profile.get('design_object', '')
    title = profile.get('title', '')
    core_problems = profile.get('core_problems', [])
    design_strategies = profile.get('design_strategies', [])

    prompt = f"""你是设计类论文结构专家。请根据以下信息，为这篇设计类毕业论文生成一个合适的章节大纲。

【设计专业】：{design_type}
【论文题目】：{title}
【设计对象】：{design_object}
【核心问题】：{json.dumps(core_problems, ensure_ascii=False)}
【初步设计策略】：{json.dumps(design_strategies, ensure_ascii=False)}

【任务要求】：
1. 生成合理的5章结构
2. 第4章"设计策略与方案"的子标题要符合该专业的设计流程和术语
3. 不同专业应有不同侧重点：
   - 服装与服饰设计：设计定位、款式设计、面料选择、色彩搭配、结构工艺等
   - 数字媒体设计：需求分析、交互设计、视觉设计、动效设计等
   - 产品设计：造型设计、功能结构、CMF设计、人机交互等
   - 室内设计：空间布局、功能分区、材质软装、灯光设计等
   - 景观设计：总体布局、分区设计、植物配置等
   - 视觉传达：标志设计、应用系统、色彩字体等

【输出格式】：
必须输出严格的JSON，格式如下：
{{
    "chapter1_intro": {{"1.1": "研究背景", "1.2": "研究问题", "1.3": "研究意义"}},
    "chapter2_theory": {{"2.1": "核心理论一", "2.2": "核心理论二", "2.3": "理论与本设计的关联"}},
    "chapter3_problem": {{"3.1": "项目概况", "3.2": "调研与分析", "3.3": "核心问题归纳"}},
    "chapter4_design": {{
        "4.1": "设计策略",
        "4.2": "【专业相关子标题】",
        "4.3": "【专业相关子标题】",
        "4.4": "【专业相关子标题】",
        "4.5": "【专业相关子标题】"
    }},
    "chapter5_conclusion": {{"5.1": "主要成果", "5.2": "不足之处", "5.3": "改进方向"}}
}}

只输出JSON，不要其他内容。"""
    response = call_llm(prompt, max_tokens=2000)
    try:
        json_match = re.search(r'\{.*\}', response, re.DOTALL)
        if json_match:
            outline = json.loads(json_match.group())
            # 确保第4章有足够的子标题
            if 'chapter4_design' in outline:
                design_ch = outline['chapter4_design']
                subs = [k for k in design_ch.keys() if k.startswith('4.') and k != '4.1']
                if len(subs) < 3:
                    for i in range(len(subs) + 1, 6):
                        design_ch[f"4.{i}"] = f"设计细化{i-1}"
            return outline
    except Exception as e:
        print(f"大纲解析失败: {e}")
    # 兜底大纲
    return _get_fallback_outline(design_type, design_object)


def _get_fallback_outline(design_type: str, design_object: str) -> dict:
    """备用大纲"""
    base = {
        "chapter1_intro": {"1.1": "研究背景", "1.2": "研究问题", "1.3": "研究意义"},
        "chapter2_theory": {"2.1": "核心理论", "2.2": "支撑理论", "2.3": "与本设计关联"},
        "chapter3_problem": {"3.1": f"{design_object}概况", "3.2": "调研分析", "3.3": "核心问题"},
        "chapter5_conclusion": {"5.1": "主要成果", "5.2": "不足之处", "5.3": "改进方向"},
    }
    if '服装' in design_type:
        base["chapter4_design"] = {
            "4.1": "设计策略", "4.2": "设计定位与灵感",
            "4.3": "款式与结构", "4.4": "面料与色彩", "4.5": "系列展示"
        }
    elif '数字媒体' in design_type or '交互' in design_type:
        base["chapter4_design"] = {
            "4.1": "设计策略", "4.2": "信息架构与交互",
            "4.3": "视觉界面", "4.4": "动效设计", "4.5": "技术实现"
        }
    elif '产品' in design_type:
        base["chapter4_design"] = {
            "4.1": "设计策略", "4.2": "造型设计",
            "4.3": "功能结构", "4.4": "CMF设计", "4.5": "体验设计"
        }
    else:
        base["chapter4_design"] = {
            "4.1": "设计策略", "4.2": "总体方案",
            "4.3": "方案细化一", "4.4": "方案细化二", "4.5": "专项设计"
        }
    return base


# ==================== 设计类图纸补齐 ====================
def repair_drawing_tags(full_text: str, profile: dict, paper_type: str = "设计") -> str:
    """LLM二次分析全文，识别图纸描述并补齐<drawing/>标签"""
    if paper_type in ("机械", "mechanical", "mech", "mj"):
        prompt = mc_drawing_analysis_prompt(full_text, profile)
    else:
        prompt = dj_drawing_analysis_prompt(full_text, profile)
    response = call_llm(prompt, max_tokens=4000)
    if not response:
        return full_text
    try:
        json_match = re.search(r'\{.*\}', response, re.DOTALL)
        if not json_match:
            return full_text
        result = json.loads(json_match.group())

        action = result.get("action", "")
        if action == "keep_existing" or action == "":
            return full_text

        existing_signatures = set()
        for d in extract_drawings_from_text(full_text):
            sig = _normalize_drawing_signature_from_parts(
                d.get("type", ""),
                d.get("title", ""),
                d.get("description", ""),
            )
            if sig:
                existing_signatures.add(sig)

        # 应用替换
        snippets = result.get("modified_text_snippets", [])
        if not snippets:
            # 如果没有替换片段，尝试从missing_drawings自动添加
            missing = result.get("missing_drawings", [])
            if missing:
                # 在文本末尾附近添加图纸
                modified = full_text.rstrip()
                base_count = len(result.get("existing_drawings", []))
                for i, d in enumerate(missing):
                    sig = _normalize_drawing_signature_from_parts(
                        d.get("type", "效果图"),
                        d.get("title", ""),
                        d.get("description", ""),
                    )
                    if sig and sig in existing_signatures:
                        continue
                    did = str(base_count + i + 1)
                    tag = f'<drawing id="{did}" type="{d.get("type", "效果图")}" title="{d.get("title", "")}" description="{d.get("description", "")}"/>'
                    modified += "\n" + tag
                    if sig:
                        existing_signatures.add(sig)
                return modified

        # 逐替换
        modified = full_text
        for s in snippets:
            orig = s.get("original", "")
            repl = s.get("replacement", "")
            if orig and repl:
                sig = _drawing_signature_from_text(repl)
                if sig and sig in existing_signatures:
                    continue
                modified = modified.replace(orig, repl, 1)
                if sig:
                    existing_signatures.add(sig)
        return modified
    except Exception as e:
        print(f"图纸补齐解析失败: {e}")
        return full_text


def normalize_mechanical_drawing_tags(full_text: str) -> str:
    """把机械类文本里的 drawing 行尽量规范化为严格单行标签。"""
    lines = full_text.splitlines()
    out = []
    seq = 0
    for idx, line in enumerate(lines):
        if "<drawing" in line.lower():
            parsed = _canonicalize_drawing_line(line, seq + 1, idx, idx + 1)
            if parsed:
                title = (parsed.get("title") or "").strip()
                desc = (parsed.get("description") or "").strip()
                if title in ("", "未命名", "图纸标题") and not desc:
                    continue
                seq += 1
                out.append(
                    f'<drawing id="{seq}" type="{parsed.get("type", "结构图")}" '
                    f'title="{parsed.get("title", "")}" description="{parsed.get("description", "")}"/>'
                )
                continue
            # 裸标签或属性不全的 drawing，直接丢弃，避免污染正文
            if re.fullmatch(r'\s*<drawing\s*/?>\s*', line, re.IGNORECASE) or re.fullmatch(r'\s*<drawing\s*/\s*>\s*', line, re.IGNORECASE):
                continue
        out.append(line)
    return "\n".join(out)


MECH_JSON_OPEN = "[[MECH_JSON]]"
MECH_JSON_CLOSE = "[[/MECH_JSON]]"


def _fallback_machine_spec(profile: dict, outline: dict = None) -> dict:
    """当 LLM 生成 machine_spec 失败时的保底 spec。"""
    title = profile.get("title", "")
    mech_object = profile.get("mech_object") or profile.get("company", "")
    mech_type = profile.get("mech_type", "") or profile.get("mechanical_type", "")
    cfg = None
    for k, v in MC_MECH_TYPE_MAP.items():
        if k in mech_type:
            cfg = v
            break
    cfg = cfg or MC_MECH_TYPE_MAP["夹具"]
    scene_lock = f"全文统一为：{mech_object}的{mech_type or '机械'}设计"
    focus_map = {
        "概述": "锁定对象、场景、任务与要求",
        "原理与方案分析": "工作原理、方案比选与约束",
        "总体设计": "总体布局、参数初定与结构关系",
        "关键部件设计": "关键部件、受力路径、尺寸与连接",
        "计算与校核": "参数计算、强度校核与结论回扣",
        "总结": "成果、局限与改进方向",
    }
    return {
        "paper_title": title,
        "mech_object": mech_object,
        "mech_type": mech_type,
        "scene_lock": scene_lock,
        "scope_boundary": [scene_lock, "只写同一套机构链", "不切换到其他设备或工艺"],
        "required_terms": [mech_object, mech_type] + list(cfg["content_points"]),
        "forbidden_terms": ["其他机械对象", "无关设备", "泛化示意图", "概览图", "示意图"],
        "fixed_entities": [mech_object, mech_type] + list(cfg["content_points"]),
        "chapter_focus": focus_map,
        "figure_policy": {
            "required": True,
            "must_be_specific": True,
            "must_include_numeric": True,
            "forbidden_generic_words": ["示意图", "概览图", "示意", "概况图", "示意性图"],
        },
        "consistency_keys": ["scene_lock", "mech_object", "mech_type", "fixed_entities", "chapter_focus", "figure_policy"],
    }


def _safe_json_load(text: str):
    if not text:
        return None
    try:
        return json.loads(text)
    except Exception:
        m = re.search(r'\{.*\}', text, re.DOTALL)
        if m:
            try:
                return json.loads(m.group())
            except Exception:
                return None
    return None


def _extract_mech_json_block(content: str):
    """返回 (正文, json_dict, raw_json_text)。"""
    if not content:
        return "", None, ""
    open_idx = content.find(MECH_JSON_OPEN)
    close_idx = content.find(MECH_JSON_CLOSE)
    if open_idx >= 0 and close_idx > open_idx:
        body = content[:open_idx].rstrip()
        raw = content[open_idx + len(MECH_JSON_OPEN):close_idx].strip()
        return body, _safe_json_load(raw), raw
    return content.strip(), None, ""


def _clean_mech_body_preserving_drawings(body: str) -> str:
    """机械章节正文清洗：保留 <drawing/> 标签不被 clean_text 吃掉。"""
    if not body:
        return ""
    placeholders = []

    def _stash(m):
        token = f"__MECH_DRAWING_{len(placeholders)}__"
        placeholders.append((token, m.group(0)))
        return token

    # 先把整行 drawing 标签藏起来，再走通用清洗
    protected = re.sub(r'(?im)^\s*<drawing\b.*?/?\s*>\s*$', _stash, body)
    protected = clean_text(protected)
    for token, raw in placeholders:
        protected = protected.replace(token, raw)
    # 再做一次局部清理：去掉可能粘连在 drawing 行前后的空行与多余空格
    protected = re.sub(r'\n{3,}', '\n\n', protected)
    return protected.strip()


def _split_headings(text: str):
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    sections = []
    for line in lines:
        if re.match(r'^\d+\.\d+(\.\d+)?\s+', line):
            sections.append(re.match(r'^(\d+\.\d+(?:\.\d+)?)\s+(.+)', line).group(1))
    return sections[:12]


def _extract_figures_from_body(body: str):
    figures = []
    seq = 0
    for idx, line in enumerate(body.splitlines()):
        if "<drawing" in line.lower():
            parsed = _canonicalize_drawing_line(line, seq + 1, idx, idx + 1)
            if not parsed:
                continue
            seq += 1
            figures.append({
                "id": str(seq),
                "type": parsed.get("type", "结构图"),
                "title": parsed.get("title", ""),
                "description": parsed.get("description", ""),
            })
    return figures


def _figure_is_specific(fig: dict) -> bool:
    text = " ".join([
        str(fig.get("title", "")),
        str(fig.get("description", "")),
    ])
    if any(g in text for g in ["示意图", "概览图", "示意", "概况图", "示意性图", "泛化"]):
        return False
    if re.search(r'\d', text):
        return True
    if re.search(r'(mm|cm|kN|MPa|kW|r/min|°|deg|φ|Φ)', text, re.IGNORECASE):
        return True
    if re.search(r'(孔|距|间距|直径|长度|高度|宽度|厚度|行程|载荷|扭矩|转速|压力|角度|数量)', text):
        return True
    return False


def _validate_mechanical_chapter(body: str, chapter_json: dict, machine_spec: dict, chapter_name: str, chapter_num: int):
    """检查章节正文、JSON块与 machine_spec 的一致性。"""
    issues = []
    spec = machine_spec or {}
    chapter_json = chapter_json or {}

    # 基础字段修复
    text_metadata = chapter_json.get("text_metadata")
    if not isinstance(text_metadata, dict):
        text_metadata = {}
        issues.append({"severity": "warn", "key": "text_metadata", "message": "缺少或损坏 text_metadata，已补全"})
    figures = chapter_json.get("figures")
    if not isinstance(figures, list):
        figures = []
        issues.append({"severity": "warn", "key": "figures", "message": "缺少或损坏 figures，已补全"})
    facts_used = chapter_json.get("facts_used")
    if not isinstance(facts_used, list):
        facts_used = []
        issues.append({"severity": "warn", "key": "facts_used", "message": "缺少或损坏 facts_used，已补全"})
    consistency_keys = chapter_json.get("consistency_keys")
    if not isinstance(consistency_keys, list):
        consistency_keys = []
        issues.append({"severity": "warn", "key": "consistency_keys", "message": "缺少或损坏 consistency_keys，已补全"})

    # machine_spec 对照
    scene_lock = str(spec.get("scene_lock", "")).strip()
    mech_object = str(spec.get("mech_object", "")).strip()
    mech_type = str(spec.get("mech_type", "")).strip()
    forbidden_terms = [str(x).strip() for x in spec.get("forbidden_terms", []) if str(x).strip()]
    required_terms = [str(x).strip() for x in spec.get("required_terms", []) if str(x).strip()]
    required_keys = [str(x).strip() for x in spec.get("consistency_keys", []) if str(x).strip()]

    full_text = f"{body}\n{json.dumps(chapter_json, ensure_ascii=False)}"
    for term in forbidden_terms:
        if term and term in full_text:
            issues.append({"severity": "fail", "key": "forbidden_terms", "message": f"出现禁止术语：{term}"})

    if mech_object and mech_object not in full_text:
        issues.append({"severity": "warn", "key": "mech_object", "message": f"正文未显式提及锁定对象：{mech_object}"})
    if mech_type and mech_type not in full_text:
        issues.append({"severity": "warn", "key": "mech_type", "message": f"正文未显式提及锁定类型：{mech_type}"})
    if scene_lock and scene_lock not in full_text:
        issues.append({"severity": "warn", "key": "scene_lock", "message": "正文未复用 machine_spec 的完整场景锁定语句"})

    # facts_used 轻量核验
    for fact in facts_used:
        if isinstance(fact, dict):
            k = str(fact.get("key") or fact.get("name") or "").strip()
            v = str(fact.get("value") or fact.get("fact") or "").strip()
            if k and k in spec and v and str(spec.get(k, "")).strip() and str(spec.get(k)).strip() not in v and v not in str(spec.get(k)):
                issues.append({"severity": "warn", "key": k, "message": f"事实值与 machine_spec 不一致：{v}"})
        elif isinstance(fact, str):
            if fact and fact not in full_text:
                issues.append({"severity": "warn", "key": "facts_used", "message": f"事实未在正文中落地：{fact}"})

    # figures 数值化要求
    normalized_figures = []
    body_figures = _extract_figures_from_body(body)
    if not figures and body_figures:
        figures = body_figures
        issues.append({"severity": "warn", "key": "figures", "message": "JSON 缺少 figures，已从正文 drawing 标签回填"})
    for fig in figures:
        if not isinstance(fig, dict):
            continue
        fig = {
            "id": str(fig.get("id", "")),
            "type": str(fig.get("type", "结构图")),
            "title": str(fig.get("title", "")),
            "description": str(fig.get("description", "")),
            **{k: v for k, v in fig.items() if k not in {"id", "type", "title", "description"}},
        }
        if not _figure_is_specific(fig):
            issues.append({"severity": "fail", "key": "figures", "message": f"图纸不够具体或缺少数值信息：{fig.get('title', '')}"})
        normalized_figures.append(fig)

    # consistency_keys 必须覆盖 machine_spec 要求
    for key in required_keys:
        if key not in consistency_keys:
            issues.append({"severity": "warn", "key": "consistency_keys", "message": f"缺少一致性键：{key}"})

    status = "pass"
    if any(i.get("severity") == "fail" for i in issues):
        status = "fail"
    elif issues:
        status = "warn"

    chapter_json["chapter_num"] = chapter_num
    chapter_json["chapter_name"] = chapter_name
    chapter_json["text_metadata"] = {
        "word_target": text_metadata.get("word_target", MC_WORD_LIMITS.get(chapter_name, 1000)),
        "sections": text_metadata.get("sections") or _split_headings(body),
        "machine_spec_keys_used": text_metadata.get("machine_spec_keys_used") or ["scene_lock", "mech_object", "mech_type", "chapter_focus", "figure_policy"],
        "scene_lock": text_metadata.get("scene_lock") or scene_lock,
        "figure_policy": text_metadata.get("figure_policy") or spec.get("figure_policy", {}),
        **{k: v for k, v in text_metadata.items() if k not in {"word_target", "sections", "machine_spec_keys_used", "scene_lock", "figure_policy"}},
    }
    chapter_json["facts_used"] = facts_used
    chapter_json["figures"] = normalized_figures
    chapter_json["consistency_keys"] = consistency_keys or required_keys
    chapter_json["consistency_validation"] = {
        "status": status,
        "issues": issues,
    }
    return chapter_json


def _render_mech_chapter(body: str, chapter_json: dict) -> str:
    block = json.dumps(chapter_json, ensure_ascii=False, indent=2)
    return f"{body.rstrip()}\n{MECH_JSON_OPEN}\n{block}\n{MECH_JSON_CLOSE}"


def _extract_mech_chapter_output(raw: str, chapter_name: str, chapter_num: int, machine_spec: dict):
    body, chapter_json, _raw_json = _extract_mech_json_block(raw)
    body = _clean_mech_body_preserving_drawings(body)
    body = re.sub(r'^第[一二三四五六\d]+章.*?\n', '', body).strip()
    body = normalize_mechanical_drawing_tags(body)
    if not isinstance(chapter_json, dict):
        chapter_json = {
            "chapter_num": chapter_num,
            "chapter_name": chapter_name,
            "text_metadata": {
                "word_target": MC_WORD_LIMITS.get(chapter_name, 1000),
                "sections": _split_headings(body),
                "machine_spec_keys_used": ["scene_lock", "mech_object", "mech_type", "chapter_focus", "figure_policy"],
                "scene_lock": machine_spec.get("scene_lock", ""),
                "figure_policy": machine_spec.get("figure_policy", {}),
            },
            "facts_used": [],
            "figures": _extract_figures_from_body(body),
            "consistency_keys": machine_spec.get("consistency_keys", []),
    "machine_spec": machine_spec if isinstance(machine_spec, dict) else {},
        }
    chapter_json = _validate_mechanical_chapter(body, chapter_json, machine_spec, chapter_name, chapter_num)
    return body, chapter_json


def normalize_design_markup(full_text: str) -> str:
    """规范设计类的 drawing/table 行，去掉裸标签与明显占位标签。"""
    lines = full_text.splitlines()
    parsed_drawings = []
    for idx, line in enumerate(lines):
        low = line.lower().strip()
        if "<drawing" in low:
            parsed = _canonicalize_drawing_line(line, idx + 1, idx, idx + 1)
            if parsed:
                title = (parsed.get("title") or "").strip()
                desc = (parsed.get("description") or "").strip()
                if title in ("", "未命名", "图纸标题") and not desc:
                    continue
                parsed["_line_idx"] = idx
                parsed_drawings.append(parsed)
    if len(parsed_drawings) > 8:
        effect_like = []
        other_like = []
        for d in parsed_drawings:
            t = " ".join([str(d.get("type", "")), str(d.get("title", "")), str(d.get("description", ""))])
            if ("效果图" in t) or ("效果" in t) or ("渲染" in t) or ("外观" in t):
                effect_like.append(d)
            else:
                other_like.append(d)
        keep = []
        keep.extend(effect_like[:5])
        keep.extend(other_like[:3])
        if len(keep) < 8:
            for d in parsed_drawings:
                if d not in keep:
                    keep.append(d)
                if len(keep) >= 8:
                    break
        keep_indices = {d["_line_idx"] for d in keep}
    else:
        keep_indices = {d["_line_idx"] for d in parsed_drawings}

    def _split_drawing_line(line: str) -> tuple[str | None, str]:
        """把同一行里的 drawing 标签和尾部正文拆开。"""
        m = re.search(r'<drawing\b.*?(?:/?>)', line, re.I | re.DOTALL)
        if not m:
            return None, line
        tag_part = line[m.start():m.end()]
        tail = line[m.end():].strip()
        return tag_part, tail

    def _strip_drawing_tag_fragment(line: str) -> str:
        """移除行内 drawing 标签碎片，保留剩余正文。"""
        stripped = re.sub(r'<drawing\b.*?(?:/?>)', '', line, flags=re.I | re.DOTALL)
        stripped = re.sub(r'\s{2,}', ' ', stripped).strip()
        return stripped

    out = []
    seq = 0
    for idx, line in enumerate(lines):
        low = line.lower().strip()
        if low in {"", "/", "/ /", "／", "／ ／"}:
            continue
        if "<drawing" in low:
            if idx not in keep_indices:
                continue
            parsed = _canonicalize_drawing_line(line, seq + 1, idx, idx + 1)
            if parsed:
                title = (parsed.get("title") or "").strip()
                desc = (parsed.get("description") or "").strip()
                if title in ("", "未命名", "图纸标题") and not desc:
                    continue
                seq += 1
                out.append(
                    f'<drawing id="{seq}" type="{parsed.get("type", "效果图")}" '
                    f'title="{title}" description="{desc}"/>'
                )
                _, tail = _split_drawing_line(line)
                if tail:
                    out.append(tail)
            else:
                tail = _strip_drawing_tag_fragment(line)
                if tail:
                    if tail in {"/", "／"}:
                        continue
                    out.append(tail)
            continue
        if low in ("<drawing/>", "<drawing / >", "<drawing />", "<table/>", "<table />"):
            continue
        if low.startswith("<table") or low.startswith("</table") or low.startswith("<tr") or low.startswith("</tr") or low.startswith("<td") or low.startswith("</td") or low.startswith("<caption"):
            continue
        out.append(line)
    cleaned = "\n".join(out)
    cleaned = re.sub(r'(?m)^\s*[\\/／]\s*$', '', cleaned)
    cleaned = re.sub(r'\n{3,}', '\n\n', cleaned)
    return cleaned.strip()


_CIVIL_PROMPT_LEAK_PATTERNS = (
    "好的，",
    "作为土木工程专业论文写作专家",
    "遵照您的要求",
    "内容将严格遵循",
    "我将为您撰写",
    "请撰写毕业论文",
    "严格遵循您提供的写作要点",
    "严格遵循您的要求",
)


def _looks_like_civil_prompt_leak(line: str) -> bool:
    text = re.sub(r"\s+", "", (line or ""))
    if not text:
        return False
    return any(pat.replace(" ", "") in text for pat in _CIVIL_PROMPT_LEAK_PATTERNS)


def _infer_civil_drawing_type(title: str, description: str = "", fallback: str = "结构图") -> str:
    """把土木 drawing 的 type 收敛到更具体、更稳定的取值。"""
    text = re.sub(r"\s+", "", f"{title} {description}")
    rules = [
        ("gantt", ["施工进度横道图", "甘特图"]),
        ("施工平面布置图", ["施工平面布置图", "总平面布置图", "平面布置图"]),
        ("建筑平面图", ["建筑平面图", "标准层平面图", "平面图"]),
        ("建筑立面图", ["正立面图", "侧立面图", "立面图"]),
        ("建筑剖面图", ["剖面图"]),
        ("水平地震作用分布图", ["水平地震作用分布图", "地震作用分布"]),
        ("框架梁配筋图", ["框架梁配筋图", "梁配筋图"]),
        ("框架柱配筋图", ["柱配筋图", "中柱配筋详图", "柱配筋详图"]),
        ("基础详图", ["基础详图", "基础图"]),
        ("节点详图", ["节点详图"]),
        ("配筋图", ["配筋详图", "配筋图"]),
        ("资源分配图", ["劳动力", "资源分配"]),
    ]
    for typ, kws in rules:
        if any(k in text for k in kws):
            return typ
    if fallback in ("", "drawing", "结构图", "图纸"):
        return "结构图"
    return fallback


def _infer_civil_table_title(chapter_num: int | None, title: str, header: str, raw_text: str, table_index: int | None = None) -> str:
    """给土木表格补更具体的标题。"""
    header_text = re.sub(r"\s+", "", f"{header} {raw_text}")
    title_text = (title or "").strip()
    generic = (not title_text) or re.fullmatch(r'表\d+', title_text) or title_text in {"表", "数据表", "统计表"}
    if not generic:
        return title_text

    if chapter_num == 3 and any(k in header_text for k in ["构件", "截面", "混凝土强度等级"]):
        return "表3-1 主要结构构件截面及材料表"
    if chapter_num == 3 and any(k in header_text for k in ["楼层", "梁左端弯矩", "跨中弯矩", "梁端弯矩"]):
        return "表3-2 竖向荷载作用下梁端弯矩计算表"
    if chapter_num == 5 and any(k in header_text for k in ["工种", "劳动力", "人数", "阶段"]):
        return "表5-2 劳动力配置表"
    if chapter_num == 5 and any(k in header_text for k in ["机械", "塔吊", "吊车", "设备"]):
        return "表5-3 主要机械配置表"
    if chapter_num:
        if table_index is None:
            table_index = 1
        if title_text.startswith(f"表{chapter_num}-"):
            return title_text
        if title_text.startswith("表"):
            tail = title_text[1:].strip()
            return f"表{chapter_num}-{table_index}" if tail in {"", "1", "2", "3", "4", "5"} else f"表{chapter_num}-{table_index} {tail}"
        return f"表{chapter_num}-{table_index} {title_text}"
    return title_text or "表1"


def _extract_first_match(text: str, patterns: list[str], default: str = "") -> str:
    for pat in patterns:
        m = re.search(pat, text, re.I | re.S)
        if m:
            if m.groups():
                vals = [g for g in m.groups() if g]
                if vals:
                    return " ".join(v.strip() for v in vals if str(v).strip()).strip()
            return m.group(0).strip()
    return default


def _prepend_paper_title(txt: str, profile: dict) -> str:
    """在纯文本论文开头补一行题目，供渲染器识别封面题目。"""
    title = str((profile or {}).get("title") or "").strip()
    if not title:
        return txt
    title_line = f"论文题目：{title}"
    body = (txt or "").lstrip()
    if body.startswith(title_line):
        return body
    return title_line + "\n\n" + body


def _build_civil_parameter_table(full_text: str, profile: dict) -> str | None:
    """当第1章缺失参数表时，按正文和 profile 兜底补一张表。"""
    text = full_text or ""
    rows = [
        ("总建筑面积", _extract_first_match(text, [r"总建筑面积为?([\d.]+)\s*平方米", r"总建筑面积约?([\d.]+)\s*平方米"], "约15200平方米")),
        ("建筑层数", _extract_first_match(text, [r"建筑主体地上(\d+)层，地下(\d+)层", r"建筑总层数为?(\d+)层"], "地上5层，地下1层")),
        ("建筑总高度", _extract_first_match(text, [r"总建筑高度为?([\d.]+)\s*m", r"建筑总高度为?([\d.]+)\s*m"], "23.4m")),
        ("标准层层高", _extract_first_match(text, [r"地上一至四层层高均为([\d.]+)\s*m", r"标准层层高([\d.]+)\s*m"], "4.2m")),
        ("柱网尺寸", _extract_first_match(text, [r"柱网尺寸主要为([0-9.×x\*]+)\s*m", r"柱网尺寸为([0-9.×x\*]+)\s*m"], "8.4m×8.4m")),
        ("地基承载力特征值", _extract_first_match(text, [r"fak为([\d.]+)\s*kPa", r"承载力特征值fak为([\d.]+)\s*kPa"], "180kPa")),
        ("基本风压", _extract_first_match(text, [r"基本风压为([\d.]+)\s*kN/m²", r"基本风压w₀=([\d.]+)\s*kN/m²"], "0.45kN/m²")),
        ("基本雪压", _extract_first_match(text, [r"基本雪压为([\d.]+)\s*kN/m²", r"基本雪压([\d.]+)\s*kN/m²"], "0.40kN/m²")),
        ("抗震设防烈度", _extract_first_match(text, [r"抗震设防烈度为(\d+)\s*度", r"设防烈度(\d+)\s*度"], "7度")),
        ("设计地震分组", _extract_first_match(text, [r"设计地震分组为([^，。；]+)", r"地震分组为([^，。；]+)"], "第二组")),
        ("场地类别", _extract_first_match(text, [r"场地类别为([^，。；]+)", r"场地类别([^，。；]+)"], "Ⅱ类")),
        ("设计使用年限", _extract_first_match(text, [r"设计使用年限为?(\d+)\s*年"], "50年")),
        ("安全等级", _extract_first_match(text, [r"建筑结构安全等级为([^，。；]+)", r"安全等级为([^，。；]+)"], "二级")),
        ("抗震等级", _extract_first_match(text, [r"框架抗震等级为([^，。；]+)", r"抗震等级为([^，。；]+)"], "三级")),
        ("结构形式", str(profile.get("structure_type") or _extract_first_match(text, [r"采用(.{2,20}?结构体系)", r"结构形式为([^，。；]+)"], "框架结构")).strip()),
        ("基础形式", _extract_first_match(text, [r"采用([^，。；]*?基础)", r"基础形式(?:拟)?采用([^，。；]+)"], "柱下独立基础")),
    ]
    if not rows:
        return None
    row_names = ",".join(name for name, value in rows if name and value)
    data = ";".join(str(value).strip() for name, value in rows if name and value)
    if not row_names or not data:
        return None
    return '<table id="1" title="表1-1 主要设计参数汇总" header="参数名称,参数值" rows="' + row_names + '" data="' + data + '"/>'


def normalize_civil_markup(full_text: str) -> str:
    """规范土木类文本，去掉提示词泄漏并收敛 drawing 标签。"""
    lines = full_text.splitlines()
    out = []
    seq = 0
    current_chapter_num: int | None = None
    current_table_idx = 0

    def _repair_civil_chart_block(block_lines: list[str], table_id: int) -> str | None:
        """把 chart 包裹的 JSON 表格块收敛成 table 标签。"""
        raw = "\n".join(block_lines)
        if "{" not in raw or "}" not in raw:
            return None
        m = re.search(r'\{.*\}', raw, re.S)
        if not m:
            return None
        try:
            payload = json.loads(m.group(0))
        except Exception:
            return None
        data = payload.get("data")
        if not isinstance(data, list) or not data:
            return None
        rows: list[list[str]] = []
        for row in data:
            if not isinstance(row, list) or not row:
                continue
            rows.append([str(cell).strip() for cell in row])
        if len(rows) < 2:
            return None
        header = ",".join(_escape_attr_text(cell) for cell in rows[0] if _escape_attr_text(cell))
        row_names: list[str] = []
        data_rows: list[str] = []
        for row in rows[1:]:
            if not row:
                continue
            row_names.append(_escape_attr_text(row[0]))
            data_rows.append("|".join(_escape_attr_text(cell) for cell in row[1:]))
        if not header or not row_names or not data_rows:
            return None
        title = _escape_attr_text(payload.get("title") or f"表{table_id}")
        return (
            f'<table id="{table_id}" title="{title}" header="{header}" '
            f'rows="{",".join(row_names)}" data="{";".join(data_rows)}"/>'
        )

    i = 0
    while i < len(lines):
        line = lines[i]
        low = line.lower().strip()
        compact = re.sub(r"\s+", "", low)
        if not compact:
            out.append(line)
            i += 1
            continue
        if _looks_like_civil_prompt_leak(line):
            i += 1
            continue
        chapter_match = re.match(r'^第\s*(\d+)\s*章\s*(.*)$', line.strip())
        if chapter_match:
            current_chapter_num = int(chapter_match.group(1))
            current_table_idx = 0
            out.append(line)
            i += 1
            continue
        if re.fullmatch(r'<drawing\s*/\s*>', low, re.I) or re.fullmatch(r'<table\s*/\s*>', low, re.I):
            i += 1
            continue
        if re.fullmatch(r'[\\/／]+', compact):
            i += 1
            continue
        if "<draing" in low or "</draing" in low or "<drowing" in low or "</drowing" in low:
            i += 1
            continue
        if low.startswith("<chart"):
            block = [line]
            j = i + 1
            while j < len(lines):
                block.append(lines[j])
                if lines[j].strip().lower().startswith("</chart"):
                    break
                j += 1
            repaired = _repair_civil_chart_block(block, seq + 1)
            if repaired:
                seq += 1
                out.append(repaired)
                i = j + 1
                continue
            i = j + 1
            continue
        if low.startswith("<table"):
            if _VALID_TABLE_TAG_RE.match(line.strip()):
                attrs = _extract_tag_attrs(line, "table") or {}
                seq += 1
                attrs["id"] = str(seq)
                current_table_idx += 1
                title = _infer_civil_table_title(current_chapter_num, attrs.get("title", ""), attrs.get("header", ""), line, current_table_idx)
                if title:
                    attrs["title"] = title
                canonical = _canonical_table_line(_repair_table_attrs(attrs)) if attrs else None
                out.append(canonical or line)
                i += 1
                continue
            block = [line]
            j = i + 1
            while j < len(lines):
                block.append(lines[j])
                if lines[j].strip().lower().startswith("</table"):
                    break
                j += 1
            current_table_idx += 1
            repaired = _repair_html_like_table_block(block, seq + 1, current_chapter_num, current_table_idx)
            if repaired:
                seq += 1
                out.append(repaired)
            i = j + 1
            continue
        if "<drawing" in low:
            parsed = _canonicalize_drawing_line(line, seq + 1, i, i + 1)
            if parsed:
                title = (parsed.get("title") or "").strip()
                desc = (parsed.get("description") or "").strip()
                if title in ("", "未命名", "图纸标题") and not desc:
                    i += 1
                    continue
                parsed["type"] = _infer_civil_drawing_type(title, desc, parsed.get("type") or "")
                seq += 1
                out.append(
                    f'<drawing id="{seq}" type="{parsed.get("type", "结构图")}" '
                    f'title="{title}" description="{desc}"/>'
                )
                i += 1
                continue
            if re.fullmatch(r'<drawing\s*/\s*>', low, re.I):
                i += 1
                continue
            tail = re.sub(r'<drawing\b.*?(?:/?>)', '', line, flags=re.I | re.DOTALL)
            tail = tail.replace('<drawing', '').strip()
            tail = re.sub(r'^[\\/／]+$', '', tail).strip()
            if tail:
                out.append(tail)
            i += 1
            continue
        if low.startswith("</dra") or low.startswith("</draw"):
            i += 1
            continue
        if re.fullmatch(r'[\\/／]\s*', line):
            i += 1
            continue
        out.append(line)
        i += 1

    cleaned = "\n".join(out)
    cleaned = re.sub(r'\n{3,}', '\n\n', cleaned)
    return cleaned.strip()


def _extract_tag_attrs(line: str, tag_name: str) -> dict | None:
    """从单行文本中提取标签属性，容忍缺失闭合符。"""
    m = re.search(rf'<{tag_name}\b', line, re.I)
    if not m:
        return None
    tag_text = line[m.start():]
    attrs = dict(re.findall(r'(\w+)\s*=\s*["\']([^"\']*)["\']', tag_text))
    return attrs or None


def _canonical_chart_line(attrs: dict) -> str | None:
    """将 chart 属性规范化为标准闭合标签。"""
    required = ("id", "title", "type", "x", "y")
    if any(not attrs.get(k) for k in required):
        return None
    parts = [
        f'id="{attrs["id"]}"',
        f'title="{attrs["title"]}"',
        f'type="{attrs["type"]}"',
        f'x="{attrs["x"]}"',
        f'y="{attrs["y"]}"',
    ]
    if attrs.get("legend"):
        parts.append(f'legend="{attrs["legend"]}"')
    if attrs.get("unit"):
        parts.append(f'unit="{attrs["unit"]}"')
    if attrs.get("data_source"):
        parts.append(f'data_source="{attrs["data_source"]}"')
    return "<chart " + " ".join(parts) + "/>"


def _canonical_table_line(attrs: dict) -> str | None:
    """将 table 属性规范化为标准闭合标签。"""
    required = ("id", "title", "header", "rows", "data")
    if any(not attrs.get(k) for k in required):
        return None
    parts = [
        f'id="{attrs["id"]}"',
        f'title="{attrs["title"]}"',
        f'header="{attrs["header"]}"',
        f'rows="{attrs["rows"]}"',
        f'data="{attrs["data"]}"',
    ]
    if attrs.get("data_source"):
        parts.append(f'data_source="{attrs["data_source"]}"')
    return "<table " + " ".join(parts) + "/>"


def _chart_to_table_attrs(chart_attrs: dict) -> dict | None:
    """把 chart 标签转成同源表格，作为管理类表格兜底。"""
    x_values = [x.strip() for x in (chart_attrs.get("x") or "").split(",") if x.strip()]
    y_series = [s.strip() for s in (chart_attrs.get("y") or "").split(";") if s.strip()]
    if not x_values or not y_series:
        return None
    legends = [s.strip() for s in (chart_attrs.get("legend") or "").split(",") if s.strip()]
    if legends and len(legends) == len(y_series):
        row_names = legends
    else:
        row_names = [f"系列{i+1}" for i in range(len(y_series))]
    data_rows = []
    for series in y_series:
        values = [v.strip() for v in series.split(",")]
        if len(values) < len(x_values):
            values.extend([""] * (len(x_values) - len(values)))
        values = values[:len(x_values)]
        data_rows.append("|".join(values))
    return {
        "id": "",
        "title": f'{chart_attrs.get("title", "数据")}明细表',
        "header": "指标," + ",".join(x_values),
        "rows": ",".join(row_names),
        "data": ";".join(data_rows),
        "data_source": chart_attrs.get("data_source", ""),
    }


def _normalize_sig_text(value: str) -> str:
    """把标签属性值压成稳定签名，便于判重。"""
    if value is None:
        return ""
    text = str(value).strip()
    text = re.sub(r"\s+", "", text)
    return text


def _chart_signature(attrs: dict) -> tuple:
    """图表签名：忽略 id 和 title，只按数据结构判重。"""
    x_vals = tuple(_normalize_sig_text(v) for v in (attrs.get("x") or "").split(",") if _normalize_sig_text(v))
    y_series = tuple(
        tuple(_normalize_sig_text(v) for v in series.split(",") if _normalize_sig_text(v))
        for series in (attrs.get("y") or "").split(";")
        if _normalize_sig_text(series)
    )
    legend = tuple(_normalize_sig_text(v) for v in (attrs.get("legend") or "").split(",") if _normalize_sig_text(v))
    unit = _normalize_sig_text(attrs.get("unit", ""))
    chart_type = _normalize_sig_text(attrs.get("type", ""))
    return (chart_type, x_vals, y_series, legend, unit)


def _table_signature(attrs: dict) -> tuple:
    """表格签名：忽略 id 和 title，只按结构化数据判重。"""
    header = tuple(_normalize_sig_text(v) for v in (attrs.get("header") or "").split(",") if _normalize_sig_text(v))
    rows = tuple(_normalize_sig_text(v) for v in (attrs.get("rows") or "").split(",") if _normalize_sig_text(v))
    data = tuple(
        tuple(_normalize_sig_text(v) for v in row.split("|") if _normalize_sig_text(v))
        for row in (attrs.get("data") or "").split(";")
        if _normalize_sig_text(row)
    )
    return (header, rows, data)


def _repair_table_attrs(attrs: dict) -> dict:
    """修复缺失 data 或 rows 挤入整行数据的表格属性。"""
    fixed = dict(attrs or {})
    header = [h.strip() for h in re.split(r"[;,，；|｜]+", fixed.get("header") or "") if h.strip()]
    if not header:
        return fixed

    table_shape = canonicalize_design_table_payload(
        fixed.get("header", ""),
        fixed.get("rows", ""),
        fixed.get("data", ""),
    )
    rows_matrix = table_shape.get("rows") or []
    if rows_matrix:
        fixed["rows"] = ",".join(row[0] for row in rows_matrix if row)
        fixed["data"] = ";".join("|".join(row[1:]) for row in rows_matrix if row)

    first_header = header[0]
    if any(k in first_header for k in ("评价", "统计", "满意度", "评分", "频次", "占比")):
        header[0] = "评价维度"
    else:
        header[0] = "对比项"
    fixed["header"] = ",".join(header)
    return fixed


def normalize_manage_markup(full_text: str) -> str:
    """规范经管类的 chart/table 行，修复闭合符、重复编号，并在缺表时补一个同源表。"""
    lines = full_text.splitlines()
    blocks = []
    current = []

    def flush_current():
        nonlocal current
        if current:
            blocks.append(current)
            current = []

    for line in lines:
        stripped = line.strip()
        if stripped == "---PAGE_BREAK---":
            flush_current()
            blocks.append([line])
            continue
        if re.match(r'^\d+\.\d+(?:\.\d+)?\s+', stripped) and current:
            flush_current()
        current.append(line)
    flush_current()

    out_lines = []
    seq = 0
    seen_chart_sigs = set()
    seen_table_sigs = set()
    fallback_added = False

    for block in blocks:
        if len(block) == 1 and block[0].strip() == "---PAGE_BREAK---":
            out_lines.append(block[0])
            continue

        entries = []
        first_chart_index = None
        first_chart_attrs = None
        has_table = False
        seen_table_in_block = False

        for idx, line in enumerate(block):
            chart_attrs = _extract_tag_attrs(line, "chart")
            if chart_attrs:
                sig = _chart_signature(chart_attrs)
                if sig in seen_chart_sigs:
                    continue
                seen_chart_sigs.add(sig)
                entries.append(("chart", line, chart_attrs))
                if first_chart_index is None:
                    first_chart_index = len(entries) - 1
                    first_chart_attrs = chart_attrs
                    # 兜底表只对“当前 chart 之前没有 table”的情况生效，避免后文表格回头压掉前文图表。
                    if (not fallback_added) and (not seen_table_in_block):
                        fallback = _chart_to_table_attrs(first_chart_attrs)
                        if fallback:
                            entries.append(("table_fallback", "", fallback))
                            fallback_added = True
                continue

            table_attrs = _extract_tag_attrs(line, "table")
            if table_attrs:
                table_attrs = _repair_table_attrs(table_attrs)
                sig = _table_signature(table_attrs)
                if sig in seen_table_sigs:
                    continue
                seen_table_sigs.add(sig)
                entries.append(("table", line, table_attrs))
                has_table = True
                seen_table_in_block = True
                continue

            entries.append(("text", line, None))

        for kind, line, attrs in entries:
            if kind == "text":
                out_lines.append(line)
                continue

            seq += 1
            attrs = dict(attrs)
            attrs["id"] = str(seq)
            if kind in ("chart",):
                canonical = _canonical_chart_line(attrs)
            else:
                canonical = _canonical_table_line(attrs)
            if canonical:
                out_lines.append(canonical)
            elif line:
                # 即使规范化失败也强制替换ID避免重复
                new_id = str(seq)
                import re as _re_for_id
                replaced = _re_for_id.sub(r'id="[^"]*"', f'id="{new_id}"', line, count=1)
                out_lines.append(replaced)
                continue

    return "\n".join(out_lines)


# ==================== 主要生成函数 ====================
def generate(profile: dict, paper_type: str, update=None) -> str:
    """生成完整论文文本（含标签）"""
    if update is None:
        def update(msg, prog): print(f"[{prog}%] {msg}")

    profile = dict(profile)  # 避免修改原对象

    if paper_type in ("管理", "经管", "manage", "mg"):
        return _generate_manage(profile, update)
    elif paper_type in ("设计", "design", "sj"):
        return _generate_design(profile, update)
    elif paper_type in ("机械", "mechanical", "mech", "mj"):
        return _generate_mechanical(profile, update)
    elif paper_type in ("土木", "civil", "cw"):
        return _generate_civil(profile, update)
    else:
        raise ValueError(f"不支持的论文类型: {paper_type}")


def _generate_manage(profile: dict, update) -> str:
    """生成经管类论文"""
    # 1. 摘要
    update("正在生成摘要...", 5)
    abstract = call_llm(mg_abstract_prompt(profile), max_tokens=600)
    abstract = clean_text(abstract)
    abstract = re.sub(r'^\*\*摘要\*\*', '', abstract.strip()).strip()
    abstract = re.sub(r'^摘要\s*', '', abstract).strip()

    # 2. 关键词
    keywords = call_llm(mg_kw_prompt(profile), max_tokens=200).strip()
    keywords = re.sub(r'^关键词[：:\s]*', '', keywords, flags=re.IGNORECASE).strip()

    # 3. 逐章生成
    chapters_content = {}
    chapter_names = [n for n, _ in MG_CHAPTERS]
    total_chapters = len(chapter_names)
    for idx, (name, num) in enumerate(MG_CHAPTERS):
        pct = 5 + int((idx + 0.5) / total_chapters * 80)
        update(f"正在生成第{num}章 {name}...", pct)
        prompt = mg_chapter_prompt(name, num, profile)
        content = call_llm(prompt, max_tokens=MG_WORD_LIMITS.get(name, 1000) * 2)
        content = clean_text(content)
        # 去除章节编号前缀
        content = re.sub(r'^第[一二三四五六\d]+章.*?\n', '', content).strip()
        chapters_content[name] = content

    # 4. 参考文献
    update("正在生成参考文献...", 88)
    refs_res = call_llm(mg_ref_prompt(profile), max_tokens=1500)
    refs = [r.strip() for r in refs_res.split('\n') if r.strip() and
            (r.strip()[0].isdigit() or r.strip().startswith('['))]
    refs = refs[:15]

    # 5. 构建TOC
    toc_lines = []
    for name, num in MG_CHAPTERS:
        toc_lines.append(f"第{num}章 {name}")
        content = chapters_content.get(name, "")
        for line in content.split('\n'):
            m2 = re.match(r'^(\d+\.\d+)\s+(.+)', line.strip())
            if m2:
                toc_lines.append(f"    {m2.group(1)} {m2.group(2)}")
            m3 = re.match(r'^(\d+\.\d+\.\d+)\s+(.+)', line.strip())
            if m3:
                toc_lines.append(f"        {m3.group(1)} {m3.group(2)}")
    toc_lines.append("参考文献")

    # 6. 合并
    txt = f"摘要\n{abstract}\n\n关键词\n{keywords}\n\n"
    txt += "---PAGE_BREAK---\n目录\n" + "\n".join(toc_lines) + "\n\n"
    txt += "---PAGE_BREAK---\n"  # 目录和正文之间加分页
    for name, num in MG_CHAPTERS:
        content = chapters_content.get(name, "")
        txt += f"第{num}章 {name}\n{content}\n\n"
        txt += "---PAGE_BREAK---\n"
    txt += "参考文献\n" + "\n".join(refs)

    # LLM格式校验（Layer 3）
    txt = _prepend_paper_title(txt, profile)
    try:
        txt = validate_tag_format(txt, "管理")
    except Exception:
        pass
    txt = normalize_manage_markup(txt)
    update("生成完成！", 100)
    return txt


def _generate_design(profile: dict, update) -> str:
    """生设计类论文"""
    # 1. 大纲
    update("正在生成大纲...", 3)
    outline = generate_outline(profile)

    # 2. 摘要
    update("正在生成摘要...", 8)
    abstract = call_llm(dj_abstract_prompt(profile), max_tokens=800)
    abstract = clean_text(abstract)
    abstract = _strip_design_abstract_prefix(abstract)

    # 3. 关键词
    keywords = call_llm(dj_kw_prompt(profile), max_tokens=200).strip()

    # 4. 逐章生成
    chapter_names = ["绪论", "理论基础", "问题发现与分析", "设计策略与方案", "总结与反思"]
    chapters_content = []
    total_ch = len(chapter_names)
    for idx, (name, num) in enumerate(zip(chapter_names, range(1, total_ch + 1))):
        pct = 10 + int((idx + 0.5) / total_ch * 75)
        update(f"正在生成第{num}章 {name}...", pct)
        prompt = dj_chapter_prompt(name, num, profile, outline)
        content = call_llm(prompt, max_tokens=SJ_WORD_LIMITS.get(name, 1000) * 2)
        content = clean_text(content)
        content = re.sub(r'^第[一二三四五六\d]+章.*?\n', '', content).strip()
        chapters_content.append((name, content))

    # 5. 图纸补齐
    update("正在分析图纸标签...", 88)
    full_text_before = ""
    for name, content in chapters_content:
        full_text_before += f"第{chapters_content.index((name, content))+1}章 {name}\n{content}\n\n"
    repaired_text = repair_drawing_tags(full_text_before, profile)
    # 如果补齐有变化，重新解析章节内容
    if repaired_text != full_text_before:
        # 按章节重新分割
        new_chapters = []
        for name, num in zip(chapter_names, range(1, len(chapter_names) + 1)):
            pat = re.compile(rf'第{num}章 {re.escape(name)}\n(.*?)(?:\n\n第{num+1}章|\n\n参考文献|$)', re.DOTALL)
            m = pat.search(repaired_text)
            if m:
                new_chapters.append((name, m.group(1).strip()))
            else:
                new_chapters.append((name, ""))
        if len(new_chapters) == len(chapter_names):
            chapters_content = new_chapters

    # 6. 参考文献
    update("正在生成参考文献...", 93)
    refs_res = call_llm(dj_ref_prompt(profile), max_tokens=1500)
    refs = [r.strip() for r in refs_res.split('\n') if r.strip() and
            (r.strip()[0].isdigit() or r.strip().startswith('['))]
    refs = refs[:15]

    # 7. TOC
    toc_lines = []
    for i, (name, content) in enumerate(chapters_content):
        toc_lines.append(f"第{i+1}章 {name}")
        for line in content.split('\n'):
            m2 = re.match(r'^(\d+\.\d+)\s+(.+)', line.strip())
            if m2:
                toc_lines.append(f"    {m2.group(1)} {m2.group(2)}")
            m3 = re.match(r'^(\d+\.\d+\.\d+)\s+(.+)', line.strip())
            if m3:
                toc_lines.append(f"        {m3.group(1)} {m3.group(2)}")
    toc_lines.append("参考文献")

    # 8. 合并
    txt = f"摘要\n{abstract}\n\n关键词\n{keywords}\n\n"
    txt += "---PAGE_BREAK---\n目录\n" + "\n".join(toc_lines) + "\n\n"
    txt += "---PAGE_BREAK---\n"  # 目录和正文之间加分页
    for i, (name, content) in enumerate(chapters_content):
        txt += f"第{i+1}章 {name}\n{content}\n\n"
        txt += "---PAGE_BREAK---\n"
    txt += "参考文献\n" + "\n".join(refs)

    # 先把明显失控的 HTML 风格表格片段收敛成标准 table 标签，再做常规校验
    txt = _repair_malformed_design_table_blocks(txt)

    # LLM格式校验（Layer 3）
    try:
        txt = validate_tag_format(txt, "设计")
    except Exception:
        pass
    txt = _prepend_paper_title(txt, profile)
    update("生成完成！", 100)
    return txt




def _check_calculation_consistency(body: str, chapter_name: str, machine_spec: dict, profile: dict) -> str:
    """检查计算与校核章节是否自洽，若发现明显数值矛盾则调用LLM重写。"""
    if chapter_name != "计算与校核":
        return body

    # 提取设计指标
    design_load = None
    load_item = machine_spec.get("load") if isinstance(machine_spec, dict) else None
    if isinstance(load_item, dict):
        design_load = load_item.get("value")
    elif isinstance(load_item, (int, float)):
        design_load = load_item
    if design_load is None and isinstance(profile, dict):
        for k in ("load", "载荷"):
            v = profile.get(k)
            if isinstance(v, dict):
                design_load = v.get("value")
            elif isinstance(v, (int, float)):
                design_load = v
            if design_load is not None:
                break

    if design_load is None:
        return body

    # 查找夹紧力/实际夹紧力数值
    clamp_values = []
    for m in re.finditer(r'(?:夹紧力|实际夹紧力|W|预紧力).*?(\d+(?:\.\d+)?)\s*(?:N|kN|牛)', body, flags=re.IGNORECASE):
        val = float(m.group(1))
        unit = m.group(0).lower()
        if 'kn' in unit or '千牛' in unit:
            val *= 1000
        clamp_values.append(val)

    if not clamp_values:
        return body

    # 如果计算出的夹紧力与设计指标相差过大(>10倍)，则重写
    max_clamp = max(clamp_values)
    if max_clamp > design_load * 10:
        print(f"  [一致性检查] 计算夹紧力 {max_clamp:.0f}N 远超设计指标 {design_load}N，正在重写...")
        fix_prompt = f"""你是一名机械设计计算审核专家。以下是一篇机械设计论文"计算与校核"章节的内容，发现其计算结果与设计指标不一致，请重写该章节，使其完全自洽。

设计指标：
- 载荷/夹紧力指标：{design_load} N

原文：
{{body}}

要求：
1. 保留所有小标题（5.1、5.2、5.3）
2. 铣削力计算、夹紧力计算、强度校核必须自洽
3. 计算得到的夹紧力应接近设计指标 {design_load} N，可通过调整安全系数、摩擦系数或接触方式实现
4. 所有公式和数值结果必须合理，不要出现明显荒谬的数值
5. 输出纯文本，不要markdown，<drawing/>标签如需保留请确保正确闭合
6. 保持原文字数规模
"""
        try:
            from core import call_llm
            fixed = call_llm(fix_prompt, max_tokens=3000)
            if fixed and len(fixed) > len(body) * 0.5:
                return fixed
        except Exception as e:
            print(f"  [一致性检查] 重写失败：{e}")
    return body

def _generate_mechanical(profile: dict, update) -> str:
    """生成机械设计类论文"""
    update("正在生成大纲...", 3)
    outline = mc_build_outline(profile)

    update("正在锁定 machine_spec...", 5)
    machine_spec_prompt = mc_machine_spec_prompt(profile, outline)
    machine_spec_raw = call_llm(machine_spec_prompt, max_tokens=1800)
    machine_spec = _safe_json_load(machine_spec_raw)
    if not isinstance(machine_spec, dict):
        machine_spec = _fallback_machine_spec(profile, outline)
    else:
        # 补齐关键字段，确保后续章节有稳定约束
        fallback_spec = _fallback_machine_spec(profile, outline)
        for key, value in fallback_spec.items():
            machine_spec.setdefault(key, value)

    update("正在生成摘要...", 8)
    abstract = call_llm(mc_abstract_prompt(profile), max_tokens=700)
    abstract = clean_text(abstract)

    update("正在生成关键词...", 10)
    keywords = call_llm(mc_kw_prompt(profile), max_tokens=200).strip()

    chapter_names = ["概述", "原理与方案分析", "总体设计", "关键部件设计", "计算与校核", "总结"]
    chapters_content = []
    total_ch = len(chapter_names)
    for idx, (name, num) in enumerate(zip(chapter_names, range(1, total_ch + 1))):
        pct = 12 + int((idx + 0.5) / total_ch * 72)
        update(f"正在生成第{num}章 {name}...", pct)
        prompt = mc_chapter_prompt(name, num, profile, outline, machine_spec)
        content = call_llm(prompt, max_tokens=MC_WORD_LIMITS.get(name, 1000) * 3)
        content = clean_text(content)
        body, chapter_json = _extract_mech_chapter_output(content, name, num, machine_spec)
        body = _check_calculation_consistency(body, name, machine_spec, profile)
        chapters_content.append((name, body, chapter_json))

    update("正在生成参考文献...", 93)
    refs_res = call_llm(mc_ref_prompt(profile), max_tokens=1500)
    refs = [r.strip() for r in refs_res.split('\n') if r.strip() and
            (r.strip()[0].isdigit() or r.strip().startswith('['))]
    refs = refs[:15]

    toc_lines = []
    for i, (name, body, _chapter_json) in enumerate(chapters_content):
        toc_lines.append(f"第{i+1}章 {name}")
        for line in body.split('\n'):
            m2 = re.match(r'^(\d+\.\d+)\s+(.+)', line.strip())
            if m2:
                toc_lines.append(f"    {m2.group(1)} {m2.group(2)}")
            m3 = re.match(r'^(\d+\.\d+\.\d+)\s+(.+)', line.strip())
            if m3:
                toc_lines.append(f"        {m3.group(1)} {m3.group(2)}")
    toc_lines.append("参考文献")

    txt = f"摘要\n{abstract}\n\n关键词\n{keywords}\n\n"
    txt += "---PAGE_BREAK---\n目录\n" + "\n".join(toc_lines) + "\n\n"
    txt += "---PAGE_BREAK---\n"
    for i, (name, body, chapter_json) in enumerate(chapters_content):
        rendered = _render_mech_chapter(body, chapter_json)
        txt += f"第{i+1}章 {name}\n{rendered}\n\n"
        txt += "---PAGE_BREAK---\n"
    txt += "参考文献\n" + "\n".join(refs)

    txt = _prepend_paper_title(txt, profile)
    txt = normalize_mechanical_drawing_tags(txt)
    try:
        txt = validate_tag_format(txt, "机械")
    except Exception:
        pass
    txt = normalize_mechanical_drawing_tags(txt)
    update("生成完成！", 100)
    return txt


def _generate_civil(profile: dict, update) -> str:
    """生成土木工程类论文"""
    update("正在生成大纲...", 3)
    outline = cv_get_chapter_outline(profile)

    update("正在生成摘要...", 8)
    abstract = call_llm(cv_abstract_prompt(profile), max_tokens=900)
    abstract = clean_text(abstract)
    abstract = re.sub(r'^\*\*摘要\*\*', '', abstract.strip()).strip()
    abstract = re.sub(r'^摘要\s*', '', abstract).strip()

    keywords = ""
    if "关键词" in abstract:
        parts = abstract.rsplit("关键词", 1)
        abstract = parts[0].strip()
        keywords = "关键词" + parts[1] if len(parts) > 1 else ""

    chapters_content = []
    chapter_names = [name for name, _ in outline]
    total_ch = len(chapter_names) or 1
    for idx, (name, num) in enumerate(outline):
        pct = 10 + int((idx + 0.5) / total_ch * 75)
        update(f"正在生成第{num}章 {name}...", pct)
        prompt = cv_chapter_prompt(name, num, profile, outline)
        content = call_llm(prompt, max_tokens=CIV_WORD_LIMITS.get(name, 1000) * 2)
        content = clean_text(content)
        content = re.sub(r'^第[一二三四五六\d]+章.*?\n', '', content).strip()
        content = normalize_civil_markup(content)
        chapters_content.append((name, content))

    update("正在生成参考文献...", 92)
    refs_res = call_llm(cv_ref_prompt(profile), max_tokens=1200)
    refs = [r.strip() for r in refs_res.split('\n') if r.strip() and
            (r.strip()[0].isdigit() or r.strip().startswith('['))]
    refs = refs[:15]

    toc_lines = []
    for i, (name, content) in enumerate(chapters_content):
        toc_lines.append(f"第{i+1}章 {name}")
        for line in content.split('\n'):
            m2 = re.match(r'^(\d+\.\d+)\s+(.+)', line.strip())
            if m2:
                toc_lines.append(f"    {m2.group(1)} {m2.group(2)}")
            m3 = re.match(r'^(\d+\.\d+\.\d+)\s+(.+)', line.strip())
            if m3:
                toc_lines.append(f"        {m3.group(1)} {m3.group(2)}")
    toc_lines.append("参考文献")

    txt = f"摘要\n{abstract}\n\n"
    if keywords:
        txt += f"{keywords}\n\n"
    txt += "---PAGE_BREAK---\n目录\n" + "\n".join(toc_lines) + "\n\n"
    txt += "---PAGE_BREAK---\n"
    for i, (name, content) in enumerate(chapters_content):
        txt += f"第{i+1}章 {name}\n{content}\n\n"
        txt += "---PAGE_BREAK---\n"
    txt += "参考文献\n" + "\n".join(refs)

    txt = normalize_civil_markup(txt)
    fallback_table = _build_civil_parameter_table(txt, profile)
    if fallback_table:
        head, sep, tail = txt.partition("---PAGE_BREAK---\n第2章")
        if sep and "<table" not in head:
            txt = head.rstrip() + "\n\n" + fallback_table + "\n\n" + sep + tail.lstrip()
    txt = _prepend_paper_title(txt, profile)
    try:
        txt = validate_tag_format(txt, "土木")
    except Exception:
        pass
    txt = normalize_civil_markup(txt)
    update("生成完成！", 100)
    return txt


# ==================== CLI入口 ====================


# ==================== LLM格式校验（Layer 3: 生成后检查） ====================

def validate_tag_format(full_text, paper_type="管理"):
    """生成论文后，用LLM轻量校验一遍标签格式正确性"""
    from core import call_llm
    tag_lines = []
    for line in full_text.split("\n"):
        if '<chart' in line or '<table' in line or '<drawing' in line:
            tag_lines.append(line.strip())
    if not tag_lines:
        return full_text
    tag_block = "\n".join(tag_lines)
    check_prompt = (
        "你是一个论文标签格式校验专家。以下是论文中的标签行，请检查每个标签：\n"
        "【检查规则】\n"
        "1. <chart/>标签：支持单系列和多系列；单系列时y用逗号分隔，多系列时y用分号分隔多个系列，legend数量必须与系列数一致\n"
        "2. <chart/>标签：多系列时每个系列的数据点数必须与x轴分类数一致；不得删维度、删线、删分类\n"
        "3. <table/>标签：rows属性只放行名，data用|和;分隔，行列数必须完整对应\n"
        "4. <drawing/>标签：id、type、title 至少要完整，且必须正确闭合（/>结尾）\n"
        "5. 所有标签必须正确闭合（/>结尾）\n"
        "6. 发现有问题的标签，输出修正后的版本\n\n"
        "【待检查的标签】\n"
        + tag_block + "\n\n"
        "【输出格式】\n"
        '严格JSON：{"issues_found": true/false, "corrections": [{"original": "...", "corrected": "...", "reason": "..."}]}'
    )
    print("  [LLM格式校验] 正在检查 %d 个标签..." % len(tag_lines))
    try:
        response = call_llm(check_prompt, max_tokens=2000)
        json_match = __import__('re').search(r'\{.*\}', response, __import__('re').DOTALL)
        if not json_match:
            print("  ⚠️ LLM validation returned no JSON, skipping")
            return full_text
        import json
        result = json.loads(json_match.group())
        if result.get("issues_found") and result.get("corrections"):
            print("  ⚠️ Found %d label format issues" % len(result['corrections']))
            modified = full_text
            for corr in result["corrections"]:
                orig = corr.get("original", "")
                repl = corr.get("corrected", "")
                if orig and repl and orig != repl and orig in modified:
                    modified = modified.replace(orig, repl, 1)
                    print("    ✅ Fixed: " + corr.get('reason', '')[:50])
            return modified
        else:
            print("  ✅ All tags valid")
            return full_text
    except Exception as e:
        print("  ⚠️ LLM validation error (skipped): " + str(e)[:80])
        return full_text

def main():
    import argparse
    parser = argparse.ArgumentParser(description="论文生成器")
    parser.add_argument("--profile", "-p", required=True, help="profile JSON 文件路径")
    parser.add_argument("--type", "-t", required=True, choices=["管理", "设计", "机械", "土木"],
                        help="论文类型：管理/设计/机械/土木")
    parser.add_argument("--output", "-o", default="", help="输出TXT路径（默认stdout）")
    args = parser.parse_args()

    with open(args.profile, 'r', encoding='utf-8') as f:
        profile = json.load(f)

    text = generate(profile, args.type)

    if args.output:
        os.makedirs(os.path.dirname(args.output) or '.', exist_ok=True)
        with open(args.output, 'w', encoding='utf-8') as f:
            f.write(text)
        print(f"论文已保存: {args.output}")
    else:
        print(text)


if __name__ == "__main__":
    main()
