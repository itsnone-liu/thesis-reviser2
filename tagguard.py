# -*- coding: utf-8 -*-
"""
tagguard.py — 确定性标签守卫（无LLM调用）
==========================================
在渲染前对 <chart/> <table/> <drawing/> 标签做：
  1. 拼写修复：<draing>/<drowing>/<taible>/<chrt> 等坏标签名 → 正确名
  2. 闭合修复：以 > 结尾但未自闭合的标签 → 补 />
  3. 属性校验与修复：
     - chart: x/y/legend 数量对齐、y 纯数字校验（非数字 → 降级 table_format）
     - table: rows 污染检测（数字/年份/百分比混入 → 重排进 data）、header/rows/data 对账
     - drawing: 四要素齐全性检查、description 长度警告
  4. 不可修复 → 降级或剔除，并全部记录，绝不静默消失

输出: (修复后文本, 结构化报告dict)
用法:
    from tagguard import audit_and_repair
    text2, report = audit_and_repair(text)
"""
import re

# 已知标签拼写错误映射（LLM 高频错拼）
_TAG_TYPOS = {
    "draing": "drawing", "drowing": "drawing", "drwaing": "drawing",
    "drawin": "drawing", "darwing": "drawing", "drawimg": "drawing",
    "taible": "table", "tablle": "table", "tabel": "table", "tabe": "table",
    "chrt": "chart", "char": "chart", "chatr": "chart", "chrat": "chart",
    "cart": "chart",
}
_VALID_TAGS = {"chart", "table", "drawing"}

_NUM_RE = re.compile(r'^-?\d+(?:\.\d+)?$')
_YEAR_RE = re.compile(r'^(?:19|20)\d{2}(?:年)?$')
_POLLUTE_RE = re.compile(r'^\d+(?:\.\d+)?\s*%?[岁年]?$')

_SPLIT_RE = re.compile(r"[;,，；|｜]+")


def _split_values(text: str) -> list:
    return [t.strip() for t in re.split(r"[;,，；]+", text or "") if t.strip()]


def _strip_dangerous(value: str) -> str:
    """去掉会破坏标签语法与属性解析的字符"""
    v = value.replace('"', "'").replace("”", "'").replace("“", "'")
    v = v.replace("<", "（").replace(">", "）")
    return v.strip()


# ==================== 1. 标签扫描（引号感知） ====================

def _scan_tag_spans(text: str):
    """扫描文本，产出 (start, end, tagname, raw, issues)。
    引号感知：属性值里的 > 与 /> 不会提前终止标签。"""
    spans = []
    i = 0
    n = len(text)
    while i < n:
        lt = text.find("<", i)
        if lt < 0:
            break
        m = re.match(r"<([A-Za-z]+)", text[lt:lt + 12])
        if not m:
            i = lt + 1
            continue
        name = m.group(1).lower()
        if name not in _VALID_TAGS and name not in _TAG_TYPOS:
            i = lt + 1
            continue
        # 引号感知扫描找标签结束
        j = lt + 1 + len(name)
        in_quote = None
        end = None
        while j < n:
            ch = text[j]
            if in_quote:
                if ch == in_quote:
                    in_quote = None
            elif ch in ('"', "“", "”"):
                in_quote = ch
            elif ch == ">" and not in_quote:
                # 找到 '>'；检查是否自闭合（前一个非空白字符是 '/'）
                k = j - 1
                while k >= lt and text[k] in " \t\r\n":
                    k -= 1
                self_closed = k >= lt and text[k] == "/"
                end = (j + 1, self_closed)
                break
            elif ch == "\n" and j - lt > 2000:
                break  # 超长无闭合，放弃
            j += 1
        if end is None:
            # 行内未闭合：取到行尾
            nl = text.find("\n", lt)
            endpos = nl if nl >= 0 else n
            spans.append((lt, endpos, name, text[lt:endpos], ["未闭合"]))
            i = endpos
            continue
        endpos, self_closed = end
        raw = text[lt:endpos]
        issues = []
        if not self_closed:
            issues.append("未自闭合")
        if name in _TAG_TYPOS:
            issues.append(f"标签名错拼({name})")
        spans.append((lt, endpos, name, raw, issues))
        i = endpos
    return spans


def _parse_attrs(raw: str) -> dict:
    return dict(re.findall(r'(\w+)=["\u201c\u201d\']([^"\u201c\u201d\']*)["\u201c\u201d\']', raw))


# ==================== 2. 各类型校验器 ====================

def _validate_chart(attrs: dict, seq: int):
    issues, actions = [], []
    x_labels = _split_values(attrs.get("x", ""))
    y_raw = (attrs.get("y", "") or "").strip()
    chart_type = attrs.get("type", "bar") or "bar"
    if not y_raw:
        return None, ["y属性为空，剔除"], "dropped"
    series = [_split_values(s) for s in re.split(r"[;；]", y_raw) if s.strip()]
    if not series:
        return None, ["y无可解析系列，剔除"], "dropped"

    # ---- gantt：x=任务列表，y=每组(start,duration)，特殊对齐规则 ----
    if chart_type == "gantt":
        for i, g in enumerate(series):
            if len(g) != 2:
                issues.append(f"第{i+1}组非start,duration两项，已截齐")
                series[i] = (g + [g[-1]] * 2)[:2]
        if not x_labels:
            x_labels = [f"任务{i+1}" for i in range(len(series))]
            issues.append("x缺失，按组数补任务名")
        n_eff = min(len(x_labels), len(series))
        if len(x_labels) != len(series):
            issues.append(f"任务数({len(x_labels)})与数据组数({len(series)})不一致，截齐为{n_eff}")
            x_labels, series = x_labels[:n_eff], series[:n_eff]
        new_attrs = {
            "id": attrs.get("id", str(seq)),
            "title": attrs.get("title", "") or f"图{seq}",
            "type": "gantt",
            "x": ",".join(x_labels),
            "y": ";".join(",".join(g) for g in series),
        }
        if attrs.get("unit"):
            new_attrs["unit"] = attrs["unit"]
        if attrs.get("data_source"):
            new_attrs["data_source"] = attrs["data_source"]
        return new_attrs, issues, ("repaired" if issues else "ok")

    # ---- 文字型图表类型：y 本来就是文字标签，不做数字校验 ----
    text_types = ("table_format", "structure")
    if chart_type not in text_types:
        non_numeric = any(not _NUM_RE.match(v.replace("%", "").replace(",", ""))
                          for s in series for v in s)
        if non_numeric:
            chart_type = "table_format"
            issues.append("y含非数字，降级为文字表格")
            actions.append("degraded_table_format")

    # x 缺失 → 从系列长度补
    if not x_labels:
        n = max(len(s) for s in series)
        x_labels = [f"类别{i+1}" for i in range(n)]
        issues.append("x缺失，按系列长度自动补类别名")
    # 数量对齐
    n_eff = min(len(x_labels), min(len(s) for s in series))
    if n_eff == 0:
        return None, ["x与y数量无法对齐，剔除"], "dropped"
    if len(x_labels) != n_eff or any(len(s) != n_eff for s in series):
        issues.append(f"x/系列长度不一致，统一截齐为{n_eff}点")
        actions.append("truncated")
        x_labels = x_labels[:n_eff]
        series = [s[:n_eff] for s in series]
    # legend 对齐
    legend = _split_values(attrs.get("legend", ""))
    if len(series) == 1 and len(legend) > 1:
        legend = legend[:1]
        issues.append("单系列但legend多项，已收敛为第一项")
    if len(series) > 1:
        if not legend:
            legend = [f"系列{i+1}" for i in range(len(series))]
            issues.append("多系列缺legend，自动补系列名")
        elif len(legend) != len(series):
            issues.append(f"legend数({len(legend)})与系列数({len(series)})不一致，已对齐")
            if len(legend) > len(series):
                legend = legend[:len(series)]
            else:
                legend = legend + [f"系列{i+1}" for i in range(len(legend), len(series))]

    new_attrs = {
        "id": attrs.get("id", str(seq)),
        "title": attrs.get("title", "") or f"图{seq}",
        "type": chart_type,
        "x": ",".join(x_labels),
        "y": ";".join(",".join(s) for s in series),
    }
    if legend:
        new_attrs["legend"] = ",".join(legend)
    if attrs.get("unit"):
        new_attrs["unit"] = attrs["unit"]
    if attrs.get("data_source"):
        new_attrs["data_source"] = attrs["data_source"]
    return new_attrs, issues, ("repaired" if issues else "ok")


def _validate_table(attrs: dict, seq: int):
    issues = []
    header_raw = attrs.get("header", "") or attrs.get("headers", "") or ""
    headers = [h.strip() for h in _SPLIT_RE.split(header_raw) if h.strip()]
    rows_raw = attrs.get("rows", "") or ""
    data_raw = attrs.get("data", "") or ""
    row_tokens = [r.strip() for r in _SPLIT_RE.split(rows_raw) if r.strip()]
    data_tokens = [t.strip() for t in _SPLIT_RE.split(data_raw) if t.strip()]

    if not headers:
        if not data_tokens and not row_tokens:
            return None, ["header/rows/data全空，剔除"], "dropped"
        headers = ["项目", "内容"]
        issues.append("header缺失，按两列补齐")

    # rows 污染：纯数字/年份/百分比出现在 rows
    polluted = [t for t in row_tokens if _POLLUTE_RE.match(t) or _YEAR_RE.match(t)]
    if polluted:
        data_rows = [r.strip() for r in re.split(r"[;；]+", data_raw) if r.strip()]
        if not data_rows:
            # data为空 → rows里塞的是整张表 → 按列数重排进data
            ncols = len(headers)
            if ncols > 1 and len(row_tokens) % ncols == 0 and len(row_tokens) >= ncols * 2:
                data_raw = ";".join(
                    "|".join(row_tokens[i:i + ncols])
                    for i in range(0, len(row_tokens), ncols))
                row_tokens, data_tokens = [], []
                issues.append(f"rows被数据污染({len(polluted)}项)，已按{ncols}列重排进data")
            else:
                row_tokens = [t for t in row_tokens if t not in polluted]
                issues.append(f"rows含{len(polluted)}项数据污染，已从rows剔除")
        elif all("|" in r or "｜" in r for r in data_rows) and len(data_rows) >= 2:
            # data为多列结构 → rows应等于每行首格；从data重建行名
            first_cells = [re.split(r"[|｜]+", r)[0].strip() for r in data_rows]
            if first_cells and all(not _POLLUTE_RE.match(c) for c in first_cells):
                row_tokens = first_cells
                issues.append(f"rows污染({len(polluted)}项)，已从data首格重建行名")
        else:
            row_tokens = [t for t in row_tokens if t not in polluted]
            issues.append(f"rows含{len(polluted)}项数据污染，已剔除")

    if row_tokens and data_tokens:
        ncols_data = max(len(headers) - 1, 1)
        need = len(row_tokens) * ncols_data
        if len(data_tokens) < need * 0.5:
            issues.append(f"data仅{len(data_tokens)}项，不足以填充{len(row_tokens)}行×{ncols_data}列")
    new_attrs = {
        "id": attrs.get("id", str(seq)),
        "title": attrs.get("title", "") or f"表{seq}",
        "header": ",".join(_strip_dangerous(h) for h in headers),
    }
    if row_tokens:
        new_attrs["rows"] = ",".join(_strip_dangerous(r) for r in row_tokens)
    if data_raw:
        # 保留原始行/列结构（|列分隔、;行分隔），只清洗危险字符
        new_attrs["data"] = ";".join(
            _strip_dangerous(r) for r in re.split(r"[;；]+", data_raw) if r.strip())
    if attrs.get("data_source"):
        new_attrs["data_source"] = attrs["data_source"]
    return new_attrs, issues, ("repaired" if issues else "ok")


def _validate_drawing(attrs: dict, seq: int):
    issues = []
    title = attrs.get("title", "")
    desc = attrs.get("description", "")
    dtype = attrs.get("type", "")
    if not title:
        issues.append("title缺失")
        title = f"图{seq}（未命名）"
    if not dtype:
        issues.append("type缺失，默认结构图")
        dtype = "结构图"
    if len(desc) < 50:
        issues.append(f"description过短({len(desc)}字<50)，生图信息可能不足")
    new_attrs = {
        "id": attrs.get("id", str(seq)),
        "type": dtype,
        "title": title,
    }
    if desc:
        new_attrs["description"] = desc
    return new_attrs, issues, ("repaired" if any("缺失" in i for i in issues) else "ok")


# ==================== 3. 序列化与主入口 ====================

def _serialize(tagname: str, attrs: dict) -> str:
    parts = [f"<{tagname}"]
    for k, v in attrs.items():
        parts.append(f'{k}="{_strip_dangerous(str(v))}"')
    parts.append("/>")
    return " ".join(parts)


def audit_and_repair(text: str):
    """主入口：校验并修复文本中的所有标签。
    返回 (修复后文本, 报告dict)。已正确且规范的标签保持原样不动。"""
    report = {
        "scanned": {"chart": 0, "table": 0, "drawing": 0},
        "fixed": {"typos": 0, "unclosed": 0, "repaired": 0, "degraded": 0, "dropped": 0},
        "details": {"charts": [], "tables": [], "drawings": []},
    }
    if not text or "<" not in text:
        return text, report

    spans = _scan_tag_spans(text)
    if not spans:
        return text, report

    replacements = []  # (start, end, new_text or None)
    for idx, (start, end, name, raw, scan_issues) in enumerate(spans, 1):
        canonical = _TAG_TYPOS.get(name, name)
        report["scanned"][canonical] += 1
        for si in scan_issues:
            if "错拼" in si:
                report["fixed"]["typos"] += 1
            elif "闭合" in si:
                report["fixed"]["unclosed"] += 1

        attrs = _parse_attrs(raw)
        entry = {"seq": idx, "title": attrs.get("title", "")[:40]}
        try:
            if canonical == "chart":
                new_attrs, issues, action = _validate_chart(attrs, idx)
                entry["kind"] = "chart"
            elif canonical == "table":
                new_attrs, issues, action = _validate_table(attrs, idx)
                entry["kind"] = "table"
            else:
                new_attrs, issues, action = _validate_drawing(attrs, idx)
                entry["kind"] = "drawing"
        except Exception as e:  # 校验器自身异常 → 保留原标签，记录
            entry.update({"action": "validator_error", "issues": [f"校验异常:{e}"]})
            report["details"][entry["kind"] + "s"].append(entry)
            continue

        entry["issues"] = issues
        entry["action"] = action
        if action == "dropped":
            replacements.append((start, end, None))
            report["fixed"]["dropped"] += 1
        elif issues:
            replacements.append((start, end, _serialize(canonical, new_attrs)))
            if action == "repaired":
                report["fixed"]["repaired"] += 1
            elif "降级" in "".join(issues):
                report["fixed"]["degraded"] += 1
        # 无issues且标签名正确、已自闭合 → 原样保留
        report["details"][entry["kind"] + "s"].append(entry)

    # 从后往前替换，避免偏移失效
    for start, end, new in sorted(replacements, key=lambda x: -x[0]):
        text = text[:start] + (new + "\n" if new else "") + text[end:]
    return text, report


def render_report_text(report: dict) -> str:
    """把报告渲染成一行人类可读摘要（用于进度日志）"""
    s = report["scanned"]; f = report["fixed"]
    if not any(s.values()):
        return ""
    parts = [f"标签守卫: chart {s['chart']} / table {s['table']} / drawing {s['drawing']}"]
    fixes = []
    if f["typos"]: fixes.append(f"错拼{f['typos']}")
    if f["unclosed"]: fixes.append(f"闭合{f['unclosed']}")
    if f["repaired"]: fixes.append(f"修复{f['repaired']}")
    if f["degraded"]: fixes.append(f"降级{f['degraded']}")
    if f["dropped"]: fixes.append(f"剔除{f['dropped']}")
    if fixes:
        parts.append("→ " + " ".join(fixes))
    return " ".join(parts)
