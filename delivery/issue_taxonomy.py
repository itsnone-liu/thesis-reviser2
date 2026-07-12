from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re
from typing import Iterable


CATEGORY_ORDER = ("code_fix", "prompt_fix", "paper_fix", "regress", "ignore")


@dataclass(frozen=True)
class IssueHit:
    category: str
    source: str
    line_number: int
    line: str
    suggested_action: str

    def as_dict(self) -> dict[str, str | int]:
        return {
            "category": self.category,
            "source": self.source,
            "line_number": self.line_number,
            "line": self.line,
            "suggested_action": self.suggested_action,
        }


_CODE_HINTS = (
    ".py",
    "core.py",
    "renderer.py",
    "generator.py",
    "batch_",
    "service.py",
    "web.py",
    "函数",
    "方法",
    "正则",
    "XML",
    "解析",
    "渲染",
    "导入",
    "缓存",
    "命名空间",
    "tab_stops",
    "finalize_docx",
    "txt_to_docx_safe",
    "add_t()",
    "CLI",
    "接口",
    "代码",
)

_PROMPT_HINTS = (
    "prompt",
    "提示词",
    "DRAWING_RULE",
    "LLM",
    "生图",
    "图表协议",
    "标签规范",
    "格式铁律",
    "约束",
    "输出规范",
    "校验 JSON",
    "prompt 强度",
    "图示 prompt",
)

_PAPER_HINTS = (
    "摘要",
    "封面",
    "目录",
    "页码",
    "分节符",
    "页脚",
    "表格",
    "图纸",
    "图表",
    "图片",
    "正文",
    "标题",
    "参考文献",
    "排版",
    "Word",
    "WPS",
    "DOCX",
    "章节",
)

_REGRESS_HINTS = (
    "重复",
    "回归",
    "回退",
    "复现",
    "再次",
    "又出现",
    "旧图",
    "旧版",
    "污染",
    "错配",
    "二次",
    "残留",
    "回填",
    "复发",
    "重渲",
    "重复 id",
    "重复图",
)

_IGNORE_HINTS = (
    "验证",
    "结果",
    "现状",
    "当前结论",
    "已完成",
    "已稳定",
    "已确认",
    "抽检通过",
    "未见",
    "正常",
    "备注",
    "待规划",
    "明日继续",
    "工作内容",
    "测试输出文件",
    "下次",
    "建议",
    "可运行",
    "已输出",
    "已生成",
    "无需",
    "计划",
    "总结",
)

_ISSUE_SIGNAL = re.compile(
    r"(问题|修复|修正|错误|缺失|失败|冲突|不一致|错位|污染|残留|空白|近白|偏小|"
    r"丢失|失效|错配|回归|复现|回退|重复|待解决|需确认|风险|不足|异常|过宽|过弱)"
)


_SUGGESTED_ACTION = {
    "code_fix": "回到对应代码路径做最小修复，并补回归验证，避免把同类逻辑再带坏。",
    "prompt_fix": "收紧提示词或输出协议，补足硬约束、示例和校验规则。",
    "paper_fix": "在成品文档层修正结构和版式，并复查生成结果是否一致。",
    "regress": "先定位最近一次可用版本，只做最小增量回退/修补，再补复测。",
    "ignore": "保留为状态/验证记录，不纳入治理问题集。",
}


def _normalize_line(line: str) -> str:
    return re.sub(r"\s+", " ", line).strip()


def _is_heading(line: str) -> bool:
    return line.lstrip().startswith("#")


def _is_code_fence(line: str) -> bool:
    stripped = line.lstrip()
    return stripped.startswith("```") or stripped.startswith("~~~")


def _is_tree_line(line: str) -> bool:
    return line.lstrip().startswith(("├", "│", "└", "┌", "┐", "┬", "┼", "─"))


def _looks_actionable(line: str) -> bool:
    if not line:
        return False
    if _ISSUE_SIGNAL.search(line):
        return True
    return False


def _score_hits(line: str, hints: tuple[str, ...]) -> int:
    lowered = line.lower()
    score = 0
    for hint in hints:
        if hint.lower() in lowered:
            score += 1
    return score


def classify_text(line: str) -> tuple[str, str]:
    text = _normalize_line(line)
    if not text:
        return "ignore", _SUGGESTED_ACTION["ignore"]

    if _score_hits(text, _REGRESS_HINTS):
        return "regress", _SUGGESTED_ACTION["regress"]

    prompt_score = _score_hits(text, _PROMPT_HINTS)
    code_score = _score_hits(text, _CODE_HINTS)
    paper_score = _score_hits(text, _PAPER_HINTS)

    if prompt_score and prompt_score >= code_score:
        return "prompt_fix", _SUGGESTED_ACTION["prompt_fix"]
    if code_score:
        return "code_fix", _SUGGESTED_ACTION["code_fix"]
    if paper_score:
        return "paper_fix", _SUGGESTED_ACTION["paper_fix"]

    if _score_hits(text, _IGNORE_HINTS) or not _looks_actionable(text):
        return "ignore", _SUGGESTED_ACTION["ignore"]

    return "paper_fix", _SUGGESTED_ACTION["paper_fix"]


def iter_issue_lines(text: str) -> Iterable[tuple[int, str]]:
    in_code_block = False
    for line_number, raw in enumerate(text.splitlines(), start=1):
        line = _normalize_line(raw)
        if not line:
            continue
        if _is_code_fence(line):
            in_code_block = not in_code_block
            continue
        if in_code_block:
            continue
        if _is_heading(line) or _is_tree_line(line):
            continue
        if _looks_actionable(line) or any(token in line for token in _CODE_HINTS + _PROMPT_HINTS + _REGRESS_HINTS):
            yield line_number, line


def scan_markdown_file(path: Path) -> list[IssueHit]:
    text = path.read_text(encoding="utf-8", errors="ignore")
    hits: list[IssueHit] = []
    source = str(path.resolve())
    for line_number, line in iter_issue_lines(text):
        category, suggested_action = classify_text(line)
        hits.append(
            IssueHit(
                category=category,
                source=source,
                line_number=line_number,
                line=line,
                suggested_action=suggested_action,
            )
        )
    return hits


def _empty_bucket() -> dict[str, object]:
    return {"count": 0, "samples": []}


def build_issue_report(paths: Iterable[Path]) -> dict:
    paths = list(paths)
    summary = {category: _empty_bucket() for category in CATEGORY_ORDER}
    issues: list[dict[str, str | int]] = []
    ignored: list[dict[str, str | int]] = []

    for path in paths:
        candidates: list[Path]
        if path.is_dir():
            candidates = [p for p in path.rglob("*") if p.suffix.lower() in {".md", ".txt"}]
        else:
            candidates = [path]

        for candidate in candidates:
            if not candidate.exists() or candidate.suffix.lower() not in {".md", ".txt"}:
                continue
            for hit in scan_markdown_file(candidate):
                bucket = summary[hit.category]
                bucket["count"] = int(bucket["count"]) + 1
                sample = {
                    "source": hit.source,
                    "line_number": hit.line_number,
                    "line": hit.line,
                    "suggested_action": hit.suggested_action,
                }
                if len(bucket["samples"]) < 8:
                    bucket["samples"].append(sample)
                if hit.category == "ignore":
                    if len(ignored) < 20:
                        ignored.append(sample)
                    continue
                issues.append(hit.as_dict())

    return {
        "targets": [str(Path(p).resolve()) for p in paths],
        "summary": summary,
        "issues": issues,
        "ignored": ignored,
        "taxonomy": {category: _SUGGESTED_ACTION[category] for category in CATEGORY_ORDER},
    }


def build_issue_summary(paths: Iterable[Path]) -> dict:
    report = build_issue_report(paths)
    return report["summary"]
