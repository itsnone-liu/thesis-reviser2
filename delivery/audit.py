from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Iterable

from .config import ROOT
from .issue_taxonomy import CATEGORY_ORDER, build_issue_report


DEFAULT_TARGETS = [
    ROOT / "修改记录.md",
    ROOT / "今日进度_20260704.md",
]


def _format_source(source: str, line_number: int) -> str:
    return f"{source}:{line_number}"


def render_markdown(report: dict) -> str:
    summary = report["summary"]
    issues = report["issues"]
    ignored = report["ignored"]

    lines: list[str] = []
    lines.append("# 交付审计报告")
    lines.append("")
    lines.append("## 目标")
    for target in report["targets"]:
        lines.append(f"- `{target}`")
    lines.append("")
    lines.append("## 分类概览")
    lines.append("| category | count |")
    lines.append("| --- | ---: |")
    for category in CATEGORY_ORDER:
        lines.append(f"| `{category}` | {summary[category]['count']} |")
    lines.append("")
    lines.append("## 治理明细")
    lines.append("| category | source | line | suggested_action |")
    lines.append("| --- | --- | --- | --- |")
    for item in issues:
        source = _format_source(item["source"], int(item["line_number"]))
        line = str(item["line"]).replace("|", "\\|")
        action = str(item["suggested_action"]).replace("|", "\\|")
        lines.append(f"| `{item['category']}` | `{source}` | {line} | {action} |")
    if not issues:
        lines.append("| `-` | `-` | `-` | `-` |")
    lines.append("")
    lines.append("## 忽略样本")
    if ignored:
        for item in ignored[:10]:
            source = _format_source(item["source"], int(item["line_number"]))
            lines.append(f"- `{source}` {item['line']} ({item['suggested_action']})")
    else:
        lines.append("- 无")
    lines.append("")
    lines.append("## 类目说明")
    for category in CATEGORY_ORDER:
        lines.append(f"- `{category}`: {report['taxonomy'][category]}")
    return "\n".join(lines)


def _load_targets(paths: Iterable[str]) -> list[Path]:
    return [Path(p) for p in paths]


def main() -> None:
    parser = argparse.ArgumentParser(description="论文系统问题梳理/抽检报告")
    parser.add_argument("--path", "-p", action="append", default=[], help="要扫描的文件或目录，可重复")
    parser.add_argument("--json-output", default="", help="输出 JSON 报告路径")
    parser.add_argument("--md-output", default="", help="输出 Markdown 报告路径")
    args = parser.parse_args()

    targets = _load_targets(args.path) if args.path else list(DEFAULT_TARGETS)
    report = build_issue_report(targets)
    markdown = render_markdown(report)
    json_text = json.dumps(report, ensure_ascii=False, indent=2)

    if args.json_output:
        Path(args.json_output).write_text(json_text, encoding="utf-8")
    if args.md_output:
        Path(args.md_output).write_text(markdown, encoding="utf-8")

    if not args.json_output and not args.md_output:
        print(markdown)
        print("")
        print("## JSON")
        print("```json")
        print(json_text)
        print("```")


if __name__ == "__main__":
    main()
