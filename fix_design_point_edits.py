#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fix_design_point_edits.py

对设计类论文修改结果做定点后处理，不调用 LLM，不重写正文。

目标：
- 去掉摘要正文前多余的“摘要/摘要：”前缀
- 清理正文里残留的 HTML 表格/图纸噪声标签
- 保持原有段落与字符格式，避免整段重排

输入目录默认：
  /root/project/workspace/论文资料/整理/设计类_修改版

输出目录默认：
  /root/project/workspace/论文资料/整理/设计类_定点修补版
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
from pathlib import Path
from typing import Iterable

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm
from docx.oxml.ns import qn


SOURCE_ROOT = Path("/root/project/workspace/论文资料/整理/设计类_修改版")
TARGET_ROOT = Path("/root/project/workspace/论文资料/整理/设计类_定点修补版")

ABSTRACT_PREFIXES = ("摘要：", "摘要:", "摘 要：", "摘 要:", "摘要 ", "摘要")
NOISE_LINE_RE = re.compile(
    r"^\s*</?(?:td|tr|table|caption|thead|tbody|th|drawing)\b[^>]*>\s*$",
    re.IGNORECASE,
)
NOISE_INLINE_RE = re.compile(
    r"</?(?:td|tr|table|caption|thead|tbody|th|drawing)\b[^>]*>",
    re.IGNORECASE,
)
TABLE_TAG_RE = re.compile(r"</?(?:table|tr|td|th|caption|thead|tbody|colgroup|col)", re.IGNORECASE)


def _delete_paragraph(paragraph) -> None:
    p = paragraph._element
    p.getparent().remove(p)


def _strip_prefix_keep_runs(paragraph, prefixes: Iterable[str]) -> bool:
    text = "".join(run.text or "" for run in paragraph.runs)
    if not text:
        return False

    matched_prefix = None
    for prefix in prefixes:
        if text.startswith(prefix):
            matched_prefix = prefix
            break
    if not matched_prefix:
        return False

    remaining_len = len(matched_prefix)
    consumed = 0
    for run in paragraph.runs:
        run_text = run.text or ""
        if not run_text:
            continue
        next_consumed = consumed + len(run_text)
        if next_consumed <= remaining_len:
            run.text = ""
            consumed = next_consumed
            continue
        if consumed < remaining_len:
            cut = remaining_len - consumed
            run.text = run_text[cut:]
        break
    return True


def _cleanup_table_like_line(line: str) -> tuple[str, bool]:
    raw = line or ""
    if not TABLE_TAG_RE.search(raw):
        return raw, False

    if raw.lstrip().lower().startswith("<table") and not re.search(r"</?(?:tr|td|th)", raw, flags=re.IGNORECASE):
        return "", True

    # 纯表格噪声直接删掉，避免留下空壳或属性碎片。
    if re.fullmatch(r"\s*</?(?:table|tr|td|th|caption|thead|tbody|colgroup|col)\b.*", raw, flags=re.IGNORECASE):
        cleaned = TABLE_TAG_RE.sub("", raw)
        cleaned = re.sub(r"[<>/]", "", cleaned).strip()
        if not re.sub(r"[\s|;:，,。；、\-]+", "", cleaned):
            return "", True

    cleaned = raw
    cleaned = re.sub(r"</?table", "\n", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"</?tr", "\n", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"</?(?:td|th)", "\t", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"</?(?:caption|thead|tbody|colgroup|col)", "\n", cleaned, flags=re.IGNORECASE)
    cleaned = cleaned.replace(">", " ")
    cleaned = cleaned.replace("/", " ")
    cleaned = re.sub(r"[ \t]+", " ", cleaned)
    cleaned = re.sub(r"\n\s+\n", "\n", cleaned)
    cleaned = re.sub(r"\n{3,}", "\n\n", cleaned)
    cleaned = cleaned.strip()
    if not cleaned or not re.sub(r"[\s|;:，,。；、\-\u3000]+", "", cleaned):
        return "", True
    return cleaned, True


def _set_run_font(run, font_name: str, font_size: int | None = None, bold: bool | None = None) -> None:
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
    if font_size is not None:
        run.font.size = Pt(font_size)
    if bold is not None:
        run.bold = bold


def _normalize_abstract_body(doc) -> None:
    start = None
    end = None
    for idx, p in enumerate(doc.paragraphs):
        text = (p.text or "").strip()
        if text == "摘要":
            start = idx + 1
            continue
        if start is not None and text.startswith("关键词"):
            end = idx
            break
    if start is None:
        return
    if end is None:
        end = min(start + 1, len(doc.paragraphs))

    body_size = 14
    for p in doc.paragraphs[end:]:
        t = (p.text or "").strip()
        if t and not t.startswith("目录"):
            if p.runs:
                size = p.runs[0].font.size
                if size is not None:
                    try:
                        body_size = int(round(size.pt))
                    except Exception:
                        pass
            break

    for p in doc.paragraphs[start:end]:
        text = (p.text or "").strip()
        if not text:
            continue
        for run in p.runs:
            if run.text:
                _set_run_font(run, "宋体", body_size, bold=False)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


def _extract_caption_text(text: str) -> str:
    cleaned = _cleanup_table_like_line(text)[0]
    cleaned = re.sub(r"^\s*caption\s*", "", cleaned, flags=re.IGNORECASE)
    cleaned = cleaned.replace("物理环境测量数据对照表", "物理环境测量数据对照表")
    return cleaned.strip()


def _is_table_block_text(text: str) -> bool:
    s = (text or "").strip()
    if not s:
        return False
    return bool(
        s.startswith("<table")
        or s.startswith("<caption")
        or s.startswith("<tr")
        or s.startswith("<td")
        or s.startswith("<th")
        or s.startswith("</table")
        or s.startswith("</tr")
        or s.startswith("</td")
        or s.startswith("</th")
    )


def _split_cells(text: str) -> list[str]:
    raw = _cleanup_table_like_line(text)[0].strip()
    if not raw:
        return []
    parts = [p.strip() for p in re.split(r"[\t|]+|\s{2,}", raw) if p.strip()]
    if len(parts) > 1:
        return parts
    return [raw]


def _insert_table_after(paragraph, rows: int, cols: int):
    table = paragraph._parent.add_table(rows=rows, cols=cols, width=Cm(16))
    paragraph._p.addnext(table._tbl)
    return table


def _rebuild_table_blocks(doc) -> None:
    blocks = []
    i = 0
    paras = doc.paragraphs
    while i < len(paras):
        text = paras[i].text or ""
        if not _is_table_block_text(text):
            i += 1
            continue
        start = i
        end = i + 1
        seen_content = False
        while end < len(paras):
            t = paras[end].text or ""
            stripped = t.strip()
            if not stripped:
                end += 1
                continue
            if _is_table_block_text(t):
                seen_content = True
                end += 1
                continue
            if seen_content:
                break
            end += 1
        blocks.append((start, end))
        i = end

    for start, end in reversed(blocks):
        block_paras = paras[start:end]
        if not block_paras:
            continue
        caption_idx = None
        rows: list[list[str]] = []
        current_row: list[str] = []
        anchor_idx = start
        has_row_marker = False
        leading_header_cells = 0
        seen_data_cells = False
        for rel_idx, p in enumerate(block_paras):
            raw = (p.text or "").strip()
            if not raw:
                continue
            if raw.startswith("<caption"):
                caption_idx = start + rel_idx
                anchor_idx = caption_idx
                caption_text = _extract_caption_text(raw)
                if caption_text:
                    p.text = caption_text
                    if p.alignment is None:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                continue
            if raw.startswith("<table") or raw.startswith("</table"):
                continue
            if raw.startswith("<tr"):
                has_row_marker = True
                if current_row:
                    rows.append(current_row)
                    current_row = []
                continue
            if raw.startswith("</tr"):
                has_row_marker = True
                if current_row:
                    rows.append(current_row)
                    current_row = []
                continue
            if "<td" in raw.lower() or "<th" in raw.lower():
                cells = _split_cells(raw)
                if not seen_data_cells and raw.lower().startswith("<th"):
                    leading_header_cells += len(cells)
                if raw.lower().startswith("<td"):
                    seen_data_cells = True
                if len(cells) > 1 and not current_row:
                    current_row.extend(cells)
                else:
                    current_row.extend(cells)
                continue
            # 已经进入表格区后，遇到纯文本但还没形成行时，尝试按单元格补入
            if rows or current_row:
                current_row.append(raw)

        if current_row:
            rows.append(current_row)

        rows = [r for r in rows if any(c.strip() for c in r)]
        if not has_row_marker and len(rows) == 1 and leading_header_cells > 1:
            flat = rows[0]
            if len(flat) > leading_header_cells:
                chunk = leading_header_cells
                header = flat[:chunk]
                body = flat[chunk:]
                rebuilt = [header]
                for i in range(0, len(body), chunk):
                    rebuilt.append(body[i:i + chunk])
                rows = rebuilt
        if not rows:
            for idx in range(end - 1, start - 1, -1):
                _delete_paragraph(paras[idx])
            continue

        cols = max(len(r) for r in rows)
        anchor_paragraph = paras[anchor_idx - 1] if anchor_idx > 0 else paras[start]
        table = _insert_table_after(anchor_paragraph, len(rows), cols)
        table.style = "Table Grid"
        for r_idx, row in enumerate(rows):
            for c_idx in range(cols):
                cell = table.rows[r_idx].cells[c_idx]
                cell.text = row[c_idx] if c_idx < len(row) else ""
                for para in cell.paragraphs:
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in para.runs:
                        if run.text:
                            _set_run_font(run, "宋体", 10)
        for idx in range(end - 1, start - 1, -1):
            if caption_idx is not None and idx == caption_idx:
                continue
            _delete_paragraph(paras[idx])


def _clean_text_lines(lines: list[str]) -> tuple[list[str], int]:
    cleaned: list[str] = []
    fixed = 0
    in_abstract = False
    abstract_fixed = False
    for raw in lines:
        line = raw.rstrip("\n")
        stripped = line.strip()
        if not stripped:
            cleaned.append(line)
            continue

        if stripped in ("摘要", "摘 要"):
            in_abstract = True
            cleaned.append(line)
            continue

        if stripped.startswith("关键词"):
            in_abstract = False

        if in_abstract and not abstract_fixed:
            new_line = line
            for prefix in ABSTRACT_PREFIXES:
                if stripped.startswith(prefix):
                    new_line = re.sub(rf"^\s*{re.escape(prefix)}\s*", "", line, count=1)
                    if new_line != line:
                        fixed += 1
                    abstract_fixed = True
                    break
            cleaned.append(new_line)
            continue

        if NOISE_LINE_RE.match(stripped):
            fixed += 1
            continue

        if TABLE_TAG_RE.search(line):
            new_line, changed = _cleanup_table_like_line(line)
            if changed:
                fixed += 1
                if new_line:
                    cleaned.extend([seg for seg in new_line.splitlines() if seg.strip()])
                continue

        if NOISE_INLINE_RE.search(line):
            new_line = NOISE_INLINE_RE.sub("", line).strip()
            if new_line:
                cleaned.append(new_line)
            fixed += 1
            continue

        cleaned.append(line)

    return cleaned, fixed


def _sanitize_docx(src: Path, dst: Path) -> dict:
    doc = Document(src)
    changes = {"abstract_prefix": 0, "noise_paragraphs": 0}

    # 先清理正文里的明显噪声段落，再处理摘要前缀
    paragraphs = list(doc.paragraphs)
    in_abstract = False
    abstract_fixed = False
    for p in paragraphs:
        text = p.text.strip()
        if not text:
            continue

        if text in ("摘要", "摘 要"):
            in_abstract = True
            continue

        if text.startswith("关键词"):
            in_abstract = False

        if NOISE_LINE_RE.match(text):
            _delete_paragraph(p)
            changes["noise_paragraphs"] += 1
            continue

        if NOISE_INLINE_RE.search(text):
            # 只删除纯噪声段落，不对混合正文做大范围改写
            if NOISE_INLINE_RE.fullmatch(text):
                _delete_paragraph(p)
                changes["noise_paragraphs"] += 1
                continue

        if in_abstract and not abstract_fixed:
            if _strip_prefix_keep_runs(p, ABSTRACT_PREFIXES):
                changes["abstract_prefix"] += 1
            abstract_fixed = True

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                cell_text = cell.text or ""
                if not cell_text:
                    continue
                if not TABLE_TAG_RE.search(cell_text) and not NOISE_INLINE_RE.search(cell_text):
                    continue
                new_text, changed = _cleanup_table_like_line(cell_text)
                if not changed:
                    continue
                changes["noise_paragraphs"] += 1
                if not new_text:
                    cell.text = ""
                else:
                    cell.text = new_text

    _rebuild_table_blocks(doc)
    _normalize_abstract_body(doc)

    dst.parent.mkdir(parents=True, exist_ok=True)
    doc.save(dst)
    return changes


def _copy_meta_files(src_dir: Path, dst_dir: Path) -> None:
    for name in ("profile.json", "cover.json", "diagnosis_report.txt", "paper.txt"):
        src = src_dir / name
        if not src.is_file():
            continue
        if name.endswith(".json"):
            data = json.loads(src.read_text(encoding="utf-8"))
            (dst_dir / name).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        elif name == "paper.txt":
            lines = src.read_text(encoding="utf-8").splitlines()
            cleaned, _ = _clean_text_lines(lines)
            (dst_dir / name).write_text("\n".join(cleaned), encoding="utf-8")
        else:
            shutil.copy2(src, dst_dir / name)


def fix_all(source_root: Path = SOURCE_ROOT, target_root: Path = TARGET_ROOT, limit: int = 0, skip_existing: bool = False):
    if not source_root.is_dir():
        raise FileNotFoundError(f"源目录不存在: {source_root}")
    target_root.mkdir(parents=True, exist_ok=True)

    folders = sorted([p for p in source_root.iterdir() if p.is_dir()])
    if limit and limit > 0:
        folders = folders[:limit]

    results = []
    for idx, src_dir in enumerate(folders, 1):
        dst_dir = target_root / src_dir.name
        docx_src = src_dir / "paper.docx"
        docx_dst = dst_dir / "paper.docx"

        if skip_existing and docx_dst.is_file():
            results.append({"folder": src_dir.name, "status": "skipped"})
            continue

        if not docx_src.is_file():
            results.append({"folder": src_dir.name, "status": "failed", "error": "missing paper.docx"})
            continue

        print(f"[{idx}/{len(folders)}] 修补 {src_dir.name}", flush=True)
        try:
            dst_dir.mkdir(parents=True, exist_ok=True)
            _copy_meta_files(src_dir, dst_dir)
            changes = _sanitize_docx(docx_src, docx_dst)
            results.append({"folder": src_dir.name, "status": "completed", **changes})
            print(f"  -> completed abstract_prefix={changes['abstract_prefix']} noise_paragraphs={changes['noise_paragraphs']}", flush=True)
        except Exception as exc:
            results.append({"folder": src_dir.name, "status": "failed", "error": str(exc)})
            print(f"  -> failed: {exc}", flush=True)

    manifest = {
        "source_root": str(source_root),
        "target_root": str(target_root),
        "count": len(results),
        "completed": sum(1 for r in results if r["status"] == "completed"),
        "failed": sum(1 for r in results if r["status"] == "failed"),
        "skipped": sum(1 for r in results if r["status"] == "skipped"),
        "items": results,
    }
    (target_root / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return manifest


def main():
    parser = argparse.ArgumentParser(description="设计类论文定点后处理修补")
    parser.add_argument("--source", default=str(SOURCE_ROOT), help="源目录")
    parser.add_argument("--target", default=str(TARGET_ROOT), help="目标目录")
    parser.add_argument("--limit", type=int, default=0, help="仅处理前 N 个目录")
    parser.add_argument("--skip-existing", action="store_true", help="跳过已存在的目标 paper.docx")
    args = parser.parse_args()

    manifest = fix_all(Path(args.source), Path(args.target), args.limit, args.skip_existing)
    print(
        f"\n完成: {manifest['completed']} | 跳过: {manifest['skipped']} | 失败: {manifest['failed']}\n"
        f"目标目录: {manifest['target_root']}",
        flush=True,
    )


if __name__ == "__main__":
    main()
