#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
batch_revise_design_papers.py — 批量修改设计类论文
=================================================
读取设计类 DOCX，调用现有设计类修改管道，输出到新的结果目录。

输出结构：
  /root/project/workspace/论文资料/整理/设计类_修改版/
    manifest.json
    {name}_{student_id}/
      profile.json
      cover.json
      diagnosis_report.txt
      paper.txt
      paper.docx
"""

import argparse
import json
import os
import sys
from copy import deepcopy
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

from core import extract_docx_text, extract_cover_info, save_diagnosis_report, normalize_cover_info
from reviser import _analyze_original_paper, _diagnose_and_reconstruct, _generate_revised
from renderer import render as render_docx


SOURCE_DIR = Path("/root/project/workspace/论文资料/整理/设计类")
OUTPUT_ROOT = Path("/root/project/workspace/论文资料/整理/设计类_修改版")
MANIFEST_PATH = OUTPUT_ROOT / "manifest.json"

# 之前确认的 6 篇需要彻底换题
TITLE_OVERRIDES = {
    "倪月华-029823410586-环境设计.docx": "基于游戏化叙事的儿童康复候诊空间设计研究",
    "廖颖-029823410656-环境设计.docx": "老工业厂房活化中的展陈空间改造设计研究",
    "张芳-029822410430-环境设计.docx": "面向夜间经济的微型市集场景营造设计研究",
    "王刚-029823410665-环境设计.docx": "宠物友好型就诊空间的环境设计研究",
    "蔡静瑶-029823410230-环境设计.docx": "高校实验教学楼公共界面的功能整合设计研究",
    "马泽宇-029821410470-环境设计.docx": "城市社区共享厨房的空间组织与使用体验设计研究",
}

REQUIRED_FILES = ("profile.json", "cover.json", "diagnosis_report.txt", "paper.txt", "paper.docx")


def _parse_identity(filename: str) -> tuple[str, str, str]:
    stem = filename[:-5] if filename.lower().endswith(".docx") else filename
    parts = stem.split("-")
    name = parts[0].strip() if len(parts) > 0 and parts[0].strip() else "未命名"
    student_id = parts[1].strip() if len(parts) > 1 and parts[1].strip() else "unknown"
    major = parts[2].strip() if len(parts) > 2 and parts[2].strip() else "环境设计"
    return name, student_id, major


def _folder_name(name: str, student_id: str) -> str:
    return f"{name}_{student_id}"


def _required_outputs_exist(folder: Path) -> bool:
    return all((folder / item).is_file() for item in REQUIRED_FILES)


def _write_json(path: Path, payload: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


def _write_manifest(records: list[dict]) -> None:
    OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)
    manifest = {
        "source_dir": str(SOURCE_DIR),
        "output_root": str(OUTPUT_ROOT),
        "count": len(records),
        "items": records,
    }
    _write_json(MANIFEST_PATH, manifest)


def _build_cover(raw_cover: dict, final_title: str, name: str, student_id: str, major: str) -> dict:
    cover = normalize_cover_info(raw_cover)
    cover["title"] = final_title
    cover["name"] = name
    cover["student_id"] = student_id
    if not cover.get("major"):
        cover["major"] = major
    return cover


def _process_one(docx_path: Path, skip_existing: bool = False) -> dict:
    name, student_id, major = _parse_identity(docx_path.name)
    folder = OUTPUT_ROOT / _folder_name(name, student_id)
    folder.mkdir(parents=True, exist_ok=True)

    record = {
        "source_file": str(docx_path),
        "output_folder": str(folder),
        "name": name,
        "student_id": student_id,
        "title": "",
        "status": "pending",
    }

    if skip_existing and _required_outputs_exist(folder):
        record["status"] = "skipped"
        record["title"] = ""
        return record

    try:
        original_text = extract_docx_text(str(docx_path))
        if not original_text:
            raise ValueError("无法读取DOCX文本")

        raw_cover = extract_cover_info(str(docx_path))
        analysis = _analyze_original_paper(original_text, "设计")
        analysis_for_pipeline = deepcopy(analysis)
        analysis_for_pipeline["major"] = analysis_for_pipeline.get("major") or major

        override_title = TITLE_OVERRIDES.get(docx_path.name, "").strip()
        if override_title:
            analysis_for_pipeline["title"] = override_title

        profile_path = folder / "profile.json"
        _write_json(profile_path, analysis_for_pipeline)

        revision_plan = _diagnose_and_reconstruct(analysis_for_pipeline, "设计")
        final_title = (
            override_title
            or revision_plan.get("revised_profile", {}).get("title", "").strip()
            or analysis_for_pipeline.get("title", "").strip()
            or analysis.get("title", "").strip()
            or docx_path.stem
        )
        revision_plan.setdefault("revised_profile", {})
        revision_plan["revised_profile"]["title"] = final_title

        report_path = folder / "diagnosis_report.txt"
        save_diagnosis_report(revision_plan, str(report_path))

        txt = _generate_revised(analysis_for_pipeline, revision_plan, "设计")
        txt_path = folder / "paper.txt"
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write(txt)

        cover = _build_cover(raw_cover, final_title, name, student_id, major)
        cover_path = folder / "cover.json"
        _write_json(cover_path, cover)

        docx_out = folder / "paper.docx"
        render_docx(str(txt_path), str(docx_out), "设计", cover_info=cover, drawing_folder=None)

        record["title"] = final_title
        record["status"] = "completed"
        return record

    except Exception as exc:
        record["status"] = "failed"
        record["error"] = str(exc)
        return record


def batch_revise_design_papers(limit: int = 0, skip_existing: bool = False) -> list[dict]:
    if not SOURCE_DIR.is_dir():
        raise FileNotFoundError(f"源目录不存在: {SOURCE_DIR}")

    OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)

    files = sorted(p for p in SOURCE_DIR.iterdir() if p.is_file() and p.suffix.lower() == ".docx")
    if limit and limit > 0:
        files = files[:limit]

    records: list[dict] = []
    for idx, path in enumerate(files, 1):
        print(f"[{idx}/{len(files)}] 正在处理: {path.name}", flush=True)
        record = _process_one(path, skip_existing=skip_existing)
        records.append(record)
        _write_manifest(records)
        print(f"  -> {record['status']}", flush=True)
        if record.get("title"):
            print(f"     title: {record['title']}", flush=True)
        if record.get("status") == "failed":
            print(f"     error: {record.get('error', '')}", flush=True)

    _write_manifest(records)
    return records


def main() -> None:
    parser = argparse.ArgumentParser(description="批量修改设计类论文")
    parser.add_argument("--limit", "-l", type=int, default=0, help="只处理前 N 篇，0 表示全部")
    parser.add_argument("--skip-existing", action="store_true", help="跳过已存在完整输出的目录")
    args = parser.parse_args()

    records = batch_revise_design_papers(limit=args.limit, skip_existing=args.skip_existing)
    completed = sum(1 for item in records if item.get("status") == "completed")
    skipped = sum(1 for item in records if item.get("status") == "skipped")
    failed = sum(1 for item in records if item.get("status") == "failed")
    print(
        f"\n完成: {completed} | 跳过: {skipped} | 失败: {failed}\n"
        f"输出目录: {OUTPUT_ROOT}",
        flush=True,
    )


if __name__ == "__main__":
    main()
