# -*- coding: utf-8 -*-
"""
batch_generate.py — 批量生成论文
================================
读取批量名单 JSON，逐个生成画像、正文 TXT 与 DOCX。

支持类型：管理、设计、机械、土木。

默认输入：
  output/batch_titles.json

默认输出：
  output/batch_run/{name}_{student_id}/
"""
import os
import json
import argparse
from pathlib import Path
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from profile import generate_profile
from generator import generate
from renderer import render


def _load_entries(path: str) -> list[dict]:
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    if isinstance(data, dict):
        data = data.get("items", data.get("students", []))
    if not isinstance(data, list):
        raise ValueError(f"批处理输入格式不正确: {path}")
    return data


def _make_folder_name(entry: dict) -> str:
    name = str(entry.get("name", "")).strip() or "未命名"
    sid = str(entry.get("student_id", "")).strip() or "unknown"
    return f"{name}_{sid}"


def _build_cover(entry: dict, profile: dict) -> dict:
    date = str(entry.get("date", "")).strip()
    year, month, day = "", "", ""
    if date and "-" in date:
        parts = date.split("-")
        if len(parts) >= 3:
            year, month, day = parts[0], str(int(parts[1])), str(int(parts[2]))
    return {
        "title": profile.get("title", entry.get("title", "")),
        "name": entry.get("name", ""),
        "student_id": entry.get("student_id", ""),
        "major": entry.get("major", profile.get("major", "")),
        "level": entry.get("level", ""),
        "advisor": entry.get("advisor", ""),
        "year": year,
        "month": month,
        "day": day,
    }


def _build_profile(entry: dict, paper_type: str) -> dict:
    title = entry.get("title", "")
    company = entry.get("company", entry.get("object", ""))
    major = entry.get("major", "")
    context = entry.get("context", "")

    profile_path = entry.get("profile_path", "")
    if profile_path and os.path.isfile(profile_path):
        with open(profile_path, "r", encoding="utf-8") as f:
            profile = json.load(f)
        profile.setdefault("title", title)
        profile.setdefault("company", company)
        profile.setdefault("major", major)
        return profile

    if paper_type in ("管理", "经管", "manage", "mg"):
        profile = generate_profile("管理", title, company, major=major)
    elif paper_type in ("设计", "design", "sj"):
        profile = generate_profile("设计", title, company, major=major, context=context)
    elif paper_type in ("土木", "civil", "cw"):
        profile = generate_profile("土木", title, company or entry.get("object_name", ""), major=major or "土木工程", context=context)
    else:
        profile = generate_profile("机械", title, company, major=major, context=context)

    profile["title"] = title
    profile["company"] = company
    if major:
        profile["major"] = major
    if context:
        profile["context"] = context
    return profile


def batch_generate(input_json: str, output_dir: str, paper_type: str = "管理",
                   limit: int = 0, skip_existing: bool = False):
    entries = _load_entries(input_json)
    if limit and limit > 0:
        entries = entries[:limit]

    base = Path(output_dir)
    base.mkdir(parents=True, exist_ok=True)

    results = []
    for idx, entry in enumerate(entries, 1):
        folder = base / _make_folder_name(entry)
        folder.mkdir(parents=True, exist_ok=True)

        txt_path = folder / "paper.txt"
        docx_path = folder / "paper.docx"
        profile_path = folder / "profile.json"
        cover_path = folder / "cover.json"

        if skip_existing and txt_path.exists() and docx_path.exists():
            results.append({"folder": str(folder), "status": "skipped"})
            continue

        print(f"[{idx}/{len(entries)}] 正在处理: {folder.name}")
        profile = _build_profile(entry, paper_type)
        with open(profile_path, "w", encoding="utf-8") as f:
            json.dump(profile, f, ensure_ascii=False, indent=2)

        cover = _build_cover(entry, profile)
        with open(cover_path, "w", encoding="utf-8") as f:
            json.dump(cover, f, ensure_ascii=False, indent=2)

        txt = generate(profile, paper_type)
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write(txt)

        render(str(txt_path), str(docx_path), paper_type, cover_info=cover, drawing_folder=None)
        results.append({"folder": str(folder), "status": "completed"})

    return results


def main():
    parser = argparse.ArgumentParser(description="批量生成论文（管理/设计/机械/土木）")
    parser.add_argument("--input", "-i", default="output/batch_titles.json", help="批量名单 JSON")
    parser.add_argument("--output-dir", "-o", default="output/batch_run", help="批量输出目录")
    parser.add_argument("--type", "-t", default="管理", choices=["管理", "设计", "机械", "土木"], help="论文类型：管理/设计/机械/土木")
    parser.add_argument("--limit", "-l", type=int, default=0, help="只处理前 N 条，0 表示全部")
    parser.add_argument("--skip-existing", action="store_true", help="跳过已存在 TXT/DOCX 的目录")
    args = parser.parse_args()

    batch_generate(args.input, args.output_dir, args.type, args.limit, args.skip_existing)


if __name__ == "__main__":
    main()
