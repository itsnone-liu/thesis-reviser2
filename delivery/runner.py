from __future__ import annotations

import json
import os
import sys
import time
import threading
import traceback
import uuid
from pathlib import Path
from typing import Any, Callable

from .config import OUTPUT_ROOT, ROOT

if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import core
from profile import generate_profile
from generator import generate as generate_text
from renderer import render as render_docx
from reviser import _analyze_original_paper, _diagnose_and_reconstruct, _generate_revised


def normalize_paper_type(value: str) -> str:
    raw = (value or "").strip()
    if raw in {"管理", "经管", "manage", "mg"}:
        return "管理"
    if raw in {"设计", "design", "sj"}:
        return "设计"
    if raw in {"机械", "mechanical", "mech", "mj"}:
        return "机械"
    if raw in {"土木", "civil", "cw"}:
        return "土木"
    return "管理"


def make_task_id(prefix: str = "task") -> str:
    return f"{prefix}_{uuid.uuid4().hex[:12]}"


def ensure_task(task_id: str, msg: str = "已创建任务", progress: int = 0) -> dict:
    core.tasks_db.setdefault(task_id, {})
    core.tasks_db[task_id].update({
        "status": "running",
        "msg": msg,
        "progress": progress,
        "files": [],
    })
    return core.tasks_db[task_id]


def update_task(task_id: str, msg: str, progress: int | None = None, status: str | None = None, files: list[str] | None = None):
    payload = {"msg": msg}
    if progress is not None:
        payload["progress"] = progress
    if status is not None:
        payload["status"] = status
    if files is not None:
        payload["files"] = files
    core.tasks_db.setdefault(task_id, {}).update(payload)


def run_generate(profile: dict, paper_type: str, output_dir: str | Path, update: Callable[[str, int], None] | None = None) -> dict:
    paper_type = normalize_paper_type(paper_type)
    folder = Path(output_dir)
    folder.mkdir(parents=True, exist_ok=True)

    def emit(msg: str, prog: int):
        if update:
            update(msg, prog)

    emit("正在生成论文文本...", 10)
    txt = generate_text(profile, paper_type, update)
    txt_path = folder / "00_完整论文.txt"
    txt_path.write_text(txt, encoding="utf-8")

    emit("正在渲染DOCX...", 70)
    docx_path = folder / "论文_含图表.docx"
    render_docx(str(txt_path), str(docx_path), paper_type, None, None, update)

    emit("完成！", 100)
    return {"txt_path": str(txt_path), "docx_path": str(docx_path)}


def run_revise(input_docx: str, paper_type: str, output_dir: str | Path, update: Callable[[str, int], None] | None = None) -> dict:
    paper_type = normalize_paper_type(paper_type)
    folder = Path(output_dir)
    folder.mkdir(parents=True, exist_ok=True)

    def emit(msg: str, prog: int):
        if update:
            update(msg, prog)

    emit("正在分析原论文...", 10)
    original_text = core.extract_docx_text(input_docx)
    cover_info = core.extract_cover_info(input_docx)
    if not original_text:
        raise ValueError("无法读取DOCX")

    emit("正在分析画像...", 25)
    analysis = _analyze_original_paper(original_text, paper_type)
    emit("正在诊断重构...", 40)
    revision_plan = _diagnose_and_reconstruct(analysis, paper_type)
    report_path = folder / "诊断报告.txt"
    core.save_diagnosis_report(revision_plan, str(report_path))

    emit("正在生成修改版...", 60)
    txt = _generate_revised(analysis, revision_plan, paper_type, update)
    txt_path = folder / "00_完整论文.txt"
    txt_path.write_text(txt, encoding="utf-8")

    emit("正在渲染DOCX...", 80)
    docx_path = folder / "论文_修改版.docx"
    render_docx(str(txt_path), str(docx_path), paper_type, cover_info, None, update)

    emit("完成！", 100)
    return {"report_path": str(report_path), "txt_path": str(txt_path), "docx_path": str(docx_path)}


def build_profile_from_request(payload: dict, paper_type: str) -> dict:
    paper_type = normalize_paper_type(paper_type)
    if paper_type == "管理":
        return generate_profile("管理", payload.get("title", ""), payload.get("company", ""), payload.get("major", ""))
    if paper_type == "设计":
        return generate_profile("设计", payload.get("title", ""), payload.get("design_object", ""), payload.get("design_type", ""), payload.get("context", ""))
    if paper_type == "机械":
        return generate_profile("机械", payload.get("title", ""), payload.get("design_object", payload.get("object", "")), payload.get("major", "机械设计"), payload.get("context", ""))
    if paper_type == "土木":
        target = payload.get("object_name") or payload.get("design_object") or payload.get("company") or payload.get("object") or payload.get("title", "")
        return generate_profile("土木", payload.get("title", ""), target, payload.get("major", "土木工程"), payload.get("context", ""))
    return generate_profile("管理", payload.get("title", ""), payload.get("company", ""), payload.get("major", ""))


def run_standard_batch(input_json: str, output_dir: str, paper_type: str = "管理", limit: int = 0, skip_existing: bool = False):
    from batch_generate import batch_generate

    return batch_generate(input_json, output_dir, paper_type, limit, skip_existing)


def run_civil_batch(argv: list[str] | None = None) -> int:
    """
    Civil batch is still a special workflow. We keep it as a dedicated entrypoint,
    but make the workspace path portable by rewriting the source before execution.
    """
    civil_script = ROOT / "batch_generate_civil.py"
    source = civil_script.read_text(encoding="utf-8")
    source = source.replace(
        'PROJECT = "/root/project/workspace/thesis-reviser"',
        f'PROJECT = r"{ROOT}"',
    )
    namespace: dict[str, Any] = {"__name__": "__main__", "__file__": str(civil_script)}
    old_argv = sys.argv[:]
    try:
        if argv is not None:
            sys.argv = [str(civil_script), *argv]
        exec(compile(source, str(civil_script), "exec"), namespace)
    finally:
        sys.argv = old_argv
    return 0


def background_task(task_id: str, fn: Callable, *args, **kwargs) -> None:
    try:
        while True:
            with core.system_lock:
                if not core.is_system_busy:
                    core.is_system_busy = True
                    break
            core.tasks_db.setdefault(task_id, {}).update({
                "status": "queue",
                "msg": "系统繁忙，等待执行",
            })
            time.sleep(2)

        fn(task_id, *args, **kwargs)
    except Exception as exc:
        core.tasks_db.setdefault(task_id, {}).update({
            "status": "error",
            "msg": f"任务失败: {exc}",
        })
        traceback.print_exc()
    finally:
        with core.system_lock:
            core.is_system_busy = False


def start_single_task(kind: str, fn: Callable, *args, **kwargs) -> str:
    task_id = make_task_id(kind)
    ensure_task(task_id, f"正在{kind}", 0)
    threading.Thread(target=background_task, args=(task_id, fn, *args), kwargs=kwargs, daemon=True).start()
    return task_id
