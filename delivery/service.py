from __future__ import annotations

import sys
from pathlib import Path

from fastapi import FastAPI, File, Form, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse

from .config import OUTPUT_ROOT, ROOT, SERVICE_HOST, SERVICE_PORT
from .schemas import GeneratePayload, ProfilePayload, TaskSnapshot, TaskStartResponse, model_to_dict
from .runner import (
    build_profile_from_request,
    normalize_paper_type,
    run_generate,
    run_revise,
    start_single_task,
)

if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import core


app = FastAPI(title="论文系统交付服务")


def _task_folder(task_id: str) -> Path:
    folder = OUTPUT_ROOT / task_id
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def _update(task_id: str, msg: str, progress: int):
    core.tasks_db.setdefault(task_id, {}).update({"msg": msg, "progress": progress})


@app.get("/", response_class=HTMLResponse)
async def home():
    return """
    <html><body style="font-family: sans-serif; padding: 24px;">
    <h1>论文系统交付服务</h1>
    <ul>
      <li><a href="/health">/health</a></li>
      <li><a href="/docs">/docs</a></li>
    </ul>
    </body></html>
    """


@app.get("/health")
async def health():
    return {"status": "ok", "busy": core.is_system_busy, "tasks": len(core.tasks_db)}


@app.post("/api/profile")
async def api_profile(data: ProfilePayload):
    payload = model_to_dict(data)
    paper_type = normalize_paper_type(payload["type"])
    profile = build_profile_from_request(payload, paper_type)
    return JSONResponse(profile)


@app.post("/api/generate", response_model=TaskStartResponse)
async def api_generate(data: GeneratePayload):
    paper_type = normalize_paper_type(data.type)
    profile = model_to_dict(data.profile)
    task_id = start_single_task("generate", _run_generate_task, profile, paper_type)
    return TaskStartResponse(status=core.tasks_db[task_id]["status"], task_id=task_id)


def _run_generate_task(task_id: str, profile: dict, paper_type: str):
    folder = _task_folder(task_id)
    def update(msg: str, prog: int):
        _update(task_id, msg, prog)
    try:
        _update(task_id, "任务已开始", 1)
        result = run_generate(profile, paper_type, folder, update)
        core.tasks_db[task_id].update({
            "status": "completed",
            "files": [result["txt_path"], result["docx_path"]],
            "msg": "完成",
            "progress": 100,
        })
    except Exception as exc:
        core.tasks_db[task_id].update({"status": "error", "msg": f"生成失败: {exc}"})
        raise


@app.post("/api/revise", response_model=TaskStartResponse)
async def api_revise(file: UploadFile = File(...), type: str = Form("管理")):
    upload_name = file.filename or "input.docx"
    upload_bytes = await file.read()
    task_id = start_single_task("revise", _run_revise_task, upload_name, upload_bytes, type)
    return TaskStartResponse(status=core.tasks_db[task_id]["status"], task_id=task_id)


def _run_revise_task(task_id: str, upload_name: str, upload_bytes: bytes, paper_type: str):
    folder = _task_folder(task_id)
    upload_path = folder / upload_name
    upload_path.write_bytes(upload_bytes)

    def update(msg: str, prog: int):
        _update(task_id, msg, prog)
    result = run_revise(str(upload_path), paper_type, folder, update)
    core.tasks_db[task_id].update({
        "status": "completed",
        "files": [result["report_path"], result["txt_path"], result["docx_path"]],
        "msg": "完成",
        "progress": 100,
    })


@app.get("/api/status/{task_id}", response_model=TaskSnapshot)
async def api_status(task_id: str):
    return core.tasks_db.get(task_id, model_to_dict(TaskSnapshot()))


@app.get("/download/{task_id}/{filename}")
async def download(task_id: str, filename: str):
    path = OUTPUT_ROOT / task_id / filename
    return FileResponse(str(path))


def main():
    import uvicorn

    uvicorn.run(app, host=SERVICE_HOST, port=SERVICE_PORT)


if __name__ == "__main__":
    main()
