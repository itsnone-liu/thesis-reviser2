# -*- coding: utf-8 -*-
"""
web.py — 论文系统 Web 服务（整合版）
"""
import sys, os, json, re, time, uuid, threading, shutil
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from fastapi import FastAPI, UploadFile, File, Form, Request
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from pydantic import BaseModel

from core import tasks_db, is_system_busy, system_lock
from profile import generate_profile
from generator import generate as gen_txt
from reviser import revise


app = FastAPI(title="论文系统")


class ProfileRequest(BaseModel):
    title: str = ""
    major: str = ""
    company: str = ""
    industry: str = ""

class DesignProfileRequest(BaseModel):
    title: str = ""
    design_type: str = ""
    design_object: str = ""
    context: str = ""

class GenRequest(BaseModel):
    type: str = "管理"
    profile: dict = None


def _run_task(task_id, fn, *args, **kwargs):
    global is_system_busy
    try:
        fn(*args, **kwargs)
    except Exception as e:
        tasks_db[task_id].update({"status": "error", "msg": f"任务失败: {str(e)}"})
        import traceback; traceback.print_exc()
    finally:
        with system_lock: is_system_busy = False

def _check_busy():
    with system_lock:
        if is_system_busy: return True
        is_system_busy = True; return False


def _gen_pipeline(task_id, profile, paper_type):
    folder = f"output/{task_id}"; os.makedirs(folder, exist_ok=True)
    def update(msg, prog): tasks_db[task_id].update({"msg": msg, "progress": prog})
    try:
        update("正在生成论文文本...", 10)
        txt = gen_txt(profile, paper_type, update)
        txt_path = os.path.join(folder, "00_完整论文.txt")
        with open(txt_path, 'w', encoding='utf-8') as f: f.write(txt)
        update("正在渲染DOCX...", 70)
        from renderer import render as render_it
        docx_path = os.path.join(folder, "论文_含图表.docx")
        render_it(txt_path, docx_path, paper_type, None, None, update)
        update("完成！", 100)
        tasks_db[task_id].update({"status": "completed", "files": [txt_path, docx_path]})
    except Exception as e:
        tasks_db[task_id].update({"status": "error", "msg": f"生成失败: {str(e)}"})
        import traceback; traceback.print_exc()

def _revise_pipeline(task_id, input_docx, paper_type):
    folder = f"output/{task_id}"; os.makedirs(folder, exist_ok=True)
    def update(msg, prog): tasks_db[task_id].update({"msg": msg, "progress": prog})
    try:
        from core import extract_docx_text, extract_cover_info, save_diagnosis_report, finalize_docx
        from reviser import _analyze_original_paper, _diagnose_and_reconstruct, _generate_revised
        from renderer import render as render_it
        update("正在分析原论文...", 10)
        original_text = extract_docx_text(input_docx)
        cover_info = extract_cover_info(input_docx)
        if not original_text: raise ValueError("无法读取DOCX")
        update("正在分析画像...", 25)
        analysis = _analyze_original_paper(original_text, paper_type)
        update("正在诊断重构...", 40)
        revision_plan = _diagnose_and_reconstruct(analysis, paper_type)
        report_path = os.path.join(folder, "诊断报告.txt")
        save_diagnosis_report(revision_plan, report_path)
        update("正在生成修改版...", 60)
        txt = _generate_revised(analysis, revision_plan, paper_type, update)
        txt_path = os.path.join(folder, "00_完整论文.txt")
        with open(txt_path, 'w', encoding='utf-8') as f: f.write(txt)
        update("正在渲染DOCX...", 80)
        docx_path = os.path.join(folder, "论文_修改版.docx")
        render_it(txt_path, docx_path, paper_type, cover_info, None, update)
        update("完成！", 100)
        tasks_db[task_id].update({"status": "completed", "files": [report_path, txt_path, docx_path]})
    except Exception as e:
        tasks_db[task_id].update({"status": "error", "msg": f"修改失败: {str(e)}"})
        import traceback; traceback.print_exc()


# ==================== HTML页面 ====================
INDEX_HTML = """<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8"><title>论文系统</title>
<script src="https://cdn.tailwindcss.com"></script></head>
<body class="bg-gray-100 p-8">
<div class="max-w-4xl mx-auto space-y-8">
<h1 class="text-3xl font-bold text-center text-blue-700">论文生成系统</h1>
<div class="grid grid-cols-2 gap-8">
<a href="/manage" class="bg-white p-8 rounded-lg shadow border-t-4 border-blue-500 hover:shadow-lg text-center">
<div class="text-5xl mb-4">📊</div><h2 class="text-xl font-bold text-blue-700 mb-2">经管类论文</h2><p class="text-gray-600">引言→现状→问题→方案</p></a>
<a href="/design" class="bg-white p-8 rounded-lg shadow border-t-4 border-purple-500 hover:shadow-lg text-center">
<div class="text-5xl mb-4">🎨</div><h2 class="text-xl font-bold text-purple-700 mb-2">设计类论文</h2><p class="text-gray-600">绪论→理论→问题→设计→总结</p></a>
</div>
<hr class="my-8">
<div class="text-center"><a href="/revise" class="bg-green-600 text-white px-8 py-4 rounded-lg text-xl hover:bg-green-700">修改论文</a></div>
</div></body></html>"""

MANAGE_HTML = """<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8"><title>经管类论文</title>
<script src="https://cdn.tailwindcss.com"></script></head>
<body class="bg-gray-100 p-8">
<div class="max-w-4xl mx-auto">
<div class="flex items-center justify-between mb-6">
<a href="/" class="text-blue-600 hover:underline">← 返回</a><h1 class="text-3xl font-bold text-blue-700">经管类论文生成</h1><div></div>
</div>
<div class="bg-white p-6 rounded-lg shadow border-t-4 border-blue-500">
<div class="grid grid-cols-2 gap-4 mb-4">
<div><label class="block text-sm font-medium mb-1">论文题目</label><input id="t_title" class="w-full border p-2 rounded" value="供应链视角下企业营运资金管理优化研究"></div>
<div><label class="block text-sm font-medium mb-1">专业</label><input id="t_major" class="w-full border p-2 rounded" value="工商管理"></div>
<div><label class="block text-sm font-medium mb-1">研究对象</label><input id="t_comp" class="w-full border p-2 rounded" value="K公司"></div>
<div><label class="block text-sm font-medium mb-1">行业</label><input id="t_ind" class="w-full border p-2 rounded" value="制造业"></div>
</div>
<button onclick="genProfile()" class="bg-gray-500 text-white px-4 py-2 rounded mb-4">1. 生成画像</button>
<div id="profile_box" class="hidden mb-4">
<label class="block text-sm font-medium mb-1 text-orange-600">【可修改】画像</label>
<textarea id="t_profile" class="w-full border p-2 rounded h-32 text-sm"></textarea>
<button onclick="startGen()" class="bg-blue-600 text-white px-4 py-2 rounded mt-2">2. 开始生成</button>
</div>
<div id="progress" class="hidden"><div id="p_msg" class="text-sm mb-2"></div><div class="w-full bg-gray-200 rounded h-3"><div id="p_bar" class="bg-blue-600 h-3 rounded" style="width:0%"></div></div></div>
<div id="queue" class="hidden mt-4 p-3 bg-red-50 text-red-700 font-bold rounded">排队...</div>
<div id="result" class="hidden mt-4 p-4 bg-green-50 rounded text-center">
<div class="text-green-700 font-bold mb-2">完成！</div>
<a id="a_txt" href="#" class="bg-blue-600 text-white px-4 py-2 rounded mr-2">下载TXT</a>
<a id="a_docx" href="#" class="bg-green-600 text-white px-4 py-2 rounded">下载DOCX</a>
</div>
</div></div>
<script>
let tid="",iv="";
async function req(u,b){const r=await fetch(u,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(b)});return r.json();}
async function poll(){const d=await(await fetch("/api/status/"+tid)).json();
document.getElementById("p_msg").innerText=d.msg;document.getElementById("p_bar").style.width=d.progress+"%";
if(d.status==="completed"){clearInterval(iv);document.getElementById("progress").classList.add("hidden");document.getElementById("result").classList.remove("hidden");
document.getElementById("a_txt").href="/download/"+tid+"/00_完整论文.txt";document.getElementById("a_docx").href="/download/"+tid+"/论文_含图表.docx";}
else if(d.status==="error"){clearInterval(iv);document.getElementById("p_msg").innerText="失败: "+d.msg;}}
async function genProfile(){const p=await req("/api/gen_profile",{title:t_title.value,major:t_major.value,company:t_comp.value,industry:t_ind.value});
document.getElementById("t_profile").value=JSON.stringify(p,null,2);document.getElementById("profile_box").classList.remove("hidden");}
async function startGen(){const profile=JSON.parse(document.getElementById("t_profile").value);
const r=await req("/api/start_gen",{type:"管理",profile:profile});
if(r.status==="queue"){document.getElementById("queue").classList.remove("hidden");iv=setInterval(startGen,3000);return;}
tid=r.task_id;document.getElementById("queue").classList.add("hidden");document.getElementById("progress").classList.remove("hidden");iv=setInterval(poll,1000);}
</script></body></html>"""

DESIGN_HTML = """<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8"><title>设计类论文</title>
<script src="https://cdn.tailwindcss.com"></script></head>
<body class="bg-gray-100 p-8">
<div class="max-w-4xl mx-auto">
<div class="flex items-center justify-between mb-6">
<a href="/" class="text-purple-600 hover:underline">← 返回</a><h1 class="text-3xl font-bold text-purple-700">设计类论文生成</h1><div></div>
</div>
<div class="bg-white p-6 rounded-lg shadow border-t-4 border-purple-500">
<div class="grid grid-cols-2 gap-4 mb-4">
<div><label class="block text-sm font-medium mb-1">论文题目</label><input id="d_title" class="w-full border p-2 rounded" value="新中式风格女装系列设计研究"></div>
<div><label class="block text-sm font-medium mb-1">设计专业</label><input id="d_type" class="w-full border p-2 rounded" value="服装设计"></div>
<div><label class="block text-sm font-medium mb-1">设计对象</label><input id="d_obj" class="w-full border p-2 rounded" value="新中式女装"></div>
<div><label class="block text-sm font-medium mb-1">使用场景</label><input id="d_ctx" class="w-full border p-2 rounded" value="日常穿着"></div>
</div>
<button onclick="genProfile()" class="bg-gray-500 text-white px-4 py-2 rounded mb-4">1. 生成画像</button>
<div id="profile_box" class="hidden mb-4">
<label class="block text-sm font-medium mb-1 text-orange-600">【可修改】画像</label>
<textarea id="t_profile" class="w-full border p-2 rounded h-32 text-sm"></textarea>
<button onclick="startGen()" class="bg-purple-600 text-white px-4 py-2 rounded mt-2">2. 开始生成</button>
</div>
<div id="progress" class="hidden"><div id="p_msg" class="text-sm mb-2"></div><div class="w-full bg-gray-200 rounded h-3"><div id="p_bar" class="bg-purple-600 h-3 rounded" style="width:0%"></div></div></div>
<div id="queue" class="hidden mt-4 p-3 bg-red-50 text-red-700 font-bold rounded">排队...</div>
<div id="result" class="hidden mt-4 p-4 bg-green-50 rounded text-center">
<div class="text-green-700 font-bold mb-2">完成！</div>
<a id="a_txt" href="#" class="bg-purple-600 text-white px-4 py-2 rounded mr-2">下载TXT</a>
<a id="a_docx" href="#" class="bg-green-600 text-white px-4 py-2 rounded">下载DOCX</a>
</div>
</div></div>
<script>
let tid="",iv="";
async function req(u,b){const r=await fetch(u,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(b)});return r.json();}
async function poll(){const d=await(await fetch("/api/status/"+tid)).json();
document.getElementById("p_msg").innerText=d.msg;document.getElementById("p_bar").style.width=d.progress+"%";
if(d.status==="completed"){clearInterval(iv);document.getElementById("progress").classList.add("hidden");document.getElementById("result").classList.remove("hidden");
document.getElementById("a_txt").href="/download/"+tid+"/00_完整论文.txt";document.getElementById("a_docx").href="/download/"+tid+"/论文_含图表.docx";}
else if(d.status==="error"){clearInterval(iv);document.getElementById("p_msg").innerText="失败: "+d.msg;}}
async function genProfile(){const p=await req("/api/gen_design_profile",{title:d_title.value,design_type:d_type.value,design_object:d_obj.value,context:d_ctx.value});
document.getElementById("t_profile").value=JSON.stringify(p,null,2);document.getElementById("profile_box").classList.remove("hidden");}
async function startGen(){const profile=JSON.parse(document.getElementById("t_profile").value);
const r=await req("/api/start_gen",{type:"设计",profile:profile});
if(r.status==="queue"){document.getElementById("queue").classList.remove("hidden");iv=setInterval(startGen,3000);return;}
tid=r.task_id;document.getElementById("queue").classList.add("hidden");document.getElementById("progress").classList.remove("hidden");iv=setInterval(poll,1000);}
</script></body></html>"""

REVISE_HTML = """<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8"><title>论文修改</title>
<script src="https://cdn.tailwindcss.com"></script></head>
<body class="bg-gray-100 p-8">
<div class="max-w-4xl mx-auto">
<div class="flex items-center justify-between mb-6">
<a href="/" class="text-blue-600 hover:underline">← 返回</a><h1 class="text-3xl font-bold text-green-700">论文修改</h1><div></div>
</div>
<div class="bg-white p-6 rounded-lg shadow border-t-4 border-green-500">
<label class="block text-sm font-medium mb-2">类型</label>
<select id="r_type" class="w-full border p-2 rounded mb-4"><option value="管理">经管类</option><option value="设计">设计类</option><option value="机械">机械类</option></select>
<label class="block text-sm font-medium mb-2">上传DOCX</label>
<input type="file" id="f_upload" accept=".docx" class="mb-4 block w-full text-sm">
<button onclick="startRevise()" class="bg-green-600 text-white px-4 py-2 rounded disabled:opacity-50" id="btn_revise" disabled>开始修改</button>
<div id="progress" class="hidden mt-4"><div id="p_msg" class="text-sm mb-2"></div><div class="w-full bg-gray-200 rounded h-3"><div id="p_bar" class="bg-green-600 h-3 rounded" style="width:0%"></div></div></div>
<div id="queue" class="hidden mt-4 p-3 bg-red-50 text-red-700 font-bold rounded">排队...</div>
<div id="result" class="hidden mt-4 p-4 bg-green-50 rounded text-center">
<div class="text-green-700 font-bold mb-2">完成！</div>
<a id="a_report" href="#" class="bg-orange-500 text-white px-4 py-2 rounded mr-2">诊断报告</a>
<a id="a_txt" href="#" class="bg-blue-600 text-white px-4 py-2 rounded mr-2">下载TXT</a>
<a id="a_docx" href="#" class="bg-green-600 text-white px-4 py-2 rounded">下载DOCX</a>
</div>
</div></div>
<script>
let tid="",iv="";
document.getElementById("f_upload").addEventListener("change",function(){document.getElementById("btn_revise").disabled=!this.files.length;});
async function poll(){const d=await(await fetch("/api/status/"+tid)).json();
document.getElementById("p_msg").innerText=d.msg;document.getElementById("p_bar").style.width=d.progress+"%";
if(d.status==="completed"){clearInterval(iv);document.getElementById("progress").classList.add("hidden");document.getElementById("result").classList.remove("hidden");
document.getElementById("a_report").href="/download/"+tid+"/诊断报告.txt";document.getElementById("a_txt").href="/download/"+tid+"/00_完整论文.txt";document.getElementById("a_docx").href="/download/"+tid+"/论文_修改版.docx";}
else if(d.status==="error"){clearInterval(iv);document.getElementById("p_msg").innerText="失败: "+d.msg;}}
async function startRevise(){const fd=new FormData();fd.append("file",document.getElementById("f_upload").files[0]);fd.append("type",document.getElementById("r_type").value);
const r=await fetch("/api/start_revise",{method:"POST",body:fd}).then(r=>r.json());
if(r.status==="queue"){document.getElementById("queue").classList.remove("hidden");iv=setInterval(startRevise,3000);return;}
tid=r.task_id;document.getElementById("queue").classList.add("hidden");document.getElementById("progress").classList.remove("hidden");iv=setInterval(poll,1000);}
</script></body></html>"""


@app.get("/", response_class=HTMLResponse)
async def index(): return INDEX_HTML

@app.get("/manage", response_class=HTMLResponse)
async def manage_page(): return MANAGE_HTML

@app.get("/design", response_class=HTMLResponse)
async def design_page(): return DESIGN_HTML

@app.get("/revise", response_class=HTMLResponse)
async def revise_page(): return REVISE_HTML


@app.post("/api/gen_profile")
async def api_gen_profile(data: ProfileRequest):
    p = generate_profile("管理", data.title, data.company, data.major)
    if not p.get("core_problems"): p["core_problems"] = ["补充问题1", "补充问题2"]
    return p

@app.post("/api/gen_design_profile")
async def api_gen_design_profile(data: DesignProfileRequest):
    p = generate_profile("设计", data.title, data.design_object, data.design_type, data.context)
    if not p.get("core_problems"): p["core_problems"] = ["补充问题1", "补充问题2", "补充问题3"]
    return p

@app.post("/api/gen_mechanical_profile")
async def api_gen_mechanical_profile(data: DesignProfileRequest):
    p = generate_profile("机械", data.title, data.design_object, data.design_type, data.context)
    if not p.get("core_problems"): p["core_problems"] = ["补充问题1", "补充问题2", "补充问题3"]
    return p

@app.post("/api/start_gen")
async def api_start_gen(data: GenRequest):
    if _check_busy(): return {"status": "queue"}
    tid = str(uuid.uuid4())
    tasks_db[tid] = {"status": "running", "progress": 0, "msg": "初始化", "files": []}
    threading.Thread(target=_run_task, args=(tid, _gen_pipeline, tid, data.profile, data.type)).start()
    return {"status": "ok", "task_id": tid}

@app.post("/api/start_revise")
async def api_start_revise(file: UploadFile = File(...), type: str = Form("管理")):
    if _check_busy(): return {"status": "queue"}
    tid = str(uuid.uuid4())
    tasks_db[tid] = {"status": "running", "progress": 0, "msg": "接收文件", "files": []}
    temp_path = f"temp_{tid}.docx"
    with open(temp_path, "wb") as f: f.write(await file.read())
    threading.Thread(target=_run_task, args=(tid, _revise_pipeline, tid, temp_path, type)).start()
    return {"status": "ok", "task_id": tid}

@app.get("/api/status/{task_id}")
async def api_status(task_id: str):
    return tasks_db.get(task_id, {"status": "not_found", "msg": "未知", "progress": 0})

@app.get("/download/{task_id}/{filename}")
async def api_download(task_id: str, filename: str):
    path = os.path.join("output", task_id, filename)
    if os.path.exists(path): return FileResponse(path, filename=filename)
    return {"error": "not found"}


if __name__ == "__main__":
    import uvicorn
    print("\n  http://localhost:8000")
    print("  manage -> /manage   design -> /design   revise -> /revise\n")
    uvicorn.run(app, host="0.0.0.0", port=8000)
