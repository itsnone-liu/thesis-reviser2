# -*- coding: utf-8 -*-
"""论文系统 V2：账号、持久化任务、法学出题、自动审计。
部署入口：python webapp.py；生产由 systemd/uvicorn 启动。
"""
from __future__ import annotations
import os, sys, json, time, uuid, queue, sqlite3, hashlib, secrets, threading, subprocess
from pathlib import Path
from datetime import datetime, timezone
from typing import Optional

ROOT = Path(__file__).resolve().parent
os.chdir(ROOT)
sys.path.insert(0, str(ROOT))
DB = ROOT / "data" / "webapp.db"
OUT = ROOT / "data" / "outputs"
DB.parent.mkdir(exist_ok=True); OUT.mkdir(parents=True, exist_ok=True)

# 兼容 systemd/非交互 shell：导入 core 前显式加载仓库 .env；绝不打印 key
_env = ROOT / ".env"
if _env.exists():
    for _line in _env.read_text(encoding="utf-8").splitlines():
        _line = _line.strip()
        if _line.startswith("export "): _line = _line[7:].strip()
        if _line and not _line.startswith("#") and "=" in _line:
            _k, _v = _line.split("=", 1)
            os.environ.setdefault(_k.strip(), _v.strip().strip('\\\"').strip("'"))

from fastapi import FastAPI, HTTPException, Request, UploadFile, File, Form
from fastapi.responses import HTMLResponse, FileResponse
from pydantic import BaseModel
from generator import generate
from renderer import render
from audit_final import audit_docx_only
import caselib
from law_pilot import polish_title, polish_title_t2, law_audit
from profile import generate_profile
from core import extract_drawings_from_text

app = FastAPI(title="个人论文生成系统 V2")
COOKIE = "thesis_sid"
DAILY_LIMIT = int(os.getenv("DAILY_LIMIT", "3"))
TASK_Q: queue.Queue[str] = queue.Queue()
STOP = object()


def now(): return datetime.now(timezone.utc).isoformat(timespec="seconds")

def db():
    c = sqlite3.connect(DB, timeout=30)
    c.row_factory = sqlite3.Row
    c.execute("PRAGMA journal_mode=WAL")
    return c

def init_db():
    with db() as c:
        c.executescript("""
        CREATE TABLE IF NOT EXISTS users(
          id INTEGER PRIMARY KEY, username TEXT UNIQUE NOT NULL, password_hash TEXT NOT NULL,
          role TEXT NOT NULL DEFAULT 'user', status TEXT NOT NULL DEFAULT 'active',
          created_at TEXT NOT NULL, last_login_at TEXT);
        CREATE TABLE IF NOT EXISTS sessions(
          token TEXT PRIMARY KEY, user_id INTEGER NOT NULL, expires_at REAL NOT NULL);
        CREATE TABLE IF NOT EXISTS tasks(
          id TEXT PRIMARY KEY, user_id INTEGER NOT NULL, kind TEXT NOT NULL,
          paper_type TEXT NOT NULL, title TEXT, profile_json TEXT, cover_json TEXT,
          status TEXT NOT NULL, progress INTEGER NOT NULL DEFAULT 0, message TEXT,
          files_json TEXT, audit_json TEXT, created_at TEXT NOT NULL,
          started_at TEXT, finished_at TEXT, error TEXT);
        CREATE TABLE IF NOT EXISTS topics(
          id TEXT PRIMARY KEY, user_id INTEGER NOT NULL, payload_json TEXT NOT NULL,
          created_at REAL NOT NULL, expires_at REAL NOT NULL, confirmed_task_id TEXT);
        CREATE TABLE IF NOT EXISTS usage(
          id INTEGER PRIMARY KEY, user_id INTEGER, task_id TEXT, event TEXT, created_at TEXT);
        CREATE INDEX IF NOT EXISTS idx_tasks_user ON tasks(user_id, created_at);
        """)
        if not c.execute("SELECT 1 FROM users WHERE username='admin'").fetchone():
            c.execute("INSERT INTO users(username,password_hash,role,created_at) VALUES(?,?,?,?)",
                      ("admin", hash_password("123456"), "admin", now()))
            print("[WARN] 已创建默认管理员 admin/123456，请上线后立即修改密码")

def hash_password(p: str) -> str:
    salt = secrets.token_bytes(16)
    return "pbkdf2$" + salt.hex() + "$" + hashlib.pbkdf2_hmac("sha256", p.encode(), salt, 180000).hex()

def check_password(p: str, encoded: str) -> bool:
    try:
        _, sh, hh = encoded.split("$", 2)
        got = hashlib.pbkdf2_hmac("sha256", p.encode(), bytes.fromhex(sh), 180000).hex()
        return secrets.compare_digest(got, hh)
    except Exception: return False

def user_from_request(request: Request, admin=False):
    sid = request.cookies.get(COOKIE)
    if not sid: raise HTTPException(401, "请先登录")
    with db() as c:
        r = c.execute("SELECT u.* FROM sessions s JOIN users u ON u.id=s.user_id WHERE s.token=? AND s.expires_at>?", (sid, time.time())).fetchone()
    if not r or r["status"] != "active": raise HTTPException(401, "登录已失效")
    if admin and r["role"] != "admin": raise HTTPException(403, "需要管理员权限")
    return dict(r)

def task_row(r):
    if not r: return None
    d = dict(r)
    for k in ("profile_json", "cover_json", "files_json", "audit_json"):
        if d.get(k):
            try: d[k[:-5] if k.endswith("_json") else k] = json.loads(d[k])
            except Exception: pass
        d.pop(k, None)
    return d

def update_task(tid, **fields):
    allowed = {"status","progress","message","files_json","audit_json","started_at","finished_at","error","title"}
    fields = {k:v for k,v in fields.items() if k in allowed}
    if not fields: return
    with db() as c:
        c.execute("UPDATE tasks SET " + ",".join(f"{k}=?" for k in fields) + " WHERE id=?", (*fields.values(), tid))

def enqueue_existing():
    with db() as c:
        c.execute("UPDATE tasks SET status='error', error='服务重启中断' WHERE status IN ('running','waiting_quota')")
        ids = [r["id"] for r in c.execute("SELECT id FROM tasks WHERE status='queue' ORDER BY created_at")]
    for tid in ids: TASK_Q.put(tid)

def is_quota_error(e):
    s = str(e).lower()
    return "1308" in s or "5小时" in s or "使用上限" in s

def _repair_civil_figure_declarations(txt: str) -> str:
    """补齐土木正文引用但生成器漏报的图表声明；只补确定性标签，不改正文事实。

    2026-09-09 验收教训：LLM 常把"表2-1 标准层主要功能房间一览表"写成孤立
    注行却不产出 <table> 标签；旧逻辑在首个引用前另插通用占位表"表N-M 土木
    工程计算表"，孤立注行原样保留 → 同号表注×2 → 审计"表注编号重复"❌。
    现在孤立注行就地转成以【注行真实标题】命名的标签；与已声明标签同号的
    孤立注行(重复注)删除；纯正文引用缺声明的仍按原方式补通用标签。
    """
    import re
    if "土木" not in txt[:3000] and "<drawing" not in txt:
        return txt
    # 清理旧版本追加的自描述补丁，避免历史正文在重试时累积双重编号。
    txt = re.sub(r'<drawing[^>]*id="repair-[^"]+"[^>]*/>', "", txt)
    txt = re.sub(r'<table[^>]*id="repair-[^"]+"[^>]*/>', "", txt)
    declared = set(re.findall(r'<drawing[^>]*title=["\']图(\d+-\d+)', txt))
    declared_tables = set(re.findall(r'<table[^>]*title=["\']表(\d+-\d+)', txt))

    # [FIGURES]自报清单在渲染时整块剥离(core.py 0908修复)，绝不能把标签转进去；
    # 但它是"真实标题"的矿藏——插入补齐标签时优先用清单里的原名，不再用通用占位名。
    manifest_titles = {}   # "图1-1" -> "建筑效果示意图"
    for blk in re.findall(r'\[FIGURES\]\s*(.*?)\[/FIGURES\]', txt, re.S):
        for ln in blk.split("\n"):
            m = re.match(r"^\s*([图表])\s*(\d{1,2}-\d{1,3})\s+(\S.{1,38}\S)\s*$", ln.strip())
            if m:
                manifest_titles[f"{m.group(1)}{m.group(2)}"] = m.group(3)

    # --- 正文孤立注行治理(仅正文；[FIGURES]清单块内的不算) ---
    cap_re = re.compile(r'^([图表])\s*(\d{1,2}-\d{1,3})\s+(\S.{1,38}\S)\s*$')
    ref_sent = re.compile(r'显示|如下|所示|可知|可以看出|列出|给出|反映了|表明|中可')
    masked = re.sub(r'\[FIGURES\]\s*.*?\[/FIGURES\]', lambda m: "\n" * m.group(0).count("\n"), txt, flags=re.S)
    lines = masked.split("\n")   # 与txt行号对齐(用空行占位保持行号)
    converted = []
    drop_idx = set()
    for i, ln in enumerate(lines):
        s = ln.strip()
        if not s or s.startswith("<") or s.endswith(("。", "，", "；", "：", "！", "？")):
            continue
        m = cap_re.match(s)
        if not m or ref_sent.search(s):
            continue
        kind, num, title_body = m.group(1), m.group(2), m.group(3)
        declared_set = declared if kind == "图" else declared_tables
        if num in declared_set:
            drop_idx.add(i)          # 重复注行：已声明同号标签，标签渲染自带注
        else:
            converted.append((i, kind, num, f"{kind}{num} {title_body}"))
            declared_set.add(num)
    next_id = 100
    for i, kind, num, real_title in converted:
        if kind == "表":
            lines[i] = (f'<table id="repair-{next_id}" title="{real_title}" '
                        f'header="项目,数值" rows="正文已引用项目,详见正文计算"/>')
        else:
            lines[i] = (f'<drawing id="repair-{next_id}" type="civil" title="{real_title}" '
                        f'description="根据正文孤立注行补齐图纸声明；具体工程参数沿用正文已确立值。"/>')
        next_id += 1
    if drop_idx or converted:
        txt = "\n".join(lines)

    # 以正文引用为准，自动发现本批次缺失的任意章节式图/表声明。
    # 插入位置必须避开[FIGURES]清单块(渲染时整块剥离，标签插进去=白插)。
    _fig_spans = [m.span() for m in re.finditer(r'\[FIGURES\]\s*.*?\[/FIGURES\]', txt, re.S)]
    def _first_ref_pos(kind_char: str, num: str):
        for m in re.finditer(rf'{kind_char}\s*{re.escape(num)}', txt):
            if any(a <= m.start() < b for a, b in _fig_spans):
                continue
            return m.start()
        return None
    ref_figs = set(re.findall(r'(?<![简插附纸样意标流路线效])图\s*(\d+-\d+)', txt))
    ref_tables = set(re.findall(r'表\s*(\d+-\d+)', txt))
    def _title_for(kind_char: str, num: str) -> str:
        real = manifest_titles.get(f"{kind_char}{num}")
        return f"{kind_char}{num} {real}" if real else f"{kind_char}{num} 土木工程{'示意图' if kind_char == '图' else '计算表'}"
    missing_figs = [(num, _title_for("图", num)) for num in sorted(ref_figs) if num not in declared]
    missing_tables = [(num, _title_for("表", num)) for num in sorted(ref_tables) if num not in declared_tables]
    if not missing_figs and not missing_tables:
        return txt
    # 标签必须插在对应正文引用之前，不能统一追加到文末；否则 registry 会把它归到末章。
    inserts = []
    for num, title in missing_figs:
        tag = f'<drawing id="repair-{next_id}" type="civil" title="{title}" description="根据正文已引用的{title}补齐图纸声明；具体工程参数沿用正文已确立值。"/>\n'
        pos = _first_ref_pos("图", num)
        if pos is not None: inserts.append((pos, tag))
        next_id += 1
    for num, title in missing_tables:
        tag = f'<table id="repair-{next_id}" title="{title}" header="项目,数值" rows="正文已引用项目,详见正文计算"/>\n'
        pos = _first_ref_pos("表", num)
        if pos is not None: inserts.append((pos, tag))
        next_id += 1
    for pos, tag in sorted(inserts, reverse=True):
        txt = txt[:pos] + tag + txt[pos:]
    return txt

_CIVIL_CANON_REFS = [
    "中华人民共和国住房和城乡建设部. 混凝土结构设计规范: GB 50010—2010[S]. 北京: 中国建筑工业出版社, 2011.",
    "中华人民共和国住房和城乡建设部. 建筑抗震设计规范: GB 50011—2010[S]. 北京: 中国建筑工业出版社, 2010.",
    "中华人民共和国住房和城乡建设部. 建筑结构荷载规范: GB 50009—2012[S]. 北京: 中国建筑工业出版社, 2012.",
    "中华人民共和国住房和城乡建设部. 建筑地基基础设计规范: GB 50007—2011[S]. 北京: 中国建筑工业出版社, 2012.",
    "中华人民共和国住房和城乡建设部. 建筑设计防火规范: GB 50016—2014(2018年版)[S]. 北京: 中国建筑工业出版社, 2015.",
    "中华人民共和国住房和城乡建设部. 混凝土结构工程施工质量验收规范: GB 50204—2015[S]. 北京: 中国建筑工业出版社, 2015.",
    "中华人民共和国住房和城乡建设部. 建筑工程施工质量验收统一标准: GB 50300—2013[S]. 北京: 中国建筑工业出版社, 2014.",
    "中华人民共和国住房和城乡建设部. 建筑施工安全检查标准: JGJ 59—2011[S]. 北京: 中国建筑工业出版社, 2011.",
    "中华人民共和国住房和城乡建设部. 建筑施工组织设计规范: GB/T 50502—2009[S]. 北京: 中国建筑工业出版社, 2009.",
    "中华人民共和国住房和城乡建设部. 高层建筑混凝土结构技术规程: JGJ 3—2010[S]. 北京: 中国建筑工业出版社, 2011.",
    "中华人民共和国住房和城乡建设部. 建筑制图标准: GB/T 50105—2010[S]. 北京: 中国建筑工业出版社, 2010.",
    "中华人民共和国住房和城乡建设部. 工程结构通用规范: GB 55001—2021[S]. 北京: 中国建筑工业出版社, 2021.",
]


def _ensure_civil_references(txt: str) -> str:
    """LLM偶尔漏写参考文献章(2026-09-09验收: refs=0)。确定性兜底：不足10条时
    补真实土木规范清单(皆为现行国家/行业标准，土木论文通用引用，不属编造)。"""
    import re
    m = re.search(r"参考文献\s*\n((?:\[?\d+\]?[^\n]+\n?)+)", txt)
    existing = []
    if m:
        existing = re.findall(r"^\s*\[?\s*(\d+)\s*\]?\.?\s*\S[^\n]*$", m.group(1), re.M)
    if len(existing) >= 10:
        return txt
    # 去掉空壳参考文献段(有标题没条目)，重建整段
    txt = re.sub(r"\n?参考文献\s*\n(?:\[?\d+\]?[^\n]*\n?)*$", "\n", txt)
    add = _CIVIL_CANON_REFS[: max(0, 12 - len(existing))]
    block = "\n参考文献\n" + "\n".join(f"[{i+1}] {r}" for i, r in enumerate(add)) + "\n"
    return txt.rstrip("\n") + "\n" + block


def _similar(a: str, b: str) -> bool:
    """标题近似判定：完全一致或一方为另一方前缀(≥4字)。"""
    if a == b:
        return True
    return len(a) >= 4 and len(b) >= 4 and (a.startswith(b) or b.startswith(a))


def _normalize_civil_numbering(txt: str) -> str:
    """图表编号治理(2026-09-09验收两起)：
    ① LLM偶发同号双标签(表3-2×2同标题) → 标题基本相同删后者；
    ② 章内跳号(图3-[1,3]) → 就地压实：第k个标签改为图C-k，标签标题与正文引用
       同步改写。只把编号往下压(k≤原号)，按目标号升序应用避免连锁错写。"""
    import re
    for kind, tag_re in (("图", r'<drawing[^>]*title="图\s*(\d+)-(\d+)[^"]*"[^>]*/>'),
                         ("表", r'<table[^>]*title="表\s*(\d+)-(\d+)[^"]*"[^>]*/>')):
        tags = [(m.start(), m.end(), int(m.group(1)), int(m.group(2)), m.group(0)) for m in re.finditer(tag_re, txt)]
        if not tags:
            continue
        # ① 同号近似同标题去重(保留先出现的)
        seen = {}
        drop_spans = []
        for st, en, c, n, raw in tags:
            title = re.search(r'title="([^"]*)"', raw).group(1)
            norm = re.sub(r"\s+", "", title.split(" ", 1)[-1])
            key = (c, n)
            if key in seen and _similar(seen[key], norm):
                drop_spans.append((st, en))
            else:
                seen[key] = norm
        for st, en in sorted(drop_spans, reverse=True):
            txt = txt[:st] + txt[en:]
        tags = [(m.start(), m.end(), int(m.group(1)), int(m.group(2))) for m in re.finditer(tag_re, txt)]
        # ② 章内压实映射(仅变更项)
        by_ch = {}
        for st, en, c, n in tags:
            by_ch.setdefault(c, []).append(n)
        mapping = []   # (old_n, new_n, c) 需要改的
        for c, nums in by_ch.items():
            if sorted(nums) == list(range(1, len(nums) + 1)) and len(set(nums)) == len(nums):
                continue
            ordered = sorted(nums)
            for k, old in enumerate(ordered, 1):
                if k != old:
                    mapping.append((old, k, c))
        # 目标号升序应用；同章内目标号互不相同且≤未处理原号，无连锁错写
        for old, new, c in sorted(mapping, key=lambda x: (x[2], x[1])):
            pat = re.compile(rf"{kind}\s*{c}\s*-\s*{old}(?![\d])")
            txt = pat.sub(f"{kind}{c}-{new}", txt)
    return txt

def run_generation(tid, profile, ptype, cover):
    folder = OUT / str(profile.get("user_id", "0")) / tid; folder.mkdir(parents=True, exist_ok=True)
    def progress(msg, pct): update_task(tid, message=msg, progress=int(pct))
    update_task(tid, status="running", started_at=now(), message="开始生成", progress=1)
    try:
        txt = generate(profile, ptype, progress)
        if ptype == "土木":
            txt = _repair_civil_figure_declarations(txt)
            txt = _normalize_civil_numbering(txt)
            txt = _ensure_civil_references(txt)
        txtpath = folder / "00_完整论文.txt"; txtpath.write_text(txt, encoding="utf-8")
        docx = folder / "论文终稿.docx"
        # 将高峰内存的渲染/审计放入独立进程；正文生成进程退出后由OS回收其全部内存。
        spec = folder / "render_spec.json"
        spec.write_text(json.dumps({"txt":str(txtpath),"docx":str(docx),"ptype":ptype,"cover":cover if ptype == "法学" else None},ensure_ascii=False),encoding="utf-8")
        proc = subprocess.Popen([sys.executable, str(ROOT / "render_task.py"), str(spec)], cwd=str(ROOT), stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
        audit = None
        for line in proc.stdout:
            try:
                event = json.loads(line)
                if event.get("message"): progress(event["message"], event.get("progress", 90))
                if event.get("done"): audit = event["audit"]
            except Exception: pass
        rc = proc.wait()
        if rc != 0 or audit is None: raise RuntimeError(f"独立渲染进程失败(rc={rc})")
        if ptype == "法学":
            cases = profile.get("cluster") or ([profile.get("case")] if profile.get("case") else [])
            audit["law"] = law_audit(txt, cases)
        files = [str(txtpath.relative_to(folder)), str(docx.relative_to(folder))]
        update_task(tid, status="done", progress=100, message="生成完成，审计已完成", files_json=json.dumps(files, ensure_ascii=False), audit_json=json.dumps(audit, ensure_ascii=False), finished_at=now())
    except Exception as e:
        if is_quota_error(e):
            update_task(tid, status="waiting_quota", message="百炼配额达到5小时上限，等待重置后重试", error=str(e))
            time.sleep(900)
            TASK_Q.put(tid)
        else:
            update_task(tid, status="error", message="生成失败", error=str(e), finished_at=now())

def worker():
    while True:
        tid = TASK_Q.get()
        if tid is STOP: return
        try:
            with db() as c: r = c.execute("SELECT * FROM tasks WHERE id=?", (tid,)).fetchone()
            if not r or r["status"] not in ("queue", "waiting_quota"): continue
            profile = json.loads(r["profile_json"]); profile["user_id"] = r["user_id"]
            run_generation(tid, profile, r["paper_type"], json.loads(r["cover_json"] or "{}"))
        finally: TASK_Q.task_done()

def start_task(user_id, ptype, profile, cover=None, title=""):
    with db() as c:
        used = c.execute("SELECT COUNT(*) FROM tasks WHERE user_id=? AND created_at LIKE ? AND status NOT IN ('error','interrupted')", (user_id, now()[:10] + "%")).fetchone()[0]
        if used >= DAILY_LIMIT: raise HTTPException(429, f"今日生成次数已达上限({DAILY_LIMIT})")
        active = c.execute("SELECT 1 FROM tasks WHERE user_id=? AND status IN ('queue','running','waiting_quota')", (user_id,)).fetchone()
        if active: raise HTTPException(409, "你已有任务在队列中")
        tid = uuid.uuid4().hex
        c.execute("INSERT INTO tasks(id,user_id,kind,paper_type,title,profile_json,cover_json,status,message,created_at) VALUES(?,?,?,?,?,?,?,?,?,?)",
                  (tid,user_id,"generate",ptype,title,json.dumps(profile,ensure_ascii=False),json.dumps(cover or {},ensure_ascii=False),"queue","排队中",now()))
        c.execute("INSERT INTO usage(user_id,task_id,event,created_at) VALUES(?,?,?,?)", (user_id,tid,"generate",now()))
    TASK_Q.put(tid); return tid

class Credentials(BaseModel): username: str; password: str
class GenerateRequest(BaseModel):
    type: str = "管理"; profile: dict = {}; cover: dict = {}
class ProfileRequest(BaseModel):
    type: str = "管理"; title: str = ""; target: str = ""; major: str = ""; context: str = ""
class TopicRequest(BaseModel): mode: str = "mix"; domain: str = ""
class ConfirmRequest(BaseModel):
    topic_id: str
    index: int = 0
    cover: dict = {}

@app.on_event("startup")
def startup():
    init_db(); enqueue_existing(); threading.Thread(target=worker, daemon=True, name="thesis-worker").start()

@app.post("/api/register")
def register(x: Credentials):
    if not (3 <= len(x.username) <= 32) or len(x.password) < 6: raise HTTPException(400,"用户名3-32字符，密码至少6位")
    try:
        with db() as c: c.execute("INSERT INTO users(username,password_hash,created_at) VALUES(?,?,?)", (x.username,hash_password(x.password),now()))
    except sqlite3.IntegrityError: raise HTTPException(409,"用户名已存在")
    return {"ok": True}

@app.post("/api/login")
def login(x: Credentials):
    with db() as c: r=c.execute("SELECT * FROM users WHERE username=?",(x.username,)).fetchone()
    if not r or r["status"] != "active" or not check_password(x.password,r["password_hash"]): raise HTTPException(401,"用户名或密码错误")
    token=secrets.token_urlsafe(32)
    with db() as c:
        c.execute("INSERT INTO sessions VALUES(?,?,?)",(token,r["id"],time.time()+7*86400)); c.execute("UPDATE users SET last_login_at=? WHERE id=?",(now(),r["id"]))
    from fastapi.responses import JSONResponse
    out=JSONResponse({"ok":True,"username":r["username"],"role":r["role"]}); out.set_cookie(COOKIE,token,max_age=7*86400,httponly=True,samesite="lax"); return out

@app.post("/api/logout")
def logout(request: Request):
    sid=request.cookies.get(COOKIE)
    if sid:
        with db() as c: c.execute("DELETE FROM sessions WHERE token=?",(sid,))
    from fastapi.responses import JSONResponse
    out=JSONResponse({"ok":True}); out.delete_cookie(COOKIE); return out

@app.get("/health")
def health():
    return {"ok": True, "service": "thesis-web-v2"}

@app.get("/api/me")
def me(request: Request):
    u=user_from_request(request); return {"id":u["id"],"username":u["username"],"role":u["role"]}

@app.get("/api/law/domains")
def law_domains(request: Request):
    user_from_request(request)
    counts={}
    for x in caselib.load_library():
        if x.get("verified"):
            d=x.get("领域") or x.get("domain") or "其他"; counts[d]=counts.get(d,0)+1
    return {"domains":[{"name":k,"count":v} for k,v in sorted(counts.items(),key=lambda z:-z[1])]}

def topic_candidate(kind, domain, lib, offset=0):
    verified=[x for x in lib if x.get("verified") and (not domain or domain in str(x.get("领域",x.get("domain",""))))]
    if kind == "T1":
        if not verified: verified=[x for x in lib if x.get("verified")]
        case=verified[offset % len(verified)]; title=polish_title(case)
        return {"law_type":"T1","title":title,"domain":case.get("领域",""),"cases":[case],"case_names":[case.get("名称","")]}
    domains=[domain] if domain else list({str(x.get("领域","")) for x in lib if x.get("verified") and x.get("领域")})
    d=domains[offset % len(domains)] if domains else ""
    cluster=caselib.cluster(d,n=4,rotate=offset) if d else []
    if len(cluster)<3: raise HTTPException(400,"所选领域没有足够的 verified 案例")
    return {"law_type":"T2","title":polish_title_t2(cluster,d),"domain":d,"cases":cluster,"case_names":[x.get("名称","") for x in cluster]}

@app.post("/api/law/topics")
def law_topics(x: TopicRequest, request: Request):
    u=user_from_request(request); lib=caselib.load_library(); kinds=["T1","T1","T2"] if x.mode=="mix" else [x.mode]*3
    arr=[topic_candidate(k,x.domain,lib,i) for i,k in enumerate(kinds)]
    tid=uuid.uuid4().hex
    with db() as c: c.execute("INSERT INTO topics VALUES(?,?,?,?,?,NULL)",(tid,u["id"],json.dumps(arr,ensure_ascii=False),time.time(),time.time()+900))
    return {"topic_id":tid,"topics":[{k:v for k,v in a.items() if k not in ("cases",)} for a in arr]}

@app.post("/api/law/topics/refresh")
def law_topics_refresh(x: TopicRequest, request: Request):
    # 候选批次有15分钟有效期；刷新直接生成新的批次并由前端替换旧批次。
    return law_topics(x, request)

@app.post("/api/law/confirm")
def law_confirm(x: ConfirmRequest, request: Request):
    u=user_from_request(request)
    with db() as c: r=c.execute("SELECT * FROM topics WHERE id=? AND user_id=? AND expires_at>?",(x.topic_id,u["id"],time.time())).fetchone()
    if not r: raise HTTPException(404,"题目已过期，请重新出题")
    arr=json.loads(r["payload_json"])
    if not (0 <= x.index < len(arr)): raise HTTPException(400,"候选题目序号无效")
    chosen=arr[x.index]
    # topic_id 对应一批候选；确认时锁定选中的案例与题目
    profile={"title":chosen["title"],"law_type":chosen["law_type"],"domain":chosen.get("domain","")}
    if chosen["law_type"]=="T1": profile["case"]=chosen["cases"][0]
    else: profile["cluster"]=chosen["cases"]
    cover={"title":chosen["title"],"name":x.cover.get("name",u["username"]),"student_id":x.cover.get("student_id",""),"major":"法学","level":x.cover.get("level","本科"),"advisor":x.cover.get("advisor",""),"year":x.cover.get("year",str(datetime.now().year)),"month":x.cover.get("month",str(datetime.now().month)),"day":x.cover.get("day","")}
    tid=start_task(u["id"],"法学",profile,cover,chosen["title"])
    with db() as c: c.execute("UPDATE topics SET confirmed_task_id=? WHERE id=?",(tid,x.topic_id))
    return {"task_id":tid,"title":chosen["title"]}

@app.post("/api/profile")
def make_profile(x: ProfileRequest, request: Request):
    user_from_request(request)
    if x.type not in ("管理", "设计", "机械", "土木"):
        raise HTTPException(400, "画像接口支持管理、设计、机械、土木")
    p = generate_profile(x.type, x.title, x.target, x.major, x.context)
    if not p.get("core_problems"): p["core_problems"] = ["待补充研究问题"]
    return p

@app.post("/api/start_gen")
def start_gen(x: GenerateRequest, request: Request):
    u=user_from_request(request)
    if x.type not in ("管理","设计","机械","土木"): raise HTTPException(400,"该接口支持管理、设计、机械、土木；法学请使用出题确认")
    tid=start_task(u["id"],x.type,x.profile,x.cover,x.profile.get("title","")); return {"task_id":tid,"status":"queue"}

@app.get("/api/status/{tid}")
def status(tid: str, request: Request):
    u=user_from_request(request)
    with db() as c:r=c.execute("SELECT * FROM tasks WHERE id=? AND user_id=?",(tid,u["id"])).fetchone()
    if not r: raise HTTPException(404,"任务不存在")
    return task_row(r)

@app.get("/api/tasks")
def my_tasks(request: Request):
    u=user_from_request(request)
    with db() as c: rows=c.execute("SELECT * FROM tasks WHERE user_id=? ORDER BY created_at DESC LIMIT 50",(u["id"],)).fetchall()
    return {"tasks":[task_row(x) for x in rows]}

@app.get("/download/{tid}/{filename}")
def download(tid: str, filename: str, request: Request):
    u=user_from_request(request)
    with db() as c:r=c.execute("SELECT * FROM tasks WHERE id=? AND (user_id=? OR ?='admin')",(tid,u["id"],u["role"])).fetchone()
    if not r: raise HTTPException(404,"任务不存在")
    base=(OUT/str(r["user_id"])/tid).resolve(); path=(base/filename).resolve()
    if base not in path.parents or not path.exists(): raise HTTPException(404,"文件不存在")
    return FileResponse(path,filename=path.name)

@app.get("/api/admin/users")
def admin_users(request: Request):
    user_from_request(request,True)
    with db() as c: rows=c.execute("SELECT u.id,u.username,u.role,u.status,u.created_at,u.last_login_at,COUNT(t.id) tasks FROM users u LEFT JOIN tasks t ON t.user_id=u.id GROUP BY u.id ORDER BY u.id").fetchall()
    return {"users":[dict(x) for x in rows]}

class UserStatusRequest(BaseModel):
    status: str

@app.post("/api/admin/users/{user_id}/status")
def admin_user_status(user_id: int, x: UserStatusRequest, request: Request):
    user_from_request(request, True)
    if x.status not in ("active", "disabled"): raise HTTPException(400, "状态只能是 active 或 disabled")
    with db() as c:
        if not c.execute("SELECT 1 FROM users WHERE id=?", (user_id,)).fetchone(): raise HTTPException(404, "用户不存在")
        c.execute("UPDATE users SET status=? WHERE id=?", (x.status, user_id))
        if x.status == "disabled": c.execute("DELETE FROM sessions WHERE user_id=?", (user_id,))
    return {"ok": True, "user_id": user_id, "status": x.status}

@app.get("/api/admin/tasks")
def admin_tasks(request: Request):
    user_from_request(request,True)
    with db() as c: rows=c.execute("SELECT t.*,u.username FROM tasks t JOIN users u ON u.id=t.user_id ORDER BY t.created_at DESC LIMIT 500").fetchall()
    return {"tasks":[task_row(x) for x in rows]}

@app.get("/api/admin/stats")
def admin_stats(request: Request):
    user_from_request(request,True)
    with db() as c:
        total=c.execute("SELECT COUNT(*) n FROM users").fetchone()["n"]; tasks=c.execute("SELECT COUNT(*) n FROM tasks").fetchone()["n"]
        by=[dict(x) for x in c.execute("SELECT paper_type,COUNT(*) count FROM tasks GROUP BY paper_type")]
        q=c.execute("SELECT COUNT(*) n FROM tasks WHERE status IN ('queue','running','waiting_quota')").fetchone()["n"]
    return {"users":total,"tasks":tasks,"by_type":by,"queue":q}

@app.get("/", response_class=HTMLResponse)
def home(): return PAGE
@app.get("/login", response_class=HTMLResponse)
def login_page(): return PAGE
@app.get("/law", response_class=HTMLResponse)
def law_page(): return PAGE
@app.get("/admin", response_class=HTMLResponse)
def admin_page(): return PAGE

PAGE = r'''<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>论文生成系统</title><style>body{font-family:Arial,"Microsoft Yahei";max-width:980px;margin:30px auto;padding:0 18px;color:#243247}nav{display:flex;gap:16px;margin-bottom:24px}button,a{padding:9px 14px;border:0;border-radius:6px;background:#2563eb;color:white;text-decoration:none;cursor:pointer}input,select,textarea{display:block;width:100%;padding:9px;margin:6px 0 12px;box-sizing:border-box}section{background:#f8fafc;border:1px solid #dbe3ee;border-radius:10px;padding:18px;margin:14px 0}.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px}.card{padding:20px;border-radius:9px;background:white;border:1px solid #dbe3ee}.muted{color:#64748b}.err{color:#b91c1c}.ok{color:#047857;white-space:pre-wrap}</style></head><body><nav><a href="/">工作台</a><a href="/law">法学智能出题</a><a href="/admin">管理端</a><button onclick="logout()">退出</button></nav><div id="app"></div><script>
const $=id=>document.getElementById(id), api=async(u,o={})=>{let r=await fetch(u,{credentials:'include',...o});let d=await r.json().catch(()=>({}));if(!r.ok)throw Error(d.detail||'请求失败');return d};
async function logout(){await api('/api/logout',{method:'POST'});location='/login'}
function login(){ $('app').innerHTML='<section><h2>登录 / 注册</h2><input id="un" placeholder="用户名"><input id="pw" type="password" placeholder="密码"><button onclick="doLogin()">登录</button> <button onclick="doReg()">注册</button><p id="msg"></p></section>' }
async function doLogin(){try{let d=await api('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username:un.value,password:pw.value})});location='/'}catch(e){msg.textContent=e.message;msg.className='err'}}
async function doReg(){try{await api('/api/register',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username:un.value,password:pw.value})});msg.textContent='注册成功，请登录'}catch(e){msg.textContent=e.message;msg.className='err'}}
async function home(){try{let m=await api('/api/me'),d=await api('/api/tasks');$('app').innerHTML='<h1>个人论文工作台</h1><p>当前用户：'+m.username+'　角色：'+m.role+'</p><div class="grid"><a class="card" href="/law">⚖️ 法学：系统出题后确认</a><a class="card" href="#" onclick="simple(\'管理\')">📊 经管论文</a><a class="card" href="#" onclick="simple(\'设计\')">🎨 设计论文</a><a class="card" href="#" onclick="simple(\'机械\')">⚙️ 机械论文</a><div class="card muted">🏗️ 土木：沿用服务接口，待个人向导完善</div></div><section><h3>我的任务</h3><pre>'+JSON.stringify(d.tasks,null,2)+'</pre></section>'}catch(e){login()}}
async function law(){try{await api('/api/me')}catch(e){return login()}$('app').innerHTML='<h1>法学智能出题</h1><section><select id="mode"><option value="mix">系统推荐</option><option value="T1">案例分析型</option><option value="T2">规范分析型</option></select><input id="domain" placeholder="可选法律领域，如劳动法"><button onclick="topics()">生成候选题目</button></section><section id="choices"></section>'}
async function topics(){try{let d=await api('/api/law/topics',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({mode:mode.value,domain:domain.value})});choices.innerHTML='<h3>请选择题目</h3>'+d.topics.map((x,i)=>'<div class="card"><b>'+x.title+'</b><p>'+x.law_type+'｜'+x.domain+'<br>'+x.case_names.join('、')+'</p><button onclick="confirmLaw(\''+d.topic_id+'\','+i+')">确认此题</button></div>').join('')}catch(e){choices.innerHTML='<p class=err>'+e.message+'</p>'}}
async function confirmLaw(tid,i){try{let d=await api('/api/law/confirm',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({topic_id:tid,index:i,cover:{}})});location='/'}catch(e){alert(e.message)}}
async function simple(t){
 try{await api('/api/start_gen',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({type:t,profile:{title:prompt('请输入论文题目（'+t+'）')||'个人论文',major:t,company:'',industry:''},cover:{}})});location='/'}catch(e){alert(e.message)}
}
async function admin(){try{await api('/api/me');let [u,t,s]=await Promise.all([api('/api/admin/users'),api('/api/admin/tasks'),api('/api/admin/stats')]);let rows=t.tasks.map(x=>'<div class="card"><b>'+x.username+' / '+(x.paper_type||'')+'</b><p>'+ (x.title||'') +'｜'+x.status+'｜'+x.progress+'%</p><p>'+((x.files||[]).map(f=>'<a href="/download/'+x.id+'/'+encodeURIComponent(f)+'" target="_blank">下载 '+f+'</a>').join('　'))+'</p><pre>'+JSON.stringify(x.audit||{},null,2)+'</pre></div>').join('');$('app').innerHTML='<h1>管理员控制台</h1><section><pre>'+JSON.stringify(s,null,2)+'</pre></section><section><h3>用户</h3><pre>'+JSON.stringify(u.users,null,2)+'</pre></section><section><h3>生成结果与任务</h3>'+rows+'</section>'}catch(e){$('app').innerHTML='<p class=err>'+e.message+'</p>'}}
let p=location.pathname;if(p=='/login')login();else if(p=='/law')law();else if(p=='/admin')admin();else home();
</script></body></html>'''

if __name__ == "__main__":
    import uvicorn
    init_db(); uvicorn.run(app, host="0.0.0.0", port=int(os.getenv("PORT","8000")))
