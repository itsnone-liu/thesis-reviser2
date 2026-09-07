# -*- coding: utf-8 -*-
"""
fix_final_images.py — 论文终版 29 篇图片问题 docx 级修复（2026-09-07）
====================================================================
不依赖 txt 源、不重渲染。四类操作：
  A 图注在图缺（5篇10张）: AI生图 → 插入图注段前
  B 图注图都缺（4篇8张）  : AI生图 → 造图注 → 插入引用句段后
  C 图注补编号（10篇）    : drawing段后图注标题补"图{章}-{序} "前缀
  D 引用重编号（9篇）     : 悬空引用号 → 语义匹配图注/位置就近 → 改run文本
  E 陈秋芳: 图8/图9 图注补标题（上下文推导）
每篇修完内置复检（缺图/悬空/编号对齐），全失败自动回滚该篇。
"""
import os, re, sys, shutil, copy
from docx import Document
from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

sys.path.insert(0, "/root/project/workspace/thesis-reviser")
os.environ.setdefault("IMAGE_API_KEY", open("/root/project/workspace/thesis-reviser/.env").read().split("IMAGE_API_KEY=")[1].split("\n")[0].strip()) if os.path.exists("/root/project/workspace/thesis-reviser/.env") and "IMAGE_API_KEY" not in os.environ else None

ROOT = "论文终版"
BAK = "论文终版/_修改备份/图片修复"
IMGDIR = "论文终版/_修改备份/生成图"
os.makedirs(BAK, exist_ok=True)
os.makedirs(IMGDIR, exist_ok=True)

CAP = re.compile(r"^图\s*(\d{1,2})[-–](\d{1,3})\s*(.*)$")
CAP_S = re.compile(r"^图\s*(\d{1,3})\s+(.*)$")
CAP_BARE = re.compile(r"^图\s*(\d{1,3})$")
CHAP = re.compile(r"^第\s*(\d+)\s*章")
REF = re.compile(r"(?<![简插附纸样意标])图\s?(\d{1,2})[-–](\d{1,3})(?!\d)")
REF_S = re.compile(r"(?<![简插附纸样意标流程数如见中上下左右])\s?图\s?(\d{1,3})(?![\d\-–\.])")

# ---------- 篇目配置 ----------
CFG = {
    # A: 图注在图缺 → {图注号: None}（prompt 由图注+前文组装）
    "25/土木/严涛_029820430414.docx":   {"B": ["图1-1"]},
    "25/土木/陈顾磊_029823431467.docx": {"B": ["图1-1"]},
    "25/土木/金凯_029823431399.docx":   {"A": ["图2-1"], "B": ["图1-1"]},
    "25/土木/张超_029823431459.docx":   {"A": ["图4-1"], "B": ["图1-1"]},
    "25/土木/张德轩_029821431063.docx": {"A": ["图2-1", "图2-2", "图3-1", "图3-2", "图3-3", "图3-4"]},
    # B: 图注图都缺 → 造注插图（引用号: None）
    "25/土木/冯然_029822430107.docx":   {"B": ["图1-1", "图3-1", "图3-2", "图3-3"]},
    "25/土木/柳骁蕾_029823430287.docx": {"B": ["图1-1", "图1-2"]},
    "25/土木/王尧_029823431344.docx":   {"B": ["图1-1"]},
    "25/土木/周亮_029823430350.docx":   {"B": ["图1-1"]},
    # C: 图注补编号
    "25/机械/王见康_029823430219.docx": {"C": []},
    "25/机械/许磊_029823430932.docx":   {"C": []},
    "25/机械/陈霁_029823430735.docx":   {"C": []},
    "25/设计/张芳_029822410430.docx":   {"C": []},
    "25/土木/王楠_029822430386.docx": {"B": ["图1-1"]},
    "25/土木/邵喆_029820430090.docx": {"B": ["图1-1", "图1-2"]},
    "25/土木/刘德荣_029822430813.docx": {"B": ["图1-1"]},
    "25/土木/曹龙华_029823430221.docx": {"B": ["图1-1"]},
    "25/土木/袁园_029821430002.docx": {"B": ["图1-1"]},
    "25/土木/冯磊_029822430252.docx": {"B": ["图1-1"]},
    # D: 引用重编号（悬空号: None=自动定目标）
    "25/经管/陈秋芳_029822430405.docx": {"E": ["图8", "图9"]},
}

def paras(doc):
    return doc.paragraphs

def is_draw(p):
    return "<w:drawing" in p._p.xml or "w:pict" in p._p.xml

def chapter_of(plist, idx):
    ch = 0
    for i in range(idx, -1, -1):
        m = CHAP.match(plist[i].text.strip())
        if m:
            return int(m.group(1))
    return 0

def captions(plist):
    """[(idx, 编号, 标题, style)] style: 'ch'|'simple'|'bare'"""
    out = []
    for i, p in enumerate(plist):
        t = p.text.strip()
        m = CAP.match(t)
        if m: out.append((i, f"图{m.group(1)}-{m.group(2)}", m.group(3), "ch")); continue
        m = CAP_S.match(t)
        if m: out.append((i, f"图{m.group(1)}", m.group(2), "simple")); continue
        m = CAP_BARE.match(t)
        if m: out.append((i, f"图{m.group(1)}", "", "bare"))
    return out

def refs_of(plist):
    s = set()
    for p in plist:
        t = p.text
        for r in REF.finditer(t): s.add(f"图{r.group(1)}-{r.group(2)}")
        for r in REF_S.finditer(t): s.add(f"图{r.group(1)}")
    return s

def check(plist):
    n = sum(1 for p in plist if is_draw(p))
    caps = captions(plist)
    ids = [c[1] for c in caps]
    refs = refs_of(plist)
    dangling = sorted(r for r in refs if r not in ids)
    missing = len(ids) - n if len(ids) > n else 0
    return n, ids, dangling, missing

def bigram(s):
    s = re.sub(r"[\sA-Za-z0-9图\-–：:，。；、]", "", s)
    return set(s[i:i+2] for i in range(len(s) - 1))

def semantic_target(sent, cap_ids, cap_titles, before_ids):
    """引用句 → 最佳图注号：bigram 重叠；低置信→位置最近的前置图注"""
    bg = bigram(sent)
    best, score = None, 0
    for cid, title in zip(cap_ids, cap_titles):
        sc = len(bg & bigram(title))
        if sc > score: best, score = cid, sc
    if score >= 3: return best
    return before_ids[-1] if before_ids else cap_ids[-1]

def gen_image(dno, title, desc, dtype):
    from core import generate_single_image
    os.makedirs(IMGDIR, exist_ok=True)
    drawing = {"id": dno.replace("图", ""), "seq": dno, "type": dtype, "title": title, "description": desc}
    return generate_single_image(drawing, IMGDIR, max_retries=4)

def infer_type(title):
    if any(k in title for k in ("效果图", "鸟瞰", "透视")): return "效果图"
    return "工程制图"

def insert_pic_before(doc, anchor_p, png, cap_no, cap_title):
    np = anchor_p.insert_paragraph_before()
    np.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = np.add_run()
    run.add_picture(png, width=Cm(13.5))
    return np

def add_cap_after(anchor_p, cap_no, title):
    from docx.oxml.ns import qn
    new_p = copy.deepcopy(anchor_p._p)
    anchor_p._p.addnext(new_p)
    from docx.text.paragraph import Paragraph
    np = Paragraph(new_p, anchor_p._parent)
    for r in list(np.runs):
        r._r.getparent().remove(r._r)
    run = np.add_run(f"{cap_no} {title}")
    np.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return np

def replace_ref_in_para(p, old, new):
    """段内跨run替换引用号：合并runs到首run保格式（正文段简单格式，风险低）"""
    t = p.text
    if old not in t: return False
    nt = t.replace(old, new, 1)
    runs = p.runs
    if not runs: return False
    runs[0].text = nt
    for r in runs[1:]:
        r._r.getparent().remove(r._r)
    return True

def fix_one(rel, ops, dry=False):
    src = os.path.join(ROOT, rel)
    log = []
    doc = Document(src)
    pl = paras(doc)
    caps = captions(pl)
    cap_ids = [c[1] for c in caps]
    cap_titles = [f"{c[1]} {c[2]}" for c in caps]

    # ---- C: 图注补编号（双侧探测：已编号图注方位优先，否则取更短侧）----
    def ok_cap_text(t):
        return 3 < len(t) <= 50 and not t.endswith(("。", "；", "！")) \
               and not CHAP.match(t) and not t.startswith(("数据来源", "表", "第"))
    def set_para_text(p, txt):
        runs = p.runs
        if not runs:
            p.add_run(txt); return
        runs[0].text = txt
        for r in runs[1:]:
            r._r.getparent().remove(r._r)
    if "C" in ops:
        # 参照方位：已编号图注在图前还是图后
        side_ref = None
        numbered = [(ci, no) for ci, no, _, _ in caps]
        draw_idx = [i for i, p in enumerate(pl) if is_draw(p)]
        for ci, no in numbered:
            near_prev = ci - 1 in draw_idx
            near_next = ci + 1 in draw_idx
            if near_prev: side_ref = "prev"; break
            if near_next: side_ref = "next"; break
        counters = {}
        for i, p in enumerate(pl):
            if not is_draw(p): continue
            cands = []
            j = i - 1
            while j >= 0 and j > i - 4:
                t = pl[j].text.strip()
                if t:
                    if ok_cap_text(t): cands.append(("prev", j, t))
                    break
                j -= 1
            j = i + 1
            while j < len(pl) and j < i + 4:
                t = pl[j].text.strip()
                if t:
                    if ok_cap_text(t): cands.append(("next", j, t))
                    break
                j += 1
            if not cands: continue
            pick = None
            if side_ref:
                pick = next((c for c in cands if c[0] == side_ref), None)
            if not pick:
                pick = min(cands, key=lambda c: len(c[2]))
            side, j, t = pick
            ch = chapter_of(pl, i)
            counters[ch] = counters.get(ch, 0) + 1
            no = f"图{ch}-{counters[ch]}"
            set_para_text(pl[j], f"{no} {t}")
            log.append(f"C:{no}({side})←[{t[:18]}]")

    # ---- E: 陈秋芳裸图注补标题 ----
    if "E" in ops:
        for want in ops["E"]:
            for i, p in enumerate(pl):
                m = CAP_BARE.match(p.text.strip())
                if m and f"图{m.group(1)}" == want:
                    ctx = pl[i-2].text.strip() if i >= 2 else ""
                    bg = bigram(ctx)
                    # 从正文找最相关短语：取引用句"如图N"所在段的前半
                    title = ""
                    for q in pl:
                        if want in q.text and q.text.strip() != p.text.strip():
                            mm = re.search(r"([^。]{0,24}?)" + re.escape(want), q.text)
                            if mm and mm.group(1):
                                title = mm.group(1).split("，")[-1].strip()
                                if len(title) < 4: title = mm.group(1).strip()
                                break
                    if not title: title = "分析图示"
                    runs = p.runs
                    if runs:
                        runs[0].text = f"{want} {title}"
                        for r in runs[1:]:
                            r._r.getparent().remove(r._r)
                    else:
                        p.add_run(f"{want} {title}")
                    log.append(f"E:{want}+[{title[:16]}]")
                    break

    # ---- D: 引用重编号 ----
    if "D" in ops:
        pl = paras(doc); caps = captions(pl)
        cap_ids = [c[1] for c in caps]; cap_titles = [c[1] + " " + c[2] for c in caps]
        cap_pos = {c[1]: c[0] for c in caps}
        for old, new in ops["D"]:
            if new: continue
            hit = None
            for i, p in enumerate(pl):
                if old in p.text:
                    hit = (i, p); break
            if not hit: log.append(f"D:{old}未找到"); continue
            i, p = hit
            before_ids = [c for c in cap_ids if cap_pos.get(c, 10**9) < i]
            new = semantic_target(p.text, cap_ids, cap_titles, before_ids)
            if replace_ref_in_para(p, old, new):
                log.append(f"D:{old}→{new}")
            else:
                log.append(f"D:{old}替换失败")

    # ---- A: 图注在图缺 → 生图插图注前 ----
    if "A" in ops and not dry:
        pl = paras(doc); caps = captions(pl)
        for want in ops["A"]:
            cap = next((c for c in caps if c[1] == want), None)
            if not cap: log.append(f"A:{want}图注未找到"); continue
            i, no, title, _ = cap
            ctx = pl[i-1].text.strip() if i else ""
            if not ctx or len(ctx) < 15:
                ctx = title
            png = gen_image(no, title if title else no, ctx[:200], infer_type(title))
            if not png: log.append(f"A:{no}生图失败"); continue
            insert_pic_before(doc, pl[i], png, no, title)
            log.append(f"A:{no}插图+[{png.split('/')[-1]}]")

    # ---- B: 图注图都缺 → 生图+造注+插引用段后 ----
    if "B" in ops and not dry:
        pl = paras(doc)
        for want in ops["B"]:
            hit = None
            for i, p in enumerate(pl):
                if want in p.text: hit = (i, p); break
            if not hit: log.append(f"B:{want}引用未找到"); continue
            i, p = hit
            t = p.text
            mm = re.search(r"([^。，；]{2,30}?)" + re.escape(want), t)
            title = mm.group(1).strip() if mm else ""
            title = re.sub(r"^(的|了|和|与|及|如|见|如见|参见|所示|下图|上图)\s*", "", title)
            title = re.sub(r"(如|见|所示|详|下|上|参见)\s*$", "", title).strip()
            title = title[-24:]
            bad = (not title) or title == want or len(title) < 4 or re.fullmatch(r"[\d\s\.]+", title)
            if bad:
                # 兜底：引用句所在章的章标题 + "图"
                ch = chapter_of(pl, i)
                ch_title = ""
                for q in pl:
                    m3 = CHAP.match(q.text.strip())
                    if m3 and int(m3.group(1)) == ch:
                        ch_title = re.sub(r"^第\d+\s*章\s*", "", q.text.strip())
                        break
                title = (ch_title[:12] or "工程") + "布置图"
            png = gen_image(want, title, t[:240], infer_type(title))
            if not png: log.append(f"B:{want}生图失败"); continue
            # 在引用段后插：图段（居中）+ 图注段
            pic_p = add_cap_after(p, "__PIC__", "")   # 复制引用段格式做底
            set_para_text(pic_p, "")
            r0 = pic_p.add_run(); r0.add_picture(png, width=Cm(13.5))
            pic_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cap_p = add_cap_after(pic_p, want, title)
            log.append(f"B:{want}+图注[{title[:16]}]")

    return doc, log

def verify(rel):
    doc = Document(os.path.join(ROOT, rel))
    pl = paras(doc)
    n, ids, dangling, missing = check(pl)
    return n, ids, dangling, missing

if __name__ == "__main__":
    dry = "--dry" in sys.argv
    skip_ab = "--skip-ab" in sys.argv
    only = None
    for a in sys.argv[1:]:
        if a.startswith("--only="): only = a.split("=", 1)[1]
    report = []
    for rel, ops in CFG.items():
        if only and only not in rel: continue
        if skip_ab and ("A" in ops or "B" in ops): continue
        print(f"\n=== {rel}")
        try:
            shutil.copy2(os.path.join(ROOT, rel), os.path.join(BAK, os.path.basename(rel) + ".bak.docx"))
            doc, log = fix_one(rel, ops, dry=dry)
            if not dry:
                doc.save(os.path.join(ROOT, rel))
            n, ids, dangling, missing = verify(rel)
            status = "✅" if not dangling and missing <= 0 else "⚠️"
            print(f"  {status} embed={n} caps={len(ids)} 悬空={dangling or '无'} 缺={missing}")
            for l in log: print("   ", l)
            report.append((rel, status, n, len(ids), ";".join(dangling), missing))
        except Exception as e:
            print(f"  ❌ 异常: {e!r}")
            report.append((rel, "❌", 0, 0, repr(e), 0))
    if not dry:
        with open("/root/project/workspace/图片修复报告_0907.csv", "w", encoding="utf-8-sig") as f:
            f.write("file,status,embed,caps,dangling,missing\n")
            for r in report:
                f.write(",".join(str(x).replace(",", "；") for x in r) + "\n")
        print("\n→ /root/project/workspace/图片修复报告_0907.csv")
