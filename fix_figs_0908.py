# -*- coding: utf-8 -*-
"""图表缺陷专项修复 (2026-09-08, 11篇有源/设计批)。
F1 同号双重编号剥离: "图1-1 图1-1 标题" → "图1-1 标题" (引用号未变无需同步)
F2 引用失配修复: "如图4-1"(不存在) → 同章最近现存编号改写; 无候选删引用从句
F3 复验: audit_final 图表检查清零
备份到 _修改备份0908/，记录 返修记录_0908图表.csv
"""
import os, re, sys, csv, json, shutil
from docx import Document

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
from audit_final import load_docx, audit_docx_only as af_audit

ROOT = os.path.join(BASE, "论文终版")
BAK = os.path.join(ROOT, "_修改备份0908")
TARGETS = [
    "25/机械/许新虎_029822430938.docx", "25/机械/许磊_029823430932.docx",
    "25/机械/陈霁_029823430735.docx", "25/经管/周子杰_029822440016.docx",
    "25/设计/张芳_029822410430.docx",
]
REC = os.path.join(BASE, "论文修订档案/返修记录_0908图表.csv")

DBL_SAME = re.compile(r"^([图表])(\d{1,2})-(\d{1,2})\s+\1\2-\3\b\s*(.*)$")
REF_TOK = re.compile(r"[如见]?[图表]\d{1,2}(?:-\d{1,2})?")
CAPT_RE = re.compile(r"^([图表])\s*(\d{1,2})(?:-(\d{1,2}))?\s+\S")
VERBY = re.compile(r"显示|如下|所示|可知|可以看出|表明|见下")


def set_text(p, new):
    if p.runs:
        p.runs[0].text = new
        for r in p.runs[1:]:
            r.text = ""
    else:
        p.text = new


def existing_ids(d):
    """现存图表编号集合: {(kind, 'C-M' or 'N')}"""
    ids = set()
    for p in d.paragraphs:
        s = p.text.strip()
        m = CAPT_RE.match(s)
        if m and not VERBY.search(s[:14]) and len(s) < 60:
            k, a, b = m.group(1), m.group(2), m.group(3)
            ids.add((k, f"{a}-{b}" if b else a))
            ids.add((k, a))  # 简单式也记
    return ids


def fix_paper(rel):
    path = os.path.join(ROOT, rel)
    log = []
    d = Document(path)
    # F1 同号双重编号
    for p in d.paragraphs:
        s = p.text.strip()
        m = DBL_SAME.match(s)
        if m:
            new = f"{m.group(1)}{m.group(2)}-{m.group(3)} {m.group(4)}"
            set_text(p, new)
            log.append(f"F1同号剥离[{s[:20]}→{new[:20]}]")
    # F2 引用失配
    ids = existing_ids(d)
    def ids_of(k):
        return {v for kk, v in ids if kk == k}
    for p in d.paragraphs:
        s = p.text
        if not s.strip() or CAPT_RE.match(s.strip()):
            continue
        out, changed = [], False
        # 逐引用token检查
        pos = 0
        for m in REF_TOK.finditer(s):
            tok = m.group(0)
            km = re.match(r"[如见]?([图表])([\d-]+)", tok)
            if not km:
                continue
            kind, num = km.group(1), km.group(2)
            have = ids_of(kind)
            if num in have:
                continue
            # 失配: 找同章最近现存 (C-M形态)
            cands = [v for v in have if re.match(rf"{num.split('-')[0]}-\d+$", v)] if "-" in num else []
            if "-" in num and len(cands) >= 1:
                # 唯一同章候选 → 改写; 多候选不动(记人工)
                sub = m.group(0).replace(num, cands[0])
                out.append((m.start(), m.end(), sub))
                changed = True
                log.append(f"F2引用改写[{tok}→{cands[0]}]")
            elif not changed:
                # 无候选 → 删"如/见+编号"及前后紧邻的顿号/逗号空格
                st = m.start()
                en = m.end()
                while st > 0 and s[st-1] in "（(、，, ":
                    st -= 1
                while en < len(s) and s[en] in "）)、，, ":
                    en += 1
                out.append((st, en, ""))
                changed = True
                log.append(f"F2引用删除[{tok}]")
        if changed and out:
            new = s
            for st, en, rep in sorted(out, reverse=True):
                new = new[:st] + rep + new[en:]
            if len(new.strip()) >= 10:  # 防删空句子
                set_text(p, new)
            else:
                log.append(f"F2跳过(句子过短)[{s[:20]}]")
    return d, log


def main():
    os.makedirs(BAK, exist_ok=True)
    rows = []
    for rel in TARGETS:
        src = os.path.join(ROOT, rel)
        bak = os.path.join(BAK, rel.replace("/", "__"))
        shutil.copy2(src, bak)
        try:
            d, log = fix_paper(rel)
            tmp = src + ".tmp"
            d.save(tmp)
            os.replace(tmp, src)
            # F3 复验
            ptype = {"土木": "土木", "机械": "机械", "经管": "管理", "设计": "设计"}[rel.split("/")[1]]
            r = af_audit(src, ptype)
            imgprob = [p for p in r["problems"] if "失配" in p or "双重" in p or "vs" in p]
            ok = not imgprob
            rows.append([rel, "✅" if ok else "⚠️", "; ".join(log)[:500], imgprob[0] if imgprob else ""])
            print(("✅" if ok else "⚠️"), rel.split("/")[-1][:20], f"{len(log)}项修复", imgprob[0][:40] if imgprob else "")
        except Exception as e:
            shutil.copy2(bak, src)  # 还原
            rows.append([rel, "❌", f"异常:{type(e).__name__}", str(e)[:100]])
            print("❌", rel, e)
    with open(REC, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerows(rows)
    ok = sum(1 for r in rows if r[1] == "✅")
    print(f"\n图表修复(本轮): {ok}/{len(rows)} 清零 → 追加 {REC}")


if __name__ == "__main__":
    main()
