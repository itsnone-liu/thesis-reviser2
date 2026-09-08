# -*- coding: utf-8 -*-
"""
batch_law.py — 法学论文批量生成驱动
====================================
输入任务清单JSON(数组), 每任务:
  {"student":"张三", "sid":"202301000001", "advisor":"李四",
   "case":"指导案例237号"}                        ← T1单案
  {"student":"王五", "sid":"202301000002", "advisor":"李四",
   "t2":true, "domain":"新就业形态"}              ← T2案例群(需库内≥3案verified)
可选: "year":"26" (输出到 论文终版/26/法学/, 默认26)
输出: 论文终版/<year>/法学/姓名_学号.docx + 逐篇审计 + 汇总报告CSV
用法: . ./.env && python batch_law.py --tasks tasks.json [--limit N] [--dry]
"""
import sys, os, json, argparse, csv, shutil
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import caselib
import law
from generator import generate
from renderer import render
from law_pilot import polish_title, polish_title_t2, law_audit
from audit_final import audit_docx_only

def run_task(t, out_dir, work_dir, used_cases=None, task_idx=0):
    student = t.get("student", "某生")
    sid = str(t.get("sid", ""))
    advisor = t.get("advisor", "某某")
    if t.get("t2"):
        domain = t.get("domain", "新就业形态")
        # 0908: 同批去重 + 稳定轮换, 同领域多篇不再拿同一案例群
        cluster = caselib.cluster(domain, exclude=used_cases, rotate=task_idx)
        if not cluster:
            return None, f"T2缺≥3个verified的'{domain}'案例(排除已用后)", []
        cases = cluster
        title = polish_title_t2(cluster, domain)
        profile = {"title": title, "law_type": "T2", "cluster": cluster, "domain": domain}
        if used_cases is not None:
            used_cases.update(str(c.get("案号", "")) for c in cases)
            used_cases.update(c.get("名称", "") for c in cases)
    else:
        case = caselib.get(t.get("case", ""))
        if not case:
            return None, f"案例库未找到: {t.get('case')}", []
        # 0908: T1同批防重(清单写重直接报错, 不静默生成两篇同案论文)
        if used_cases is not None:
            if str(case.get("案号", "")) in used_cases or case.get("名称", "") in used_cases:
                return None, f"T1案例本批已用过: {case.get('案号') or case.get('名称', '')}", []
            used_cases.add(str(case.get("案号", "")))
            used_cases.add(case.get("名称", ""))
        cases = [case]
        title = polish_title(case)
        profile = {"title": title, "law_type": "T1", "case": case}

    txt = generate(profile, "法学")
    txt_path = os.path.join(work_dir, f"{student}_{sid}.txt")
    open(txt_path, "w", encoding="utf-8").write(txt)

    cover = {"title": title, "name": student, "student_id": sid, "major": "法学",
             "level": "本科", "advisor": advisor, "year": "2026", "month": "5", "day": ""}
    docx_name = f"{student}_{sid}.docx" if sid else f"{student}.docx"
    docx_path = os.path.join(out_dir, docx_name)
    render(txt_path, docx_path, paper_type="法学", cover_info=cover)
    os.remove(txt_path)
    rp = docx_path + ".report.json"
    if os.path.exists(rp):   # 渲染回执不入终版目录
        os.remove(rp)

    probs = law_audit(txt, cases)
    return docx_path, title, probs

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--tasks", required=True)
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--dry", action="store_true")
    args = ap.parse_args()

    tasks = json.load(open(args.tasks, encoding="utf-8"))
    if args.limit:
        tasks = tasks[:args.limit]
    if args.dry:
        for i, t in enumerate(tasks):
            print(f"[dry {i+1}] {t.get('student')} {'T2:' + t.get('domain','') if t.get('t2') else 'T1:' + t.get('case','')}")
        return

    rows = []
    used_cases = set()   # 0908: 全批案例去重账本(案号+名称)
    for i, t in enumerate(tasks):
        year = str(t.get("year", "26"))
        out_dir = os.path.join("论文终版", year, "法学")
        os.makedirs(out_dir, exist_ok=True)
        print(f"── [{i+1}/{len(tasks)}] {t.get('student', '某生')} "
              f"{'T2:' + t.get('domain','') if t.get('t2') else 'T1:' + t.get('case','')}", flush=True)
        try:
            docx_path, info, probs = run_task(t, out_dir, "/tmp",
                                              used_cases=used_cases, task_idx=i)
            if docx_path is None:
                print(f"   ⚠️ 失败: {info}")
                rows.append({"student": t.get("student"), "sid": t.get("sid"), "file": "",
                             "title": info, "gen_audit": ";".join(probs), "final_verdict": "❌"})
                continue
            # docx级终审计(评阅必达项)
            fin = audit_docx_only(docx_path, "法学")
            v = fin["verdict"]
            print(f"   ✅ {os.path.basename(docx_path)} | {info[:30]} | 终审计{v} {fin['stats'].get('law_body_words','')}字")
            rows.append({"student": t.get("student"), "sid": t.get("sid"), "file": docx_path,
                         "title": info, "gen_audit": ";".join(probs),
                         "final_verdict": v, "final_problems": ";".join(fin["problems"])})
        except Exception as e:
            print(f"   ❌ 异常: {e}")
            rows.append({"student": t.get("student"), "sid": t.get("sid"), "file": "",
                         "title": "", "gen_audit": f"异常:{e}", "final_verdict": "❌"})

    if rows:
        with open("/tmp/law_batch_report.csv", "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=["student", "sid", "file", "title",
                                              "gen_audit", "final_verdict", "final_problems"])
            w.writeheader()
            w.writerows(rows)
        ok = sum(1 for r in rows if r["final_verdict"] == "✅")
        print(f"\n完成: {ok}/{len(rows)} ✅ | 报告: /tmp/law_batch_report.csv")

if __name__ == "__main__":
    main()
