# -*- coding: utf-8 -*-
"""
law_pilot.py — 法学论文单篇生成驱动(样稿/试验用)
==================================================
与其它专业同一管线: profile → generator.generate(profile, "法学") → renderer.render
案例来自 caselib(案例先行), 题目由案例经LLM润色产生。
用法: . ./.env && python law_pilot.py --case "指导案例237号" [--t2 --domain 新就业形态]
     [--student 李某 --sid 202301000000 --advisor 某某] [--skip-gen]
"""
import sys, os, json, argparse, re
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import caselib
import law
from generator import generate
from renderer import render
from core import call_llm, clean_text

OUT_DIR = "output/law_pilot"

def polish_title(case: dict) -> str:
    rule = law.derive_title_t1(case)
    prompt = f"""以下是一个真实的最高人民法院指导性案例:
名称: {case.get('名称')}
案由: {case.get('案由')}
争议焦点: {case.get('争议焦点')}

请为本科毕业论文(案例分析型)拟定规范题目, 要求:
- 主标题概括本案核心法律问题(10至20字), 副标题"——以{case.get('名称')}为例";
- 法学论文题目风格, 不口语化;
- 只输出题目本身(不要书名号、不要引号、不要解释)。"""
    raw = call_llm(prompt, max_tokens=200)
    t = clean_text(raw).strip().strip('《》"“”').splitlines()[0][:60]
    if t and "——" not in t and "为例" not in t:   # T1必须有副标题
        t = f"{t}——以{case.get('名称')}为例"
    return t or rule

def polish_title_t2(cluster: list, domain: str) -> str:
    names = "、".join(c.get("名称", "")[:18] for c in cluster[:4])
    prompt = f"""以下{len(cluster)}起真实的{domain}领域典型案例(均为最高人民法院指导性案例):
{names}

请为本科毕业论文(规范分析型: 由案例引出问题、分析并提完善建议)拟定规范题目, 要求:
- 主标题概括这些案例共同反映的法律问题(12至22字);
- 法学论文题目风格, 不口语化, 不出现具体案件名;
- 只输出题目本身(不要书名号、不要引号、不要解释)。"""
    raw = call_llm(prompt, max_tokens=200)
    t = clean_text(raw).strip().strip('《》"“”').splitlines()[0][:40]
    return t or law.derive_title_t2(cluster, domain)

def law_audit(txt: str, cases: list) -> list:
    """法学期审计: 字数/摘要/关键词/文献/案例一致性"""
    probs = []
    body = re.sub(r'---PAGE_BREAK---', '', txt)
    body = re.sub(r'(?s)目录\n.*?\n(?=\n*第1章)', '', body, count=1)   # 剥目录区, 防审计误配
    m = re.search(r'第1章.*?(?=参考文献|$)', body, re.S)
    if not m:
        probs.append("未找到正文区")
    else:
        n = len(re.sub(r'\s|第\d+章[^\n]*|\d+\.\d+(\.\d+)?[^\n]*', '', m.group(0)))
        probs.append(f"OK正文{n}字" if n >= 8000 else f"❌正文{n}字 < 8000")
    ab = re.search(r'摘要\n(.+?)(?=\n)', body, re.S)
    if ab:
        n = len(re.sub(r'\s', '', ab.group(1)))
        probs.append(f"OK摘要{n}字" if 300 <= n <= 500 else f"❌摘要{n}字 不在300-500")
    kw = re.search(r'关键词\n(.+)', body)
    if kw:
        n = len([k for k in re.split(r'[;；]', kw.group(1)) if k.strip()])
        probs.append(f"OK关键词{n}个" if 3 <= n <= 5 else f"❌关键词{n}个 不在3-5")
    rn = len(re.findall(r'^\[\d+\]', body, re.M))
    probs.append(f"OK文献{rn}条" if rn >= 10 else f"❌文献{rn}条 < 10")
    cons = caselib.check_consistency(body, cases)
    probs += cons if cons else ["OK案例一致性(案号/特征词/来源)"]
    return probs

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--case", default="指导案例237号")
    ap.add_argument("--t2", action="store_true", help="规范分析型(案例群)")
    ap.add_argument("--domain", default="新就业形态")
    ap.add_argument("--student", default="李某")
    ap.add_argument("--sid", default="202301000000")
    ap.add_argument("--advisor", default="某某")
    ap.add_argument("--skip-gen", action="store_true")
    args = ap.parse_args()

    os.makedirs(OUT_DIR, exist_ok=True)
    log = lambda m: print(m, flush=True)
    txt_path = f"{OUT_DIR}/paper.txt"
    meta_path = f"{OUT_DIR}/meta.json"

    if args.skip_gen and os.path.exists(txt_path) and os.path.exists(meta_path):
        meta = json.load(open(meta_path))
        case = caselib.get(args.case) or {}
        cases = [case] if not args.t2 else caselib.cluster(args.domain)
    else:
        if args.t2:
            cluster = caselib.cluster(args.domain)
            if not cluster:
                sys.exit(f"T2需要≥3个已核的'{args.domain}'案例")
            log(f"① 案例群({len(cluster)}案): {[c['案号'] for c in cluster]}")
            title = polish_title_t2(cluster, args.domain)
            profile = {"title": title, "law_type": "T2", "cluster": cluster, "domain": args.domain}
            cases = cluster
        else:
            case = caselib.get(args.case)
            if not case:
                sys.exit(f"案例库未找到: {args.case}")
            log(f"① 案例装载: {case['案号']} {case['名称'][:24]}")
            title = polish_title(case)
            profile = {"title": title, "law_type": "T1", "case": case}
            cases = [case]
        log(f"② 拟题: {title}")
        log("③ 生成(与其它专业同一管线 generator.generate):")
        txt = generate(profile, "法学")
        open(txt_path, "w", encoding="utf-8").write(txt)
        json.dump({"title": title, "case": args.case, "t2": args.t2, "domain": args.domain},
                  open(meta_path, "w"), ensure_ascii=False, indent=1)
        log(f"④ TXT: {txt_path} (全篇{len(txt)}字符)")

    meta = json.load(open(meta_path))
    log("⑤ 法学期审计:")
    for p in law_audit(open(txt_path, encoding='utf-8').read(), cases):
        log(f"   {p}")

    log("⑥ 渲染DOCX(同一渲染模块):")
    cover = {"title": meta["title"], "name": args.student, "student_id": args.sid,
             "major": "法学", "level": "本科", "advisor": args.advisor,
             "year": "2026", "month": "5", "day": ""}
    docx_path = f"{OUT_DIR}/{meta['title'][:30]}.docx".replace("——", "-")
    render(txt_path, docx_path, paper_type="法学", cover_info=cover)
    log(f"✅ 样稿: {docx_path}")

if __name__ == "__main__":
    main()
