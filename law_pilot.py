# -*- coding: utf-8 -*-
"""
law_pilot.py — 法学论文单篇生成驱动(样稿用)
============================================
流程: 案例库取案 → LLM拟题(案例先行) → 分章生成(素材注入) → 摘要/关键词
     → 确定性参考文献 → 组装TXT(渲染契约同其他专业) → renderer渲染DOCX
用法: python law_pilot.py --case "指导案例237号" [--student 李某 --sid 202301000000 --advisor 某某]
"""
import sys, os, json, argparse, re, time
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import caselib
import law
from prompts.law import build_chapter_prompt, build_abstract_prompt, build_keywords_prompt, build_references
from core import call_llm, clean_text
from renderer import render

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
    return t or rule

def gen_chapters(title: str, case: dict, log):
    parts, prev = [], ""
    for chapter, seq, limit in law.LAW_T1_CHAPTERS:
        log(f"  第{seq}章 {chapter} (~{limit}字) ...")
        t0 = time.time()
        body = clean_text(call_llm(build_chapter_prompt(title, case, chapter, seq, limit, prev),
                                   max_tokens=max(4096, int(limit * 2.2)), retry=3))
        body = re.sub(r'^(第?\d+章[^\n]*|一、[^\n]*)\n+', '', body).strip()
        parts.append((f"第{seq}章 {chapter}", body))
        prev = f"{chapter}: " + body[:800]
        log(f"    {len(body)}字 / {time.time()-t0:.0f}s")
    return parts

def assemble(title: str, case: dict, chapters, abstract: str, kws: str, refs: list) -> str:
    PB = "\n\n---PAGE_BREAK---\n\n"
    toc = "\n".join(h for h, _ in chapters)
    blocks = [f"摘要\n{abstract}", f"关键词\n{kws}", f"目录\n{toc}"]
    blocks += [f"{h}\n{body}" for h, body in chapters]
    blocks.append("参考文献\n" + "\n".join(f"[{i}] {r}" for i, r in enumerate(refs, 1)))
    return PB.join(blocks)

def law_audit(txt: str, case: dict) -> list:
    """法学期审计: 字数/摘要/关键词/文献/案例一致性"""
    probs = []
    body = re.sub(r'---PAGE_BREAK---', '', txt)
    # 正文字数: 剥离摘要/关键词/目录/参考文献
    m = re.search(r'第1章.*?(?=参考文献|$)', body, re.S)
    if not m:
        probs.append("未找到正文区")
    else:
        n = len(re.sub(r'\s|第\d+章[^\n]*|(一|二|三|四|五|六)(、|\.)[^\n]*|\d\.[^\n]*', '', m.group(0)))
        (probs.append if n < 8000 else lambda x: None)(f"正文{n}字 < 8000") if n < 8000 else None
        if n >= 8000: probs.append(f"OK正文{n}字")
    ab = re.search(r'摘要\n(.+?)(?=\n)', body, re.S)
    if ab:
        n = len(re.sub(r'\s', '', ab.group(1)))
        probs.append(f"OK摘要{n}字" if 300 <= n <= 500 else f"摘要{n}字 不在300-500")
    kw = re.search(r'关键词\n(.+)', body)
    if kw:
        n = len([k for k in re.split(r'[;；]', kw.group(1)) if k.strip()])
        probs.append(f"OK关键词{n}个" if 3 <= n <= 5 else f"关键词{n}个 不在3-5")
    rn = len(re.findall(r'^\[\d+\]', body, re.M))
    probs.append(f"OK文献{rn}条" if rn >= 10 else f"文献{rn}条 < 10")
    probs += caselib.check_consistency(body, [case])
    return probs

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--case", default="指导案例237号")
    ap.add_argument("--student", default="李某")
    ap.add_argument("--sid", default="202301000000")
    ap.add_argument("--advisor", default="某某")
    ap.add_argument("--skip-gen", action="store_true", help="只重渲染已有TXT")
    args = ap.parse_args()

    case = caselib.get(args.case)
    if not case:
        sys.exit(f"案例库未找到: {args.case}")
    os.makedirs(OUT_DIR, exist_ok=True)
    log = lambda m: print(m, flush=True)

    txt_path = f"{OUT_DIR}/paper.txt"
    if args.skip_gen and os.path.exists(txt_path):
        title = json.load(open(f"{OUT_DIR}/meta.json"))["title"]
    else:
        log(f"① 案例装载: {case['案号']} {case['名称'][:24]}")
        title = polish_title(case)
        log(f"② 拟题: {title}")
        log("③ 分章生成:")
        chapters = gen_chapters(title, case, log)
        log("④ 摘要/关键词:")
        abstract = clean_text(call_llm(build_abstract_prompt(title, case), max_tokens=900))
        if not (280 <= len(re.sub(r'\s', '', abstract)) <= 500):
            abstract = clean_text(call_llm(
                build_abstract_prompt(title, case) + "\n\n注意: 上次输出超长, 本次必须压缩到450字以内, 宁精勿滥。",
                max_tokens=800))
            log(f"   摘要超长已重写: {len(re.sub(chr(92)+'s','',abstract))}字")
        kws = clean_text(call_llm(build_keywords_prompt(title, abstract), max_tokens=120)).strip()
        refs = build_references(case)
        log(f"   摘要{len(abstract)}字 | 关键词: {kws} | 文献{len(refs)}条")
        txt = assemble(title, case, chapters, abstract, kws, refs)
        open(txt_path, "w", encoding="utf-8").write(txt)
        json.dump({"title": title, "case": case.get("案号"), "keywords": kws},
                  open(f"{OUT_DIR}/meta.json", "w"), ensure_ascii=False, indent=1)
        log(f"⑤ TXT: {txt_path} (全篇{len(txt)}字符)")

    log("⑥ 法学期审计:")
    probs = law_audit(open(txt_path, encoding="utf-8").read(), case)
    for p in probs: log(f"   {p}")

    log("⑦ 渲染DOCX:")
    meta = json.load(open(f"{OUT_DIR}/meta.json"))
    cover = {"title": meta["title"], "name": args.student, "student_id": args.sid,
             "major": "法学", "level": "本科", "advisor": args.advisor,
             "year": "2026", "month": "5", "day": ""}
    docx_path = f"{OUT_DIR}/{meta['title'][:30]}.docx".replace("——", "-")
    render(txt_path, docx_path, paper_type="管理", cover_info=cover)
    log(f"✅ 样稿: {docx_path}")

if __name__ == "__main__":
    main()
