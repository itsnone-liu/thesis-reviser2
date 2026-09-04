# -*- coding: utf-8 -*-
"""
batch_25.py — 25年题目不一致批次：按表格生成50篇论文
=================================================
数据源: 生成样稿/25年论文题目不一致筛选结果(1).xlsx
要求:
  1) 题目按表格 LWTM 原文生成
  2) 关键词覆盖为表格 LWGJC（分号统一为中文；）
  3) 封面用表格信息（姓名/学号/专业/指导教师/题目，层次专升本，日期2025年12月16日）
特性: 断点续跑(完成即跳过) / 429配额自愈(休眠15分钟重试) / 每篇自动终审→CSV
用法: cd thesis-reviser && . ./.env && python3 batch_25.py [--limit N] [--only 学号,...] [--dry]
"""
import sys, os, re, json, time, csv, argparse, traceback

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

XLSX = "/root/project/workspace/生成样稿/25年论文题目不一致筛选结果(1).xlsx"
OUT_DIR = "/root/project/workspace/25年批量论文"
STATE = os.path.join(OUT_DIR, "_state.json")
CSV_PATH = os.path.join(OUT_DIR, "批量终审.csv")

TYPE_MAP = {
    "机械工程": "机械", "土木工程": "土木", "环境设计": "设计",
    "财务管理": "管理", "人力资源管理": "管理", "旅游管理": "管理",
}

QUOTA_MARKS = ("429", "1308", "Too Many Requests", "rate limit", "RateLimitError")


def load_students():
    import openpyxl
    wb = openpyxl.load_workbook(XLSX)
    ws = wb["25题目不一致名单"]
    out = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r[0]:
            continue
        out.append({
            "name": str(r[0]).strip(), "sid": str(r[6]).strip(),
            "major": str(r[8]).strip(), "degree": str(r[10] or "").strip(),
            "advisor": str(r[12] or "").strip(),
            "title": str(r[14] or "").strip(),
            "keywords_raw": str(r[15] or "").strip(),
            "date": str(r[11] or "20251216").strip(),
        })
    return out


def norm_keywords(kw: str) -> str:
    terms = [t.strip() for t in re.split(r"[;；]", kw) if t.strip()]
    return "；".join(terms)


def extract_target(title: str) -> str:
    m = re.search(r'以(.+?)(?:项目|工程)?为例', title)
    if m:
        t = m.group(1).strip()
    elif re.search(r'——(.+)$', title):
        t = re.search(r'——(.+)$', title).group(1).strip()
    else:
        t = title
    # 剥掉研究性后缀，让"设计对象"是对象本身
    t = re.sub(r'(的设计与实现|的设计|应用实践|的应用|设计与装配技术研究|分析与优化设计研究|的研究)+$', '', t)
    return t or title


def cover_of(st: dict) -> dict:
    d = st["date"]
    return {
        "title": st["title"], "name": st["name"], "student_id": st["sid"],
        "major": st["major"], "level": "专升本", "advisor": st["advisor"],
        "year": d[:4] if len(d) >= 4 else "2025",
        "month": d[4:6].lstrip("0") if len(d) >= 6 else "12",
        "day": d[6:8].lstrip("0") if len(d) >= 8 else "16",
    }


def override_keywords(text: str, kw_norm: str) -> str:
    """生成文本的'关键词：'行 → 表格关键词（保证与学位系统一致）"""
    new_line = f"关键词：{kw_norm}"
    if re.search(r'(?m)^关键词[：:]', text):
        return re.sub(r'(?m)^关键词[：:].*$', new_line, text, count=1)
    # 没有关键词行 → 插在摘要块后
    return re.sub(r'(?m)^(\s*关键词)', new_line, text, count=1) if re.search(
        r'(?m)^\s*关键词', text) else text.replace('摘要\n', '摘要\n', 1)


def is_quota_error(e: Exception) -> bool:
    msg = str(e)
    return any(k in msg for k in QUOTA_MARKS)


def load_state():
    if os.path.exists(STATE):
        with open(STATE, encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_state(st):
    with open(STATE, "w", encoding="utf-8") as f:
        json.dump(st, f, ensure_ascii=False, indent=1)


def run_one(st: dict, idx: int, total: int):
    from profile import generate_profile
    from generator import generate
    import renderer
    from audit_docx import audit_pair

    ptype = TYPE_MAP[st["major"]]
    base = os.path.join(OUT_DIR, f"{st['sid']}_{st['name']}_{st['major']}")
    txt_path, docx_path = base + ".txt", base + ".docx"
    kw_norm = norm_keywords(st["keywords_raw"])
    tag = f"[{idx+1}/{total} {st['name']}·{st['major']}·{ptype}]"

    # 已完成且终审通过(无占位/缺图) → 跳过
    if os.path.exists(docx_path) and os.path.exists(txt_path):
        row = audit_pair(txt_path, docx_path, ptype)
        if row["verdict"] == "✅" and "占位图" not in row["problems"] and "图缺" not in row["problems"]:
            print(f"{tag} 已完成，跳过 ({row['verdict']})", flush=True)
            return {"status": "done", "verdict": row["verdict"]}

    print(f"{tag} 题目: {st['title']}", flush=True)
    print(f"{tag} 关键词(表格): {kw_norm}", flush=True)

    # 1) 画像 + 2) 生成全文（已有txt则复用，只重渲染——省LLM配额）
    target = extract_target(st["title"])
    if os.path.exists(txt_path) and os.path.getsize(txt_path) > 5000:
        full_text = open(txt_path, encoding="utf-8").read()
        print(f"{tag} 复用已生成txt {len(full_text)}字", flush=True)
    else:
        # 土木类偏工程对象(如公路测量)时，给图纸强提示，防止LLM一张图都不发
        context = "论文正文必须包含图纸标签(总平面/纵断面/横断面/结构构造等)" if ptype == "土木" else ""
        if ptype == "设计":
            profile = generate_profile(ptype, st["title"], target, major=st["major"])
        else:
            profile = generate_profile(ptype, st["title"], target, major=st["major"], context=context)
        full_text = generate(profile, ptype, update=lambda m, p: None)
        # 图纸下限：设计类管线(机械/土木/设计)图纸<4张 → 带强提示重生成一次
        import re as _re
        if ptype in ("机械", "土木", "设计") and len(_re.findall(r'<drawing\b', full_text)) < 4:
            print(f"{tag} 图纸仅{len(_re.findall(r'<drawing', full_text))}张(<4)，带图纸要求重生成...", flush=True)
            profile["context"] = (profile.get("context", "") +
                                  "；论文必须包含至少4个<drawing>图纸标签(含原理图/布置图/结构图)")
            full_text = generate(profile, ptype, update=lambda m, p: None)
        # 3) 关键词覆盖
        full_text = override_keywords(full_text, kw_norm)
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write(full_text)
        print(f"{tag} 生成完成 {len(full_text)}字, 关键词行已覆盖", flush=True)

    # 4) 渲染（封面=表格信息；图纸目录独立缓存）
    drawing_folder = os.path.join(OUT_DIR, f"drawings_{st['sid']}")
    renderer.render(txt_path, docx_path, paper_type=ptype,
                    cover_info=cover_of(st), drawing_folder=drawing_folder,
                    update=lambda m, p: None)

    # 5) 终审
    row = audit_pair(txt_path, docx_path, ptype)
    print(f"{tag} 终审 {row['verdict']} 图{row['drawings_in_txt']}/{row['images_in_docx']} "
          f"表{row['tables_in_txt']}/{row['tables_in_docx']} {row['problems'][:120]}", flush=True)
    return {"status": "done", "verdict": row["verdict"], "problems": row["problems"][:200]}


def append_csv(st, result):
    new = not os.path.exists(CSV_PATH)
    with open(CSV_PATH, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["学号", "姓名", "专业", "类型", "题目", "关键词", "状态", "终审", "问题"])
        w.writerow([st["sid"], st["name"], st["major"], TYPE_MAP[st["major"]],
                    st["title"], norm_keywords(st["keywords_raw"]),
                    result.get("status", ""), result.get("verdict", ""),
                    result.get("problems", "")])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=0, help="只跑前N篇(测试)")
    ap.add_argument("--only", default="", help="只跑指定学号(逗号分隔)")
    ap.add_argument("--retry-failed", action="store_true", help="重试之前失败的")
    args = ap.parse_args()

    os.makedirs(OUT_DIR, exist_ok=True)
    students = load_students()
    state = load_state()

    # 试点先行：四类各取第1篇最先跑，验证管线，再跑其余
    by_type = {}
    for i, s in enumerate(students):
        t = TYPE_MAP[s["major"]]
        by_type.setdefault(t, []).append(i)
    pilots = [idxs[0] for idxs in by_type.values()]
    order = pilots + [i for i in range(len(students)) if i not in pilots]

    if args.only:
        keep = set(x.strip() for x in args.only.split(",") if x.strip())
        order = [i for i, s in enumerate(students) if s["sid"] in keep]
    if args.limit:
        order = order[:args.limit]

    print(f"批量任务: 共{len(students)}人, 本轮执行{len(order)}人 "
          f"(试点先行: {[students[i]['name'] for i in order[:len(pilots)]]})", flush=True)

    done = fail = 0
    for i in order:
        st = students[i]
        key = st["sid"]
        if not args.retry_failed and state.get(key, {}).get("status") == "done":
            done += 1
            continue
        attempts = 0
        while True:
            attempts += 1
            try:
                result = run_one(st, i, len(students))
                state[key] = result
                save_state(state)
                append_csv(st, result)
                done += 1
                break
            except KeyboardInterrupt:
                print("人工中断", flush=True)
                return 1
            except Exception as e:
                if is_quota_error(e):
                    print(f"[{st['name']}] 配额受限(429)，休眠15分钟后重试...", flush=True)
                    time.sleep(15 * 60)
                    attempts -= 1  # 配额等待不计入失败次数
                    continue
                print(f"[{st['name']}] 第{attempts}次失败: {e}", flush=True)
                traceback.print_exc()
                if attempts >= 3:
                    state[key] = {"status": "failed", "error": str(e)[:300]}
                    save_state(state)
                    append_csv(st, {"status": "failed", "problems": str(e)[:200]})
                    fail += 1
                    break
                time.sleep(30)
    print(f"\n本轮结束: 成功{done} 失败{fail} (输出: {OUT_DIR})", flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
