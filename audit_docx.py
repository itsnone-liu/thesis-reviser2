# -*- coding: utf-8 -*-
"""
audit_docx.py — 论文自动终审（原"DOCX终点审计"全面升级）
======================================================
不依赖人工抽查，机器自动发现：图没进文档、表残缺、空图、占位图、
数值冲突、图号引用悬空、章节字数不足、参考文献不足、百分比合计错、
图注编号跳号。批完即刻对账，❌即人工只需看一眼报告。

用法:
  python audit_docx.py paper.txt paper.docx                 # 单篇审计
  python audit_docx.py --dir 输出目录/ --type 设计          # 目录批量审计
  python audit_docx.py --dir 目录/ --csv report.csv         # 批量并导出CSV
退出码: 0=通过(可有⚠️警告), 1=存在❌阻断性问题
"""
import sys, os, re, csv, json, zipfile, argparse, tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from core import tolerant_extract_charts, tolerant_extract_tables, extract_drawings_from_text

try:
    from tagguard import audit_and_repair as _guard_repair
except Exception:
    _guard_repair = None

from PIL import Image

# 各类型每章最低字数（低于即⚠️警告）
_MIN_CHAPTER_WORDS = {"管理": 800, "设计": 900, "机械": 800, "土木": 800}


def _near_blank(path: str) -> bool:
    """更严的空白判定（≥99%近白）——只抓真正的白板图，
    避免把留白较多的正常 matplotlib 图表误报为空白。"""
    try:
        img = Image.open(path).convert("RGB")
        px = img.getdata()
        total = img.size[0] * img.size[1] or 1
        white = sum(1 for r, g, b in px if r >= 248 and g >= 248 and b >= 248)
        return white / total >= 0.99
    except Exception:
        return False


def _check_text_quality(text: str, warns: list, paper_type: str):
    """TXT正文级检查：引用悬空/章节字数/参考文献/摘要关键词/百分比合计"""
    # 1) 字面占位引用残留（守卫已修则此处为0；>0说明有漏网）
    literal_refs = re.findall(r'[如见由][图表][Xx×?？]', text)
    if literal_refs:
        warns.append(f"正文{len(literal_refs)}处未填编号的图/表占位引用")

    # 2) 图号引用超出实际图数（"见图7"但全文只有5张图）
    n_fig_total = len(re.findall(r'<drawing\b', text)) + len(re.findall(r'<chart\b', text))
    fig_refs = [int(m.group(1)) for m in re.finditer(r'[如见由]图(\d+)(?![-\d])', text)]
    if fig_refs and n_fig_total:
        over = sorted(set(r for r in fig_refs if r > n_fig_total))
        if over:
            warns.append(f"图号引用超界{over} vs 实际图数{n_fig_total}")

    # 3) 章节字数（先剔除目录区，避免目录条目被当章节体误报）
    body = text
    m_toc = re.search(r'^目录\s*$', body, re.M)
    if m_toc:
        pb = body.find('---PAGE_BREAK---', m_toc.end())
        if pb >= 0:
            body = body[:m_toc.start()] + body[pb:]
    parts = re.split(r'(?m)^(第[一二三四五六七八九十\d]+章[^\n。！？]{0,28})$', body)
    minw = _MIN_CHAPTER_WORDS.get(paper_type, 800)
    for i in range(1, len(parts) - 1, 2):
        heading, chbody = parts[i], parts[i + 1]
        wc = len(re.sub(r'<[^>]+>|---PAGE_BREAK---|\s', '', chbody))
        # 绪论/结论/总结类章节本身较短，阈值放宽
        th = 500 if re.search(r'绪论|引言|结论|总结|展望|参考文献', heading) else minw
        if wc < th:
            warns.append(f"{heading[:10]}仅{wc}字(<{th})")

    # 4) 摘要/关键词/参考文献
    if not re.search(r'^摘要', text, re.M):
        warns.append("缺少摘要")
    if not re.search(r'^关键词', text, re.M):
        warns.append("缺少关键词")
    if '参考文献' in text:
        refs_sec = text.split('参考文献')[-1]
        n_refs = len(re.findall(r'\[\d+\]', refs_sec))
        if n_refs and n_refs < 10:
            warns.append(f"参考文献仅{n_refs}条(<10)")
        elif not n_refs:
            warns.append("参考文献无[1][2]编号条目")

    # 5) 百分比列合计（data列对齐存在两种惯例：含行名列/不含；
    #    分组表按连续段每段≈100%判定，两种对齐任一合法即放行）
    for m in re.finditer(r'<table\b[^>]*?header="([^"]*)"(?:[^>]*?rows="([^"]*)")?[^>]*?data="([^"]*)"', text):
        header, rows_attr, data = m.group(1), m.group(2) or "", m.group(3) or ""
        headers = [h.strip() for h in header.split(',')]
        data_rows = [re.split(r'[|｜]+', r2) for r2 in re.split(r'[;；]+', data) if r2.strip()]
        if not data or len(headers) < 2:
            continue

        def _try_offset(offset):
            """按data列偏移offset对齐header[offset:]，返回(列名, 列值)或None"""
            eff = headers[offset:]
            pct_idx = [i for i, h in enumerate(eff) if '%' in h or '占比' in h or '比例' in h]
            if not pct_idx:
                return None
            ci = pct_idx[0]
            vals = []
            for cells in data_rows:
                if ci < len(cells):
                    mm = re.search(r'([\d.]+)', cells[ci])
                    if mm and float(mm.group(1)) <= 100:
                        vals.append(float(mm.group(1)))
            if len(vals) < 3:
                return None
            return eff[ci], vals

        def _groupable(vals):
            n = len(vals)
            ok = [False] * (n + 1)
            ok[0] = True
            for i in range(1, n + 1):
                for j in range(i):
                    if ok[j] and 90 <= sum(vals[j:i]) <= 110:
                        ok[i] = True
                        break
            return ok[n]

        tried = []   # (列名, 列值) 两种对齐都试
        for offset in (0, 1):
            r = _try_offset(offset)
            if r:
                tried.append(r)
                if _groupable(r[1]):
                    break
        else:
            # 两种对齐都无法按100%分组 → 报最接近的那次
            if tried:
                name, vals = tried[0]
                warns.append(f"表\"{name[:8]}\"占比列无法按100%分组(合计{sum(vals):.0f}%)")


def audit_pair(txt_path: str, docx_path: str, paper_type: str = "管理") -> dict:
    """审计一对 txt/docx，返回对账dict"""
    row = {
        "txt": os.path.basename(txt_path),
        "docx": os.path.basename(docx_path),
        "charts_in_txt": 0, "tables_in_txt": 0, "drawings_in_txt": 0,
        "images_in_docx": 0, "tables_in_docx": 0, "blank_images": 0,
        "receipt": "", "consistency": "", "verdict": "✅", "problems": "",
    }
    if not os.path.exists(docx_path):
        row["verdict"] = "❌"
        row["problems"] = "DOCX不存在"
        return row

    warns = []
    with open(txt_path, "r", encoding="utf-8") as f:
        text = f.read()
    # 按守卫修复后的口径计数（错拼/坏标签在渲染前已被归一化）
    if _guard_repair is not None:
        try:
            text, _ = _guard_repair(text)
        except Exception:
            pass
    row["charts_in_txt"] = len(tolerant_extract_charts(text))
    row["tables_in_txt"] = len(tolerant_extract_tables(text))
    row["drawings_in_txt"] = len(extract_drawings_from_text(text))

    with zipfile.ZipFile(docx_path) as z:
        names = z.namelist()
        media = [n for n in names if n.startswith("word/media/")]
        row["images_in_docx"] = len(media)
        with z.open("word/document.xml") as f:
            docxml = f.read().decode("utf-8", errors="ignore")
        row["tables_in_docx"] = docxml.count("<w:tbl>")

        # 渲染回执（若有，直接引用其失败明细）
        receipt_path = docx_path + ".report.json"
        if os.path.exists(receipt_path):
            try:
                with open(receipt_path, "r", encoding="utf-8") as f:
                    rc = json.load(f)
                c_, t_, d_ = rc["charts"], rc["tables"], rc["drawings"]
                row["receipt"] = (f"chart {c_['ok']}/{c_['total']}"
                                  f" table {t_['ok']}/{t_['total']}"
                                  f" drawing {d_['ok']}/{d_['total']}")
                if c_["failed"] or t_["failed"]:
                    warns.append(f"回执降级/失败chart{len(c_['failed'])}·table{len(t_['failed'])}")
                if d_["missing_image"]:
                    warns.append(f"占位图{len(d_['missing_image'])}张")
                cons = rc.get("consistency")
                if cons and cons.get("conflicts"):
                    row["consistency"] = "; ".join(
                        f"{c['name']}:{c['values'][:40]}" for c in cons["conflicts"])
                    unfixed = sum(1 for c in cons["conflicts"] if not c.get("fixed"))
                    if unfixed:
                        warns.append(f"数值冲突未修{unfixed}")
            except Exception:
                pass

        # 空白图检测（≥99%近白的真白板才报警）
        for m in media:
            suffix = os.path.splitext(m)[1].lower()
            if suffix not in (".png", ".jpg", ".jpeg"):
                continue
            try:
                with tempfile.NamedTemporaryFile(suffix=suffix, delete=False) as tf:
                    tf.write(z.read(m))
                    tmp = tf.name
                if _near_blank(tmp):
                    row["blank_images"] += 1
                os.unlink(tmp)
            except Exception:
                pass

        # 图注/表注编号连续性（图1,图2...不跳号不重复）
        for kind in ("图", "表"):
            caps = sorted(set(int(m.group(1)) for m in re.finditer(
                rf'>{kind}(\d+)[ 　]', docxml)))
            if caps and caps != list(range(1, len(caps) + 1)):
                warns.append(f"{kind}注编号不连续{caps[:10]}")

    # 正文级检查
    _check_text_quality(text, warns, paper_type)

    problems = []
    expect_imgs = row["drawings_in_txt"] + row["charts_in_txt"]
    if expect_imgs > row["images_in_docx"]:
        problems.append(f"图缺{expect_imgs - row['images_in_docx']}张"
                        f"(期望{expect_imgs}/实际{row['images_in_docx']})")
    if row["tables_in_txt"] > row["tables_in_docx"]:
        problems.append(f"表缺{row['tables_in_txt'] - row['tables_in_docx']}个")
    if row["blank_images"]:
        problems.append(f"空白图{row['blank_images']}张")
    if problems:
        row["verdict"] = "❌"
        row["problems"] = "; ".join(problems)
    if warns:
        row["problems"] = (row["problems"] + "; " if row["problems"] else "") + \
            "⚠️" + "; ⚠️".join(warns[:6])
    return row


def collect_pairs(directory: str):
    pairs = []
    for fn in sorted(os.listdir(directory)):
        if fn.endswith(".txt"):
            docx = os.path.join(directory, fn[:-4] + ".docx")
            txt = os.path.join(directory, fn)
            if os.path.exists(docx):
                pairs.append((txt, docx))
            else:
                pairs.append((txt, txt[:-4] + ".docx"))  # 审计会标记DOCX不存在
    return pairs


def main():
    ap = argparse.ArgumentParser(description="论文自动终审")
    ap.add_argument("txt", nargs="?", help="论文TXT路径")
    ap.add_argument("docx", nargs="?", help="论文DOCX路径")
    ap.add_argument("--dir", dest="directory", help="批量审计目录")
    ap.add_argument("--type", dest="paper_type", default="管理",
                    help="论文类型(管理/设计/机械/土木)，影响章节字数阈值")
    ap.add_argument("--csv", dest="csv_out", help="导出CSV路径")
    args = ap.parse_args()

    if args.directory:
        pairs = collect_pairs(args.directory)
        if not pairs:
            print("目录中未找到 txt/docx 对")
            return 1
    elif args.txt and args.docx:
        pairs = [(args.txt, args.docx)]
    else:
        print(__doc__)
        return 1

    rows = []
    bad = 0
    for txt, docx in pairs:
        r = audit_pair(txt, docx, args.paper_type)
        rows.append(r)
        if r["verdict"] == "❌":
            bad += 1
        print(f"{r['verdict']} {r['txt'][:36]:<38} "
              f"图{r['drawings_in_txt']}/{r['images_in_docx']} "
              f"表{r['tables_in_txt']}/{r['tables_in_docx']} "
              f"空白{r['blank_images']} {r['receipt']} {r['problems'][:160]}")

    n_warn = sum(1 for r in rows if "⚠️" in r["problems"])
    print(f"\n共{len(rows)}篇: ✅{len(rows) - bad - n_warn} ⚠️{n_warn} ❌{bad}"
          f"{' ← 存在阻断性问题，需人工处理' if bad else ''}")
    if args.csv_out and rows:
        with open(args.csv_out, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
            w.writeheader()
            w.writerows(rows)
        print(f"CSV已导出: {args.csv_out}")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
