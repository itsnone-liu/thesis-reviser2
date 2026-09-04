# -*- coding: utf-8 -*-
"""
audit_docx.py — DOCX 终点审计（方案F）
======================================
对照 TXT 标签数与 DOCX 实际产出（图片数/表格数/空白图），批完即刻对账，
不用人工抽查"图到底进没进文档"。

用法:
  python audit_docx.py paper.txt paper.docx          # 单篇审计
  python audit_docx.py --dir 输出目录/               # 目录内 *.txt+*.docx 批量审计
  python audit_docx.py --dir 目录/ --csv report.csv  # 批量并导出CSV
退出码: 0=全部通过, 1=存在缺口/空白图
"""
import sys, os, re, csv, json, zipfile, argparse, tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from core import tolerant_extract_charts, tolerant_extract_tables, extract_drawings_from_text

try:
    from tagguard import audit_and_repair as _guard_repair
except Exception:
    _guard_repair = None

from PIL import Image


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


def audit_pair(txt_path: str, docx_path: str) -> dict:
    """审计一对 txt/docx，返回对账dict"""
    row = {
        "txt": os.path.basename(txt_path),
        "docx": os.path.basename(docx_path),
        "charts_in_txt": 0, "tables_in_txt": 0, "drawings_in_txt": 0,
        "images_in_docx": 0, "tables_in_docx": 0, "blank_images": 0,
        "receipt": "", "verdict": "✅", "problems": "",
    }
    if not os.path.exists(docx_path):
        row["verdict"] = "❌"
        row["problems"] = "DOCX不存在"
        return row

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

    problems = []
    expect_imgs = row["drawings_in_txt"] + row["charts_in_txt"]
    # 占位图也计入 images_in_docx，所以数量上应 ≥ 期望；缺口=期望>实际
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
    ap = argparse.ArgumentParser(description="DOCX终点审计")
    ap.add_argument("txt", nargs="?", help="论文TXT路径")
    ap.add_argument("docx", nargs="?", help="论文DOCX路径")
    ap.add_argument("--dir", dest="directory", help="批量审计目录")
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
        r = audit_pair(txt, docx)
        rows.append(r)
        if r["verdict"] != "✅":
            bad += 1
        print(f"{r['verdict']} {r['txt'][:40]:<42} "
              f"图{r['drawings_in_txt']}/{r['images_in_docx']} "
              f"表{r['tables_in_txt']}/{r['tables_in_docx']} "
              f"空白{r['blank_images']} {r['receipt']} {r['problems']}")

    print(f"\n共{len(rows)}篇: ✅{len(rows) - bad} ❌{bad}")
    if args.csv_out and rows:
        with open(args.csv_out, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
            w.writeheader()
            w.writerows(rows)
        print(f"CSV已导出: {args.csv_out}")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
