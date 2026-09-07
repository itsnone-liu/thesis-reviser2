# -*- coding: utf-8 -*-
"""
renderer.py — 论文渲染管道（无LLM调用）
======================================
解析 TXT 中的标签，生成图表/图片，输出 DOCX。
渲染 = 标签解析 → matplotlib图表 + AI图片 → 组装DOCX（统一封面/分节符/页码）

支持类型：管理、设计、机械、土木。

使用：
  python renderer.py --input paper.txt --type 管理 -o paper.docx
  python renderer.py --input paper.txt --type 设计 -o paper.docx --cover cover.json --drawing-folder ./drawings
"""
import sys, os, re, json, argparse, shutil
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from core import (
    txt_to_docx_safe,
    extract_cover_info, extract_drawings_from_text,
    generate_all_images,
    extract_mech_json_blocks, strip_mech_json_blocks,
    merge_mech_json_into_drawings,
)


def render(txt_path: str, docx_path: str, paper_type: str = "管理",
           cover_info: dict = None, drawing_folder: str = None,
           update=None):
    """
    渲染 TXT → DOCX
    - txt_path: 输入的TXT路径（含<chart/>等标签）
    - docx_path: 输出的DOCX路径
    - paper_type: "管理"、"设计"、"机械" 或 "土木"
    - cover_info: 封面信息字典（可选）
    - drawing_folder: 设计图图片存储目录（可选，设计类用）
    - update: 进度回调
    """
    if update is None:
        def update(msg, prog): print(f"[{prog}%] {msg}")

    # 对设计类，检查是否已有图片或需要生成
    drawing_images = {}
    if paper_type in ("设计", "design", "sj", "机械", "mechanical", "mech", "mj", "土木", "civil", "cw"):
        # 读取并清理机械结构块，避免 [[MECH_JSON]] 泄漏到 DOCX，同时保留结构化字段
        update("正在读取并解析图纸标签...", 5)
        with open(txt_path, 'r', encoding='utf-8') as f:
            raw_text = f.read()
        mech_blocks = extract_mech_json_blocks(raw_text)
        full_text = strip_mech_json_blocks(raw_text)
        # 预渲染审计层：hard 问题（标签变体/LaTeX残留/裸行图注等）直接阻断，
        # 绝不带伤渲染（审计只报告不改写；改写是下面 tagguard 的职责）
        try:
            from audit_txt import audit_txt as _audit_txt_body
            _r = _audit_txt_body(full_text)
            if _r["hard"]:
                raise ValueError(
                    "txt预渲染审计未通过(hard): " + "; ".join(_r["hard"]) +
                    " — 请先修复txt（audit_txt.py 报告）再渲染")
        except ImportError:
            pass
        # Core's consistency gates consume the complete article, including
        # civil papers whose legacy type tuple did not forward source_text.
        os.environ["THESIS_SOURCE_TEXT"] = full_text
        # 标签守卫前置：先规范化标签（补缺括号/错拼等），生图与排版用同一份文本，
        # 否则生图提取器可能漏掉格式瑕疵的drawing → 排版时才发现缺图
        try:
            from tagguard import audit_and_repair
            full_text, _g = audit_and_repair(full_text)
        except Exception:
            pass
        drawings = extract_drawings_from_text(full_text)
        drawings = merge_mech_json_into_drawings(drawings, mech_blocks)

        if drawing_folder and os.path.isdir(drawing_folder):
            # 从已有图片目录加载（断点续跑：只加载已生成的）
            update("正在加载设计图纸...", 8)
            for d in drawings:
                # 尝试找匹配文件
                img_key = str(d.get("seq") or d.get("id"))
                img_name = f"drawing_{img_key}.png"
                img_path = os.path.join(drawing_folder, img_name)
                if os.path.exists(img_path):
                    drawing_images[img_key] = img_path
                else:
                    # 尝试模糊匹配
                    for fname in os.listdir(drawing_folder):
                        if img_key in fname or d['title'][:5] in fname:
                            drawing_images[img_key] = os.path.join(drawing_folder, fname)
                            break
            update(f"已加载 {len(drawing_images)} 张设计图", 12)
            # 补生缺失图纸：目录里没有的必须现场生成，绝不静默占位
            missing = [d for d in drawings
                       if str(d.get("seq") or d.get("id")) not in drawing_images]
            if missing:
                update(f"补生缺失图纸 {len(missing)} 张...", 13)
                new_imgs = generate_all_images(
                    missing, drawing_folder, max_workers=1,
                    source_text=full_text if paper_type in ("机械", "mechanical", "mech", "mj") else "",
                )
                drawing_images.update(new_imgs)
                update(f"补生完成，共 {len(drawing_images)} 张设计图", 20)
        else:
            # 需要生成图片
            if drawings:
                img_dir = drawing_folder or os.path.join(
                    os.path.dirname(docx_path) or '.',
                    f"drawings_{os.path.splitext(os.path.basename(docx_path))[0]}"
                )
                if not drawing_folder and os.path.isdir(img_dir):
                    shutil.rmtree(img_dir, ignore_errors=True)
                update(f"正在生成 {len(drawings)} 张设计图纸...", 15)
                drawing_images = generate_all_images(
                    drawings, img_dir, max_workers=1,
                    source_text=full_text if paper_type in ("机械", "mechanical", "mech", "mj") else "",
                )
                update(f"已生成 {len(drawing_images)} 张设计图", 20)
            else:
                update("未发现设计图纸标签", 10)

    # 调用 core 的 TXT→DOCX 主函数
    update("正在排版DOCX...", 30)
    txt_to_docx_safe(
        txt_path, docx_path,
        update=lambda m, p: update(m, 30 + int(p * 0.6)),
        cover_info=cover_info,
        drawing_images=drawing_images
    )

    update(f"渲染完成！DOCX 已保存: {docx_path}", 100)


def main():
    parser = argparse.ArgumentParser(description="论文渲染器 - TXT转DOCX（管理/设计/机械/土木）")
    parser.add_argument("--input", "-i", required=True, help="输入的TXT文件路径")
    parser.add_argument("--output", "-o", required=True, help="输出的DOCX文件路径")
    parser.add_argument("--type", "-t", default="管理", choices=["管理", "设计", "机械", "土木"],
                        help="论文类型：管理/设计/机械/土木")
    parser.add_argument("--cover", "-c", default="", help="封面信息JSON文件（可选）")
    parser.add_argument("--drawing-folder", "-d", default="",
                        help="设计图图片目录（可选，为空则自动生成）")
    args = parser.parse_args()

    if not os.path.isfile(args.input):
        print(f"错误：输入文件不存在 {args.input}")
        return

    cover_info = None
    if args.cover:
        if os.path.isfile(args.cover):
            with open(args.cover, 'r', encoding='utf-8') as f:
                cover_info = json.load(f)
            print(f"已加载封面信息: {args.cover}")
        else:
            print(f"警告：封面文件不存在 {args.cover}")
    else:
        stem = os.path.splitext(args.input)[0]
        candidates = [
            stem + ".cover.json",
            stem + "_cover.json",
            stem.replace(".txt", "") + "_cover.json",
            os.path.join(os.path.dirname(args.input), "cover.json"),
        ]
        for cand in candidates:
            if os.path.isfile(cand):
                with open(cand, 'r', encoding='utf-8') as f:
                    cover_info = json.load(f)
                print(f"已自动加载封面信息: {cand}")
                break

    drawing_folder = args.drawing_folder or None

    def update(msg, prog):
        print(f"[{prog}%] {msg}")

    render(args.input, args.output, args.type, cover_info, drawing_folder, update)


if __name__ == "__main__":
    main()
