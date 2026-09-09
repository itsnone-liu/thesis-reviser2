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
    image_quota_warning,
)


def _emit_image_quota_warning(img_dir, update):
    """生图额度耗尽时向任务进度推送警告（webapp 会把它写进任务消息，用户端可见）。"""
    try:
        if not img_dir:
            return
        info = image_quota_warning(img_dir)
        if info:
            update(f"⚠️ 生图额度已满（{info.get('reason', '')}）：图纸已用本地确定性工程图兜底，"
                   f"可稍后重跑本任务恢复GPT生图", 22)
    except Exception:
        pass


def _write_render_meta(docx_path, drawings_dir):
    """把本篇实际图纸目录记到 docx 旁，render_task 审计时据此读额度告警标记。"""
    try:
        if not drawings_dir:
            return
        meta_path = str(docx_path) + ".render_meta.json"
        with open(meta_path, "w", encoding="utf-8") as f:
            json.dump({"drawings_dir": drawings_dir}, f, ensure_ascii=False)
    except Exception:
        pass


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
        # 标签守卫先规范化可修复的标签；随后对“实际送入渲染器”的文本做硬审计。
        # 旧顺序先审计再修复，会把 draing/缺闭合等可确定修复的问题直接阻断，
        # 同时 tagguard 异常被吞掉会造成“看似继续、实际漏图”。
        try:
            from tagguard import audit_and_repair
            full_text, _g = audit_and_repair(full_text)
        except Exception as exc:
            raise RuntimeError(f"标签守卫失败，阻断渲染: {exc}") from exc
        try:
            from audit_txt import audit_txt as _audit_txt_body
            _r = _audit_txt_body(full_text)
            if _r["hard"]:
                raise ValueError(
                    "txt预渲染审计未通过(hard): " + "; ".join(_r["hard"]) +
                    " — 请先修复txt（audit_txt.py 报告）再渲染")
        except ImportError as exc:
            raise RuntimeError(f"TXT审计模块不可用，阻断渲染: {exc}") from exc
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
                new_imgs = generate_all_images(missing, drawing_folder, max_workers=1)
                drawing_images.update(new_imgs)
                update(f"补生完成，共 {len(drawing_images)} 张设计图", 20)
                _emit_image_quota_warning(drawing_folder, update)
                _write_render_meta(docx_path, drawing_folder)
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
                drawing_images = generate_all_images(drawings, img_dir, max_workers=1)
                update(f"已生成 {len(drawing_images)} 张设计图", 20)
                _emit_image_quota_warning(img_dir, update)
                _write_render_meta(docx_path, img_dir)
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
