from __future__ import annotations

import argparse
from pathlib import Path

from .config import OUTPUT_ROOT
from .runner import run_civil_batch, run_standard_batch


def main():
    parser = argparse.ArgumentParser(description="论文系统批量交付入口")
    parser.add_argument("--workflow", choices=["standard", "civil"], default="standard", help="批处理工作流")
    parser.add_argument("--input", "-i", default=str(OUTPUT_ROOT / "batch_titles.json"), help="标准批处理输入 JSON")
    parser.add_argument("--output-dir", "-o", default=str(OUTPUT_ROOT / "batch_run"), help="输出目录")
    parser.add_argument("--type", "-t", default="管理", choices=["管理", "设计", "机械"], help="论文类型")
    parser.add_argument("--limit", "-l", type=int, default=0, help="只处理前 N 条")
    parser.add_argument("--skip-existing", action="store_true", help="跳过已存在结果")
    parser.add_argument("--civil-args", nargs=argparse.REMAINDER, help="传给土木批处理脚本的额外参数")
    args = parser.parse_args()

    if args.workflow == "civil":
        run_civil_batch(args.civil_args or [])
        return

    run_standard_batch(args.input, args.output_dir, args.type, args.limit, args.skip_existing)


if __name__ == "__main__":
    main()

