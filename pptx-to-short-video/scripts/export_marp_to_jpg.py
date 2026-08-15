# -*- coding: utf-8 -*-
"""
将 Marp Markdown 每页导出为单独 JPG，保存到 img_upload/ 供剪映使用。

前置：需安装 Chrome/Chromium/Edge/Firefox（Marp CLI 依赖）

用法：
  python export_marp_to_jpg.py [Markdown路径]
  默认查找：技能目录或 10-Marp_slides 下的 short-video-slides.md
"""

import subprocess
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPT_DIR.parent


def find_md() -> Path | None:
    for d in (SKILL_DIR, SKILL_DIR.parent.parent / "10-Marp_slides"):
        p = d / "short-video-slides.md"
        if p.is_file():
            return p
    return None


def main():
    md_path = Path(sys.argv[1]) if len(sys.argv) > 1 else find_md()
    if not md_path or not md_path.is_file():
        print("未找到 short-video-slides.md，请指定路径：")
        print("  python export_marp_to_jpg.py <Markdown路径>")
        sys.exit(1)

    out_dir = md_path.parent / "img_upload"
    out_dir.mkdir(parents=True, exist_ok=True)

    try:
        subprocess.run(
            ["npx", "-y", "@marp-team/marp-cli", str(md_path.name), "--images", "--output", str(out_dir), "--allow-local-files"],
            check=True,
            cwd=md_path.parent,
            timeout=180,
        )
        print(f"已导出到: {out_dir}")
    except (subprocess.CalledProcessError, FileNotFoundError, subprocess.TimeoutExpired) as e:
        print("导出失败。可手动操作：")
        print("  1. 在 Cursor 中打开 short-video-slides.md")
        print("  2. 使用 Marp 扩展「Export slide deck」→ 选择 JPG")
        print("  3. 将导出的图片放入 img_upload/")
        print(f"错误: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
