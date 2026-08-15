# -*- coding: utf-8 -*-
"""Convert legacy .ppt to .pptx via PowerPoint COM (Windows only)."""

import os
import sys


def convert_ppt_to_pptx(src: str, dst: str) -> None:
    import win32com.client

    src = os.path.abspath(src)
    dst = os.path.abspath(dst)
    if not os.path.exists(src):
        raise FileNotFoundError(src)

    pp = win32com.client.Dispatch("PowerPoint.Application")
    pp.Visible = 0
    try:
        pres = pp.Presentations.Open(src, WithWindow=False)
        # 24 = ppSaveAsOpenXMLPresentation
        pres.SaveAs(dst, 24)
        pres.Close()
    finally:
        pp.Quit()


def main():
    if len(sys.argv) < 3:
        print("用法: python convert_ppt_to_pptx.py <输入.ppt> <输出.pptx>")
        sys.exit(1)
    convert_ppt_to_pptx(sys.argv[1], sys.argv[2])
    print(f"Converted -> {sys.argv[2]}")


if __name__ == "__main__":
    main()
