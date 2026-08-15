# -*- coding: utf-8 -*-
"""
Clone slides from source PPT into target via PowerPoint COM (Windows).
Preserves GROUP/LINE composite diagrams that python-pptx cannot.

Usage:
  python clone_slides_com.py --target work.pptx --source source.pptx \
    --pages 30,51 --after-title "价值交付的全生命周期"
"""

import argparse
import os
import sys


def norm(s):
    import re
    s = re.sub(r"\s+", "", s or "")
    return s.lower()


def get_slide_title(slide):
    for i in range(1, slide.Shapes.Count + 1):
        sh = slide.Shapes(i)
        if sh.HasTextFrame and sh.TextFrame.HasText:
            t = sh.TextFrame.TextRange.Text.strip().split("\n")[0]
            if t and sh.Top < 900000:
                return t
    return ""


def find_slide_by_title(pres, title):
    for i in range(1, pres.Slides.Count + 1):
        t = get_slide_title(pres.Slides(i))
        if norm(title) == norm(t) or title in t:
            return i
    return None


def clone_slides(target_path, source_path, pages, after_title=None, after_index=None):
    import win32com.client

    target_path = os.path.abspath(target_path)
    source_path = os.path.abspath(source_path)

    pp = win32com.client.Dispatch("PowerPoint.Application")
    pp.Visible = 1
    target = pp.Presentations.Open(target_path, WithWindow=False)
    source = pp.Presentations.Open(source_path, WithWindow=False)

    if after_index is None:
        after_index = find_slide_by_title(target, after_title or "")
        if after_index is None:
            after_index = target.Slides.Count
    insert_at = after_index + 1

    cloned = []
    for p in pages:
        if p < 1 or p > source.Slides.Count:
            print(f"Skip invalid page {p}")
            continue
        source.Slides(p).Copy()
        target.Slides.Paste(Index=insert_at)
        cloned.append({"source_page": p, "target_page": insert_at})
        insert_at += 1

    target.Save()
    source.Close()
    target.Close()
    pp.Quit()
    return cloned


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--target", required=True)
    parser.add_argument("--source", required=True)
    parser.add_argument("--pages", required=True, help="comma-separated 1-based pages")
    parser.add_argument("--after-title", default=None)
    parser.add_argument("--after-index", type=int, default=None)
    args = parser.parse_args()

    if sys.platform != "win32":
        print("COM clone requires Windows + PowerPoint")
        sys.exit(1)

    pages = [int(x.strip()) for x in args.pages.split(",") if x.strip()]
    result = clone_slides(args.target, args.source, pages, args.after_title, args.after_index)
    for r in result:
        print(f"Cloned source p{r['source_page']} -> target p{r['target_page']}")
    print("Done. Update titles/notes manually or run revise_from_spec --phase notes")


if __name__ == "__main__":
    main()
