# -*- coding: utf-8 -*-
"""아이콘 파일 출력: SVG(남색) + PNG 256px(남색·흰색) + 목록."""
import os, re, io, sys
import cairosvg
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from icons import ICONS, CATS

ROOT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "03_아이콘")
NAVY = "#1F3A5F"


def fname(k, ic):
    nm = re.sub(r"[()·,]", "", ic.name).strip()
    nm = re.sub(r"\s+", "_", nm)
    return f"{k:02d}_{nm}_{ic.key}"


def main():
    for sub in ("svg", "png_남색", "png_흰색"):
        os.makedirs(os.path.join(ROOT, sub), exist_ok=True)
    rows = []
    for k, ic in enumerate(ICONS, 1):
        fn = fname(k, ic)
        svg = ic.svg(NAVY)
        with open(os.path.join(ROOT, "svg", fn + ".svg"), "w", encoding="utf-8") as f:
            f.write(svg)
        cairosvg.svg2png(bytestring=svg.encode(), write_to=os.path.join(ROOT, "png_남색", fn + ".png"),
                         output_width=256, output_height=256)
        cairosvg.svg2png(bytestring=ic.svg("#FFFFFF").encode(), write_to=os.path.join(ROOT, "png_흰색", fn + ".png"),
                         output_width=256, output_height=256)
        rows.append((k, ic.name, ic.key, dict(CATS)[ic.cat]))
    print(len(rows), "icons →", ROOT)
    return rows


if __name__ == "__main__":
    main()
