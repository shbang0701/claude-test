# -*- coding: utf-8 -*-
"""아이콘 전체 미리보기 (검수용)."""
import sys, io
import cairosvg
from PIL import Image, ImageDraw, ImageFont
from icons import ICONS, CATS

def render(ic, px=96, color="#1F3A5F"):
    png = cairosvg.svg2png(bytestring=ic.svg(color).encode(), output_width=px, output_height=px)
    return Image.open(io.BytesIO(png)).convert("RGBA")

def sheet(path, px=96, cols=10):
    font = ImageFont.truetype("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc", 13)
    cell_w, cell_h = px + 44, px + 50
    small = ImageFont.truetype("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc", 11)
    rows = (len(ICONS) + cols - 1) // cols
    im = Image.new("RGB", (cols * cell_w, rows * cell_h), "white")
    d = ImageDraw.Draw(im)
    for k, ic in enumerate(ICONS):
        x, y = (k % cols) * cell_w, (k // cols) * cell_h
        d.rectangle([x + 20, y + 6, x + 20 + px, y + 6 + px], outline="#E4E9EF")
        g = render(ic, px)
        im.paste(g, (x + 20, y + 6), g)
        d.text((x + 6, y + px + 10), ic.name[:10], fill="#1E2A36", font=font)
        d.text((x + 6, y + px + 28), f"{k+1:02d} · {ic.key}", fill="#8C99A8", font=small)
    im.save(path)
    print(len(ICONS), "icons")

if __name__ == "__main__":
    sheet(sys.argv[1])
