# -*- coding: utf-8 -*-
"""문서에 들어갈 자리표시 이미지(화면 캡처 자리 등)를 만든다."""
import os, sys
from PIL import Image, ImageDraw, ImageFont
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import design as D

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "assets")
FONT = "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc"

def rgb(h): return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

def dashed_rect(d, box, color, width=3, dash=18, gap=12):
    x0, y0, x1, y1 = box
    def run(a, b, horiz, fixed):
        x = a
        while x < b:
            x2 = min(x + dash, b)
            if horiz: d.line([(x, fixed), (x2, fixed)], fill=color, width=width)
            else:     d.line([(fixed, x), (fixed, x2)], fill=color, width=width)
            x += dash + gap
    run(x0, x1, True, y0); run(x0, x1, True, y1)
    run(y0, y1, False, x0); run(y0, y1, False, x1)

def center(d, text, font, cx, y, fill):
    w = d.textbbox((0, 0), text, font=font)[2]
    d.text((cx - w / 2, y), text, font=font, fill=fill)

def placeholder(path, w, h, title, hint, chrome=False):
    im = Image.new("RGB", (w, h), rgb(D.SURFACE))
    d = ImageDraw.Draw(im)
    m = int(h * 0.035)
    if chrome:
        bar = int(h * 0.09)
        d.rectangle([m, m, w - m, m + bar], fill=rgb("E8ECF0"))
        for i, c in enumerate(("C9D2DA", "C9D2DA", "C9D2DA")):
            cx = m + bar * 0.55 + i * bar * 0.55
            r = bar * 0.16
            d.ellipse([cx - r, m + bar / 2 - r, cx + r, m + bar / 2 + r], fill=rgb(c))
        d.rectangle([m, m, w - m, h - m], outline=rgb("D5DBE1"), width=2)
        dashed_rect(d, (m + bar + m // 2, m + bar + m // 2, w - m - m // 2, h - m - m // 2), rgb("C6D0D9"), 2)
        ty = m + bar
    else:
        dashed_rect(d, (m, m, w - m, h - m), rgb("C6D0D9"), 3)
        ty = 0
    f1 = ImageFont.truetype(FONT, int(h * 0.075))
    f2 = ImageFont.truetype(FONT, int(h * 0.045))
    cy = ty + (h - ty) / 2
    center(d, title, f1, w / 2, cy - h * 0.075, rgb("6B7683"))
    center(d, hint,  f2, w / 2, cy + h * 0.025, rgb("97A1AC"))
    im.save(path, dpi=(220, 220))
    print("  ", os.path.basename(path), im.size)

def logo(path, w=660, h=170):
    im = Image.new("RGB", (w, h), rgb("FFFFFF"))
    d = ImageDraw.Draw(im)
    dashed_rect(d, (3, 3, w - 3, h - 3), rgb("C6D0D9"), 2, 14, 10)
    f = ImageFont.truetype(FONT, 40)
    center(d, "로고 자리", f, w / 2, h / 2 - 26, rgb("8A94A0"))
    im.save(path, dpi=(220, 220))
    print("  ", os.path.basename(path), im.size)

if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    print("자리표시 이미지 생성:")
    placeholder(os.path.join(OUT, "screen_placeholder.png"), 1600, 900,
                "화면 · 이미지 자리", "그림 선택 → 마우스 오른쪽 → 그림 바꾸기", chrome=True)
    placeholder(os.path.join(OUT, "diagram_placeholder.png"), 1600, 760,
                "도식 · 그림 자리", "PPT에서 복사해 붙여넣기(도형 상태로 유지)")
    placeholder(os.path.join(OUT, "photo_placeholder.png"), 1200, 900,
                "사진 자리", "그림 바꾸기로 교체")
    logo(os.path.join(OUT, "logo_placeholder.png"))
