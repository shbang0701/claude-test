# -*- coding: utf-8 -*-
"""library 파일을 한 장짜리 미리보기 이미지로 만든다(찾아 쓰기 쉽게)."""
import os, sys, glob, subprocess, shutil, math
from PIL import Image, ImageDraw, ImageFont
HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import design as D

LIB = os.path.join(ROOT, "library")
OUT = os.path.join(ROOT, "preview")
TMP = "/tmp/_preview"
FONT = "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc"

def rgb(h): return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

def render(path, dpi=60):
    d = os.path.join(TMP, os.path.basename(path))
    shutil.rmtree(d, ignore_errors=True); os.makedirs(d, exist_ok=True)
    subprocess.run(["soffice", "--headless", "--norestore", "--convert-to", "pdf",
                    "--outdir", d, path], check=True, capture_output=True)
    pdf = glob.glob(os.path.join(d, "*.pdf"))[0]
    subprocess.run(["pdftoppm", "-r", str(dpi), "-png", pdf, os.path.join(d, "p")], check=True)
    return sorted(glob.glob(os.path.join(d, "p-*.png")))

def sheet(pngs, title, sub, cols, out, thumb_w=360):
    f_t = ImageFont.truetype(FONT, 30)
    f_s = ImageFont.truetype(FONT, 17)
    f_n = ImageFont.truetype(FONT, 15)
    ims = [Image.open(p) for p in pngs]
    ratio = ims[0].height / ims[0].width
    tw = thumb_w; th = int(tw * ratio)
    rows = math.ceil(len(ims) / cols)
    pad, gap, head, cap = 34, 18, 96, 26
    W = pad * 2 + cols * tw + (cols - 1) * gap
    H = head + pad + rows * (th + cap + gap)
    out_im = Image.new("RGB", (W, H), (255, 255, 255))
    d = ImageDraw.Draw(out_im)
    d.rectangle([0, 0, W, 6], fill=rgb(D.PRIMARY))
    d.text((pad, 30), title, font=f_t, fill=rgb(D.PRIMARY))
    d.text((pad, 68), sub, font=f_s, fill=rgb(D.INK_SOFT))
    for i, im in enumerate(ims):
        r, c = divmod(i, cols)
        x = pad + c * (tw + gap)
        y = head + pad + r * (th + cap + gap)
        im2 = im.resize((tw, th), Image.LANCZOS)
        out_im.paste(im2, (x, y))
        d.rectangle([x, y, x + tw - 1, y + th - 1], outline=rgb(D.LINE), width=1)
        d.text((x + 2, y + th + 5), f"{i+1}", font=f_n, fill=rgb(D.INK_FAINT))
    out_im.save(out, optimize=True)
    print("  ", os.path.basename(out), out_im.size)
    return out

SPEC = [
    ("01_문서_템플릿.docx", "01 문서 템플릿", "표지 · 목차 · 본문 구조가 갖춰진 문서 뼈대", 4, 330),
    ("02_Word_블록.docx", "02 Word 블록", "표현 유형별로 복사해 쓰는 문서 조각 + 서식 팁", 6, 300),
    ("03_슬라이드_요소.pptx", "03 슬라이드 요소", "장면 조합 19장 + 도형 도구상자 9장", 4, 400),
    ("04_Excel_시각서식.xlsx", "04 Excel 시각 서식", "셀 서식으로 만드는 시각 요소 · 함수 · 리본 기능", 4, 400),
]

if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    print("미리보기 생성:")
    for fn, title, sub, cols, tw in SPEC:
        p = os.path.join(LIB, fn)
        if not os.path.exists(p): continue
        pngs = render(p)
        sheet(pngs, title, sub, cols, os.path.join(OUT, os.path.splitext(fn)[0] + ".png"), tw)
