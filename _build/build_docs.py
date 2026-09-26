# -*- coding: utf-8 -*-
"""빠른 사용 가이드 PDF(엑셀 사용법·기준 + PPT 사용법·기준) + 아이콘 목록 이미지."""
import os, sys, tempfile, shutil
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pypdf import PdfReader, PdfWriter
from xlkit import new_workbook, postprocess
from xlguide import build_guide, build_standard
import icon_sheet
import lo

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def to_pdf(path, outdir):
    return lo.convert(path, "pdf", outdir)


def main():
    tmp = tempfile.mkdtemp()
    wb = new_workbook()
    build_guide(wb)
    build_standard(wb)
    xp = os.path.join(tmp, "guide.xlsx")
    wb.save(xp)
    postprocess(xp)
    xpdf = to_pdf(xp, tmp)
    pp = os.path.join(tmp, "deck.pptx")
    shutil.copy(os.path.join(ROOT, "02_PPT_구성요소", "IT인프라_PPT_구성요소_라이브러리.pptx"), pp)
    ppdf = to_pdf(pp, tmp)
    w = PdfWriter()
    for p in PdfReader(xpdf).pages:
        w.add_page(p)
    rd = PdfReader(ppdf)
    for i in (1, 2):          # PPT 2쪽(사용법·목차), 3쪽(디자인 기준)
        w.add_page(rd.pages[i])
    out = os.path.join(ROOT, "00_빠른_사용_가이드.pdf")
    with open(out, "wb") as f:
        w.write(f)
    icon_sheet.sheet(os.path.join(ROOT, "03_아이콘", "아이콘_목록.png"), px=96, cols=10)
    shutil.rmtree(tmp, ignore_errors=True)
    print(out)


if __name__ == "__main__":
    main()
