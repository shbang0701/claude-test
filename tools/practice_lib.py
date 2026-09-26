# -*- coding: utf-8 -*-
"""실무 가이드 3종 공통 도우미 — 난이도 표시, 기준값 표, 조작 단계."""
import os, sys
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import design as D
import docx_lib as L
import blocks as K
from docx_lib import (P, R, C, V, U, B, TBL, TR, TC, th, td, PCT, sid, FLD, IMG, TAB,
                      spacing, ind, pbdr, bdr, pt, cm, nonum, CELLMAR)
from blocks import (spacer, rule, big_title, h1, h2, h3, h4, para, lead, note, small,
                    checks, bullets, substeps, callout, info_table, data_table, step, body,
                    cmd, out_block, lab)
from build_docx import NOBORDER, plain_table, summary_points

LV = {"기본": (D.PRIMARY_MID, D.PRIMARY_TINT), "심화": (D.INK_FAINT, D.SURFACE)}

def topic(title, level, summary=None):
    """소제목 + 난이도 꼬리표 + 한 줄 요약."""
    c, bg = LV[level]
    x = P([R(title), R("   "), R(f"[{level}]", sz=8.5, b=True, color=c, fill=bg)],
          style="Heading3")
    if summary: x += small(summary)
    return x

def steps(items):
    return "".join(P(t, style=sid("step")) for t in items)

def menu(*path):
    """메뉴 경로를 UI 서식으로."""
    return U(" › ".join(path))

def kv(title, rows, w=(30, 70)):
    return K.tcap(title) + data_table(["항목", "기준"], rows, list(w))

def why(text):
    return callout("info", text)

def keys(rows, caption="단축키"):
    return data_table(["단축키", "하는 일"], rows, [26, 74], caption=caption)

def build(body_parts, path, title, header_left, header_right, footer_left, subject):
    body_xml = "".join(body_parts) + L.sect_pr()
    L.build_document(body_xml, path, header_left=header_left, header_right=header_right,
                     footer_left=footer_left, title=title, subject=subject)
    return path

def cover(kicker, title, sub, intro, contents):
    """가이드 첫 장."""
    b = [spacer(10),
         P([R(kicker, sz=9, b=True, color=D.PRIMARY_MID, track=60)], sp=spacing(after=pt(6))),
         P([R(title)], style="Title"),
         P([R(sub)], style="Subtitle"),
         rule(D.PRIMARY_MID, 8, before=10, after=14),
         lead(intro),
         P([R("[기본]", sz=8.5, b=True, color=LV["기본"][0], fill=LV["기본"][1]),
            R("  먼저 익힌다. 여기까지면 대부분의 문서는 만들 수 있다.      ", sz=10),
            R("[심화]", sz=8.5, b=True, color=LV["심화"][0], fill=LV["심화"][1]),
            R("  필요할 때 한 번씩 본다.", sz=10)],
           sp=spacing(after=pt(10))),
         K.tcap("이 가이드에 있는 것"),
         data_table(["장", "내용", "난이도"], contents, [10, 70, 20])]
    return "".join(b)
