# -*- coding: utf-8 -*-
"""재사용 블록 조립기 — 두 Word 문서가 같은 블록을 쓰도록 한 곳에 모았다."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import re
import design as D
import docx_lib as L
from docx_lib import P, R, C, V, U, B, TBL, TR, TC, th, td, PCT, sid, FLD, IMG

CALLOUT = {
    "caution": ("주의", D.WARN),
    "warn":    ("경고", D.DANGER),
    "ok":      ("확인", D.OK),
    "info":    ("참고", D.NOTE),
}

def callout(kind, *runs, label=None):
    lb, color = CALLOUT[kind]
    if label is not None: lb = label
    body = []
    for x in runs:
        body.append(x if isinstance(x, str) and x.startswith("<w:") else R(x))
    return P([R(lb, b=True, color=color), R("  ")] + body, style=sid(kind))

_VAR = re.compile(r"(<[^<>\n]{1,40}>)")

def mono_runs(text):
    """명령 블록 안의 <바꿔 넣을 값>을 Doc-치환값 서식으로 구분해 준다."""
    out = []
    for seg in _VAR.split(text):
        if not seg: continue
        out.append(R(seg, style=sid("var")) if _VAR.fullmatch(seg) else R(seg))
    return out

def cmd(*lines, label=None):
    """명령 블록. 여러 줄은 한 문단 안의 줄바꿈(Shift+Enter)으로 넣는다."""
    out = lab(label) if label else ""
    return out + P(mono_runs("\n".join(lines)), style=sid("cmd"))

def out_block(*lines, label=None):
    o = lab(label) if label else ""
    return o + P(R("\n".join(lines)), style=sid("out"))

def lab(text):
    return P(R(text), style=sid("label"))

def step(text_runs, **kw):
    return P(text_runs, style=sid("step"), **kw)

def body(text_runs, **kw):
    return P(text_runs, style=sid("stepbody"), **kw)

def note(text_runs, **kw):
    return P(text_runs, style=sid("note"), **kw)

def small(text_runs, **kw):
    return P(text_runs, style=sid("small"), **kw)

def checks(items):
    return "".join(P(i, style=sid("check")) for i in items)

def bullets(items):
    return "".join(P(i, style=sid("bullet")) for i in items)

def substeps(items):
    return "".join(P(i, style=sid("substep")) for i in items)

CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫"

def guides(items, start=0):
    """화면 조작 안내. 번호는 고정 기호라 블록을 여러 개 써도 서로 영향을 주지 않는다."""
    out = []
    for i, t in enumerate(items):
        out.append(P([R(CIRCLED[(start + i) % len(CIRCLED)], b=True, color=D.PRIMARY_MID),
                      L.TAB(), R(t)], style=sid("guide")))
    return "".join(out)

# ── 표 ─────────────────────────────────────────────────────────────────────
def info_table(pairs, lw=26):
    """라벨 | 값 2열 정보표(Doc-정보표)."""
    rows = []
    for k, v in pairs:
        rows.append(TR([TC(td(k), w=PCT(lw), valign="center"),
                        TC(td(v), w=PCT(100 - lw), valign="center")]))
    return TBL(rows, [lw, 100 - lw], style="DocTableInfo",
               look='<w:tblLook w:val="0680" w:firstRow="0" w:lastRow="0" w:firstColumn="1" '
                    'w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>')

def data_table(headers, rows, widths, caption=None, align=None):
    """머리글 있는 기본 표(Doc-표). rows 의 각 칸은 문자열 또는 run XML."""
    out = tcap(caption) if caption else ""
    hr = TR([TC(th(h), w=PCT(w), valign="center") for h, w in zip(headers, widths)], header=True)
    body_rows = []
    for r in rows:
        cells = []
        for i, c in enumerate(r):
            jc = (align or [None] * len(r))[i]
            cells.append(TC(td(c, jc=jc) if jc else td(c), w=PCT(widths[i]), valign="top"))
        body_rows.append(TR(cells))
    return out + TBL([hr] + body_rows, widths)

def tcap(text):
    return P(R(text), style=sid("tcap"))

def fig_caption(text, kind="그림"):
    return P([R(f"{kind} "), FLD(f" SEQ {kind} \\* ARABIC ", "1"), R("  " + text)], style="Caption")

def screen(rid, caption, guide_items=None, w=14.0, h=None, ratio=9 / 16.0):
    h = h or round(w * ratio, 2)
    x = P(IMG(rid, w, h, "screen"), style=sid("fig")) + fig_caption(caption)
    if guide_items: x += guides(guide_items)
    return x

# ── 문서 조립용 공통 도우미 ────────────────────────────────────────────────
from docx_lib import spacing, ind, pbdr, bdr, pt, cm, nonum

def spacer(h_pt, after=0):
    return P('', sp=spacing(after=after, line=pt(h_pt), rule="exact"))

def rule(color=D.LINE, sz=6, before=0, after=8):
    return P('', sp=spacing(before=pt(before), after=pt(after), line=pt(2), rule="exact"),
             border=pbdr(bottom=bdr("bottom", sz, color, 0)))

def big_title(text, pagebreak=False, sub=None):
    """목차·표지처럼 차례에 넣지 않는 큰 제목(직접 서식)."""
    x = P([R(text, sz=16.5, b=True, color=D.PRIMARY, track=-8)],
          sp=spacing(before=0, after=pt(10), line=300), pagebreak=pagebreak,
          border=pbdr(bottom=bdr("bottom", 6, "C9D6E2", 8)))
    if sub: x += P([R(sub, sz=9.5, color=D.INK_SOFT)], sp=spacing(after=pt(8)))
    return x

def h1(t, num=True):  return P(R(t), style="Heading1", extra="" if num else nonum())
def h2(t, num=True):  return P(R(t), style="Heading2", extra="" if num else nonum())
def h3(t, num=True):  return P(R(t), style="Heading3", extra="" if num else nonum())
def h4(t):            return P(R(t), style="Heading4")

def para(*runs, **kw):
    return P([x if isinstance(x, str) and x.startswith("<w:") else R(x) for x in runs], **kw)

def lead(text, **kw):
    return P(R(text), style=sid("lead"), **kw)

def quote(text, source=None):
    x = P(R(text), style=sid("quote"))
    if source: x += P(R("— " + source), style=sid("quote"), sp=spacing(before=0, after=pt(6)))
    return x

def entry(name, when):
    """블록 라이브러리의 한 항목: 제목 + 한 줄 설명."""
    return h2(name) + (note(when) if when else "")
