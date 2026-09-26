# -*- coding: utf-8 -*-
"""Word(.docx) 생성 엔진.

- styles.xml / numbering.xml 을 직접 정의해 '스타일로 관리되는' 문서를 만든다.
- 본문은 XML 문자열로 조립한 뒤 python-docx 패키지에 삽입한다(포장·이미지·헤더는 python-docx 담당).
"""
import os, sys
from xml.sax.saxutils import escape as _esc

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import design as D
from design import cm, pt, hp, eighth

W = 'w:'
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
      'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
      'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
      'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
      'mc:Ignorable="w14"')

# 스타일 ID -> 사용자에게 보이는 이름
S = {
    "step":     ("ITStep",        "IT-단계"),
    "stepbody": ("ITStepBody",    "IT-단계설명"),
    "substep":  ("ITSubStep",     "IT-단계하위"),
    "bullet":   ("ITBullet",      "IT-글머리"),
    "check":    ("ITCheck",       "IT-체크"),
    "cmd":      ("ITCmd",         "IT-명령"),
    "out":      ("ITOut",         "IT-출력"),
    "label":    ("ITLabel",       "IT-블록라벨"),
    "fig":      ("ITFig",         "IT-그림"),
    "guide":    ("ITScreenGuide", "IT-화면안내"),
    "note":     ("ITWriteNote",   "IT-작성안내"),
    "thead":    ("ITTHead",       "IT-표머리"),
    "tbody":    ("ITTBody",       "IT-표본문"),
    "tcap":     ("ITTableCap",    "IT-표제목"),
    "small":    ("ITSmall",       "IT-보조설명"),
    "caution":  ("ITCaution",     "IT-주의"),
    "warn":     ("ITWarn",        "IT-경고"),
    "ok":       ("ITOk",          "IT-확인"),
    "info":     ("ITInfo",        "IT-참고"),
    # 글자 스타일
    "code":     ("ITCode",        "IT-코드"),
    "var":      ("ITVar",         "IT-치환값"),
    "ui":       ("ITUi",          "IT-UI"),
    "strong":   ("ITStrong",      "IT-강조"),
}
def sid(key): return S[key][0]

IND = cm(0.7)          # 단계 본문 기준 들여쓰기
IND2 = cm(1.4)

# ───────────────────────────── 저수준 조각 ──────────────────────────────────
def esc(t): return _esc(t)
def esc_attr(t): return _esc(str(t), {'"': "&quot;", "'": "&apos;"})

def bdr(tag, sz=4, color=D.LINE, space=0, val="single"):
    return f'<w:{tag} w:val="{val}" w:sz="{sz}" w:space="{space}" w:color="{color}"/>'

def pbdr(top=None, left=None, bottom=None, right=None, between=None):
    parts = [x for x in (top, left, bottom, right, between) if x]
    return f'<w:pBdr>{"".join(parts)}</w:pBdr>' if parts else ""

def shd(fill): return f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>'

def spacing(before=None, after=None, line=None, rule="auto"):
    a = []
    if before is not None: a.append(f'w:before="{before}"')
    if after is not None:  a.append(f'w:after="{after}"')
    if line is not None:   a.append(f'w:line="{line}" w:lineRule="{rule}"')
    return f'<w:spacing {" ".join(a)}/>' if a else ""

def ind(left=None, hanging=None, right=None, first=None):
    a = []
    if left is not None:    a.append(f'w:left="{left}"')
    if right is not None:   a.append(f'w:right="{right}"')
    if hanging is not None: a.append(f'w:hanging="{hanging}"')
    if first is not None:   a.append(f'w:firstLine="{first}"')
    return f'<w:ind {" ".join(a)}/>' if a else ""

def rfonts(ascii_=D.FONT_LATIN, ea=D.FONT_EA, mono=False):
    if mono: ascii_, ea = D.FONT_MONO, D.FONT_MONO_EA
    return f'<w:rFonts w:ascii="{ascii_}" w:eastAsia="{ea}" w:hAnsi="{ascii_}" w:cs="{ascii_}"/>'

def rpr(font=None, mono=False, b=False, sz=None, color=None, fill=None, track=None,
        caps=False, noproof=False, style=None, u=False, i=False):
    x = []
    if style: x.append(f'<w:rStyle w:val="{style}"/>')
    if mono or font: x.append(rfonts(font or D.FONT_LATIN, D.FONT_EA, mono))
    if b: x.append('<w:b/>')
    if i: x.append('<w:i/>')
    if caps: x.append('<w:caps/>')
    if noproof: x.append('<w:noProof/>')
    if color: x.append(f'<w:color w:val="{color}"/>')
    if track is not None: x.append(f'<w:spacing w:val="{track}"/>')
    if sz: x.append(f'<w:sz w:val="{hp(sz)}"/><w:szCs w:val="{hp(sz)}"/>')
    if u: x.append('<w:u w:val="single"/>')
    if fill: x.append(shd(fill))
    return f'<w:rPr>{"".join(x)}</w:rPr>' if x else ""

def R(text="", **kw):
    """글자 실행(run). \\n 은 줄바꿈(같은 문단 안)."""
    body = ""
    for i, seg in enumerate(str(text).split("\n")):
        if i: body += '<w:br/>'
        if seg: body += f'<w:t xml:space="preserve">{esc(seg)}</w:t>'
    return f'<w:r>{rpr(**kw)}{body}</w:r>'

def C(text):   return R(text, style=sid("code"))
def V(text):   return R(text, style=sid("var"))
def U(text):   return R(text, style=sid("ui"))
def B(text):   return R(text, style=sid("strong"))
def TAB():     return '<w:r><w:tab/></w:r>'

def FLD(instr, cached="1", **kw):
    """간단 필드(PAGE, NUMPAGES, SEQ, TOC 등)."""
    return (f'<w:fldSimple w:instr="{esc_attr(instr)}"><w:r>{rpr(**kw)}'
            f'<w:t>{esc(cached)}</w:t></w:r></w:fldSimple>')

def P(runs="", style=None, numpr=None, keepnext=False, keeplines=False, pagebreak=False,
      sp=None, indent=None, border=None, fill=None, jc=None, ctx=False, outline=None,
      tabs=None, extra=""):
    if isinstance(runs, (list, tuple)): runs = "".join(runs)
    elif isinstance(runs, str) and not runs.startswith("<w:"):
        runs = R(runs) if runs else ""
    x = []
    if style: x.append(f'<w:pStyle w:val="{style}"/>')
    if keepnext: x.append('<w:keepNext/>')
    if keeplines: x.append('<w:keepLines/>')
    if pagebreak: x.append('<w:pageBreakBefore/>')
    if numpr: x.append(numpr)
    if border: x.append(border)
    if fill: x.append(shd(fill))
    if tabs: x.append(tabs)
    if sp: x.append(sp)
    if indent: x.append(indent)
    if ctx: x.append('<w:contextualSpacing/>')
    if jc: x.append(f'<w:jc w:val="{jc}"/>')
    if outline is not None: x.append(f'<w:outlineLvl w:val="{outline}"/>')
    x.append(extra)
    ppr = f'<w:pPr>{"".join(x)}</w:pPr>' if any(x) else ""
    return f'<w:p>{ppr}{runs}</w:p>'

def numpr(ilvl, numid): return f'<w:numPr><w:ilvl w:val="{ilvl}"/><w:numId w:val="{numid}"/></w:numPr>'
def nonum(): return '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>'

def BR_PAGE(): return '<w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:br w:type="page"/></w:r></w:p>'
def EMPTY(after=0, sz=8): return P('', sp=spacing(after=after, line=pt(sz), rule="exact"))

# ───────────────────────────── 표 ──────────────────────────────────────────
def TC(paras, w=None, fill=None, span=None, valign="center", mar=None, borders=None, vmerge=None):
    if isinstance(paras, str): paras = [paras]
    x = []
    if w: x.append(f'<w:tcW w:w="{w}" w:type="pct"/>')
    else: x.append('<w:tcW w:w="0" w:type="auto"/>')
    if span: x.append(f'<w:gridSpan w:val="{span}"/>')
    if vmerge: x.append(f'<w:vMerge w:val="{vmerge}"/>' if vmerge != "cont" else '<w:vMerge/>')
    if borders: x.append(borders)
    if fill: x.append(shd(fill))
    if mar: x.append(mar)
    if valign: x.append(f'<w:vAlign w:val="{valign}"/>')
    return f'<w:tc><w:tcPr>{"".join(x)}</w:tcPr>{"".join(paras)}</w:tc>'

def TR(cells, header=False, cantsplit=True, height=None):
    x = []
    if cantsplit: x.append('<w:cantSplit/>')
    if header: x.append('<w:tblHeader/>')
    if height: x.append(f'<w:trHeight w:val="{height}" w:hRule="atLeast"/>')
    trpr = f'<w:trPr>{"".join(x)}</w:trPr>' if x else ""
    return f'<w:tr>{trpr}{"".join(cells)}</w:tr>'

def TBL(rows, widths, style="ITTable", indent=0, look='<w:tblLook w:val="0620" w:firstRow="1" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>',
        borders=None, cellmar=None, layout="fixed"):
    """widths: 각 열의 백분율(합 100)."""
    grid = "".join(f'<w:gridCol w:w="{int(w * 94.68)}"/>' for w in widths)  # A4 본문폭 기준 근사
    x = [f'<w:tblStyle w:val="{style}"/>' if style else "",
         '<w:tblW w:w="5000" w:type="pct"/>',
         f'<w:tblInd w:w="{indent}" w:type="dxa"/>' if indent else "",
         borders or "", f'<w:tblLayout w:type="{layout}"/>', cellmar or "", look]
    return (f'<w:tbl><w:tblPr>{"".join(x)}</w:tblPr><w:tblGrid>{grid}</w:tblGrid>'
            f'{"".join(rows)}</w:tbl>')

PCT = lambda v: int(round(v * 50))   # % -> pct 단위(5000 = 100%)

CELLMAR = ('<w:tblCellMar><w:top w:w="72" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>'
           '<w:bottom w:w="72" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>')

def th(text, **kw):  return P(text if text.startswith("<w:") else R(text), style=sid("thead"), **kw)
def td(text, **kw):  return P(text if isinstance(text, str) and text.startswith("<w:") else R(text), style=sid("tbody"), **kw)

# ───────────────────────────── 이미지 ──────────────────────────────────────
EMU = 360000  # 1cm
_img_id = [10]
def IMG(rid, w_cm, h_cm, name="image"):
    _img_id[0] += 1
    i = _img_id[0]
    cx, cy = int(w_cm * EMU), int(h_cm * EMU)
    return (f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
            f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'<wp:docPr id="{i}" name="{name}{i}"/>'
            f'<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>'
            f'<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            f'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:nvPicPr><pic:cNvPr id="{i}" name="{name}{i}"/><pic:cNvPicPr/></pic:nvPicPr>'
            f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
            f'</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')

# ───────────────────────────── 스타일 정의 ─────────────────────────────────
def _style(sid_, name, typ="paragraph", based=None, nxt=None, ppr="", rpr_="",
           qformat=True, ui=None, custom=True, link=None):
    x = [f'<w:name w:val="{name}"/>']
    if based: x.append(f'<w:basedOn w:val="{based}"/>')
    if nxt: x.append(f'<w:next w:val="{nxt}"/>')
    if link: x.append(f'<w:link w:val="{link}"/>')
    if ui is not None: x.append(f'<w:uiPriority w:val="{ui}"/>')
    if qformat: x.append('<w:qFormat/>')
    if ppr: x.append(f'<w:pPr>{ppr}</w:pPr>')
    if rpr_: x.append(f'<w:rPr>{rpr_}</w:rPr>')
    cu = ' w:customStyle="1"' if custom else ''
    return f'<w:style w:type="{typ}"{cu} w:styleId="{sid_}">{"".join(x)}</w:style>'

def _rp(**kw):
    """스타일 안에서 쓰는 rPr 내용(태그 없이)."""
    s = rpr(**kw)
    return s[len('<w:rPr>'):-len('</w:rPr>')] if s else ""

def _callout(key, name, accent, bg, bdrc):
    return _style(sid(key), name, based="Normal", nxt="Normal", ui=20,
        ppr=(pbdr(bdr("top", 4, bdrc, 6), bdr("left", 18, accent, 8),
                  bdr("bottom", 4, bdrc, 6), bdr("right", 4, bdrc, 8),
                  bdr("between", 0, bdrc, 0, val="nil")) +
             shd(bg) + spacing(before=pt(6), after=pt(6), line=310) +
             ind(left=IND, right=0) + '<w:contextualSpacing/>'),
        rpr_=_rp(sz=10, color=D.INK))

def styles_xml():
    lat = ('<w:latentStyles w:defLockedState="0" w:defUIPriority="99" w:defSemiHidden="1" '
           'w:defUnhideWhenUsed="1" w:defQFormat="0" w:count="376">' +
           "".join(f'<w:lsdException w:name="{n}" w:semiHidden="0" w:uiPriority="{p}" '
                   f'w:unhideWhenUsed="0" w:qFormat="1"/>'
                   for n, p in [("Normal", 0), ("heading 1", 9), ("heading 2", 9), ("heading 3", 9),
                                ("heading 4", 9), ("caption", 35), ("Title", 10), ("Subtitle", 11),
                                ("Strong", 22), ("Hyperlink", 99), ("Normal Table", 99)]) +
           '</w:latentStyles>')

    docdef = (
        '<w:docDefaults><w:rPrDefault><w:rPr>' + rfonts() +
        f'<w:color w:val="{D.INK}"/><w:sz w:val="{hp(10.5)}"/><w:szCs w:val="{hp(10.5)}"/>'
        '<w:lang w:val="en-US" w:eastAsia="ko-KR" w:bidi="ar-SA"/>'
        '</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>'
        '<w:widowControl/><w:wordWrap/><w:autoSpaceDE/><w:autoSpaceDN/><w:snapToGrid w:val="0"/>'
        + spacing(after=pt(6), line=336) + '<w:jc w:val="left"/>'
        '</w:pPr></w:pPrDefault></w:docDefaults>')

    st = []
    # 기본
    st.append('<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
              '<w:qFormat/></w:style>')
    st.append('<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">'
              '<w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/>'
              '<w:unhideWhenUsed/></w:style>')
    st.append('<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/>'
              '<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>'
              '<w:tblPr><w:tblInd w:w="0" w:type="dxa"/>' + CELLMAR + '</w:tblPr></w:style>')
    st.append('<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/>'
              '<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style>')

    # 표지
    st.append(_style("Title", "Title", based="Normal", nxt="Normal", ui=10, custom=False,
        ppr=spacing(before=0, after=pt(6), line=300) + '<w:contextualSpacing/>',
        rpr_=_rp(sz=25, b=True, color=D.PRIMARY, track=-12)))
    st.append(_style("Subtitle", "Subtitle", based="Normal", nxt="Normal", ui=11, custom=False,
        ppr=spacing(before=0, after=pt(2), line=300),
        rpr_=_rp(sz=11.5, color=D.INK_SOFT)))

    # 제목 1~4 (제목 1은 새 쪽에서 시작)
    st.append(_style("Heading1", "heading 1", based="Normal", nxt="Normal", ui=9, custom=False, link=None,
        ppr=('<w:keepNext/><w:keepLines/><w:pageBreakBefore/>' + numpr(0, 1) +
             pbdr(bottom=bdr("bottom", 6, "C9D6E2", 8)) +
             spacing(before=0, after=pt(10), line=300) + '<w:outlineLvl w:val="0"/>'),
        rpr_=_rp(sz=16.5, b=True, color=D.PRIMARY, track=-8)))
    st.append(_style("Heading2", "heading 2", based="Normal", nxt="Normal", ui=9, custom=False,
        ppr=('<w:keepNext/><w:keepLines/>' + numpr(1, 1) +
             spacing(before=pt(17), after=pt(5), line=300) + '<w:outlineLvl w:val="1"/>'),
        rpr_=_rp(sz=13, b=True, color=D.PRIMARY, track=-6)))
    st.append(_style("Heading3", "heading 3", based="Normal", nxt="Normal", ui=9, custom=False,
        ppr=('<w:keepNext/><w:keepLines/>' + numpr(2, 1) +
             spacing(before=pt(12), after=pt(4), line=300) + '<w:outlineLvl w:val="2"/>'),
        rpr_=_rp(sz=11.5, b=True, color=D.INK)))
    st.append(_style("Heading4", "heading 4", based="Normal", nxt="Normal", ui=9, custom=False,
        ppr='<w:keepNext/><w:keepLines/>' + spacing(before=pt(9), after=pt(3)) + '<w:outlineLvl w:val="3"/>',
        rpr_=_rp(sz=10.5, b=True, color=D.INK_SOFT)))

    # 절차
    st.append(_style(sid("step"), S["step"][1], based="Normal", nxt=sid("stepbody"), ui=12,
        ppr=('<w:keepNext/><w:keepLines/>' + numpr(3, 1) +
             spacing(before=pt(11), after=pt(3), line=310) + ind(left=IND, hanging=IND)),
        rpr_=_rp(sz=11, b=True, color=D.INK)))
    st.append(_style(sid("stepbody"), S["stepbody"][1], based="Normal", nxt=sid("stepbody"), ui=12,
        ppr=spacing(before=0, after=pt(4), line=320) + ind(left=IND)))
    st.append(_style(sid("substep"), S["substep"][1], based="Normal", nxt=sid("substep"), ui=12,
        ppr=numpr(4, 1) + spacing(before=0, after=pt(3), line=310) + ind(left=cm(1.25), hanging=cm(0.55))))
    st.append(_style(sid("bullet"), S["bullet"][1], based="Normal", nxt=sid("bullet"), ui=12,
        ppr=numpr(0, 2) + spacing(before=0, after=pt(3), line=310) + ind(left=cm(1.25), hanging=cm(0.55))))
    st.append(_style(sid("check"), S["check"][1], based="Normal", nxt=sid("check"), ui=12,
        ppr=numpr(0, 3) + spacing(before=0, after=pt(3), line=310) + ind(left=cm(1.25), hanging=cm(0.55))))
    st.append(_style(sid("guide"), S["guide"][1], based="Normal", nxt=sid("guide"), ui=13,
        ppr=(nonum() + spacing(before=0, after=pt(2), line=300) +
             ind(left=cm(1.25), hanging=cm(0.55))),
        rpr_=_rp(sz=10)))

    # 명령 / 출력
    st.append(_style(sid("cmd"), S["cmd"][1], based="Normal", nxt=sid("cmd"), ui=14,
        ppr=('<w:keepLines/>' +
             pbdr(bdr("top", 4, D.CODE_BDR, 5), bdr("left", 18, D.PRIMARY_MID, 8),
                  bdr("bottom", 4, D.CODE_BDR, 5), bdr("right", 4, D.CODE_BDR, 8),
                  bdr("between", 0, D.CODE_BDR, 0, val="nil")) +
             shd(D.CODE_BG) + '<w:suppressAutoHyphens/>' +
             spacing(before=pt(3), after=pt(3), line=270) + ind(left=IND) + '<w:contextualSpacing/>'),
        rpr_=_rp(mono=True, sz=9.5, color=D.CODE_TXT, noproof=True)))
    st.append(_style(sid("out"), S["out"][1], based="Normal", nxt=sid("out"), ui=14,
        ppr=('<w:keepLines/>' +
             pbdr(bdr("top", 4, D.OUT_BDR, 5), bdr("left", 4, D.OUT_BDR, 8),
                  bdr("bottom", 4, D.OUT_BDR, 5), bdr("right", 4, D.OUT_BDR, 8),
                  bdr("between", 0, D.OUT_BDR, 0, val="nil")) +
             shd(D.OUT_BG) + '<w:suppressAutoHyphens/>' +
             spacing(before=pt(1), after=pt(3), line=265) + ind(left=IND) + '<w:contextualSpacing/>'),
        rpr_=_rp(mono=True, sz=9, color=D.OUT_TXT, noproof=True)))
    st.append(_style(sid("label"), S["label"][1], based="Normal", nxt="Normal", ui=14,
        ppr='<w:keepNext/>' + spacing(before=pt(6), after=pt(1), line=260) + ind(left=IND),
        rpr_=_rp(sz=8.5, b=True, color=D.INK_FAINT, track=16)))

    # 그림 / 캡션 / 표
    st.append(_style(sid("fig"), S["fig"][1], based="Normal", nxt="Caption", ui=15,
        ppr='<w:keepNext/>' + spacing(before=pt(6), after=pt(2), line=240) + ind(left=IND)))
    st.append(_style("Caption", "caption", based="Normal", nxt="Normal", ui=35, custom=False,
        ppr='<w:keepLines/>' + spacing(before=0, after=pt(8), line=280) + ind(left=IND),
        rpr_=_rp(sz=9, color=D.INK_SOFT)))
    st.append(_style(sid("tcap"), S["tcap"][1], based="Normal", nxt="Normal", ui=15,
        ppr='<w:keepNext/><w:keepLines/>' + spacing(before=pt(9), after=pt(3), line=280),
        rpr_=_rp(sz=9.5, b=True, color=D.INK_SOFT)))
    st.append(_style(sid("thead"), S["thead"][1], based="Normal", nxt=sid("tbody"), ui=16,
        ppr='<w:keepNext/>' + nonum() + spacing(before=pt(2), after=pt(2), line=265),
        rpr_=_rp(sz=9.5, b=True, color=D.PRIMARY)))
    st.append(_style(sid("tbody"), S["tbody"][1], based="Normal", nxt=sid("tbody"), ui=16,
        ppr=nonum() + spacing(before=pt(2), after=pt(2), line=270),
        rpr_=_rp(sz=9.5)))

    # 콜아웃 4종
    st.append(_callout("caution", S["caution"][1], D.WARN, D.WARN_BG, D.WARN_BDR))
    st.append(_callout("warn",    S["warn"][1],    D.DANGER, D.DANGER_BG, D.DANGER_BDR))
    st.append(_callout("ok",      S["ok"][1],      D.OK, D.OK_BG, D.OK_BDR))
    st.append(_callout("info",    S["info"][1],    D.NOTE, D.NOTE_BG, D.NOTE_BDR))

    # 보조
    st.append(_style(sid("small"), S["small"][1], based="Normal", nxt="Normal", ui=18,
        ppr=spacing(before=0, after=pt(6), line=290) + ind(left=IND),
        rpr_=_rp(sz=9, color=D.INK_SOFT)))
    st.append(_style(sid("note"), S["note"][1], based="Normal", nxt=sid("note"), ui=18,
        ppr=(pbdr(left=bdr("left", 6, "B9C4CE", 8, val="dotted")) +
             spacing(before=pt(4), after=pt(4), line=290) + ind(left=IND) + '<w:contextualSpacing/>'),
        rpr_=_rp(sz=9.5, color="7A8590")))

    # 글자 스타일
    st.append(_style(sid("code"), S["code"][1], typ="character", based="DefaultParagraphFont", ui=1,
        rpr_=_rp(mono=True, sz=9.5, color=D.CODE_TXT, fill=D.CODE_BG, noproof=True)))
    st.append(_style(sid("var"), S["var"][1], typ="character", based="DefaultParagraphFont", ui=1,
        rpr_=_rp(mono=True, sz=9.5, b=True, color=D.VAR, fill=D.VAR_BG, noproof=True)))
    st.append(_style(sid("ui"), S["ui"][1], typ="character", based="DefaultParagraphFont", ui=1,
        rpr_=_rp(b=True, color=D.PRIMARY)))
    st.append(_style(sid("strong"), S["strong"][1], typ="character", based="DefaultParagraphFont", ui=1,
        rpr_=_rp(b=True, color=D.INK)))

    # 표 스타일 2종
    tbl_common = ('<w:tblPr><w:tblStyleRowBandSize w:val="1"/>'
                  '<w:tblBorders>' + bdr("top", 12, D.PRIMARY) + bdr("left", 0, "auto", val="nil") +
                  bdr("bottom", 8, D.PRIMARY_MID) + bdr("right", 0, "auto", val="nil") +
                  bdr("insideH", 4, D.LINE) + bdr("insideV", 0, "auto", val="nil") + '</w:tblBorders>'
                  + CELLMAR + '</w:tblPr>')
    st.append('<w:style w:type="table" w:customStyle="1" w:styleId="ITTable">'
              '<w:name w:val="IT-표"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="19"/><w:qFormat/>'
              f'<w:rPr>{_rp(sz=9.5)}</w:rPr>' + tbl_common +
              '<w:tblStylePr w:type="firstRow"><w:pPr><w:keepNext/></w:pPr>'
              f'<w:rPr>{_rp(b=True, color=D.PRIMARY)}</w:rPr>'
              '<w:tcPr><w:tcBorders>' + bdr("bottom", 8, D.PRIMARY_MID) + '</w:tcBorders>'
              + shd(D.PRIMARY_TINT) + '</w:tcPr></w:tblStylePr>'
              '<w:tblStylePr w:type="band1Horz"><w:tcPr>' + shd(D.SURFACE) + '</w:tcPr></w:tblStylePr>'
              '</w:style>')
    st.append('<w:style w:type="table" w:customStyle="1" w:styleId="ITTableInfo">'
              '<w:name w:val="IT-정보표"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="19"/><w:qFormat/>'
              f'<w:rPr>{_rp(sz=9.5)}</w:rPr>'
              '<w:tblPr><w:tblBorders>' + bdr("top", 4, D.LINE) + bdr("left", 0, "auto", val="nil") +
              bdr("bottom", 4, D.LINE) + bdr("right", 0, "auto", val="nil") +
              bdr("insideH", 4, D.LINE_SOFT) + bdr("insideV", 0, "auto", val="nil") + '</w:tblBorders>'
              + CELLMAR + '</w:tblPr>'
              '<w:tblStylePr w:type="firstCol">'
              f'<w:rPr>{_rp(b=True, color=D.INK_SOFT)}</w:rPr>'
              '<w:tcPr>' + shd(D.SURFACE) + '</w:tcPr></w:tblStylePr>'
              '</w:style>')

    return docdef + lat + "".join(st)

# ───────────────────────────── 번호 정의 ───────────────────────────────────
def _lvl(i, fmt, text, style=None, left=0, hanging=0, suff=None, rpr_="", start=1, restart=None, jc="left"):
    x = [f'<w:start w:val="{start}"/>']
    if restart is not None: x.append(f'<w:lvlRestart w:val="{restart}"/>')
    if style: x.append(f'<w:pStyle w:val="{style}"/>')
    x.append(f'<w:numFmt w:val="{fmt}"/>')
    x.append(f'<w:lvlText w:val="{text}"/>')
    if suff: x.append(f'<w:suff w:val="{suff}"/>')
    x.append(f'<w:lvlJc w:val="{jc}"/>')
    x.append(f'<w:pPr>{ind(left=left, hanging=hanging) if hanging else ind(left=left, first=0)}</w:pPr>')
    if rpr_: x.append(f'<w:rPr>{rpr_}</w:rPr>')
    # 요소 순서: start, numFmt, lvlRestart, pStyle, ... -> 스키마 순서에 맞춰 재배열
    order = ['<w:start', '<w:numFmt', '<w:lvlRestart', '<w:pStyle', '<w:isLgl', '<w:suff',
             '<w:lvlText', '<w:lvlPicBulletId', '<w:legacy', '<w:lvlJc', '<w:pPr', '<w:rPr']
    x.sort(key=lambda e: next(k for k, o in enumerate(order) if e.startswith(o)))
    return f'<w:lvl w:ilvl="{i}">' + "".join(x) + '</w:lvl>'

def numbering_xml():
    mono_small = _rp(mono=False, b=True, color=D.PRIMARY_MID)
    # 절차: 0~2 = 제목1~3(보이지 않음) → 제목이 바뀌면 단계 번호가 1부터 다시 시작
    EM = "\u2003"
    proc = [
        # 제목 1·2 만 번호를 보이고, 제목 3 은 번호 없이 '단계 번호 다시 시작' 기준으로만 쓴다.
        # (제목 3 에 번호를 주면, 제목 3 없이 단계를 쓴 절에서 번호가 하나씩 밀린다.)
        _lvl(0, "decimal", "%1." + EM, style="Heading1", left=0, suff="nothing"),
        _lvl(1, "decimal", "%1.%2" + EM, style="Heading2", left=0, suff="nothing"),
        _lvl(2, "none", "", style="Heading3", left=0, suff="nothing"),
        _lvl(3, "decimal", "%4.", style=sid("step"), left=IND, hanging=IND, rpr_=mono_small),
        _lvl(4, "lowerLetter", "%5.", style=sid("substep"), left=cm(1.25), hanging=cm(0.55),
             rpr_=_rp(color=D.INK_SOFT)),
    ]
    for i in range(5, 9):
        proc.append(_lvl(i, "bullet", "–", left=cm(1.2 + 0.5 * (i - 5)), hanging=cm(0.4)))
    bullets = [_lvl(0, "bullet", "•", style=sid("bullet"), left=cm(1.2), hanging=cm(0.5),
                    rpr_=_rp(color=D.PRIMARY_MID))]
    bullets += [_lvl(i, "bullet", "–", left=cm(1.7 + 0.45 * (i - 1)), hanging=cm(0.45))
                for i in range(1, 9)]
    checks = [_lvl(0, "bullet", "□", style=sid("check"), left=cm(1.2), hanging=cm(0.5),
                   rpr_=_rp(color=D.PRIMARY_MID))]
    checks += [_lvl(i, "bullet", "□", left=cm(1.7), hanging=cm(0.5)) for i in range(1, 9)]

    def an(idx, lvls, name):
        return (f'<w:abstractNum w:abstractNumId="{idx}"><w:nsid w:val="1A2B3C{idx}0"/>'
                f'<w:multiLevelType w:val="{"hybridMultilevel" if idx else "multilevel"}"/>'
                f'<w:tmpl w:val="04090023"/>' + "".join(lvls) + '</w:abstractNum>')
    body = (an(0, proc, "절차") + an(1, bullets, "글머리") + an(2, checks, "체크"))
    nums = "".join(f'<w:num w:numId="{i+1}"><w:abstractNumId w:val="{i}"/></w:num>' for i in range(3))
    return body + nums

# ───────────────────────────── 문서 포장 ───────────────────────────────────
PG_W, PG_H = 11906, 16838          # A4
MAR = dict(top=cm(2.3), right=cm(2.0), bottom=cm(2.1), left=cm(2.3), header=cm(1.25), footer=cm(1.2))
TEXT_W = PG_W - MAR["left"] - MAR["right"]

def sect_pr(with_hf=True, first=False):
    refs = ""
    return (f'<w:sectPr>{refs}<w:pgSz w:w="{PG_W}" w:h="{PG_H}"/>'
            f'<w:pgMar w:top="{MAR["top"]}" w:right="{MAR["right"]}" w:bottom="{MAR["bottom"]}" '
            f'w:left="{MAR["left"]}" w:header="{MAR["header"]}" w:footer="{MAR["footer"]}" w:gutter="0"/>'
            f'<w:cols w:space="425"/><w:docGrid w:type="default" w:linePitch="360"/></w:sectPr>')

def P_sect(**kw):
    """구역 나누기용 빈 문단."""
    return f'<w:p><w:pPr><w:spacing w:after="0" w:line="20" w:lineRule="exact"/>{sect_pr(**kw)}</w:pPr></w:p>'

def _settings_extra(settings_el):
    from docx.oxml.ns import qn
    def get(tag):
        el = settings_el.find(qn(tag))
        return el
    # 필드 자동 갱신(목차/그림 번호/쪽 번호)
    if get('w:updateFields') is None:
        from docx.oxml import parse_xml
        from docx.oxml.ns import nsdecls
        el = parse_xml(f'<w:updateFields {nsdecls("w")} w:val="true"/>')
        compat = settings_el.find(qn('w:compat'))
        (compat.addprevious(el) if compat is not None else settings_el.append(el))

def build_document(body_xml, out_path, images=None, header_left="", header_right="",
                   footer_left="", title="", subject="", creator="IT 인프라 문서 라이브러리"):
    """body_xml: 본문 XML(문자열). images: {키: 파일경로} → build 전에 rId 확보용으로 쓰려면
    get_image_rids()를 먼저 호출한다."""
    from docx import Document
    from docx.oxml import parse_xml
    doc = _DOC[0]
    body = doc.element.body
    for ch in list(body): body.remove(ch)
    frag = parse_xml(f'<w:root {NS}>{body_xml}</w:root>')
    for ch in list(frag): body.append(ch)

    # 머리글/바닥글 (2번째 구역에만)
    tab_r = f'<w:tabs><w:tab w:val="right" w:pos="{TEXT_W}"/></w:tabs>'
    secs = doc.sections
    tgt = secs[-1]
    hdr = tgt.header; hdr.is_linked_to_previous = False
    ftr = tgt.footer; ftr.is_linked_to_previous = False
    def fill(part_el, xml):
        for ch in list(part_el): part_el.remove(ch)
        fr = parse_xml(f'<w:root {NS}>{xml}</w:root>')
        for ch in list(fr): part_el.append(ch)
    fill(hdr._element, P(
        [R(header_left, sz=8.5, color=D.INK_FAINT), TAB(), R(header_right, sz=8.5, color=D.INK_FAINT)],
        sp=spacing(after=0, line=260), tabs=tab_r,
        border=pbdr(bottom=bdr("bottom", 4, D.LINE, 4))))
    fill(ftr._element, P(
        [R(footer_left, sz=8.5, color=D.INK_FAINT), TAB(),
         FLD(" PAGE ", "1", sz=8.5, color=D.INK_SOFT, b=True),
         R(" / ", sz=8.5, color=D.INK_FAINT),
         FLD(" NUMPAGES ", "1", sz=8.5, color=D.INK_FAINT)],
        sp=spacing(before=pt(2), after=0, line=260), tabs=tab_r))
    if len(secs) > 1:      # 표지 구역에는 머리글/바닥글 없음
        for s in secs[:-1]:
            s.header.is_linked_to_previous = False
            s.footer.is_linked_to_previous = False
            fill(s.header._element, P('', sp=spacing(after=0, line=20, rule="exact")))
            fill(s.footer._element, P('', sp=spacing(after=0, line=20, rule="exact")))

    cp = doc.core_properties
    cp.title, cp.subject, cp.author = title, subject, creator
    cp.last_modified_by = creator
    cp.comments = "IT 인프라 문서 작성 라이브러리"
    doc.save(out_path)
    return out_path

_DOC = [None]

def new_document():
    """스타일/번호 정의를 교체한 빈 문서를 준비한다."""
    from docx import Document
    from docx.oxml import parse_xml
    doc = Document()
    # styles.xml 교체
    st = doc.styles.element
    for ch in list(st): st.remove(ch)
    for ch in list(parse_xml(f'<w:root {NS}>{styles_xml()}</w:root>')): st.append(ch)
    # numbering.xml 교체
    nm = doc.part.numbering_part.element
    for ch in list(nm): nm.remove(ch)
    for ch in list(parse_xml(f'<w:root {NS}>{numbering_xml()}</w:root>')): nm.append(ch)
    _settings_extra(doc.settings.element)
    _strip_unused(doc)
    _fix_theme(doc)
    _DOC[0] = doc
    return doc

def _strip_unused(doc):
    """Word 2010 호환용 stylesWithEffects, 빈 customXml 등 쓰지 않는 부분을 뺀다."""
    part = doc.part
    for rid, rel in list(part.rels.items()):
        t = rel.reltype
        if t.endswith("/stylesWithEffects") or "customXml" in t:
            try: part.drop_rel(rid)
            except Exception: pass

def _fix_theme(doc):
    """테마 글꼴·주색을 라이브러리 값으로 바꾼다(새 표·도형·글상자가 같은 글꼴이 되도록)."""
    import re
    for prt in doc.part.package.iter_parts():
        if str(prt.partname).endswith("theme1.xml"):
            try: xml = prt.blob.decode("utf-8")
            except Exception: return
            xml = re.sub(r'(<a:(?:majorFont|minorFont)>\s*)<a:latin[^/]*/>\s*<a:ea[^/]*/>',
                         lambda m: (f'{m.group(1)}<a:latin typeface="{D.FONT_LATIN}"/>'
                                    f'<a:ea typeface="{D.FONT_EA}"/>'), xml)
            for name, val in (("accent1", D.PRIMARY), ("accent2", D.PRIMARY_MID),
                              ("accent3", D.INK_SOFT), ("accent4", D.OK),
                              ("accent5", D.WARN), ("accent6", D.DANGER)):
                xml = re.sub(r'<a:%s>\s*<a:srgbClr val="[0-9A-Fa-f]{6}"/>\s*</a:%s>' % (name, name),
                             '<a:%s><a:srgbClr val="%s"/></a:%s>' % (name, val, name), xml)
            prt._blob = xml.encode("utf-8")
            return

def add_image(doc, path):
    rid, _ = doc.part.get_or_add_image(path)
    return rid
