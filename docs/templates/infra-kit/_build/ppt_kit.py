# -*- coding: utf-8 -*-
"""PPT 공통 도구: 16:9 프레젠테이션(테마 색·글꼴 = 맑은 고딕) + 도형/텍스트/아이콘/연결선/표 헬퍼.

원칙
- 모든 색은 RGB 직접 지정(테마 색 X) → 다른 회사 양식에 붙여도 색이 바뀌지 않는다.
- 모든 글꼴은 '맑은 고딕' 직접 지정 → Office 2016 기본 글꼴, 양식이 달라도 동일.
- 요소 하나 = 그룹 하나(이름 지정) → 한 번 클릭으로 선택·복사.
- 도형 스타일(p:style) 제거 → 그림자·테마 효과가 끼어들지 않음.
"""
import copy
import datetime
from lxml import etree
from pptx import Presentation
from pptx.util import Emu, Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

FONT = "맑은 고딕"
EMU = 914400
SW, SH = 13.333, 7.5

P = dict(
    ink="1E2A36", navy="1F3A5F", s700="3A4856", s600="4E5D6C", s500="677789", s400="8C99A8", s300="B4BFCB",
    s200="D5DCE4", s150="E4E9EF", s100="EEF2F6", s50="F6F8FA", white="FFFFFF",
    svc="2F6DB5", mgmt="0B7C7C", bak="377F3F", san="BF5B14", ic="7A55B3", con="5E6B78", pwr="C73E3A",
    wan="3A4856",
    # 옅은 바탕(영역·카드)
    t_svc="EAF1FA", t_mgmt="E5F2F2", t_bak="EAF3EB", t_san="FBEFE6", t_ic="F1ECF8", t_pwr="FAEAE9", t_ext="F1F3F5",
    # 상태
    ok="2E8B57", warn="D98E04", crit="C73E3A", info="2F6DB5", t_ok="E6F3EC", t_warn="FDF3DE", t_crit="FAE7E6",
    t_info="E8F0FA",
)


def I(v):
    return int(round(v * EMU))


def rgb(h):
    return RGBColor.from_string(h)


# ── 프레젠테이션 ───────────────────────────────────────────
THEME_COLORS = dict(dk1="1E2A36", lt1="FFFFFF", dk2="1F3A5F", lt2="EEF2F6", accent1="2F6DB5", accent2="BF5B14",
                    accent3="377F3F", accent4="7A55B3", accent5="0B7C7C", accent6="C73E3A", hlink="2F6DB5",
                    folHlink="7A55B3")


def _fix_theme(prs):
    theme_part = prs.slide_master.part.part_related_by(RT.THEME)
    root = etree.fromstring(theme_part.blob)
    a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    root.set("name", "IT 인프라 문서")
    cs = root.find(f".//{{{a}}}clrScheme")
    cs.set("name", "IT 인프라")
    for tag, val in THEME_COLORS.items():
        el = cs.find(f"{{{a}}}{tag}")
        for ch in list(el):
            el.remove(ch)
        s = etree.SubElement(el, f"{{{a}}}srgbClr")
        s.set("val", val)
    fs = root.find(f".//{{{a}}}fontScheme")
    fs.set("name", "IT 인프라 맑은 고딕")
    for which in ("majorFont", "minorFont"):
        f = fs.find(f"{{{a}}}{which}")
        f.find(f"{{{a}}}latin").set("typeface", FONT)
        f.find(f"{{{a}}}ea").set("typeface", FONT)
        for fnt in f.findall(f"{{{a}}}font"):
            if fnt.get("script") == "Hang":
                fnt.set("typeface", FONT)
    # 효과 스타일에서 그림자 제거
    for es in root.iter(f"{{{a}}}effectStyle"):
        for ch in list(es):
            es.remove(ch)
        etree.SubElement(es, f"{{{a}}}effectLst")
    theme_part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _scale_xfrm(el, sx):
    for off in el.iter(qn("a:off")):
        off.set("x", str(int(int(off.get("x")) * sx)))
    for ext in el.iter(qn("a:ext")):
        if ext.getparent().tag == qn("a:xfrm"):
            ext.set("cx", str(int(int(ext.get("cx")) * sx)))


def new_presentation(footer="IT 인프라 문서 구성요소 라이브러리"):
    prs = Presentation()
    # 기본 템플릿에 남아 있는 파일 정보(작성자 · 설명 · 2013년 날짜)를 정리
    cp = prs.core_properties
    cp.title, cp.subject, cp.author, cp.last_modified_by, cp.comments = footer, "", "", "", ""
    cp.revision = 1
    cp.created = cp.modified = datetime.datetime.now(datetime.timezone.utc).replace(tzinfo=None, microsecond=0)
    old_w = prs.slide_width
    prs.slide_width = I(SW)
    prs.slide_height = I(SH)
    sx = prs.slide_width / old_w
    _fix_theme(prs)
    master = prs.slide_master
    _scale_xfrm(master._element, sx)
    for lay in prs.slide_layouts:
        _scale_xfrm(lay._element, sx)
    # 마스터 제목 스타일: 24pt 굵게, 왼쪽 정렬, 잉크색
    tstyle = master._element.find(qn("p:txStyles")).find(qn("p:titleStyle"))
    l1 = tstyle.find(qn("a:lvl1pPr"))
    l1.set("algn", "l")
    d = l1.find(qn("a:defRPr"))
    d.set("sz", "2400")
    d.set("b", "1")
    for ch in list(d):
        d.remove(ch)
    sf = etree.SubElement(d, qn("a:solidFill"))
    etree.SubElement(sf, qn("a:srgbClr")).set("val", P["ink"])
    etree.SubElement(d, qn("a:latin")).set("typeface", "+mj-lt")
    etree.SubElement(d, qn("a:ea")).set("typeface", "+mj-ea")
    etree.SubElement(d, qn("a:cs")).set("typeface", "+mj-cs")
    # 마스터 제목 자리 위치
    for sp in master.placeholders:
        if sp.placeholder_format.type == 1:  # TITLE
            sp.left, sp.top, sp.width, sp.height = I(0.5), I(0.42), I(12.33), I(0.62)
            bp = sp._element.find(qn("p:txBody")).find(qn("a:bodyPr"))
            bp.set("anchor", "b")
            bp.set("lIns", "0")
            bp.set("rIns", "0")
            bp.set("tIns", "0")
            bp.set("bIns", "0")
    # 제목만 레이아웃(5)·빈 화면(6)에 바닥글·쪽번호(필드) 넣기
    for idx in (5,):
        lay = prs.slide_layouts[idx]
        for sp in list(lay.placeholders):
            if sp.placeholder_format.type == 1:
                sp.left, sp.top, sp.width, sp.height = I(0.5), I(0.42), I(12.33), I(0.62)
        _layout_footer(lay, footer)
    return prs


def _layout_footer(layout, footer):
    """레이아웃에 바닥글 텍스트 + 쪽 번호 필드 (슬라이드 복사 시 따라가지 않음)."""
    tree = layout.shapes._spTree
    ids = [int(e.get("id")) for e in tree.iter(qn("p:cNvPr"))]
    nid = max(ids + [1]) + 1

    def tb(idx, name, x, y, w, h, algn, run_xml):
        xml = (f'<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
               f'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
               f'<p:nvSpPr><p:cNvPr id="{idx}" name="{name}"/><p:cNvSpPr txBox="1"/><p:nvPr userDrawn="1"/></p:nvSpPr>'
               f'<p:spPr><a:xfrm><a:off x="{I(x)}" y="{I(y)}"/><a:ext cx="{I(w)}" cy="{I(h)}"/></a:xfrm>'
               f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
               f'<p:txBody><a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"><a:noAutofit/></a:bodyPr>'
               f'<a:lstStyle/><a:p><a:pPr algn="{algn}"/>{run_xml}</a:p></p:txBody></p:sp>')
        tree.append(etree.fromstring(xml))

    rpr = (f'<a:rPr lang="ko-KR" sz="800"><a:solidFill><a:srgbClr val="{P["s500"]}"/></a:solidFill>'
           f'<a:latin typeface="{FONT}"/><a:ea typeface="{FONT}"/></a:rPr>')
    tb(nid, "바닥글", 0.5, 7.08, 8, 0.25, "l", f'<a:r>{rpr}<a:t>{footer}</a:t></a:r>')
    tb(nid + 1, "쪽 번호", 11.83, 7.08, 1.0, 0.25, "r",
       f'<a:fld id="{{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}}" type="slidenum">{rpr}<a:t>‹#›</a:t></a:fld>')


# ── 텍스트 ──────────────────────────────────────────────────
def _zero_margins(tf, l=0, r=0, t=0, b=0):
    tf.margin_left, tf.margin_right, tf.margin_top, tf.margin_bottom = I(l), I(r), I(t), I(b)


def set_font(run, size=None, bold=None, color=None, font=FONT, italic=None, strike=False):
    f = run.font
    if size is not None:
        f.size = Pt(size)
    if bold is not None:
        f.bold = bold
    if italic is not None:
        f.italic = italic
    if color:
        f.color.rgb = rgb(color)
    rpr = run._r.get_or_add_rPr()
    rpr.set("lang", "ko-KR")
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = rpr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rpr, qn(tag))
        el.set("typeface", font)
    if strike:
        rpr.set("strike", "sngStrike")


def _para(tf, content, size, bold, color, align=None, first=True, spacing=None, font=FONT):
    """content: str | [(text, {size,bold,color})...]"""
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    if align:
        p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[align]
    if spacing:
        p.space_after = Pt(spacing)
    runs = content if isinstance(content, list) else [(content, {})]
    for t, st in runs:
        r = p.add_run()
        r.text = t
        set_font(r, st.get("size", size), st.get("bold", bold), st.get("color", color), st.get("font", font),
                 st.get("italic"), st.get("strike", False))
    return p


def text(sl, x, y, w, h, content, size=11, bold=False, color=None, align="l", anchor="t", wrap=True, name=None,
         spacing=None, line_spacing=None, margins=(0, 0, 0, 0)):
    """content: str | 줄 목록(각 줄 = str 또는 run 목록)."""
    tb = sl.shapes.add_textbox(I(x), I(y), I(w), I(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.auto_size = None
    _zero_margins(tf, *margins)
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
    lines = content if (isinstance(content, list) and content and not isinstance(content[0], tuple)) else [content]
    for k, ln in enumerate(lines):
        p = _para(tf, ln, size, bold, color or P["ink"], align, first=(k == 0), spacing=spacing)
        if line_spacing:
            p.line_spacing = line_spacing
    if name:
        tb.name = name
    return tb


# ── 도형 ────────────────────────────────────────────────────
DASH = {"solid": None, "dash": "dash", "sysDash": "sysDash", "sysDot": "sysDot", "lgDash": "lgDash",
        "dashDot": "dashDot", "lgDashDot": "lgDashDot"}


def _strip_style(sh):
    st = sh._element.find(qn("p:style"))
    if st is not None:
        sh._element.remove(st)


def line_fmt(ln_el_owner, color=None, w=0.75, dash=None, head=None, tail=None, cap=None):
    """a:ln 직접 작성 (선 색/두께/점선/화살표)."""
    spPr = ln_el_owner._element.spPr
    ln = spPr.find(qn("a:ln"))
    if ln is not None:
        spPr.remove(ln)
    ln = etree.SubElement(spPr, qn("a:ln"))
    if color is None:
        etree.SubElement(ln, qn("a:noFill"))
        return
    ln.set("w", str(int(w * 12700)))
    if cap:
        ln.set("cap", cap)
    sf = etree.SubElement(ln, qn("a:solidFill"))
    etree.SubElement(sf, qn("a:srgbClr")).set("val", color)
    if dash and dash != "solid":
        etree.SubElement(ln, qn("a:prstDash")).set("val", dash)
    etree.SubElement(ln, qn("a:round"))
    if head:
        e = etree.SubElement(ln, qn("a:headEnd"))
        e.set("type", head)
        e.set("w", "med")
        e.set("len", "med")
    if tail:
        e = etree.SubElement(ln, qn("a:tailEnd"))
        e.set("type", tail)
        e.set("w", "med")
        e.set("len", "med")


def shape(sl, prst, x, y, w, h, fill=None, line=None, lw=0.75, dash=None, radius=None, name=None,
          txt=None, size=10, bold=False, color=None, align="c", anchor="m", margins=(0.04, 0.04, 0.02, 0.02),
          wrap=True, adj=None):
    sh = sl.shapes.add_shape(prst, I(x), I(y), I(w), I(h))
    _strip_style(sh)
    if fill:
        sh.fill.solid()
        sh.fill.fore_color.rgb = rgb(fill)
    else:
        sh.fill.background()
    line_fmt(sh, line, lw, dash)
    if radius is not None and prst == MSO_SHAPE.ROUNDED_RECTANGLE:
        sh.adjustments[0] = max(0.0, min(0.5, radius / min(w, h)))
    if adj:
        for k, v in adj.items():
            sh.adjustments[k] = v
    tf = sh.text_frame
    tf.word_wrap = wrap
    _zero_margins(tf, *margins)
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
    if txt is not None:
        lines = txt if isinstance(txt, list) and txt and not isinstance(txt[0], tuple) else [txt]
        for k, ln in enumerate(lines):
            _para(tf, ln, size, bold, color or P["ink"], align, first=(k == 0))
    if name:
        sh.name = name
    return sh


def rect(sl, x, y, w, h, **kw):
    return shape(sl, MSO_SHAPE.RECTANGLE, x, y, w, h, **kw)


def rrect(sl, x, y, w, h, r=0.08, **kw):
    return shape(sl, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, radius=r, **kw)


def oval(sl, x, y, w, h, **kw):
    return shape(sl, MSO_SHAPE.OVAL, x, y, w, h, **kw)


def line(sl, x1, y1, x2, y2, color=None, w=1.0, dash=None, head=None, tail=None, name=None):
    c = sl.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, I(x1), I(y1), I(x2), I(y2))
    _strip_style(c)
    line_fmt(c, color or P["s400"], w, dash, head, tail)
    if name:
        c.name = name
    return c


def connect(sl, a, ai, b, bi, kind="elbow", color=None, w=1.5, dash=None, head=None, tail=None, name=None, adj=None,
            bulge=None):
    """두 도형의 연결점을 잇는 연결선(도형을 옮기면 따라감). 연결점: 0 위, 1 왼쪽, 2 아래, 3 오른쪽 (사각형 기준)."""
    ct = {"elbow": MSO_CONNECTOR.ELBOW, "straight": MSO_CONNECTOR.STRAIGHT}[kind]
    c = sl.shapes.add_connector(ct, 0, 0, I(1), I(1))
    _strip_style(c)
    c.begin_connect(a, ai)
    c.end_connect(b, bi)
    line_fmt(c, color or P["s500"], w, dash, head, tail)
    if bulge and kind == "elbow":
        # 같은 쪽 연결점끼리(예: 오른쪽↔오른쪽) 이을 때 바깥으로 꺾이도록: 끝점을 살짝 옮기고 adj1을 크게
        d = 0.02
        c.end_x = c.end_x + I(d)
        adj = (bulge / d) * (1 if ai == 3 else -1)
    if adj is not None and kind == "elbow":
        av = c._element.spPr.find(qn("a:prstGeom")).find(qn("a:avLst"))
        for ch in list(av):
            av.remove(ch)
        gd = etree.SubElement(av, qn("a:gd"))
        gd.set("name", "adj1")
        gd.set("fmla", f"val {int(adj * 100000)}")
    if name:
        c.name = name
    return c


def group(sl, shapes, name):
    g = sl.shapes.add_group_shape(shapes)
    g.name = name
    return g


# ── 아이콘 (네이티브 도형) ───────────────────────────────────
_ICON_CACHE = {}


def icon(sl, key, x, y, size, color=None, name=None):
    from icons import ICONS
    if not _ICON_CACHE:
        for ic in ICONS:
            _ICON_CACHE[ic.key] = ic
    ic = _ICON_CACHE[key]
    sh = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, I(x), I(y), I(size), I(size))
    _strip_style(sh)
    spPr = sh._element.spPr
    prst = spPr.find(qn("a:prstGeom"))
    cg = etree.fromstring('<root xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                          + ic.custgeom() + "</root>")[0]
    prst.addprevious(cg)
    spPr.remove(prst)
    sh.fill.solid()
    sh.fill.fore_color.rgb = rgb(color or P["navy"])
    line_fmt(sh, None)
    cnv = sh._element.find(qn("p:nvSpPr")).find(qn("p:cNvSpPr"))
    lk = etree.SubElement(cnv, qn("a:spLocks"))
    lk.set("noChangeAspect", "1")
    sh.name = name or f"아이콘-{ic.name}"
    return sh


# ── 슬라이드 ────────────────────────────────────────────────
def slide(prs, title=None, eyebrow=None, layout=5, bg=None):
    sl = prs.slides.add_slide(prs.slide_layouts[layout])
    if title is not None and sl.shapes.title is not None:
        sl.shapes.title.text = title
        for p in sl.shapes.title.text_frame.paragraphs:
            for r in p.runs:
                set_font(r, None, None, None)
    if eyebrow:
        text(sl, 0.5, 0.2, 9, 0.22, eyebrow, size=9, bold=True, color=P["svc"], name="분류")
    if bg:
        b = sl.background.fill
        b.solid()
        b.fore_color.rgb = rgb(bg)
    return sl


def notes(sl, s):
    sl.notes_slide.notes_text_frame.text = s


# ── 표 ─────────────────────────────────────────────────────
def _cell_border(cell, color, w=0.75, sides=("L", "R", "T", "B")):
    tcPr = cell._tc.get_or_add_tcPr()
    for s in ("L", "R", "T", "B"):
        tag = qn(f"a:ln{s}")
        old = tcPr.find(tag)
        if old is not None:
            tcPr.remove(old)
    for s in ("L", "R", "T", "B"):
        ln = etree.SubElement(tcPr, qn(f"a:ln{s}"))
        if s in sides and color:
            ln.set("w", str(int(w * 12700)))
            sf = etree.SubElement(ln, qn("a:solidFill"))
            etree.SubElement(sf, qn("a:srgbClr")).set("val", color)
        else:
            ln.set("w", "0")
            etree.SubElement(ln, qn("a:noFill"))
    # tcPr 자식 순서: lnL lnR lnT lnB ... solidFill 은 뒤
    fill = tcPr.find(qn("a:solidFill"))
    if fill is not None:
        tcPr.remove(fill)
        tcPr.append(fill)


def table(sl, x, y, col_w, rows, row_h=0.3, header=True, size=9.5, hsize=9.5, name=None, head_fill=None,
          zebra=True, align=None, fills=None, bold_cols=(), colors=None):
    """rows: 2차원 목록(첫 행 = 머리글). fills/colors: {(r,c): hex} 셀 채움/글자색."""
    nr, nc = len(rows), len(col_w)
    gf = sl.shapes.add_table(nr, nc, I(x), I(y), I(sum(col_w)), I(row_h * nr))
    tbl = gf.table
    tblPr = tbl._tbl.tblPr
    for attr in ("firstRow", "bandRow"):
        tblPr.set(attr, "0")
    sid = tblPr.find(qn("a:tableStyleId"))
    if sid is not None:
        sid.text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"  # 스타일 없음, 눈금 없음
    for j, w in enumerate(col_w):
        tbl.columns[j].width = I(w)
    for i in range(nr):
        tbl.rows[i].height = I(row_h)
        for j in range(nc):
            cell = tbl.cell(i, j)
            v = rows[i][j]
            is_h = header and i == 0
            cell.margin_left = I(0.07)
            cell.margin_right = I(0.05)
            cell.margin_top = I(0.02)
            cell.margin_bottom = I(0.02)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            a = (align[j] if align else "l")
            txt = "" if v is None else v
            _para(tf, txt if isinstance(txt, list) else str(txt), hsize if is_h else size,
                  True if is_h or j in bold_cols else False,
                  P["white"] if (is_h and (head_fill or P["navy"]) not in (P["s100"], P["s150"])) else (colors or {}).get((i, j), P["ink"]),
                  a)
            f = (fills or {}).get((i, j))
            if is_h:
                f = head_fill or P["navy"]
            elif f is None and zebra and i % 2 == 0:
                f = P["s50"]
            cell.fill.solid() if f else cell.fill.background()
            if f:
                cell.fill.fore_color.rgb = rgb(f)
            _cell_border(cell, P["s200"], 0.75, sides=("B",) if not is_h else ())
    if name:
        gf.name = name
    return gf
