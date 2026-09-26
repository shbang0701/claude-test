# -*- coding: utf-8 -*-
"""Excel 도면용 격자 캔버스 + 셀 스타일 (openpyxl).

모든 시트는 같은 격자(열 너비 2.0 / 행 높이 12pt ≒ 4.2mm 정사각형)를 쓴다.
그래서 어느 시트에서 복사한 부품이든 다른 시트에 붙이면 크기가 같다.
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, NamedStyle
from openpyxl.utils import get_column_letter as COL
from openpyxl.worksheet.pagebreak import Break
from openpyxl.formatting.rule import FormulaRule

FONT = "맑은 고딕"
MONO = "Consolas"

GRID_W = 2.0      # 열 너비(맑은 고딕 11 기준 16px ≒ 4.2mm)
GRID_H = 12.0     # 행 높이 pt (≒ 4.2mm)
NCOLS = 70        # 격자를 명시적으로 지정할 열 수

# ── 팔레트 (Excel·PPT 공통) ─────────────────────────────────
P = dict(
    ink="1E2A36", navy="1F3A5F", s700="3A4856", s600="4E5D6C", s500="677789",
    s400="8C99A8", s300="B4BFCB", s200="D5DCE4", s150="E4E9EF", s100="EEF2F6",
    s50="F6F8FA", white="FFFFFF",
    svc="2F6DB5", mgmt="0B7C7C", bak="377F3F", san="BF5B14", ic="7A55B3",
    con="5E6B78", pwr="C73E3A",
    tape="FFF1BF", warn_bg="F8D3D0", warn_fg="9B1C17",
    disk="C3CCD6",
)

# 용도(연결 종류) → 색, 표에 쓰는 이름
USE = [
    ("svc", "서비스"), ("mgmt", "관리"), ("bak", "백업"), ("san", "스토리지"),
    ("ic", "인터커넥트"), ("con", "콘솔"), ("pwr", "전원"),
]
USE_NAME = dict(USE)
USE_KEY = {v: k for k, v in USE}
USE_LIST = [n for _, n in USE] + ["기타"]


def side(style="thin", color=None):
    return Side(style=style, color=color or P["s400"])


WHITE_MED = Side(style="medium", color=P["white"])
SEP = Side(style="medium", color="E4E9EF")   # 포트·디스크 구분선 = 모듈 바탕색
NOSIDE = Side(style=None)


def fill(c):
    return PatternFill("solid", fgColor=c)


def font(sz=8, b=False, col=None, name=FONT, i=False):
    return Font(name=name, size=sz, bold=b, italic=i, color=col or P["ink"])


def align(h="left", v="center", wrap=False, rot=0, shrink=False, indent=0):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap, text_rotation=rot,
                     shrink_to_fit=shrink, indent=indent)


# ── 셀 스타일(홈 > 셀 스타일 목록에 표시됨) ─────────────────────────
def _port_style(name, color):
    # 테두리는 스타일에 넣지 않는다(적용해도 기존 테두리 유지) → NO_BORDER_STYLES
    return NamedStyle(name=name, font=font(8, True, P["white"]), fill=fill(color),
                      alignment=align("center"), number_format="General")


NO_BORDER_STYLES = {f"포트 {n}" for _, n in USE} | {"포트 미연결", "디스크 장착", "디스크 빈칸", "모듈 채움",
                                                     "포트명", "모듈 제목", "번호 배지", "구역 제목", "메모",
                                                     "RAID 표시"}


NO_FILL_STYLES = {"RAID 표시"}   # 적용해도 바탕색(모듈 회색 등)은 그대로


def _defs():
    d = []
    for k, n in USE:
        d.append(_port_style(f"포트 {n}", P[k]))
    d.append(NamedStyle(name="포트 미연결", font=font(8, False, P["s400"]), fill=fill(P["white"]),
                        alignment=align("center")))
    d.append(NamedStyle(name="포트명", font=font(6.5, False, P["s500"]), alignment=align("center")))
    d.append(NamedStyle(name="모듈 제목", font=font(7, True, P["ink"]), alignment=align("left", indent=0)))
    d.append(NamedStyle(name="모듈 채움", fill=fill(P["s150"])))
    # 디스크: 칸 안에 용량을 작은 글씨로 적는다 (빈 베이 = '빈')
    d.append(NamedStyle(name="디스크 장착", font=font(6, False, P["ink"]), fill=fill(P["disk"]),
                        alignment=align("center", "center", wrap=True)))
    d.append(NamedStyle(name="디스크 빈칸", font=font(6, False, P["s400"]), fill=fill(P["white"]),
                        alignment=align("center", "center", wrap=True)))
    d.append(NamedStyle(name="RAID 표시", font=font(6.5, True, P["ink"]), alignment=align("left")))
    d.append(NamedStyle(name="번호 배지", font=font(8, True, P["white"]), fill=fill(P["navy"]),
                        alignment=align("center")))
    d.append(NamedStyle(name="케이블 라벨", font=font(8.5, False, P["ink"], MONO), fill=fill(P["tape"]),
                        border=Border(bottom=side("thin", P["s200"])), alignment=align("left", indent=0)))
    d.append(NamedStyle(name="구역 제목", font=font(9, True, P["navy"]), alignment=align("left")))
    d.append(NamedStyle(name="메모", font=font(7, False, P["s500"]), alignment=align("left")))
    d.append(NamedStyle(name="표 머리글", font=font(8, True, P["ink"]), fill=fill(P["s100"]),
                        border=Border(top=side("thin", P["s400"]), bottom=side("thin", P["s400"])),
                        alignment=align("left")))
    d.append(NamedStyle(name="표 본문", font=font(8.5, False, P["ink"]),
                        border=Border(bottom=side("thin", P["s200"])), alignment=align("left")))
    d.append(NamedStyle(name="표 번호", font=font(8.5, True, P["ink"]),
                        border=Border(bottom=side("thin", P["s200"])), alignment=align("center")))
    return d


def new_workbook(title=None):
    wb = Workbook()
    wb.properties.creator = None          # 파일 정보(작성자)는 비워 둠 — 회사에서 저장하면 사용자 이름이 들어감
    wb.properties.title = title
    wb._named_styles["Normal"].font = Font(name=FONT, size=11)
    for ns in _defs():
        wb.add_named_style(ns)
    wb.remove(wb.active)
    return wb


class Canvas:
    """격자 시트 래퍼. 셀 병합은 finalize()에서 한꺼번에 처리."""

    def __init__(self, wb, title, tab=None, zoom=100):
        ws = wb.create_sheet(title)
        self.ws = ws
        ws.sheet_format.defaultColWidth = GRID_W
        ws.sheet_format.baseColWidth = 2
        ws.sheet_format.defaultRowHeight = GRID_H
        ws.sheet_format.customHeight = True
        ws.sheet_view.showGridLines = False
        ws.sheet_view.zoomScale = zoom
        for c in range(1, NCOLS + 1):
            ws.column_dimensions[COL(c)].width = GRID_W
        if tab:
            ws.sheet_properties.tabColor = tab
        self.merges = []
        self.max_row = 1
        self.edges = {}   # (r,c) -> {side: Side}  모듈 외곽선 위치(포트가 덮어쓰지 않도록)

    # ── 기본 셀 조작 ───────────────────────────
    def c(self, r, c):
        self.max_row = max(self.max_row, r)
        return self.ws.cell(r, c)

    def h(self, r, pt):
        self.ws.row_dimensions[r].height = pt

    def put(self, r, c, v=None, style=None, f=None, fl=None, al=None, nf=None):
        cell = self.c(r, c)
        if style:
            cell.style = style
        if v is not None:
            cell.value = v
        if f is not None:
            cell.font = f
        if fl is not None:
            cell.fill = fl
        if al is not None:
            cell.alignment = al
        if nf is not None:
            cell.number_format = nf
        return cell

    def text(self, r, c, v, sz=7, b=False, col=None, h="left", v_="center", name=FONT, i=False,
             wrap=False, rot=0, indent=0):
        return self.put(r, c, v, f=font(sz, b, col, name, i), al=align(h, v_, wrap, rot, indent=indent))

    def add_border(self, r, c, left=None, right=None, top=None, bottom=None):
        cell = self.c(r, c)
        b = cell.border
        cell.border = Border(left=left or b.left, right=right or b.right,
                             top=top or b.top, bottom=bottom or b.bottom)

    def set_border(self, r, c, left=NOSIDE, right=NOSIDE, top=NOSIDE, bottom=NOSIDE):
        self.c(r, c).border = Border(left=left, right=right, top=top, bottom=bottom)

    def fill_rect(self, r1, c1, r2, c2, color):
        for r in range(r1, r2 + 1):
            for cc in range(c1, c2 + 1):
                self.c(r, cc).fill = fill(color)

    def style_rect(self, r1, c1, r2, c2, style):
        for r in range(r1, r2 + 1):
            for cc in range(c1, c2 + 1):
                self.c(r, cc).style = style

    def outline(self, r1, c1, r2, c2, s, record=True):
        for cc in range(c1, c2 + 1):
            self.add_border(r1, cc, top=s)
            self.add_border(r2, cc, bottom=s)
            if record:
                self.edges.setdefault((r1, cc), {})["top"] = s
                self.edges.setdefault((r2, cc), {})["bottom"] = s
        for r in range(r1, r2 + 1):
            self.add_border(r, c1, left=s)
            self.add_border(r, c2, right=s)
            if record:
                self.edges.setdefault((r, c1), {})["left"] = s
                self.edges.setdefault((r, c2), {})["right"] = s

    def cell_box(self, r, c, base=None):
        """포트·디스크 칸 테두리: 기본 흰 굵은 선, 단 모듈 외곽선 위이면 외곽선 유지."""
        base = base or SEP
        e = self.edges.get((r, c), {})
        self.c(r, c).border = Border(left=e.get("left", base), right=e.get("right", base),
                                     top=e.get("top", base), bottom=e.get("bottom", base))

    def clear_rect(self, r1, c1, r2, c2):
        for r in range(r1, r2 + 1):
            for cc in range(c1, c2 + 1):
                cell = self.c(r, cc)
                cell.style = "Normal"
                cell.value = None

    def merge(self, r1, c1, r2, c2):
        if (r1, c1) != (r2, c2):
            self.merges.append((r1, c1, r2, c2))

    def finalize(self, extra_rows=40):
        ws = self.ws
        for r in range(1, self.max_row + extra_rows):
            if ws.row_dimensions[r].height is None:
                ws.row_dimensions[r].height = GRID_H
        for r1, c1, r2, c2 in self.merges:
            a = ws.cell(r1, c1)
            # 병합 영역 바깥 테두리를 모서리 셀에서 모아 앵커에 지정 (openpyxl이 가장자리에 복사)
            lb = ws.cell(r1, c1).border.left
            rb = ws.cell(r1, c2).border.right
            tb = ws.cell(r1, c1).border.top
            bb = ws.cell(r2, c1).border.bottom
            a.border = Border(left=lb, right=rb, top=tb, bottom=bb)
            ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
        self.merges = []

    # ── 인쇄 설정 ───────────────────────────────
    def page_setup(self, last_col, last_row, title_rows="1:3", landscape=True, breaks=(),
                   footer_left="&F  ·  &A", fit_width=True):
        ws = self.ws
        ws.print_area = f"B1:{COL(last_col)}{last_row}"
        ws.page_setup.orientation = "landscape" if landscape else "portrait"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        if fit_width:
            # 100% 인쇄가 기본. 1페이지 너비 맞춤은 너비가 넘칠 때만 줄이는 안전장치.
            ws.sheet_properties.pageSetUpPr.fitToPage = True
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0
        else:
            ws.page_setup.scale = 100
        m = ws.page_margins
        m.left = m.right = 0.39   # 10mm
        m.top = m.bottom = 0.47   # 12mm
        m.header = m.footer = 0.2
        ws.print_options.horizontalCentered = True
        if title_rows:
            ws.print_title_rows = title_rows
        # 글꼴 코드는 넣지 않음(통합 문서 기본 글꼴 = 맑은 고딕을 따름)
        ws.oddFooter.left.text = footer_left
        ws.oddFooter.left.size = 7
        ws.oddFooter.right.text = "&P / &N"
        ws.oddFooter.right.size = 7
        for b in breaks:
            ws.row_breaks.append(Break(id=b))


def add_cf(ws, rng, formula, fill_color=None, font_color=None, bold=None, stop=False):
    kw = {}
    if fill_color:
        kw["fill"] = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
    if font_color or bold is not None:
        kw["font"] = Font(color=font_color, bold=bold)
    ws.conditional_formatting.add(rng, FormulaRule(formula=[formula], stopIfTrue=stop, **kw))


def postprocess(path):
    """저장 후 styles.xml 수정: NO_BORDER_STYLES 셀 스타일은 '테두리 포함' 해제."""
    import zipfile, shutil, os, re
    from lxml import etree
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/styles.xml":
                root = etree.fromstring(data)
                xfs = root.find("m:cellStyleXfs", ns)
                for cs in root.find("m:cellStyles", ns):
                    if cs.get("name") in NO_BORDER_STYLES:
                        xf = xfs[int(cs.get("xfId"))]
                        xf.set("applyBorder", "0")
                        if cs.get("name") in NO_FILL_STYLES:
                            xf.set("applyFill", "0")
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            zout.writestr(item, data)
    shutil.move(tmp, path)
