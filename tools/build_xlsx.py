# -*- coding: utf-8 -*-
"""library/04_Excel_시각서식.xlsx
   도형을 쓰지 않고 행 높이 · 열 너비 · 테두리 · 채우기 · 조건부 서식으로 만드는 시각 요소와,
   실무에서 자주 쓰는 함수 · 리본 기능을 미리 적용해 둔 파일."""
import os, sys
HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import design as D
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import (DataBarRule, IconSetRule, ColorScaleRule, CellIsRule,
                                      FormulaRule)
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.workbook.defined_name import DefinedName

OUT = os.path.join(ROOT, "library", "04_Excel_시각서식.xlsx")
F = D.FONT_EA

# ── 기본 도구 ──────────────────────────────────────────────────────────────
def font(size=10, bold=False, color=D.INK, italic=False, name=None):
    return Font(name=name or F, size=size, bold=bold, color="FF" + color, italic=italic)

def fill(c): return PatternFill("solid", fgColor="FF" + c)
def side(c=D.LINE, style="thin"): return Side(style=style, color="FF" + c)
def al(h="left", v="center", wrap=True, indent=0):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap, indent=indent)

def put(ws, r, c, v=None, size=10, bold=False, color=D.INK, bg=None, h="left", v_=" center",
        wrap=False, indent=0, border=None, fmt=None, italic=False, name=None):
    cell = ws.cell(r, c)
    if v is not None: cell.value = v
    cell.font = font(size, bold, color, italic, name)
    cell.alignment = al(h, "center", wrap, indent)
    if bg: cell.fill = fill(bg)
    if border: cell.border = border
    if fmt: cell.number_format = fmt
    return cell

def merge(ws, r1, c1, r2, c2):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)

def boxed(ws, r1, c1, r2, c2, color=D.LINE, style="thin", bg=None):
    """셀 범위를 테두리로 감싼다(도형 없이 상자 만들기)."""
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cur = ws.cell(r, c)
            b = Border(
                top=side(color, style) if r == r1 else cur.border.top,
                bottom=side(color, style) if r == r2 else cur.border.bottom,
                left=side(color, style) if c == c1 else cur.border.left,
                right=side(color, style) if c == c2 else cur.border.right)
            cur.border = b
            if bg: cur.fill = fill(bg)

def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w

def heights(ws, spec):
    for r, h in spec.items(): ws.row_dimensions[r].height = h

def title_block(ws, title, desc):
    ws.sheet_view.showGridLines = False
    put(ws, 1, 1, title, 15, True, D.PRIMARY)
    ws.row_dimensions[1].height = 26
    put(ws, 2, 1, desc, 9, color=D.INK_SOFT)
    ws.row_dimensions[2].height = 16
    ws.row_dimensions[3].height = 6

def sect(ws, row, text, col=1, span=8):
    put(ws, row, col, text, 11, True, D.PRIMARY)
    ws.row_dimensions[row].height = 22
    for c in range(col, col + span):
        ws.cell(row, c).border = Border(bottom=side(D.PRIMARY_MID, "thin"))
    return row + 1

def tip(ws, row, text, col=1, color=D.INK_FAINT):
    put(ws, row, col, text, 9, color=color)
    ws.row_dimensions[row].height = 17
    return row + 1

def page(ws, landscape=True, title_rows=None, header=None, fit=1, fit_h=0):
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = fit; ws.page_setup.fitToHeight = fit_h
    ws.page_margins.left = ws.page_margins.right = 0.5
    ws.page_margins.top = 0.6; ws.page_margins.bottom = 0.55
    ws.page_margins.header = ws.page_margins.footer = 0.3
    if title_rows: ws.print_title_rows = title_rows
    ws.oddHeader.left.text = header or ws.title
    ws.oddHeader.left.size = 9; ws.oddHeader.left.color = "888888"
    ws.oddFooter.right.text = "&P / &N"; ws.oddFooter.right.size = 9
    ws.oddFooter.right.color = "888888"

HEAD = Border(top=side(D.PRIMARY, "medium"), bottom=side(D.PRIMARY_MID, "thin"))
ROW = Border(bottom=side(D.LINE_SOFT))
ENDR = Border(bottom=side(D.PRIMARY_MID, "thin"))

def table(ws, row, col, headers, rows, widths_=None, aligns=None, h=20, head_h=22):
    n = len(headers)
    for j, htxt in enumerate(headers):
        c = put(ws, row, col + j, htxt, 9.5, True, D.PRIMARY, D.PRIMARY_TINT,
                "center" if (aligns or ["l"] * n)[j] == "c" else "left", indent=1)
        c.border = HEAD
    ws.row_dimensions[row].height = head_h
    for i, r in enumerate(rows):
        rr = row + 1 + i
        for j in range(n):
            a = (aligns or ["l"] * n)[j]
            c = put(ws, rr, col + j, r[j] if j < len(r) else "", 10,
                    h={"l": "left", "c": "center", "r": "right"}[a], indent=1 if a == "l" else 0)
            c.border = ROW
        ws.row_dimensions[rr].height = h
    last = row + len(rows)
    for j in range(n): ws.cell(last, col + j).border = ENDR
    return last

# ══ 1. 안내 ════════════════════════════════════════════════════════════════
def sh_guide(wb):
    ws = wb.active; ws.title = "안내"
    title_block(ws, "Excel 시각 서식 도구",
                "도형을 쓰지 않고 셀 서식만으로 만드는 시각 요소와, 실무 함수 · 리본 기능을 미리 적용해 둔 파일.")
    widths(ws, {"A": 3, "B": 20, "C": 42, "D": 34, "E": 16, "F": 14, "G": 14, "H": 14, "I": 14})
    r = 5
    r = sect(ws, r, "시트 안내", 2, 3)
    r = table(ws, r, 2, ["시트", "무엇이 들어 있나", "이럴 때 연다"],
              [["셀 시각요소", "카드 · 배지 · 구분선 · 진행 막대 · 아이콘 · 셀로 그린 도식", "보고서를 보기 좋게 만들 때"],
               ["표 서식", "머리글 · 줄무늬 · 합계 강조 · 인쇄용 표 네 가지", "표를 정리할 때"],
               ["대시보드", "지표 카드 + 막대 + 상태를 한 화면에 모은 예", "한 장 요약이 필요할 때"],
               ["일정표", "조건부 서식으로 기간을 칠하는 간트 표", "일정을 보여줄 때"],
               ["실무 함수", "자주 쓰는 함수 18개의 수식과 결과", "수식이 기억나지 않을 때"],
               ["리본 기능", "틀 고정 · 필터 · 드롭다운 · 조건부 서식이 적용된 연습 시트", "기능 위치를 찾을 때"]],
              aligns=["l", "l", "l"], h=22)
    r += 2
    r = sect(ws, r, "쓰는 방법", 2, 3)
    for t in ["시트째로 쓰기 — 시트 탭 오른쪽 클릭 › 이동/복사 › 복사본 만들기 › 대상 통합 문서 선택",
              "일부만 쓰기 — 범위를 복사한 뒤, 붙여넣을 곳에서 오른쪽 클릭 › 선택하여 붙여넣기 › 서식",
              "값은 그대로 두고 모양만 바꾸기 — 서식 복사(빗자루) 단추를 두 번 누르면 여러 곳에 연속 적용",
              "행 높이 · 열 너비도 함께 옮기려면 행/열 전체를 선택해 복사한다",
              "〈 〉 안의 글자는 실제 값으로 바꿔 넣는다"]:
        r = tip(ws, r, "· " + t, 2, D.INK_SOFT)
    r += 1
    r = sect(ws, r, "색과 크기 기준", 2, 3)
    sw = [("주색 · 머리글 글자", D.PRIMARY), ("보조 · 강조선", D.PRIMARY_MID), ("머리글 바탕", D.PRIMARY_TINT),
          ("연한 바탕 · 줄무늬", D.SURFACE), ("표 선", D.LINE), ("정상", D.ST_OK), ("주의", D.ST_WARN),
          ("위험", D.ST_BAD)]
    r0 = r
    for i, (nm, c) in enumerate(sw):
        rr = r0 + i
        put(ws, rr, 2, "██████", 11, color=c, indent=1)   # 글자로 만든 색 견본
        put(ws, rr, 3, nm, 10)
        put(ws, rr, 4, "#" + c, 9.5, color=D.INK_FAINT, name="Consolas")
        ws.row_dimensions[rr].height = 18
    r = r0 + len(sw) + 1
    r = tip(ws, r, "글꼴은 맑은 고딕 하나만 쓴다. 제목 15 · 구역 제목 11 · 본문 10 · 주석 9pt.", 2)
    r = tip(ws, r, "행 높이는 머리글 22, 본문 20, 구분선 6 을 기준으로 한다.", 2)
    page(ws, landscape=True, header="Excel 시각 서식 도구 · 안내")
    return ws

# ══ 2. 셀 시각요소 ═════════════════════════════════════════════════════════
def sh_visual(wb):
    ws = wb.create_sheet("셀 시각요소")
    title_block(ws, "셀로 만드는 시각 요소",
                "도형을 쓰지 않는다. 행 높이 · 열 너비 · 테두리 · 채우기 · 조건부 서식만으로 만든다.")
    widths(ws, {"A": 2.6, "B": 14, "C": 14, "D": 14, "E": 3, "F": 14, "G": 14, "H": 14,
                "I": 3, "J": 46, "K": 12})
    r = 5
    # 2-1 카드
    r = sect(ws, r, "① 지표 카드", 2, 7)
    put(ws, r, 10, "만드는 법", 10, True, D.PRIMARY)
    r += 1
    c0 = r
    cards = [("〈전체 건수〉", "1,240", "건", "〈전월 대비 +8%〉", D.PRIMARY),
             ("〈처리율〉", "98.2", "%", "〈목표 95%〉", D.ST_OK),
             ("〈지연 건수〉", "12", "건", "〈목표 10건 이하〉", D.ST_BAD)]
    for i, (name, val, unit, sub, color) in enumerate(cards):
        c = 2 + i * 3 + (1 if i else 0) * 0  # B, E?  -> 2,5,8 ... 아래에서 조정
    cols = [2, 6, 10]
    cols = [2, 6, 2]  # placeholder (아래에서 실제 배치)
    for i, (name, val, unit, sub, color) in enumerate(cards):
        cc = 2 + i * 2
        put(ws, c0, cc, name, 9, color=D.INK_FAINT)
        put(ws, c0 + 1, cc, val, 20, True, color, h="left")
        put(ws, c0 + 1, cc + 1, unit, 10, color=D.INK_SOFT, h="left")
        put(ws, c0 + 2, cc, sub, 9, color=D.INK_SOFT)
        merge(ws, c0, cc, c0, cc + 1); merge(ws, c0 + 2, cc, c0 + 2, cc + 1)
        boxed(ws, c0, cc, c0 + 2, cc + 1, D.LINE)
        for rr in range(c0, c0 + 3):
            ws.cell(rr, cc).alignment = al(indent=1)
    heights(ws, {c0: 18, c0 + 1: 30, c0 + 2: 18})
    put(ws, c0, 10, "① 3×2 칸을 테두리로 묶는다  ② 숫자 칸 행 높이를 30으로 키우고 20pt 굵게"
                    "  ③ 단위는 옆 칸에 10pt 회색  ④ 위·아래 칸에 이름과 비교 기준", 9,
        color=D.INK_SOFT, wrap=True)
    merge(ws, c0, 10, c0 + 2, 11)
    ws.cell(c0, 10).alignment = al(v="top", wrap=True, indent=1)
    r = c0 + 4
    # 2-2 배지
    r = sect(ws, r, "② 상태 배지 · 태그", 2, 7)
    r += 1
    badges = [("정상", D.ST_OK, D.OK_BG), ("주의", D.ST_WARN, D.WARN_BG),
              ("위험", D.ST_BAD, D.DANGER_BG), ("〈태그〉", D.PRIMARY, D.PRIMARY_TINT)]
    for i, (t, fg, bg) in enumerate(badges):
        cc = 2 + i * 2
        put(ws, r, cc, t, 10, True, fg, bg, "center")
        merge(ws, r, cc, r, cc + 1)
        boxed(ws, r, cc, r, cc + 1, bg)
    ws.row_dimensions[r].height = 22
    put(ws, r, 10, "글자색과 바탕색을 짝으로 쓴다. 홈 › 채우기 색 / 글꼴 색. "
                   "가운데 맞춤 + 굵게 + 양옆 칸 병합이면 배지처럼 보인다.", 9, color=D.INK_SOFT, wrap=True)
    merge(ws, r, 10, r + 1, 11); ws.cell(r, 10).alignment = al(v="top", wrap=True, indent=1)
    r += 3
    # 2-3 구분선
    r = sect(ws, r, "③ 구분선과 여백", 2, 7)
    r += 1
    put(ws, r, 2, "〈굵은 구분선 — 행 높이 6 + 주색 채우기〉", 9, color=D.INK_FAINT)
    r += 1
    for c in range(2, 9): put(ws, r, c, "", bg=D.PRIMARY_MID)
    ws.row_dimensions[r].height = 6
    r += 1
    put(ws, r, 2, "〈얇은 구분선 — 행 높이 4 + 연한 채우기〉", 9, color=D.INK_FAINT)
    r += 1
    for c in range(2, 9): put(ws, r, c, "", bg=D.LINE)
    ws.row_dimensions[r].height = 4
    put(ws, r - 3, 10, "빈 행의 높이를 4~6 으로 줄이고 색을 채우면 선이 된다. "
                       "표와 표 사이에는 높이 8~10 의 빈 행을 두어 숨통을 틔운다.", 9,
        color=D.INK_SOFT, wrap=True)
    merge(ws, r - 3, 10, r - 1, 11); ws.cell(r - 3, 10).alignment = al(v="top", wrap=True, indent=1)
    r += 2
    # 2-4 진행 막대
    r = sect(ws, r, "④ 진행 막대 — 두 가지 방법", 2, 7)
    r += 1
    put(ws, r, 2, "항목", 9.5, True, D.PRIMARY, D.PRIMARY_TINT, indent=1)
    put(ws, r, 3, "값", 9.5, True, D.PRIMARY, D.PRIMARY_TINT, "center")
    put(ws, r, 4, "REPT 함수 막대", 9.5, True, D.PRIMARY, D.PRIMARY_TINT, indent=1)
    put(ws, r, 6, "조건부 서식 데이터 막대", 9.5, True, D.PRIMARY, D.PRIMARY_TINT, "center")
    merge(ws, r, 4, r, 5)
    for c in range(2, 7): ws.cell(r, c).border = HEAD
    hr = r
    data = [("〈항목 A〉", 0.82), ("〈항목 B〉", 0.45), ("〈항목 C〉", 0.15)]
    for i, (nm, v) in enumerate(data):
        rr = r + 1 + i
        put(ws, rr, 2, nm, 10, indent=1)
        put(ws, rr, 3, v, 10, h="center", fmt="0%")
        f = ws.cell(rr, 4)
        f.value = f'=REPT("■",ROUND(C{rr}*10,0))&" "&TEXT(C{rr},"0%")'
        f.font = font(10, False, D.PRIMARY_MID); f.alignment = al(indent=1)
        merge(ws, rr, 4, rr, 5)
        put(ws, rr, 6, v, 10, h="center", fmt="0%")
        for c in range(2, 7): ws.cell(rr, c).border = ROW
        ws.row_dimensions[rr].height = 20
    for c in range(2, 7): ws.cell(r + 3, c).border = ENDR
    ws.conditional_formatting.add(f"F{hr+1}:F{hr+3}",
        DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                    color=D.PRIMARY_MID, showValue=True, minLength=None, maxLength=None))
    put(ws, hr, 10, "REPT 막대 — 수식으로 ■ 를 반복한다. 인쇄·복사에 강하다.\n"
                    "데이터 막대 — 범위 선택 › 홈 › 조건부 서식 › 데이터 막대. 값이 바뀌면 길이도 바뀐다.", 9,
        color=D.INK_SOFT, wrap=True)
    merge(ws, hr, 10, hr + 3, 11); ws.cell(hr, 10).alignment = al(v="top", wrap=True, indent=1)
    r = hr + 5
    # 2-5 아이콘 · 색조
    r = sect(ws, r, "⑤ 아이콘 집합 · 색조", 2, 7)
    r += 1
    put(ws, r, 2, "구분", 9.5, True, D.PRIMARY, D.PRIMARY_TINT, indent=1)
    put(ws, r, 3, "아이콘", 9.5, True, D.PRIMARY, D.PRIMARY_TINT, "center")
    put(ws, r, 4, "색조", 9.5, True, D.PRIMARY, D.PRIMARY_TINT, "center")
    put(ws, r, 5, "O · △ · × 표기", 9.5, True, D.PRIMARY, D.PRIMARY_TINT, "center")
    merge(ws, r, 5, r, 6)
    for c in range(2, 7): ws.cell(r, c).border = HEAD
    ir = r
    vals = [("〈좋음〉", 95, 95, "○"), ("〈보통〉", 70, 70, "△"), ("〈나쁨〉", 35, 35, "×")]
    for i, (nm, a, b, mark) in enumerate(vals):
        rr = r + 1 + i
        put(ws, rr, 2, nm, 10, indent=1)
        put(ws, rr, 3, a, 10, h="center")
        put(ws, rr, 4, b, 10, h="center")
        put(ws, rr, 5, mark, 11, True,
            D.ST_OK if mark == "○" else (D.ST_WARN if mark == "△" else D.ST_BAD), h="center")
        merge(ws, rr, 5, rr, 6)
        for c in range(2, 7): ws.cell(rr, c).border = ROW
        ws.row_dimensions[rr].height = 20
    for c in range(2, 7): ws.cell(r + 3, c).border = ENDR
    ws.conditional_formatting.add(f"C{ir+1}:C{ir+3}",
        IconSetRule("3TrafficLights1", "num", [0, 60, 90], showValue=True))
    ws.conditional_formatting.add(f"D{ir+1}:D{ir+3}",
        ColorScaleRule(start_type="num", start_value=0, start_color="FF" + D.DANGER_BG,
                       mid_type="num", mid_value=60, mid_color="FF" + D.WARN_BG,
                       end_type="num", end_value=100, end_color="FF" + D.OK_BG))
    put(ws, ir, 10, "아이콘 · 색조 — 범위 선택 › 홈 › 조건부 서식 › 아이콘 집합 / 색조.\n"
                    "기준값은 ‘규칙 관리 › 규칙 편집’에서 숫자로 직접 정한다.\n"
                    "O △ × 는 그냥 글자다. 색만 다르게 주면 인쇄에서도 잘 보인다.", 9,
        color=D.INK_SOFT, wrap=True)
    merge(ws, ir, 10, ir + 3, 11); ws.cell(ir, 10).alignment = al(v="top", wrap=True, indent=1)
    r = ir + 5
    # 2-6 셀로 그린 도식
    r = sect(ws, r, "⑥ 셀로 그린 도식", 2, 7)
    r += 1
    put(ws, r, 10, "열 너비를 2~3 으로 줄이면 셀이 모눈종이가 된다.\n"
                   "칸을 병합하고 테두리를 두르면 상자가, 화살표는 글자(→ ↓)를 넣으면 된다.\n"
                   "도형과 달리 셀에 붙어 있어 행·열을 옮겨도 흐트러지지 않는다.", 9,
        color=D.INK_SOFT, wrap=True)
    merge(ws, r, 10, r + 3, 11); ws.cell(r, 10).alignment = al(v="top", wrap=True, indent=1)
    dr = r
    steps = [("〈준비〉", D.PRIMARY_TINT, D.PRIMARY), ("〈진행〉", D.PRIMARY_TINT, D.PRIMARY),
             ("〈확인〉", D.SURFACE, D.INK)]
    for i, (t, bg, fg) in enumerate(steps):
        cc = 2 + i * 3
        merge(ws, dr, cc, dr + 1, cc + 1)
        put(ws, dr, cc, t, 10, True, fg, bg, "center")
        boxed(ws, dr, cc, dr + 1, cc + 1, D.PRIMARY_MID)
        if i < 2:
            put(ws, dr, cc + 2, "→", 12, True, D.PRIMARY_MID, h="center")
            merge(ws, dr, cc + 2, dr + 1, cc + 2)
    heights(ws, {dr: 18, dr + 1: 18})
    put(ws, dr + 2, 4, "↓", 12, True, D.PRIMARY_MID, h="center")
    merge(ws, dr + 3, 2, dr + 4, 7)
    put(ws, dr + 3, 2, "〈아래로 이어지는 단계도 같은 방법으로 만든다〉", 10, color=D.INK_SOFT, h="center")
    boxed(ws, dr + 3, 2, dr + 4, 7, D.LINE)
    page(ws, landscape=True, header="셀로 만드는 시각 요소", fit_h=1)
    return ws

# ══ 3. 표 서식 ═════════════════════════════════════════════════════════════
def sh_tables(wb):
    ws = wb.create_sheet("표 서식")
    title_block(ws, "표 서식 네 가지", "쓰임에 맞는 표를 골라 복사한다. 선은 적게, 여백은 넉넉히.")
    widths(ws, {"A": 2.6, "B": 18, "C": 22, "D": 13, "E": 13, "F": 13, "G": 3, "H": 46})
    r = 5
    r = sect(ws, r, "① 기본 표 — 가장 많이 쓴다", 2, 5)
    put(ws, r, 8, "만드는 법", 10, True, D.PRIMARY)
    put(ws, r + 1, 8, "머리글만 연한 색으로 채우고 아래 선을 굵게 둔다. 세로선은 넣지 않는다.\n"
                      "홈 › 테두리 › 다른 테두리 에서 아래쪽 선만 지정한다.", 9, color=D.INK_SOFT, wrap=True)
    merge(ws, r + 1, 8, r + 3, 8); ws.cell(r + 1, 8).alignment = al(v="top", wrap=True, indent=1)
    r += 1
    last = table(ws, r, 2, ["구분", "항목", "〈값 1〉", "〈값 2〉", "비고"],
                 [["〈구분 A〉", "〈항목 이름〉", 120, 100, "〈메모〉"],
                  ["〈구분 A〉", "〈항목 이름〉", 98, 102, ""],
                  ["〈구분 B〉", "〈항목 이름〉", 45, 38, ""]],
                 aligns=["l", "l", "r", "r", "l"])
    r = last + 2
    r = sect(ws, r, "② 줄무늬 표 — 행이 많을 때", 2, 5)
    put(ws, r, 8, "줄무늬는 색을 직접 칠하지 말고 조건부 서식 수식으로 준다.\n"
                  "새 규칙 › 수식 사용 › =MOD(ROW(),2)=0 › 서식 › 채우기", 9, color=D.INK_SOFT, wrap=True)
    merge(ws, r, 8, r + 2, 8); ws.cell(r, 8).alignment = al(v="top", wrap=True, indent=1)
    r += 1
    z = table(ws, r, 2, ["구분", "항목", "〈값 1〉", "〈값 2〉", "비고"],
              [[f"〈구분 {chr(65+i//2)}〉", "〈항목 이름〉", 100 - i * 7, 90 - i * 5, ""] for i in range(4)],
              aligns=["l", "l", "r", "r", "l"])
    ws.conditional_formatting.add(f"B{r+1}:F{z}",
        FormulaRule(formula=[f"MOD(ROW(),2)=0"], fill=fill(D.SURFACE)))
    r = z + 2
    r = sect(ws, r, "③ 합계가 있는 표", 2, 5)
    put(ws, r, 8, "합계 행은 위쪽에 얇은 선을 긋고 굵게 쓴다. 색은 쓰지 않아도 충분하다.\n"
                  "합계는 =SUM(위 범위) 로 넣어 두면 값이 바뀌어도 따라온다.", 9, color=D.INK_SOFT, wrap=True)
    merge(ws, r, 8, r + 2, 8); ws.cell(r, 8).alignment = al(v="top", wrap=True, indent=1)
    r += 1
    t = table(ws, r, 2, ["구분", "항목", "〈1월〉", "〈2월〉", "합계"],
              [["〈구분 A〉", "〈항목〉", 120, 135, f"=SUM(D{r+1}:E{r+1})"],
               ["〈구분 B〉", "〈항목〉", 80, 76, f"=SUM(D{r+2}:E{r+2})"],
               ["〈구분 C〉", "〈항목〉", 45, 52, f"=SUM(D{r+3}:E{r+3})"]],
              aligns=["l", "l", "r", "r", "r"])
    sr = t + 1
    put(ws, sr, 2, "합계", 10, True, D.PRIMARY, indent=1)
    put(ws, sr, 3, "", 10)
    for j, col in enumerate("DEF"):
        c = put(ws, sr, 4 + j, f"=SUM({col}{r+1}:{col}{t})", 10, True, D.PRIMARY, h="right")
    for c in range(2, 7):
        ws.cell(sr, c).border = Border(top=side(D.PRIMARY_MID, "thin"), bottom=side(D.PRIMARY_MID, "thin"))
    ws.row_dimensions[sr].height = 22
    r = sr + 2
    r = sect(ws, r, "④ 인쇄용 표 — 나눠 주는 문서", 2, 5)
    put(ws, r, 8, "인쇄물은 모든 칸에 얇은 선을 넣는 편이 읽기 쉽다.\n"
                  "페이지 레이아웃 › 인쇄 제목 › 반복할 행 에 머리글 행을 지정한다.", 9,
        color=D.INK_SOFT, wrap=True)
    merge(ws, r, 8, r + 2, 8); ws.cell(r, 8).alignment = al(v="top", wrap=True, indent=1)
    r += 1
    pr = r
    p = table(ws, r, 2, ["구분", "항목", "〈값 1〉", "〈값 2〉", "확인"],
              [["〈구분〉", "〈항목 이름〉", "", "", ""] for _ in range(4)],
              aligns=["l", "l", "r", "r", "c"])
    for rr in range(pr, p + 1):
        for c in range(2, 7):
            ws.cell(rr, c).border = Border(top=side(D.LINE), bottom=side(D.LINE),
                                           left=side(D.LINE), right=side(D.LINE))
    for c in range(2, 7):
        ws.cell(pr, c).border = Border(top=side(D.PRIMARY, "medium"), bottom=side(D.PRIMARY_MID),
                                       left=side(D.LINE), right=side(D.LINE))
    page(ws, landscape=True, header="표 서식", fit_h=1)
    return ws

# ══ 4. 대시보드 ════════════════════════════════════════════════════════════
def sh_dashboard(wb):
    ws = wb.create_sheet("대시보드")
    title_block(ws, "한 장 요약 (대시보드)",
                "〈기준일 2026-01-01〉 지표 · 막대 · 상태를 한 화면에 모은 예. 숫자만 바꿔 쓴다.")
    widths(ws, {"A": 2.6, "B": 13, "C": 7, "D": 13, "E": 7, "F": 13, "G": 7, "H": 13, "I": 7,
                "J": 3, "K": 40})
    r = 5
    cards = [("〈전체 건수〉", 1240, "건", "〈전월 1,148〉", D.PRIMARY, "#,##0"),
             ("〈처리율〉", 0.982, "", "〈목표 95%〉", D.ST_OK, "0.0%"),
             ("〈평균 소요〉", 2.4, "일", "〈전월 2.9일〉", D.PRIMARY, "#,##0.0"),
             ("〈지연 건수〉", 12, "건", "〈목표 10 이하〉", D.ST_BAD, "#,##0")]
    for i, (nm, val, unit, sub, color, fmt) in enumerate(cards):
        cc = 2 + i * 2
        put(ws, r, cc, nm, 9, color=D.INK_FAINT, indent=1)
        put(ws, r + 1, cc, val, 18, True, color, indent=1, fmt=fmt)
        put(ws, r + 1, cc + 1, unit, 10, color=D.INK_SOFT)
        put(ws, r + 2, cc, sub, 9, color=D.INK_SOFT, indent=1)
        merge(ws, r, cc, r, cc + 1); merge(ws, r + 2, cc, r + 2, cc + 1)
        boxed(ws, r, cc, r + 2, cc + 1, D.LINE)
    heights(ws, {r: 18, r + 1: 30, r + 2: 18})
    r += 4
    hr = sect(ws, r, "항목별 현황", 2, 8)
    for cspec, h, a in ((2, "항목", "l"), (4, "건수", "c"), (5, "비율", "c"), (6, "막대", "l"), (8, "상태", "c")):
        c = put(ws, hr, cspec, h, 9.5, True, D.PRIMARY, D.PRIMARY_TINT,
                {"l": "left", "c": "center"}[a], indent=1 if a == "l" else 0)
    merge(ws, hr, 2, hr, 3); merge(ws, hr, 6, hr, 7); merge(ws, hr, 8, hr, 9)
    for c in range(2, 10):
        ws.cell(hr, c).border = HEAD
        if not ws.cell(hr, c).fill.fgColor.rgb: ws.cell(hr, c).fill = fill(D.PRIMARY_TINT)
    ws.row_dimensions[hr].height = 22
    rows = [("〈항목 A〉", 420, 0.92, "정상"), ("〈항목 B〉", 310, 0.78, "정상"),
            ("〈항목 C〉", 260, 0.61, "주의"), ("〈항목 D〉", 150, 0.44, "주의"),
            ("〈항목 E〉", 100, 0.22, "위험")]
    for i, (nm, cnt, rate, st) in enumerate(rows):
        rr = hr + 1 + i
        put(ws, rr, 2, nm, 10, indent=1); merge(ws, rr, 2, rr, 3)
        put(ws, rr, 4, cnt, 10, h="center", fmt="#,##0")
        put(ws, rr, 5, rate, 10, h="center", fmt="0%")
        f = ws.cell(rr, 6)
        f.value = f'=REPT("■",ROUND(E{rr}*10,0))'
        f.font = font(10, False, D.PRIMARY_MID); f.alignment = al(indent=1)
        merge(ws, rr, 6, rr, 7)
        col = {"정상": (D.ST_OK, D.OK_BG), "주의": (D.ST_WARN, D.WARN_BG), "위험": (D.ST_BAD, D.DANGER_BG)}[st]
        put(ws, rr, 8, st, 9.5, True, col[0], col[1], "center"); merge(ws, rr, 8, rr, 9)
        ws.cell(rr, 9).fill = fill(col[1])
        for c in range(2, 10): ws.cell(rr, c).border = ROW
        ws.row_dimensions[rr].height = 20
    last = hr + len(rows)
    for c in range(2, 10): ws.cell(last, c).border = ENDR
    ws.conditional_formatting.add(f"D{hr+1}:D{last}",
        DataBarRule(start_type="num", start_value=0, end_type="num", end_value=500,
                    color=D.PRIMARY_TINT, showValue=True))
    # 오른쪽 — 확인할 것
    k = sect(ws, r, "확인할 것", 11, 1)
    items = [("〈지연 건수 증가〉", "〈원인 확인 후 조치 · 01-10까지〉", D.ST_BAD),
             ("〈항목 E 비율 낮음〉", "〈담당 배정 조정 · 01-15까지〉", D.ST_WARN),
             ("〈전체 처리율 목표 달성〉", "〈현행 유지〉", D.ST_OK)]
    for i, (t, d2, c) in enumerate(items):
        y = k + i * 3
        put(ws, y, 11, t, 10.5, True, D.INK, indent=2)
        put(ws, y + 1, 11, d2, 9.5, color=D.INK_SOFT, indent=2)
        boxed(ws, y, 11, y + 1, 11, D.LINE)
        for rr2 in (y, y + 1):
            b = ws.cell(rr2, 11).border
            ws.cell(rr2, 11).border = Border(left=side(c, "thick"), right=b.right,
                                             top=b.top, bottom=b.bottom)
    r2 = max(last, k + 9) + 2
    tip(ws, r2, "※ 〈기준일 · 산출 방법 · 제외 대상을 여기에 적는다. 숫자만 있으면 읽는 사람이 각자 해석한다.〉", 2)
    tip(ws, r2 + 1, "※ 막대는 REPT 함수, 건수 칸은 조건부 서식 데이터 막대. 상태 칸은 채우기 색만 바꾼 것이다.", 2)
    ws.freeze_panes = "A5"
    page(ws, landscape=True, header="한 장 요약")
    return ws

# ══ 5. 일정표 ══════════════════════════════════════════════════════════════
def sh_gantt(wb):
    ws = wb.create_sheet("일정표")
    title_block(ws, "일정표 (간트)",
                "시작 · 종료 주차만 입력하면 막대가 자동으로 칠해진다. 도형을 그리지 않는다.")
    widths(ws, {"A": 2.6, "B": 22, "C": 10, "D": 8, "E": 8, "F": 9})
    for i in range(12):
        ws.column_dimensions[get_column_letter(7 + i)].width = 4.4
    ws.column_dimensions[get_column_letter(20)].width = 3
    ws.column_dimensions[get_column_letter(21)].width = 44
    r = 5
    put(ws, r, 2, "현재 주차", 9.5, True, D.PRIMARY, indent=1)
    put(ws, r, 3, 5, 10, True, D.PRIMARY, D.PRIMARY_TINT, "center")
    put(ws, r, 4, "← 이 숫자를 바꾸면 기준선이 옮겨진다", 9, color=D.INK_FAINT)
    merge(ws, r, 4, r, 8)
    ws.row_dimensions[r].height = 20
    r += 2
    hr = r
    heads = ["작업", "담당", "시작", "종료", "진행률"]
    for j, h in enumerate(heads):
        c = put(ws, hr, 2 + j, h, 9.5, True, D.PRIMARY, D.PRIMARY_TINT,
                "center" if j >= 2 else "left", indent=1 if j < 2 else 0)
        c.border = HEAD
    for i in range(12):
        c = put(ws, hr, 7 + i, i + 1, 9, True, D.PRIMARY, D.PRIMARY_TINT, "center")
        c.border = HEAD
    ws.row_dimensions[hr].height = 22
    tasks = [("〈준비 단계〉", "〈이름〉", 1, 3, 1.0), ("〈설계 · 검토〉", "〈이름〉", 2, 5, 0.8),
             ("〈본 작업〉", "〈이름〉", 4, 9, 0.45), ("〈검증〉", "〈이름〉", 8, 10, 0.1),
             ("〈정리 · 보고〉", "〈이름〉", 10, 12, 0.0)]
    for i, (nm, who, st, en, pct) in enumerate(tasks):
        rr = hr + 1 + i
        put(ws, rr, 2, nm, 10, indent=1)
        put(ws, rr, 3, who, 10, indent=1)
        put(ws, rr, 4, st, 10, h="center")
        put(ws, rr, 5, en, 10, h="center")
        put(ws, rr, 6, pct, 10, h="center", fmt="0%")
        for c in range(2, 7): ws.cell(rr, c).border = ROW
        for c in range(7, 19):
            cell = ws.cell(rr, c)
            cell.border = Border(bottom=side(D.LINE_SOFT), right=side("EFF2F5"))
        ws.row_dimensions[rr].height = 22
    last = hr + len(tasks)
    for c in range(2, 19): ws.cell(last, c).border = ENDR
    rng = f"G{hr+1}:R{last}"
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f"AND(G${hr}>=$D{hr+1},G${hr}<=$E{hr+1})"], fill=fill(D.PRIMARY_MID), stopIfTrue=False))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f"G${hr}=$C${hr-2}"], fill=fill(D.WARN_BG), stopIfTrue=False))
    ws.conditional_formatting.add(f"F{hr+1}:F{last}",
        DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                    color=D.PRIMARY_TINT, showValue=True))
    put(ws, hr, 21, "만드는 법", 10, True, D.PRIMARY)
    put(ws, hr + 1, 21,
        "① 주차 머리글(1~12)과 시작 · 종료 칸을 숫자로 둔다\n"
        "② 막대 범위를 선택하고 홈 › 조건부 서식 › 새 규칙 › 수식 사용\n"
        "③ 수식  =AND(G$8>=$D9,G$8<=$E9)   ← 머리글 행은 행 고정($8), 시작 · 종료는 열 고정($D,$E)\n"
        "④ 서식 › 채우기 에서 막대 색을 고른다\n"
        "⑤ 기준선은 규칙을 하나 더 만들어  =G$8=$C$5  로 준다\n"
        "⑥ 진행률 칸에는 조건부 서식 › 데이터 막대\n\n"
        "주차 대신 날짜로 쓰려면 머리글에 날짜를 넣고 같은 수식을 쓰면 된다.", 9,
        color=D.INK_SOFT, wrap=True)
    merge(ws, hr + 1, 21, hr + 12, 21)
    ws.cell(hr + 1, 21).alignment = al(v="top", wrap=True, indent=1)
    ws.freeze_panes = f"G{hr+1}"
    page(ws, landscape=True, title_rows=f"{hr}:{hr}", header="일정표")
    return ws

# ══ 6. 실무 함수 ═══════════════════════════════════════════════════════════
def sh_formula(wb):
    ws = wb.create_sheet("실무 함수")
    title_block(ws, "자주 쓰는 함수",
                "왼쪽 자료를 대상으로 실제 계산되는 수식이다. 수식 칸을 그대로 복사해 쓰면 된다.")
    widths(ws, {"A": 2.6, "B": 12, "C": 10, "D": 10, "E": 12, "F": 12, "G": 3,
                "H": 26, "I": 44, "J": 14, "K": 30})
    r = 5
    r = sect(ws, r, "자료", 2, 5)
    dr = r
    data = [("〈김하늘〉", "〈영업〉", "A", 1200000, "2026-01-05", "완료"),
            ("〈이바다〉", "〈영업〉", "B", 850000, "2026-01-12", "진행"),
            ("〈박구름〉", "〈지원〉", "A", 430000, "2026-01-18", "완료"),
            ("〈최나무〉", "〈지원〉", "C", 1750000, "2026-02-02", "진행"),
            ("〈정바람〉", "〈개발〉", "B", 920000, "2026-02-11", "보류"),
            ("〈한소리〉", "〈개발〉", "A", 2100000, "2026-02-20", "완료")]
    last = table(ws, dr, 2, ["이름", "부서", "구분", "금액", "날짜", "상태"],
                 [[a, b, c, d, e, f] for a, b, c, d, e, f in data],
                 aligns=["l", "c", "c", "r", "c", "c"])
    for rr in range(dr + 1, last + 1):
        ws.cell(rr, 5).number_format = "#,##0"
    r = last + 2
    r = sect(ws, r, "함수", 2, 5)
    fr = r
    d1, d2 = dr + 1, last
    rows = [
        ("이름으로 금액 찾기", f'=VLOOKUP("〈박구름〉",$B${d1}:$F${d2},4,FALSE)', "찾을 값이 첫 열에 있어야 한다"),
        ("왼쪽 값도 찾기", f'=INDEX($B${d1}:$B${d2},MATCH(2100000,$E${d1}:$E${d2},0))', "VLOOKUP 이 안 될 때"),
        ("조건 하나로 세기", f'=COUNTIF($C${d1}:$C${d2},"〈영업〉")', "몇 건인지"),
        ("조건 둘로 세기", f'=COUNTIFS($C${d1}:$C${d2},"〈개발〉",$G${d1}:$G${d2},"완료")', "부서 + 상태"),
        ("조건으로 더하기", f'=SUMIF($C${d1}:$C${d2},"〈지원〉",$E${d1}:$E${d2})', "해당 건만 합계"),
        ("조건 둘로 더하기", f'=SUMIFS($E${d1}:$E${d2},$D${d1}:$D${d2},"A",$G${d1}:$G${d2},"완료")', "구분 + 상태"),
        ("조건으로 평균", f'=AVERAGEIF($D${d1}:$D${d2},"A",$E${d1}:$E${d2})', "0 으로 나누면 오류"),
        ("오류 숨기기", f'=IFERROR(VLOOKUP("없는이름",$B${d1}:$F${d2},4,FALSE),"확인 필요")', "#N/A 대신 글자"),
        ("조건 나누기", f'=IF($E${d1}>=1000000,"큰 건","일반")', "IF 는 3단계까지만"),
        ("여러 조건", f'=IF($E${d1}>=2000000,"상",IF($E${d1}>=1000000,"중","하"))', "네 단계가 넘으면 표로"),
        ("숫자를 글자로", f'=TEXT($E${d1},"#,##0")&"원"', "보고서 문장에 넣을 때"),
        ("날짜 형식", f'=TEXT($F${d1},"yyyy-mm-dd(aaa)")', "요일까지"),
        ("월말 날짜", f'=EOMONTH($F${d1},0)', "0 은 이번 달, 1 은 다음 달"),
        ("경과 일수", f'=$F${d2}-$F${d1}', "날짜끼리 빼면 일수"),
        ("글자 자르기", f'=LEFT($B${d1},3)&"…"', "긴 이름 줄이기"),
        ("글자 합치기", f'=$C${d1}&" / "&$B${d1}', "& 로 잇는다"),
        ("반올림", f'=ROUND($E${d1}/10000,1)&"만원"', "자리수는 두 번째 인수"),
        ("순위", f'=RANK($E${d1},$E${d1}:$E${d2})', "큰 값이 1위"),
    ]
    for j, h in enumerate(["하고 싶은 것", "수식", "결과", "메모"]):
        c = put(ws, fr, 8 + j, h, 9.5, True, D.PRIMARY, D.PRIMARY_TINT,
                "center" if j == 2 else "left", indent=1)
        c.border = HEAD
    ws.row_dimensions[fr].height = 22
    for i, (what, formula, memo) in enumerate(rows):
        rr = fr + 1 + i
        put(ws, rr, 8, what, 10, indent=1)
        c = put(ws, rr, 9, formula, 9.5, color=D.CODE_TXT, name="Consolas", indent=1)
        res = ws.cell(rr, 10); res.value = formula
        res.font = font(10, True, D.PRIMARY); res.alignment = al("center")
        put(ws, rr, 11, memo, 9, color=D.INK_SOFT, indent=1)
        for cc in range(8, 12): ws.cell(rr, cc).border = ROW
        ws.row_dimensions[rr].height = 20
    for cc in range(8, 12): ws.cell(fr + len(rows), cc).border = ENDR
    r2 = last + 2
    put(ws, r2, 2, "알아 두면 좋은 것", 11, True, D.PRIMARY)
    ws.row_dimensions[r2].height = 22
    for c in range(2, 7): ws.cell(r2, c).border = Border(bottom=side(D.PRIMARY_MID))
    for i, t in enumerate([
            "$ 표시는 고정이다. F4 를 누르면 $A$1 → A$1 → $A1 순서로 바뀐다.",
            "범위를 복사해 쓸 때는 조건 범위에 $ 를 붙여 둬야 어긋나지 않는다.",
            "찾는 값이 없으면 #N/A 가 뜬다. IFERROR 로 감싼다.",
            "숫자처럼 보이는 글자는 계산되지 않는다. 데이터 › 텍스트 나누기 › 마침 으로 고친다.",
            "수식이 보이게 하려면 수식 › 수식 표시(Ctrl+`).",
            "SUBTOTAL(109, 범위) 는 필터로 숨긴 행을 빼고 합한다."]):
        put(ws, r2 + 1 + i, 2, "· " + t, 9.5, color=D.INK_SOFT)
        merge(ws, r2 + 1 + i, 2, r2 + 1 + i, 6)
        ws.row_dimensions[r2 + 1 + i].height = 18
    page(ws, landscape=True, header="자주 쓰는 함수", fit_h=1)
    return ws

# ══ 7. 리본 기능 ═══════════════════════════════════════════════════════════
def sh_ribbon(wb):
    ws = wb.create_sheet("리본 기능")
    title_block(ws, "미리 켜 둔 기능들",
                "이 시트에는 아래 기능이 이미 적용되어 있다. 직접 눌러 보고, 같은 방법으로 내 파일에 적용한다.")
    widths(ws, {"A": 2.6, "B": 12, "C": 12, "D": 12, "E": 12, "F": 12, "G": 12, "H": 3,
                "I": 18, "J": 30, "K": 26})
    r = 5
    hr = r
    heads = ["날짜", "부서", "담당", "구분", "금액", "상태"]
    for j, h in enumerate(heads):
        c = put(ws, hr, 2 + j, h, 9.5, True, D.PRIMARY, D.PRIMARY_TINT,
                "center" if j != 0 else "center", indent=0)
        c.border = HEAD
    ws.row_dimensions[hr].height = 22
    import datetime as dt
    depts = ["〈영업〉", "〈지원〉", "〈개발〉"]
    people = ["〈김하늘〉", "〈이바다〉", "〈박구름〉", "〈최나무〉"]
    kinds = ["A", "B", "C"]
    states = ["완료", "진행", "보류"]
    n = 18
    for i in range(n):
        rr = hr + 1 + i
        put(ws, rr, 2, dt.date(2026, 1 + i // 9, 1 + (i * 3) % 27), 10, h="center", fmt="yyyy-mm-dd")
        put(ws, rr, 3, depts[i % 3], 10, h="center")
        put(ws, rr, 4, people[i % 4], 10, h="center")
        put(ws, rr, 5, kinds[i % 3], 10, h="center")
        put(ws, rr, 6, 300000 + (i * 137000) % 1900000, 10, h="right", fmt="#,##0")
        put(ws, rr, 7, states[i % 3], 10, h="center")
        for c in range(2, 8): ws.cell(rr, c).border = ROW
        ws.row_dimensions[rr].height = 20
    last = hr + n
    for c in range(2, 8): ws.cell(last, c).border = ENDR
    # 상태 드롭다운(이름 정의 사용)
    for i, st in enumerate(states):
        put(ws, hr + 1 + i, 13, st, 10, h="center")
    put(ws, hr, 13, "상태목록", 9, True, D.INK_FAINT, h="center")
    wb.defined_names.add(DefinedName("상태목록", attr_text=f"'리본 기능'!$M${hr+1}:$M${hr+3}"))
    dv = DataValidation(type="list", formula1="=상태목록", allow_blank=True)
    dv.prompt = "목록에서 고른다"; dv.promptTitle = "상태"
    ws.add_data_validation(dv); dv.add(f"G{hr+1}:G{last}")
    # 조건부 서식
    for key, fg, bg in (("완료", D.ST_OK, D.OK_BG), ("진행", D.ST_WARN, D.WARN_BG),
                        ("보류", D.INK_FAINT, D.SURFACE)):
        ws.conditional_formatting.add(f"G{hr+1}:G{last}",
            CellIsRule(operator="equal", formula=[f'"{key}"'], fill=fill(bg), font=font(10, True, fg)))
    ws.conditional_formatting.add(f"F{hr+1}:F{last}",
        CellIsRule(operator="greaterThanOrEqual", formula=["1500000"], font=font(10, True, D.ST_BAD)))
    ws.auto_filter.ref = f"B{hr}:G{last}"
    ws.freeze_panes = f"B{hr+1}"
    # 설명
    put(ws, hr, 9, "기능", 9.5, True, D.PRIMARY, D.PRIMARY_TINT, indent=1)
    put(ws, hr, 10, "어디서 켜나", 9.5, True, D.PRIMARY, D.PRIMARY_TINT, indent=1)
    put(ws, hr, 11, "이 시트에서 확인할 곳", 9.5, True, D.PRIMARY, D.PRIMARY_TINT, indent=1)
    for c in range(9, 12): ws.cell(hr, c).border = HEAD
    feats = [
        ("틀 고정", "보기 › 틀 고정 › 틀 고정", "아래로 내려도 머리글이 남는다"),
        ("자동 필터", "데이터 › 필터 (Ctrl+Shift+L)", "머리글의 ▼ 를 눌러 골라 본다"),
        ("드롭다운 목록", "데이터 › 데이터 유효성 검사 › 목록", "상태 칸을 누르면 목록이 뜬다"),
        ("이름 정의", "수식 › 이름 관리자", "목록 범위를 ‘상태목록’으로 정의해 두었다"),
        ("조건부 서식", "홈 › 조건부 서식 › 셀 강조 규칙", "상태 색, 금액 150만 이상 빨강"),
        ("표시 형식", "홈 › 표시 형식 (Ctrl+1)", "금액 #,##0 · 날짜 yyyy-mm-dd"),
        ("인쇄 제목", "페이지 레이아웃 › 인쇄 제목", "쪽마다 머리글 행이 반복된다"),
        ("페이지 맞춤", "페이지 레이아웃 › 너비: 1페이지", "가로로 잘리지 않는다"),
        ("정렬", "데이터 › 정렬 (여러 기준 가능)", "부서 › 금액 순으로 정렬해 본다"),
        ("중복 제거", "데이터 › 중복된 항목 제거", "복사본에서 먼저 시험한다"),
        ("텍스트 나누기", "데이터 › 텍스트 나누기", "숫자처럼 보이는 글자를 고칠 때"),
        ("빠른 채우기", "데이터 › 빠른 채우기 (Ctrl+E)", "규칙을 한 번 보여 주면 나머지를 채운다"),
    ]
    for i, (a, b, c) in enumerate(feats):
        rr = hr + 1 + i
        put(ws, rr, 9, a, 10, True, D.INK, indent=1)
        put(ws, rr, 10, b, 9.5, color=D.INK_SOFT, indent=1)
        put(ws, rr, 11, c, 9.5, color=D.INK_SOFT, indent=1)
        for cc in range(9, 12): ws.cell(rr, cc).border = ROW
        ws.row_dimensions[rr].height = 20
    for cc in range(9, 12): ws.cell(hr + len(feats), cc).border = ENDR
    tip(ws, last + 2, "※ M 열의 ‘상태목록’은 드롭다운의 원본이다. 인쇄 영역 밖이라 출력에는 나오지 않는다.", 2)
    ws.print_area = f"A1:L{last + 3}"
    page(ws, landscape=True, title_rows=f"{hr}:{hr}", header="미리 켜 둔 기능들")
    return ws

def build():
    wb = Workbook()
    sh_guide(wb); sh_visual(wb); sh_tables(wb); sh_dashboard(wb)
    sh_gantt(wb); sh_formula(wb); sh_ribbon(wb)
    wb.properties.title = "Excel 시각 서식 도구"
    wb.properties.creator = "문서 작성 라이브러리"
    wb.active = 0
    wb.save(OUT)
    return OUT

if __name__ == "__main__":
    print("Excel 생성:", build())
