# -*- coding: utf-8 -*-
"""library/04_운영문서_표모음.xlsx 생성 — 기록·점검·비교용 실무 시트."""
import os, sys
HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
import design as D
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule
from openpyxl.worksheet.properties import PageSetupProperties

OUT = os.path.join(ROOT, "library", "04_운영문서_표모음.xlsx")
F = D.FONT_EA

def font(size=10, bold=False, color=D.INK, italic=False):
    return Font(name=F, size=size, bold=bold, color="FF" + color, italic=italic)

def fill(c): return PatternFill("solid", fgColor="FF" + c)
def side(c=D.LINE, style="thin"): return Side(style=style, color="FF" + c)
def al(h="left", v="center", wrap=True, indent=0):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap, indent=indent)

HEAD_FILL = fill(D.PRIMARY_TINT)
SUB_FILL = fill(D.SURFACE)
B_BOTTOM = Border(bottom=side(D.LINE_SOFT))
B_HEAD = Border(top=side(D.PRIMARY, "medium"), bottom=side(D.PRIMARY_MID, "thin"))
B_END = Border(top=side(D.PRIMARY_MID, "thin"))

def sheet_head(ws, title, desc, width_cols):
    ws.sheet_view.showGridLines = False
    ws["A1"] = title; ws["A1"].font = font(15, True, D.PRIMARY); ws["A1"].alignment = al(wrap=False)
    ws.row_dimensions[1].height = 26
    ws["A2"] = desc; ws["A2"].font = font(9, color=D.INK_SOFT); ws["A2"].alignment = al(wrap=False)
    ws.row_dimensions[2].height = 16
    ws.row_dimensions[3].height = 6
    for i, w in enumerate(width_cols, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

def meta_block(ws, row, pairs, label_w=None, span=2):
    """라벨 | 값 정보 블록."""
    for i, (k, v) in enumerate(pairs):
        r = row + i
        c1 = ws.cell(r, 1, k); c1.font = font(9.5, True, D.INK_SOFT); c1.fill = SUB_FILL
        c1.alignment = al(indent=1); c1.border = Border(bottom=side(D.LINE_SOFT))
        c2 = ws.cell(r, 2, v); c2.font = font(10); c2.alignment = al(indent=1)
        c2.border = Border(bottom=side(D.LINE_SOFT))
        if span > 2:
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=span)
            for c in range(3, span + 1):
                ws.cell(r, c).border = Border(bottom=side(D.LINE_SOFT))
        ws.row_dimensions[r].height = 20
    return row + len(pairs)

def table(ws, row, headers, rows, aligns=None, heights=22, head_h=24, spans=None):
    """spans: 각 열이 차지하는 실제 열 수(병합). 예: [2,2,1,1,1]"""
    n = len(headers)
    spans = spans or [1] * n
    starts, col = [], 1
    for sp in spans:
        starts.append(col); col += sp
    total = col - 1
    aligns = aligns or ["l"] * n
    def put(r, j, v, is_head):
        c0 = starts[j]; sp = spans[j]
        c = ws.cell(r, c0, v)
        if sp > 1:
            ws.merge_cells(start_row=r, start_column=c0, end_row=r, end_column=c0 + sp - 1)
        a = aligns[j]
        c.font = font(9.5, True, D.PRIMARY) if is_head else font(10)
        c.alignment = al({"l": "left", "c": "center", "r": "right"}[a], indent=1 if a == "l" else 0)
        for k in range(c0, c0 + sp):
            cc = ws.cell(r, k)
            if is_head:
                cc.fill = HEAD_FILL; cc.border = B_HEAD
            else:
                cc.border = B_BOTTOM
        return c
    for j, h in enumerate(headers): put(row, j, h, True)
    ws.row_dimensions[row].height = head_h
    for i, r in enumerate(rows):
        rr = row + 1 + i
        for j in range(n): put(rr, j, r[j] if j < len(r) else "", False)
        ws.row_dimensions[rr].height = heights
    last = row + len(rows)
    for j in range(1, total + 1):
        ws.cell(last, j).border = Border(bottom=side(D.PRIMARY_MID, "thin"))
    return last

def note(ws, row, text, col=1):
    c = ws.cell(row, col, text); c.font = font(9, color=D.INK_FAINT); c.alignment = al(wrap=False)
    ws.row_dimensions[row].height = 18
    return row + 1

def status_dv(ws, rng, options, fills=None):
    dv = DataValidation(type="list", formula1='"' + ",".join(options) + '"', allow_blank=True)
    dv.error = "목록에서 고르세요"; dv.promptTitle = "선택"; dv.prompt = " / ".join(options)
    ws.add_data_validation(dv); dv.add(rng)
    palette = fills or {"정상": (D.OK, D.OK_BG), "주의": (D.WARN, D.WARN_BG), "위험": (D.DANGER, D.DANGER_BG)}
    for key, (fc, bg) in palette.items():
        if key in options:
            ws.conditional_formatting.add(rng, CellIsRule(
                operator="equal", formula=['"%s"' % key], fill=fill(bg), font=font(10, True, fc)))

def page(ws, landscape=True, fit=1, title_rows=None, area=None, header=""):
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = fit; ws.page_setup.fitToHeight = 0
    ws.page_margins.left = ws.page_margins.right = 0.5
    ws.page_margins.top = 0.6; ws.page_margins.bottom = 0.55
    ws.page_margins.header = ws.page_margins.footer = 0.3
    if title_rows: ws.print_title_rows = title_rows
    if area: ws.print_area = area
    ws.oddHeader.left.text = header or ws.title
    ws.oddHeader.left.size = 9; ws.oddHeader.left.color = "888888"
    ws.oddFooter.right.text = "&P / &N"; ws.oddFooter.right.size = 9
    ws.oddFooter.right.color = "888888"

# ══ 1. 안내 ════════════════════════════════════════════════════════════════
def sh_guide(wb):
    ws = wb.active; ws.title = "안내"
    sheet_head(ws, "운영 문서 표 모음", "기록하고 넘겨주는 표만 모았다. 필요한 시트를 복사해 쓰고, 쓰지 않는 시트는 지운다.",
               [22, 34, 34, 20, 18, 18, 18, 18])
    r = 5
    r = table(ws, r, ["시트", "이럴 때 쓴다", "기록하는 사람", "비고"],
              [["작업계획·결과", "작업 전 계획과 작업 후 결과를 한 장에 남길 때", "작업 수행자", "작업 건별로 시트 복사"],
               ["점검 체크리스트", "일·주·월 정기 점검 결과", "운영 담당", "월 단위로 시트 복사"],
               ["장비·IP 관리대장", "장비와 주소를 한곳에 모을 때", "운영 담당", "변경 시 즉시 수정"],
               ["장애 기록", "장애 발생부터 복구까지 시간순 기록", "대응 담당", "장애 건별로 시트 복사"],
               ["방안 비교", "선택지를 같은 기준으로 비교할 때", "검토자", "점수는 자동 계산"],
               ["인수인계 목록", "인계 항목과 진행 상태", "인계자·인수자", "완료율 자동 계산"]],
              aligns=["l", "l", "l", "l"], heights=22)
    r += 2
    ws.cell(r, 1, "쓰는 규칙").font = font(11, True, D.PRIMARY); r += 1
    for t in ["〈 〉 로 표시된 칸은 실제 값으로 바꿔 넣는다.",
              "회색 글씨의 예시 행은 지우고 쓴다.",
              "판정·상태 칸은 목록에서 고른다. 고르면 색이 자동으로 바뀐다.",
              "행을 추가할 때는 기존 행을 복사해 붙여 넣으면 서식과 목록이 함께 따라온다.",
              "인쇄는 [보기] → [페이지 나누기 미리 보기]로 잘리는 곳을 먼저 확인한다."]:
        c = ws.cell(r, 1, "· " + t); c.font = font(10, color=D.INK_SOFT); c.alignment = al(wrap=False)
        ws.row_dimensions[r].height = 19; r += 1
    r += 1
    ws.cell(r, 1, "공통 표 서식 (복사해서 쓰는 빈 표)").font = font(11, True, D.PRIMARY); r += 1
    note(ws, r, "아래 표를 통째로 복사해 다른 시트·통합 문서에 붙여 넣으면 같은 서식으로 이어 쓸 수 있다."); r += 1
    table(ws, r, ["〈머리글〉", "〈머리글〉", "〈머리글〉", "〈머리글〉"],
          [["", "", "", ""] for _ in range(4)], aligns=["l", "l", "c", "l"])
    page(ws, landscape=True, header="운영 문서 표 모음 · 안내")
    return ws

# ══ 2. 작업계획·결과 ═══════════════════════════════════════════════════════
def sh_work(wb):
    ws = wb.create_sheet("작업계획·결과")
    sheet_head(ws, "작업 계획 · 결과", "작업 전에 계획을 채우고, 작업 후 같은 표에 결과를 적는다.",
               [10, 32, 13, 13, 13, 11, 30])
    r = meta_block(ws, 5, [("작업명", "〈작업 이름〉"), ("대상", "〈장비 · 서비스〉"),
                           ("일시", "〈2026-01-01 22:00 ~ 23:30〉"), ("수행", "〈소속 / 이름〉"),
                           ("서비스 영향", "〈중단 15분 / 없음〉"), ("승인", "〈승인자 / 승인일〉")],
                   span=7)
    r += 2
    hdr_row = r
    r2 = table(ws, r, ["순서", "작업 내용", "계획 시각", "완료 시각", "담당", "결과", "특이 사항"],
               [[1, "〈사전 백업〉", "〈22:00〉", "", "〈이름〉", "", ""],
                [2, "〈설정 변경〉", "〈22:20〉", "", "〈이름〉", "", ""],
                [3, "〈서비스 재기동〉", "〈22:40〉", "", "〈이름〉", "", ""],
                [4, "〈정상 확인〉", "〈23:00〉", "", "〈이름〉", "", ""],
                [5, "", "", "", "", "", ""], [6, "", "", "", "", "", ""]],
               aligns=["c", "l", "c", "c", "c", "c", "l"], heights=22)
    status_dv(ws, f"F{hdr_row+1}:F{r2}", ["정상", "주의", "실패", "미실시"],
              {"정상": (D.OK, D.OK_BG), "주의": (D.WARN, D.WARN_BG),
               "실패": (D.DANGER, D.DANGER_BG), "미실시": (D.INK_FAINT, D.SURFACE)})
    r = r2 + 2
    ws.cell(r, 1, "작업 후 확인").font = font(11, True, D.PRIMARY); r += 1
    chk = table(ws, r, ["확인 항목", "기준", "결과", "확인자", "비고"],
                [["〈서비스 기동 상태〉", "〈active (running)〉", "", "", ""],
                 ["〈주요 기능 동작〉", "〈정상 응답〉", "", "", ""],
                 ["〈오류 로그〉", "〈신규 오류 없음〉", "", "", ""],
                 ["〈되돌리기 필요 여부〉", "〈불필요〉", "", "", ""]],
                aligns=["l", "l", "c", "c", "l"], spans=[2, 2, 1, 1, 1])
    status_dv(ws, f"E{r+1}:E{chk}", ["정상", "주의", "위험"])
    r = chk + 2
    ws.cell(r, 1, "총평").font = font(11, True, D.PRIMARY)
    ws.cell(r, 2, "〈결과를 한두 문장으로. 계획과 달라진 점, 후속 조치를 함께 적는다.〉").font = font(10)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    ws.cell(r, 2).alignment = al(indent=1)
    ws.row_dimensions[r].height = 30
    for c in range(2, 8): ws.cell(r, c).border = Border(bottom=side(D.LINE))
    ws.freeze_panes = f"A{hdr_row+1}"
    page(ws, landscape=True, title_rows=f"{hdr_row}:{hdr_row}", header="작업 계획 · 결과")
    return ws

# ══ 3. 점검 체크리스트 ═════════════════════════════════════════════════════
def sh_check(wb):
    ws = wb.create_sheet("점검 체크리스트")
    sheet_head(ws, "정기 점검 체크리스트", "점검 주기별 항목과 결과. 기준을 벗어나면 조치 내용을 반드시 남긴다.",
               [8, 14, 30, 26, 12, 12, 30])
    r = meta_block(ws, 5, [("점검 기간", "〈2026-01〉"), ("점검자", "〈소속 / 이름〉"),
                           ("확인자", "〈소속 / 이름〉")], span=7)
    r += 1
    ws.cell(r, 1, "요약").font = font(10, True, D.PRIMARY)
    ws.cell(r, 1).alignment = al(indent=1)
    c = ws.cell(r, 2)
    f0, f1 = r + 2, r + 21
    c.value = (f'="정상 "&COUNTIF(F{f0}:F{f1},"정상")&"   ·   주의 "&COUNTIF(F{f0}:F{f1},"주의")'
               f'&"   ·   위험 "&COUNTIF(F{f0}:F{f1},"위험")'
               f'&"   ·   미점검 "&COUNTBLANK(F{f0}:F{f1})')
    c.font = font(11, True, D.PRIMARY); c.alignment = al(indent=1)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
    ws.row_dimensions[r].height = 22
    r += 2
    hdr_row = r
    rows = [["일", "〈서비스〉", "〈서비스 기동 상태〉", "〈active (running)〉", "", "", ""],
            ["일", "〈서버〉", "〈디스크 사용률〉", "〈80% 미만〉", "", "", ""],
            ["일", "〈서버〉", "〈CPU · 메모리 사용률〉", "〈평균 70% 미만〉", "", "", ""],
            ["일", "〈로그〉", "〈오류 로그 발생 건수〉", "〈신규 오류 없음〉", "", "", ""],
            ["주", "〈백업〉", "〈백업 성공 여부〉", "〈최근 7일 모두 성공〉", "", "", ""],
            ["주", "〈보안〉", "〈접속 실패 기록〉", "〈비정상 시도 없음〉", "", "", ""],
            ["월", "〈용량〉", "〈로그 보관 용량〉", "〈기준 용량 이내〉", "", "", ""],
            ["월", "〈이중화〉", "〈대기 장비 상태〉", "〈정상 대기〉", "", "", ""],
            ["월", "〈문서〉", "〈구성 정보 최신화〉", "〈변경 반영 완료〉", "", "", ""]]
    rows += [["", "", "", "", "", "", ""] for _ in range(11)]
    last = table(ws, r, ["주기", "구분", "점검 항목", "정상 기준", "점검일", "판정", "조치 내용"],
                 rows, aligns=["c", "l", "l", "l", "c", "c", "l"], heights=22)
    status_dv(ws, f"F{hdr_row+1}:F{last}", ["정상", "주의", "위험"])
    dv = DataValidation(type="list", formula1='"일,주,월,분기"', allow_blank=True)
    ws.add_data_validation(dv); dv.add(f"A{hdr_row+1}:A{last}")
    ws.freeze_panes = f"A{hdr_row+1}"
    ws.auto_filter.ref = f"A{hdr_row}:G{last}"
    page(ws, landscape=True, title_rows=f"{hdr_row}:{hdr_row}", header="정기 점검 체크리스트")
    return ws

# ══ 4. 장비·IP 관리대장 ════════════════════════════════════════════════════
def sh_asset(wb):
    ws = wb.create_sheet("장비·IP 관리대장")
    sheet_head(ws, "장비 · IP 관리대장", "장비와 주소를 한곳에 모은다. 변경이 생기면 바로 고치고 변경일을 남긴다.",
               [16, 12, 18, 18, 16, 14, 14, 14, 22])
    r = 5
    hdr_row = r
    ex = [["〈web-01〉", "〈웹〉", "〈10.0.0.11〉", "〈10.0.9.11〉", "〈IDC A · R12〉",
           "〈Rocky 9〉", "〈운영〉", "〈이름〉", "〈2026-01-01〉"],
          ["〈was-01〉", "〈WAS〉", "〈10.0.0.21〉", "〈10.0.9.21〉", "〈IDC A · R12〉",
           "〈Rocky 9〉", "〈운영〉", "〈이름〉", "〈2026-01-01〉"],
          ["〈db-01〉", "〈DB〉", "〈10.0.0.31〉", "〈10.0.9.31〉", "〈IDC A · R13〉",
           "〈Rocky 9〉", "〈운영〉", "〈이름〉", "〈2026-01-01〉"]]
    ex += [["", "", "", "", "", "", "", "", ""] for _ in range(14)]
    last = table(ws, r, ["호스트명", "구분", "서비스 IP", "관리 IP", "위치", "OS · 버전",
                         "상태", "담당", "최근 변경일"],
                 ex, aligns=["l", "c", "l", "l", "l", "l", "c", "c", "c"])
    status_dv(ws, f"G{hdr_row+1}:G{last}", ["운영", "대기", "점검", "폐기"],
              {"운영": (D.OK, D.OK_BG), "대기": (D.PRIMARY_MID, D.PRIMARY_TINT),
               "점검": (D.WARN, D.WARN_BG), "폐기": (D.INK_FAINT, D.SURFACE)})
    ws.freeze_panes = f"B{hdr_row+1}"
    ws.auto_filter.ref = f"A{hdr_row}:I{last}"
    note(ws, last + 2, "※ 실제 IP·계정 정보가 들어가므로 보관 위치와 열람 범위를 정해 둔다.")
    page(ws, landscape=True, title_rows=f"{hdr_row}:{hdr_row}", header="장비 · IP 관리대장")
    return ws

# ══ 5. 장애 기록 ═══════════════════════════════════════════════════════════
def sh_incident(wb):
    ws = wb.create_sheet("장애 기록")
    sheet_head(ws, "장애 기록", "시간순으로 남긴다. 추측과 확인된 사실을 섞지 않는다.",
               [14, 16, 34, 26, 14, 14, 26])
    r = meta_block(ws, 5, [("장애명", "〈한 줄 요약〉"), ("대상", "〈시스템 · 서비스〉"),
                           ("발생 / 인지", "〈14:20 / 14:25〉"), ("복구", "〈15:05〉"),
                           ("영향", "〈전체 사용자 45분 중단〉"), ("등급", "〈1급 / 2급 / 3급〉")], span=7)
    r += 2
    hdr_row = r
    rows = [["〈14:20〉", "〈발생〉", "〈증상 — 관측된 사실만〉", "", "", "", ""],
            ["〈14:25〉", "〈인지〉", "〈알람 · 신고 경로〉", "", "〈이름〉", "", ""],
            ["〈14:31〉", "〈조치〉", "〈수행한 조치〉", "〈조치 결과〉", "〈이름〉", "", ""],
            ["〈14:52〉", "〈원인 확인〉", "〈확인 내용〉", "〈확인 근거 — 로그·지표〉", "〈이름〉", "", ""],
            ["〈15:05〉", "〈복구〉", "〈정상 확인 방법〉", "〈정상 확인 결과〉", "〈이름〉", "", ""]]
    rows += [["", "", "", "", "", "", ""] for _ in range(8)]
    last = table(ws, r, ["시각", "구분", "내용", "결과 · 근거", "담당", "사실/추정", "비고"],
                 rows, aligns=["c", "c", "l", "l", "c", "c", "l"], heights=24)
    dv = DataValidation(type="list", formula1='"발생,인지,조치,원인 확인,복구,보고"', allow_blank=True)
    ws.add_data_validation(dv); dv.add(f"B{hdr_row+1}:B{last}")
    status_dv(ws, f"F{hdr_row+1}:F{last}", ["사실", "추정"],
              {"사실": (D.OK, D.OK_BG), "추정": (D.WARN, D.WARN_BG)})
    r = last + 2
    ws.cell(r, 1, "마무리").font = font(11, True, D.PRIMARY); r += 1
    fin = table(ws, r, ["구분", "내용", "담당", "기한", "비고"],
                [["원인", "〈직접 원인 한 줄〉", "", "", ""],
                 ["조치", "〈수행한 조치와 정상 확인 방법〉", "", "", ""],
                 ["재발 방지", "〈무엇을 어떻게 바꿀 것인가〉", "〈이름〉", "〈2026-01-31〉", ""],
                 ["공유", "〈누구에게 언제 알렸는가〉", "", "", ""]],
                aligns=["l", "l", "c", "c", "l"], heights=24, spans=[1, 2, 1, 1, 2])
    ws.freeze_panes = f"A{hdr_row+1}"
    page(ws, landscape=True, title_rows=f"{hdr_row}:{hdr_row}", header="장애 기록")
    return ws

# ══ 6. 방안 비교 ═══════════════════════════════════════════════════════════
def sh_compare(wb):
    ws = wb.create_sheet("방안 비교")
    sheet_head(ws, "방안 비교", "같은 기준으로 비교한다. 가중치와 점수를 넣으면 합계가 자동 계산된다.",
               [26, 12, 16, 16, 16, 30])
    r = meta_block(ws, 5, [("검토 항목", "〈무엇을 정하려는가〉"), ("결정 기한", "〈2026-01-31〉"),
                           ("검토자", "〈소속 / 이름〉")], span=6)
    r += 2
    hdr_row = r
    crit = [["〈비용〉", 30], ["〈작업 시간〉", 15], ["〈서비스 영향〉", 25],
            ["〈운영 편의〉", 15], ["〈위험도〉", 15]]
    rows = [[c, w, "", "", "", "〈판단 근거〉"] for c, w in crit]
    last = table(ws, r, ["기준", "가중치(%)", "방안 A", "방안 B", "방안 C", "판단 근거"],
                 rows, aligns=["l", "c", "c", "c", "c", "l"], heights=24)
    for rr in range(hdr_row + 1, last + 1):
        for col in "CDE":
            c = ws[f"{col}{rr}"]
            c.font = font(10); c.alignment = al("center")
    tot = last + 1
    ws.cell(tot, 1, "가중 합계").font = font(10, True, D.PRIMARY)
    ws.cell(tot, 1).alignment = al(indent=1)
    ws.cell(tot, 2, f"=SUM(B{hdr_row+1}:B{last})").font = font(10, True, D.INK_SOFT)
    ws.cell(tot, 2).alignment = al("center")
    for i, col in enumerate("CDE"):
        c = ws.cell(tot, 3 + i)
        c.value = f"=SUMPRODUCT($B${hdr_row+1}:$B${last},{col}{hdr_row+1}:{col}{last})/100"
        c.font = font(12, True, D.PRIMARY); c.alignment = al("center"); c.number_format = "0.0"
    for j in range(1, 7):
        ws.cell(tot, j).fill = fill(D.PRIMARY_TINT)
        ws.cell(tot, j).border = Border(top=side(D.PRIMARY_MID, "thin"), bottom=side(D.PRIMARY_MID, "thin"))
    ws.row_dimensions[tot].height = 26
    ws.cell(tot, 6, "〈점수 1~5 를 넣으면 자동 계산〉").font = font(9, color=D.INK_FAINT)
    ws.cell(tot, 6).alignment = al(indent=1)
    r = tot + 2
    ws.cell(r, 1, "선정 결과").font = font(11, True, D.PRIMARY)
    ws.cell(r, 2, "〈선택한 방안과 이유를 한두 문장으로. 점수만으로 정하지 않는다.〉").font = font(10)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    ws.cell(r, 2).alignment = al(indent=1); ws.row_dimensions[r].height = 30
    for c in range(2, 7): ws.cell(r, c).border = Border(bottom=side(D.LINE))
    note(ws, r + 2, "※ 점수는 1(나쁨) ~ 5(좋음). 가중치 합계는 100 이 되도록 맞춘다.")
    page(ws, landscape=False, title_rows=f"{hdr_row}:{hdr_row}", header="방안 비교")
    return ws

# ══ 7. 인수인계 목록 ═══════════════════════════════════════════════════════
def sh_handover(wb):
    ws = wb.create_sheet("인수인계 목록")
    sheet_head(ws, "인수인계 목록", "넘겨줄 것과 남은 일을 한 장에. 완료율은 자동으로 계산된다.",
               [14, 30, 30, 12, 14, 14, 26])
    r = meta_block(ws, 5, [("대상 업무", "〈시스템 · 서비스〉"), ("인계자", "〈소속 / 이름〉"),
                           ("인수자", "〈소속 / 이름〉"), ("기간", "〈2026-01-01 ~ 01-31〉")], span=7)
    r += 1
    prog = r
    ws.cell(r, 1, "완료율").font = font(10, True, D.PRIMARY)
    r += 2
    hdr_row = r
    rows = [["계정·권한", "〈계정 목록과 권한 범위〉", "〈보관 위치 · 전달 방법〉", "", "", "", ""],
            ["접속 정보", "〈접속 경로 · 점프 서버〉", "〈접속 확인까지 완료〉", "", "", "", ""],
            ["운영 문서", "〈운영 매뉴얼 · 구성도〉", "〈최신본 위치〉", "", "", "", ""],
            ["정기 점검", "〈점검 항목과 주기〉", "〈실제 수행 1회 동행〉", "", "", "", ""],
            ["장애 대응", "〈연락 체계 · 과거 이력〉", "〈에스컬레이션 기준 공유〉", "", "", "", ""],
            ["모니터링", "〈알람 수신 대상 변경〉", "〈수신 확인까지〉", "", "", "", ""],
            ["미결 과제", "〈진행 중인 작업〉", "〈현재 상태와 다음 할 일〉", "", "", "", ""],
            ["", "", "", "", "", "", ""], ["", "", "", "", "", "", ""]]
    last = table(ws, r, ["구분", "인계 항목", "확인 방법", "상태", "완료일", "담당", "남은 일 · 비고"],
                 rows, aligns=["l", "l", "l", "c", "c", "c", "l"], heights=24)
    status_dv(ws, f"D{hdr_row+1}:D{last}", ["완료", "진행", "예정", "보류"],
              {"완료": (D.OK, D.OK_BG), "진행": (D.WARN, D.WARN_BG),
               "예정": (D.INK_FAINT, D.SURFACE), "보류": (D.DANGER, D.DANGER_BG)})
    c = ws.cell(prog, 2)
    c.value = (f'=IF(COUNTA(B{hdr_row+1}:B{last})=0,"",'
               f'TEXT(COUNTIF(D{hdr_row+1}:D{last},"완료")/COUNTA(B{hdr_row+1}:B{last}),"0%")'
               f'&"  (완료 "&COUNTIF(D{hdr_row+1}:D{last},"완료")&" / 전체 "&COUNTA(B{hdr_row+1}:B{last})&")")')
    c.font = font(12, True, D.PRIMARY); c.alignment = al(indent=1)
    ws.row_dimensions[prog].height = 22
    ws.freeze_panes = f"A{hdr_row+1}"
    note(ws, last + 2, "※ ‘완료’는 인수자가 직접 해 보고 확인한 상태를 뜻한다. 설명만 한 것은 ‘진행’으로 둔다.")
    page(ws, landscape=True, title_rows=f"{hdr_row}:{hdr_row}", header="인수인계 목록")
    return ws

def build():
    wb = Workbook()
    sh_guide(wb); sh_work(wb); sh_check(wb); sh_asset(wb); sh_incident(wb); sh_compare(wb); sh_handover(wb)
    wb.properties.title = "운영 문서 표 모음"
    wb.properties.creator = "IT 인프라 문서 라이브러리"
    wb.active = 0
    wb.save(OUT)
    return OUT

if __name__ == "__main__":
    print("Excel 생성:", build())
