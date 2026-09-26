# -*- coding: utf-8 -*-
"""사용법 · 기준 시트."""
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Border
from xlkit import Canvas, P, side, fill, font, FONT, MONO, USE, align
from xlparts import port, pname, module, slot, legend_colors, legend_ids, badge, KSTYLE, _tw

TAB_GUIDE = "8C99A8"


def _head(cv, title, sub, doc="장비 전·후면도 라이브러리"):
    cv.h(1, 24)
    cv.text(1, 2, title, sz=16, b=True)
    cv.text(1, 2 + max(6, int(_tw(title, 16)) + 2), sub, sz=10, col=P["s700"])
    cv.text(1, 63, doc, sz=8, col=P["s500"], h="right")
    for c in range(2, 64):
        cv.add_border(1, c, bottom=side("medium", P["navy"]))
    cv.h(2, 6)
    cv.h(3, 4)
    return 4


def _sec(cv, r, no, text):
    cv.put(r, 2, no, style="번호 배지")
    cv.text(r, 3, "  " + text, sz=9.5, b=True, col=P["navy"])
    for c in range(3, 64):
        cv.add_border(r, c, bottom=side("thin", P["s300"]))
    cv.h(r, 15)
    return r + 1


def _rich(cv, r, c, parts, sz=8):
    """parts: [(text, style)] style: None | 'b' | 'k'(key/굵게 남색) | 'g'(회색) | 'm'(고정폭)"""
    blocks = []
    for t, s in parts:
        kw = dict(rFont=FONT, sz=sz, color=P["ink"])
        if s == "b":
            kw["b"] = True
        elif s == "k":
            kw.update(b=True, color=P["navy"])
        elif s == "g":
            kw["color"] = P["s500"]
        elif s == "m":
            kw.update(rFont=MONO, color=P["ink"])
        blocks.append(TextBlock(InlineFont(**kw), t))
    cv.c(r, c).value = CellRichText(*blocks)
    cv.c(r, c).alignment = align("left", "center")


def _lines(cv, r, items, c=3, sz=8, step_badges=False):
    """items: str | [(text, style)...] | ('-', str) 들여쓰기"""
    for i, it in enumerate(items):
        cv.h(r, 13)
        indent = 0
        if isinstance(it, tuple) and it and it[0] == "-":
            indent, it = 2, it[1]
        cc = c + indent
        if step_badges:
            cv.put(r, cc, i + 1, style="포트 서비스")
            cv.cell_box(r, cc)
            cc += 1
            prefix = "  "
        else:
            prefix = "‣ " if indent else "· "
        if isinstance(it, list):
            _rich(cv, r, cc, [(prefix, "g")] + it, sz)
        else:
            cv.text(r, cc, prefix + it, sz=sz, col=P["ink"])
        r += 1
    return r


def build_guide(wb, rack=False):
    cv = Canvas(wb, "사용법", TAB_GUIDE)
    r = _head(cv, "사용법", "5분 요약 — 복사 · 수정 · 번호 · 랙 관리 · 인쇄")
    r += 1
    r = _sec(cv, r, 0, "파일 구성과 탭 색")
    r = _lines(cv, r, [
        [("1_장비도면_라이브러리.xlsx", "k"), ("  사용법 · 기준 · 부품 · 빈 양식 13종 · 작성 예시 11종", None)],
        [("2_랙예시_DC1-R02.xlsx", "k"), ("  랙 1대를 파일 1개로 관리하는 예 (랙 실장도 + 장비 시트 13개, 케이블 라벨 양쪽 일치)", None)],
        [("3_랙_빈양식.xlsx", "k"), ("  새 랙 파일을 만들 때 복사해서 시작", None)],
        [("탭 색", "b"), ("  회색 = 안내   주황 = 부품   파랑 = 빈 양식(양식-)   초록 = 작성 예시(예시-)", None)],
    ])
    r += 1
    r = _sec(cv, r, 1, "새 장비 시트 만들기 — 6단계")
    r = _lines(cv, r, [
        [("가장 비슷한 ", None), ("예시-", "k"), (" 또는 ", None), ("양식-", "k"),
         (" 시트 탭 우클릭 → [이동/복사] → 대상 파일(랙 파일) 선택 → [복사본 만들기] 체크 → 확인", None)],
        [("시트 이름 바꾸기: 탭 더블클릭 → ", None), ("U20 SRV-01", "m"), (" 처럼 U 위치 + 장비명", None)],
        [("1~2행 머리글 입력: 장비명 · 모델 · 설치 위치 · 랙·U · 구분 · 시리얼 · 관리 IP · 작성일", None)],
        [("부품 바꾸기: [부품] 시트에서 부품 범위 복사 → 장비 시트에서 놓을 자리의 ", None), ("왼쪽 위 칸", "b"),
         (" 클릭 → 붙여넣기 (같은 크기면 그대로 덮어써짐)", None)],
        [("연결된 포트 칸에 ", None), ("도면 번호", "b"), (" (1, 2, 3 …) 입력 → [홈 > 셀 스타일]에서 '포트 서비스' 등 용도 색 선택", None)],
        [("아래 표에 번호별 정보 입력 (용도·면은 목록 선택) → ", None), ("번호 칸이 분홍색", "b"),
         ("이면 도면에 그 번호가 없거나 두 번 이상 있음", None)],
    ], step_badges=True)
    r += 1
    r = _sec(cv, r, 2, "부품 다루기")
    r = _lines(cv, r, [
        [("교체", "b"), (": 슬롯 모듈은 19칸(1/3 폭) · 29칸(1/2 폭)으로 크기가 통일 → 같은 크기 부품을 그 자리에 붙이면 끝", None)],
        [("삭제", "b"), (": 범위 선택 → [홈 > 지우기 > 모두 지우기]  (Delete 키는 글자만 지움) — 장비 바탕이 흰색이라 지운 자리가 자연스러움", None)],
        [("포트 색", "b"), (": 포트 칸 선택 → 셀 스타일 클릭 (테두리는 그대로). 되돌리기 = '포트 미연결'", None)],
        [("포트 늘리기", "b"), (": 옆 포트와 아래 포트명 칸을 함께 복사해 붙이기.  ", None), ("연속 번호", "b"),
         (": 1, 2 입력 → 두 칸 선택 → 채우기 핸들 끌기", None)],
        [("크기가 없는 모듈", "b"), (": 범위 선택 → 셀 스타일 '모듈 채움' → [테두리 > 바깥쪽 테두리]  (선 색 회색)", None)],
        [("열 너비 · 행 높이는 바꾸지 않기", "b"), (" — 모든 시트가 같은 격자(1칸 ≈ 4.2mm)라서 어디에 붙여도 크기가 같음", None)],
    ])
    r += 1
    r = _sec(cv, r, 3, "도면 번호 · 포트명 · 케이블 라벨 — 서로 다른 세 가지")
    cv.h(r, 14)
    legend_ids(cv, r, 3)
    r += 1
    r = _lines(cv, r, [
        [("도면 번호", "k"), (" = 색 칸 안 흰 숫자. 장비 한 대 안에서 겹치지 않게 1부터, 전면 → 후면 · 위 → 아래 · 왼쪽 → 오른쪽 순서", None)],
        ("-", "스위치 · 패치패널 · PDU처럼 번호가 인쇄된 장비는 인쇄된 번호를 그대로 도면 번호로 사용"),
        ("-", "슬롯마다 포트 번호가 반복되는 섀시는 슬롯 × 100 + 포트  (예: 슬롯 2의 3번 포트 = 203)"),
        [("포트명", "k"), (" = 장비·OS에 표기된 이름 (iLO, LOM1, Eth1/21 …). 도면에서는 포트 아래 회색 작은 글씨, 표에서는 '포트명' 칸", None)],
        [("케이블 라벨", "k"), (" = 케이블에 실제로 붙이는 라벨. 표의 노란 칸에만 적고 도면에는 적지 않음. 케이블 양쪽 장비의 표에 같은 라벨", None)],
    ])
    r += 1
    r = _sec(cv, r, 4, "다른 파일로 복사")
    r = _lines(cv, r, [
        [("시트 통째로 (권장)", "b"), (": 탭 우클릭 → [이동/복사] → 대상 파일 → [복사본 만들기]. 서식 · 인쇄 설정 · 목록 · 번호 확인이 모두 유지", None)],
        [("보고서에 도면만", "b"), (": 도면 범위 선택 → [홈 > 복사 ▾ > 그림으로 복사] → 워드 · PPT · 다른 엑셀에 붙여넣기 (모양 그대로, 수정은 원본에서)", None)],
        [("셀째로 다른 시트에", "b"), (": 받는 시트도 같은 격자여야 크기가 같음 → 양식 시트를 먼저 복사해 두고 그 위에 붙이기", None)],
        [("표만", "b"), (": 표 범위 복사 → 받는 곳에서 [값 붙여넣기]. 병합 칸에는 여러 칸을 한 번에 붙일 수 없으니 한 열씩", None)],
    ])
    r += 1
    r = _sec(cv, r, 5, "랙별로 묶어 관리 (장비 200대 → 랙 파일 10~20개)")
    r = _lines(cv, r, [
        [("랙 1대 = 파일 1개", "b"), (" (예: DC1_R02.xlsx). 첫 시트 [랙] = 실장도 + 장비 목록, 나머지 시트 = 장비 1대씩", None)],
        [("장비 추가", "b"), (": ① 장비 시트를 랙 파일로 복사  ② [랙] 목록 빈 줄에 U · 높이 · 면 · 장비명 · 모델 · 구분 · 시트 이름 입력 → 실장도 자동 표시 + ▶ 링크", None)],
        [("장비 이동 · 제거", "b"), (": 목록의 U 숫자만 고치면 실장도가 따라감. 제거는 목록 줄 내용 지우기 + 시트 삭제", None)],
        [("탭 순서", "b"), ("는 랙 위쪽 장비부터 (탭 끌어서 정렬). 랙이 적으면 한 파일에 [랙] 시트를 여러 개(랙_R01, 랙_R02…) 두어도 됨", None)],
    ])
    r += 1
    r = _sec(cv, r, 6, "인쇄")
    r = _lines(cv, r, [
        [("모든 시트: A4 가로 · 배율 100% · 여백 좌우 10mm / 위아래 12mm · 바닥글(파일명 · 시트명 · 쪽 번호) 설정 완료", None)],
        [("Ctrl+P 미리보기로 확인. ", None), ("배율을 줄이지 말 것", "b"), (" — 장비마다 글자 크기가 달라짐 (너비가 넘칠 때만 자동으로 조금 줄어듦)", None)],
        [("표가 길면 다음 쪽으로 이어지고 1~3행 제목이 반복됨. ", None), ("표 줄 추가", "b"),
         (": 빈 줄의 행 번호를 눌러 복사 → 아래 행 번호 우클릭 [복사한 셀 삽입] (병합·확인 기능 유지)", None)],
        [("랙 파일 전체: [인쇄 > 설정 > 전체 통합 문서 인쇄] → 랙 실장도 + 모든 장비가 차례로. PDF는 [다른 이름으로 저장 > PDF]", None)],
    ])
    r += 1
    r = _sec(cv, r, 7, "자주 하는 실수")
    r = _lines(cv, r, [
        [("부품이 작거나 크게 붙음 → 받는 시트의 열 너비 · 행 높이가 다름. 양식 시트를 복사해서 시작", None)],
        [("'병합된 셀의 일부를 변경할 수 없습니다' → 병합 칸에 여러 칸을 붙여넣음. 한 칸씩 입력", None)],
        [("번호 칸이 분홍색 → 그 번호가 도면에 없거나 두 번 이상 있음", None)],
        [("도형 · 그림은 쓰지 않음 → 전부 셀이라 복사 · 행 삽입 · 인쇄해도 어긋나지 않음", None)],
    ])
    cv.finalize()
    cv.page_setup(63, r, title_rows="1:3", breaks=[])
    return cv


def build_standard(wb):
    cv = Canvas(wb, "기준", TAB_GUIDE)
    r = _head(cv, "도면 기준", "장비가 달라도 인쇄 크기 · 모양이 같도록 정한 규칙 (한 장 요약)")
    r += 1
    # 1. 크기 표
    r = _sec(cv, r, 1, "격자와 크기")
    rows = [
        ("격자", "1칸 = 약 4.2mm 정사각형 (열 너비 2.0 · 행 높이 12pt)", "모든 시트 동일 — 바꾸지 않음"),
        ("장비 폭", "19인치 랙 장비 = 61칸 (C~BK열, 양 끝 1칸은 랙 귀)", "하프 폭 30칸 · 타워 29칸 (앞·뒤 나란히)"),
        ("장비 높이", "1U 5칸 · 2U 10칸 · 3U 이상 U당 4칸 (4U 16칸)", "한 면 최대 38칸 = 한 쪽"),
        ("모듈", "높이 2칸 (제목+포트 / 포트명) · 1칸 (띠형) · 3칸 (제목 / 포트 / 포트명)", "폭 1/3 = 19칸 · 1/2 = 29칸 · 전체 = 59칸"),
        ("간격", "모듈 사이 1칸 · 외곽선 안쪽 위아래 1칸", "위아래로 쌓는 슬롯은 0칸"),
        ("포트", "2칸 × 1행 (RJ45 · SFP · LC · USB · C14)", "넓은 포트 3칸 (QSFP · SAS HD · D-sub · C20) · 포트 사이 1칸, 고밀도 0칸"),
        ("디스크", "SFF 2 × 6칸 (1U는 2 × 3칸) · LFF 14 × 2칸", "장착 = 회색, 빈 베이 = 흰색"),
    ]
    for a, b, c in [("항목", "기준", "비고")] + rows:
        hdr = a == "항목"
        cv.h(r, 14)
        for x in range(3, 64):
            cv.add_border(r, x, bottom=side("thin", P["s200"] if not hdr else P["s400"]))
            if hdr:
                cv.c(r, x).fill = fill(P["s100"])
        cv.text(r, 3, a, sz=8, b=True, col=P["navy"] if not hdr else P["ink"])
        cv.text(r, 10, b, sz=8, b=hdr)
        cv.text(r, 40, c, sz=7.5 if not hdr else 8, col=P["s500"] if not hdr else P["ink"], b=hdr)
        r += 1
    r += 1
    # 2. 글자
    r = _sec(cv, r, 2, "글자 (모두 맑은 고딕, 케이블 라벨만 Consolas)")
    samples = [("장비명", 16, True, P["ink"], "SRV-01"), ("모델", 11, False, P["s700"], "DL380 Gen10"),
               ("구역 탭", 8, True, P["white"], "전 면"), ("모듈 제목", 7, True, P["ink"], "S1 · NIC 2P"),
               ("도면 번호", 8, True, P["white"], "12"), ("포트명", 6.5, False, P["s500"], "LOM1"),
               ("표 본문", 8.5, False, P["ink"], "TOR-A  Eth1/21"), ("케이블 라벨", 8.5, False, P["ink"], "R02-D001")]
    x = 3
    for name, sz, b, col, s in samples:
        cv.text(r, x, f"{name} {sz}pt", sz=7, col=P["s500"])
        cell = cv.text(r + 1, x, s, sz=sz, b=b, col=col, name=MONO if name == "케이블 라벨" else FONT)
        if name == "구역 탭":
            cv.c(r + 1, x).fill = fill(P["navy"])
            cv.c(r + 1, x + 1).fill = fill(P["navy"])
        if name == "도면 번호":
            cv.c(r + 1, x).fill = fill(P["svc"])
            cv.c(r + 1, x + 1).fill = fill(P["svc"])
            cv.c(r + 1, x).alignment = align("center")
            cv.merge(r + 1, x, r + 1, x + 1)
        if name == "케이블 라벨":
            for xx in range(x, x + 5):
                cv.c(r + 1, xx).fill = fill(P["tape"])
        x += 7 if name not in ("모델",) else 10
    cv.h(r + 1, 24)
    r += 3
    # 3. 선
    r = _sec(cv, r, 3, "선")
    lines = [("장비 외곽", side("medium", P["s700"]), "굵은 선 · 진회색 #3A4856"),
             ("모듈", side("thin", P["s400"]), "가는 선 · 회색 #8C99A8"),
             ("빈 슬롯", side("dashed", P["s300"]), "점선 · 연회색 #B4BFCB"),
             ("표 줄", side("thin", P["s200"]), "가는 선 · #D5DCE4")]
    x = 3
    for name, s, desc in lines:
        cv.text(r, x, name, sz=8, b=True)
        for xx in range(x, x + 10):
            cv.add_border(r + 1, xx, bottom=s)
        cv.text(r + 2, x, desc, sz=7, col=P["s500"])
        x += 15
    r += 4
    # 4. 색
    r = _sec(cv, r, 4, "색 = 용도 (포트 칸 · 표 번호 칸 · PPT 연결선 공통)")
    desc = {"svc": "업무 · 데이터 (이더넷)", "mgmt": "관리 · OOB · BMC", "bak": "백업망", "san": "SAN · FC · SAS · iSCSI",
            "ic": "HA · 스택 · 피어링크 · IB", "con": "콘솔 · 시리얼", "pwr": "전원 (PSU · PDU)"}
    x = 3
    col_i = 0
    rr = r
    for k, n in USE + [("none", "미연결")]:
        cv.put(rr, x, style=KSTYLE[k])
        cv.put(rr, x + 1, style=KSTYLE[k])
        cv.cell_box(rr, x)
        cv.cell_box(rr, x + 1)
        if k == "none":
            for xx in (x, x + 1):
                cv.c(rr, xx).border = Border(left=side("thin", P["s400"]), right=side("thin", P["s400"]),
                                             top=side("thin", P["s400"]), bottom=side("thin", P["s400"]))
        cv.merge(rr, x, rr, x + 1)
        cv.text(rr, x + 2, n, sz=8, b=True)
        hexv = "#" + (P[k] if k != "none" else "FFFFFF")
        cv.text(rr, x + 7, hexv, sz=7, col=P["s500"], name=MONO)
        cv.text(rr, x + 11, desc.get(k, "연결 안 됨 · 로컬 I/O"), sz=7.5, col=P["s700"])
        cv.h(rr, 14)
        rr += 1
        col_i += 1
        if col_i == 4:
            x, rr = 34, r
    r += 5
    # 5. 쪽 구성·인쇄
    r = _sec(cv, r, 5, "쪽 구성 · 인쇄")
    r = _lines(cv, r, [
        [("1~2U", "b"), (": 전면 + 후면 + 표(약 14줄) 한 쪽   ", None), ("3~4U", "b"), (": 전면 + 후면 한 쪽, 표 다음 쪽   ", None),
         ("5U 이상", "b"), (": 전면 · 후면 · 표 각각 한 쪽   ", None), ("스위치류", "b"), (": 2단 표", None)],
        [("A4 가로 · 배율 100% (너비가 넘칠 때만 1쪽 너비 맞춤으로 자동 축소) · 여백 좌우 10mm, 위아래 12mm", None)],
        [("1~3행(장비명·정보) 모든 쪽에 반복 · 바닥글 = 파일명 · 시트명 · 쪽 번호 자동", None)],
        [("셀 스타일(홈 > 셀 스타일 > 사용자 지정): 포트 서비스 ~ 포트 전원 · 포트 미연결 · 포트명 · 모듈 채움 · 디스크 장착 · 디스크 빈칸 · 번호 배지 · 케이블 라벨 · 표 머리글 · 표 본문 · 표 번호", None)],
    ])
    cv.finalize()
    cv.page_setup(63, r, title_rows="1:3", breaks=[])
    return cv


def build_index(wb, entries):
    """목차 시트: 모든 시트로 가는 링크. entries = [(구역, 시트이름, 설명)]"""
    from openpyxl.worksheet.hyperlink import Hyperlink
    cv = Canvas(wb, "목차", TAB_GUIDE)
    r = _head(cv, "목차", "시트를 클릭하면 이동합니다  ·  탭 색: 회색 안내 · 주황 부품 · 파랑 빈 양식 · 초록 작성 예시")
    r += 1
    cols = {"안내": 2, "빈 양식": 2, "작성 예시": 33}
    groups = {}
    for sec, sh, desc in entries:
        groups.setdefault(sec, []).append((sh, desc))
    pos = {"안내": (r, 2), "빈 양식": (r + 6, 2), "작성 예시": (r, 33)}
    for sec, items in groups.items():
        rr, cc = pos[sec]
        cv.text(rr, cc, sec, sz=10, b=True, col=P["navy"])
        for x in range(cc, cc + 30):
            cv.add_border(rr, x, bottom=side("thin", P["s300"]))
        cv.h(rr, 16)
        rr += 1
        for sh, desc in items:
            cv.h(rr, 15)
            cell = cv.c(rr, cc)
            cell.value = "▶ " + sh
            cell.hyperlink = Hyperlink(ref=cell.coordinate, location=f"'{sh}'!A1", display=sh)
            cell.font = font(9, True, P["svc"])
            cv.text(rr, cc + 11, desc, sz=8.5, col=P["s700"])
            rr += 1
    cv.finalize()
    cv.page_setup(63, r + 6 + len(groups.get("빈 양식", [])) + 2, title_rows="1:3", breaks=[])
    return cv
