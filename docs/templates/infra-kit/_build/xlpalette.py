# -*- coding: utf-8 -*-
"""부품 시트: 복사해서 쓰는 셀 부품 모음 (구역별, 페이지 넘김 자동)."""
from openpyxl.worksheet.pagebreak import Break
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from xlkit import Canvas, P, side, fill, font, FONT
from xlparts import _tw
from xlparts import (port, pname, module, slot, vmod, strip, psu, fan, sff_bays, lff_bays, lff_bay, port_block,
                     ports_row, badge, ear_frame, side_tab, legend_colors, legend_ids, conn_table, TABLE_FULL,
                     TABLE_HALF, note, title)
import xldevices as D

TAB_PARTS = "D9822B"
PAGE_ROWS = 40


def vslot(cv, r, c, h=10, ttl="1", nports=2):
    module(cv, r, c, h, 3, ttl)
    for i in range(nports):
        port(cv, r + 2 + i * 2, c + 1)
    return h, 3


def _single(kind, name="P1", w=2, label=None):
    def f(cv, r, c):
        port(cv, r, c, kind, None, w, label=label)
        if name:
            pname(cv, r + 1, c, name, w)
        return 2, w
    return f


def _frame(h, w, tab="전  면"):
    def f(cv, r, c):
        side_tab(cv, r, c, h, tab)
        ear_frame(cv, r, c + 1, h, w)
        return h, w + 1
    return f


def _tbl(cols, n=2):
    def f(cv, r, c):
        conn_table(cv, r, c, cols, [], n)
        return n + 1, sum(w for _, w in cols)
    return f


def _legend(cv, r, c):
    w = legend_colors(cv, r, c)
    return 1, w


def _ids(cv, r, c):
    w = legend_ids(cv, r, c)
    return 1, w


def _tape(cv, r, c):
    for x in range(c, c + 7):
        cv.put(r, x, style="케이블 라벨")
    cv.c(r, c).value = "R02-D001"
    cv.merge(r, c, r, c + 6)
    return 1, 7


def _memo(cv, r, c):
    cv.put(r, c, "※ 메모: 7pt 회색 — 도면 아래 설명", style="메모")
    return 1, 16


def _dir(cv, r, c):
    cv.text(r, c, "◀ 왼쪽", sz=7, b=True, col=P["s500"])
    cv.text(r, c + 20, "오른쪽 ▶", sz=7, b=True, col=P["s500"], h="right")
    cv.text(r, c + 7, "(보는 방향 기준)", sz=6.5, col=P["s400"])
    return 1, 21


def _blade(name):
    def f(cv, r, c):
        if name:
            module(cv, r, c, 14, 7, "1")
            cell = cv.c(r + 2, c + 1)
            cell.value = name
            cell.font = font(7, False, P["s700"])
            from xlkit import align
            cell.alignment = align("center", "center", rot=90)
            cv.merge(r + 2, c + 1, r + 10, c + 5)
            port(cv, r + 12, c + 2, "none", None, 3)
        else:
            module(cv, r, c, 14, 7, "2", empty=True)
            cv.text(r + 6, c, "빈 베이", sz=6.5, col=P["s400"], h="center")
            cv.merge(r + 6, c, r + 6, c + 6)
        return 14, 7
    return f


def _icbay(cv, r, c):
    module(cv, r, c, 3, 29, "BAY 1 · 인터커넥트 8P")
    ports_row(cv, r + 1, c + 1, [("none", None, f"X{i}") for i in range(1, 9)], gap=1)
    return 3, 29


def _oa(cv, r, c):
    module(cv, r, c, 3, 29, "OA 1 · 관리 모듈")
    ports_row(cv, r + 1, c + 1, [("none", None, "MGMT"), ("none", None, "", 2, "USB"), ("none", None, "", 3, "CON")], gap=2)
    return 3, 29


def _linecard(cv, r, c):
    module(cv, r, c, 4, 59, None)
    title(cv, r, c, "SLOT 1 · 48P")
    port_block(cv, r + 1, c + 9, 24, 2, 101)
    cv.text(r + 3, c, "번호 101–148", sz=6, col=P["s500"])
    return 4, 59


def _embedded(cv, r, c):
    specs = [("none", None, "MGMT"), ("none", None, "SVC")] + [("none", None, f"ETH{i}") for i in range(4)] + \
            [("none", None, f"SAS{i}", 3) for i in range(2)]
    slot(cv, r, c, 34, "임베디드", specs)
    return 2, 34


def _sw_block(kind):
    def f(cv, r, c):
        if kind == "sfp":
            module(cv, r, c, 3, 48, "1–48 · SFP")
            port_block(cv, r + 1, c, 24, 2, 1)
            return 3, 48
        if kind == "rj45":
            module(cv, r, c, 3, 51, "1–48 · RJ45")
            port_block(cv, r + 1, c, 24, 2, 1, gap_every=6)
            return 3, 51
        if kind == "qsfp":
            module(cv, r, c, 3, 6, "49–54")
            port_block(cv, r + 1, c, 3, 2, 49)
            return 3, 6
        if kind == "grp8":
            module(cv, r, c, 3, 8, "0–7")
            port_block(cv, r + 1, c, 4, 2, 0, order="rowwise")
            return 3, 8
        if kind == "pp24":
            module(cv, r, c, 2, 55, "1–24 · 패치패널")
            port_block(cv, r + 1, c + 4, 24, 1, 1, order="rowwise", gap_every=6)
            return 2, 55
        if kind == "mc":
            module(cv, r, c, 3, 3, None)
            cv.text(r, c, "관리", sz=6.5, b=True)
            port(cv, r + 1, c + 1, "none", None, 2, label="M")
            port(cv, r + 2, c + 1, "none", None, 2, label="C")
            return 3, 3
    return f


def _raid_demo(cv, r, c):
    """RAID 묶음 표시 예: SFF 6베이 중 1–2 = RAID1, 3–6 = RAID5."""
    return sff_bays(cv, r, c, 6, 6, 1, bh=6, ttl=None, cap=["480G SSD"] * 2 + ["1.2T SAS"] * 4,
                    raid=[(0, 1, "RAID1 · OS"), (2, 5, "RAID5 · DATA")])


SECTIONS = [
    ("A. 포트 1개  —  포트 2칸 × 1행, 아래 1행은 포트명 (색 = 용도, 숫자 = 도면 번호)", [
        ("A01", "포트 · 서비스", "2×2", _single("svc")),
        ("A02", "포트 · 관리", "2×2", _single("mgmt", "iLO")),
        ("A03", "포트 · 백업", "2×2", _single("bak")),
        ("A04", "포트 · 스토리지", "2×2", _single("san")),
        ("A05", "포트 · 인터커넥트", "2×2", _single("ic")),
        ("A06", "포트 · 콘솔", "2×2", _single("con", "CON")),
        ("A07", "포트 · 전원", "2×2", _single("pwr", "C14")),
        ("A08", "포트 · 미연결", "2×2", _single("none")),
        ("A09", "로컬 I/O (글자)", "2×1", _single("none", None, 2, "USB")),
        ("A10", "넓은 포트 (QSFP·SAS HD·C20)", "3×2", _single("none", "Q1", 3)),
        ("A11", "넓은 포트 (D-sub·VGA)", "3×1", _single("none", None, 3, "VGA")),
        ("A12", "번호 배지", "1×1", lambda cv, r, c: badge(cv, r, c, 1)),
    ]),
    ("B. 카드·모듈 (포트 포함)  —  2행형: 제목+포트 / 포트명 · 3행형: 제목 / 포트 / 포트명", [
        ("B01", "카드 2포트", "19×2", lambda cv, r, c: slot(cv, r, c, 19, "S1 · NIC 2P", [("none", None, "P1"), ("none", None, "P2")])),
        ("B02", "카드 4포트", "19×2", lambda cv, r, c: slot(cv, r, c, 19, "S1 · NIC 4P", [("none", None, f"P{i}") for i in range(1, 5)])),
        ("B03", "카드 넓은 포트 2개", "19×2", lambda cv, r, c: slot(cv, r, c, 19, "S1 · HCA/SAS 2P", [("none", None, "P1", 3), ("none", None, "P2", 3)])),
        ("B04", "카드 8포트", "29×2", lambda cv, r, c: slot(cv, r, c, 29, "S1 · 8P", [("none", None, f"{i}") for i in range(1, 9)])),
        ("B05", "온보드 LAN 4포트", "12×2", lambda cv, r, c: slot(cv, r, c, 12, "LOM", [("none", None, str(i)) for i in range(1, 5)], gap=0)),
        ("B06", "기본 I/O (USB·VGA·관리)", "16×2", lambda cv, r, c: slot(cv, r, c, 16, "I/O", [("none", None, "", 2, "USB"), ("none", None, "", 3, "VGA"), ("none", None, "iLO")])),
        ("B07", "OCP·FlexLOM 2포트", "9×2", lambda cv, r, c: slot(cv, r, c, 9, "OCP", [("none", None, "1"), ("none", None, "2")])),
        ("B08", "띠형 카드 (1U용)", "19×1", lambda cv, r, c: strip(cv, r, c, 19, "S1 · NIC 2P", [("none", None, ""), ("none", None, "")])),
        ("B09", "관리 포트 (3행형)", "6×3", lambda cv, r, c: vmod(cv, r, c, 6, "관리", [("none", None, "MGMT")])),
        ("B10", "관리+콘솔 (3행형)", "10×3", lambda cv, r, c: vmod(cv, r, c, 10, "관리", [("none", None, "MGMT"), ("none", None, "CON")])),
    ]),
    ("C. 빈 슬롯·모듈 틀  —  점선 = 비어 있음 · 크기가 없으면 범위 선택 → 셀 스타일 '모듈 채움' + 바깥쪽 테두리", [
        ("C01", "빈 슬롯 1/3 폭", "19×2", lambda cv, r, c: slot(cv, r, c, 19, "S2  빈 슬롯", empty=True)),
        ("C02", "빈 슬롯 1/2 폭", "29×2", lambda cv, r, c: slot(cv, r, c, 29, "S2  빈 슬롯", empty=True)),
        ("C03", "모듈 틀 1/3 폭", "19×2", lambda cv, r, c: module(cv, r, c, 2, 19, "모듈")),
        ("C04", "모듈 틀 1/2 폭 3행", "29×3", lambda cv, r, c: module(cv, r, c, 3, 29, "모듈")),
        ("C05", "세로 슬롯 (타워·세로형)", "3×10", lambda cv, r, c: vslot(cv, r, c, 10, "1", 2)),
        ("C06", "빈 슬롯 전체 폭", "59×2", lambda cv, r, c: slot(cv, r, c, 59, "빈 슬롯 (전체 폭)", empty=True)),
    ]),
    ("D. 전원·냉각", [
        ("D01", "PSU (C14)", "9×2", lambda cv, r, c: psu(cv, r, c, 9, "PSU 1")),
        ("D02", "PSU (C20·넓은 인렛)", "10×2", lambda cv, r, c: psu(cv, r, c, 10, "PSU 1", inlet="C20", iw=3)),
        ("D03", "PSU 3행형", "10×3", lambda cv, r, c: psu(cv, r, c, 10, "PSU 1", h=3)),
        ("D04", "PSU 포트 없음 (섀시 전면)", "9×4", lambda cv, r, c: psu(cv, r, c, 9, "PSU 1", noport=True, h=4)),
        ("D05", "AC 입력 (C20)", "9×3", lambda cv, r, c: psu(cv, r, c, 9, "AC 1", inlet="C20", iw=3, h=3)),
        ("D06", "팬", "6×3", lambda cv, r, c: fan(cv, r, c, 6, 3, "FAN 1")),
        ("D07", "큰 팬 (섀시)", "11×4", lambda cv, r, c: fan(cv, r, c, 11, 4, "FAN 1")),
    ]),
    ("E. 디스크·베이  —  칸 안 작은 글씨 = 용량(한 칸 쓰고 복사) · 빈 베이 = '빈' · RAID = 굵은 바깥쪽 테두리 + 아래 칸에 표시", [
        ("E01", "SFF 2.5\" 8베이 (2U 박스)", "18×8", lambda cv, r, c: sff_bays(cv, r, c, 8, 4, 1, bh=6, ttl="BOX 1", cap="1.2T SAS")),
        ("E02", "SFF 10베이 (1U)", "22×3", lambda cv, r, c: (module(cv, r, c, 3, 22, None), sff_bays(cv, r, c + 1, 10, 4, 1, bh=3, pad=False, cap="1.2T"))[0]),
        ("E03", "SFF 디스크 1개", "2×6", lambda cv, r, c: sff_bays(cv, r, c, 1, 1, 1, bh=6, pad=False, cap="1.2T SAS")),
        ("E04", "LFF 3.5\" 디스크 1개", "14×2", lambda cv, r, c: lff_bay(cv, r, c, 14, 1, True, "8T NL-SAS")),
        ("E05", "LFF 12베이 (2U)", "58×8", lambda cv, r, c: lff_bays(cv, r, c, 4, 3, 6, bw=14, cap="8T NL-SAS")),
        ("E06", "RAID 묶음 표시 (예)", "14×8", _raid_demo),
    ]),
    ("F. 스위치·패치패널 포트 블록  —  칸 안 숫자 = 장비에 인쇄된 포트 번호, 연결되면 셀 스타일로 색 지정", [
        ("F01", "SFP 48포트 (홀수 위/짝수 아래)", "48×3", _sw_block("sfp")),
        ("F02", "업링크 QSFP 6포트", "6×3", _sw_block("qsfp")),
        ("F03", "관리·콘솔", "3×3", _sw_block("mc")),
        ("F04", "RJ45 48포트 (6포트 묶음)", "51×3", _sw_block("rj45")),
        ("F05", "8포트 그룹 (SAN)", "8×3", _sw_block("grp8")),
        ("F06", "패치패널 24포트", "55×2", _sw_block("pp24")),
    ]),
    ("G. 섀시 요소", [
        ("G01", "블레이드 베이", "7×14", _blade("BL460c Gen10")),
        ("G02", "빈 블레이드 베이", "7×14", _blade(None)),
        ("G03", "인터커넥트 베이 8포트", "29×3", _icbay),
        ("G04", "관리 모듈 (OA·CMC)", "29×3", _oa),
        ("G05", "컨트롤러 임베디드 포트", "34×2", _embedded),
        ("G06", "라인카드 48포트", "59×4", _linecard),
    ]),
    ("H. 장비 외곽 틀  —  B열 세로 탭 + C~BK 외곽(19인치 = 61칸). 높이: 1U 5칸 · 2U 10칸 · 3U↑ U당 4칸", [
        ("H01", "1U 틀", "62×5", _frame(5, 61)),
        ("H02", "하프 폭 1U 틀", "31×5", _frame(5, 30, "전면")),
        ("H03", "2U 틀", "62×10", _frame(10, 61, "후  면")),
    ]),
    ("I. 범례·표시·표", [
        ("I01", "용도 색 범례", "한 줄", _legend),
        ("I02", "표기 구분 범례", "한 줄", _ids),
        ("I03", "케이블 라벨 칸", "7×1", _tape),
        ("I04", "메모 줄", "한 줄", _memo),
        ("I05", "방향 표시", "21×1", _dir),
        ("I06", "연결 표 (전체 폭)", "62×3", _tbl(TABLE_FULL, 2)),
        ("I07", "연결 표 (반쪽 · 2단용)", "30×3", _tbl(TABLE_HALF, 2)),
    ]),
]




def _cap_w(it):
    code, name, size = it[0], it[1], it[2]
    import math
    return math.ceil(_tw(code + " " + name + "  " + size, 7)) + 1


def _estimate(fn):
    """부품 크기(h, w) 계산용: 임시 캔버스에 그려 본다."""
    from xlkit import new_workbook
    wb = new_workbook()
    cv = Canvas(wb, "tmp")
    return fn(cv, 5, 5)


def build_palette(wb):
    cv = Canvas(wb, "부품", TAB_PARTS)
    cv.h(1, 24)
    cv.text(1, 2, "부품 모음", sz=16, b=True)
    cv.text(1, 12, "복사해서 붙이는 셀 부품  ·  1칸 ≈ 4.2mm 정사각형", sz=10, col=P["s700"])
    cv.text(1, 63, "장비 전·후면도 라이브러리", sz=8, col=P["s500"], h="right")
    for c in range(2, 64):
        cv.add_border(1, c, bottom=side("medium", P["navy"]))
    cv.h(2, 16)
    cv.text(2, 2, "쓰는 법", sz=7.5, b=True, col=P["navy"])
    cv.text(2, 6, "① 부품 범위 선택(캡션 아래 칸만) → Ctrl+C   ② 장비 시트에서 놓을 자리의 왼쪽 위 칸 클릭 → Ctrl+V   "
                  "③ 숫자(도면 번호)·포트명 입력 → 셀 스타일로 색 지정", sz=7.5, col=P["ink"])
    cv.h(3, 5)
    r = 4
    page_start = 4
    breaks = []

    def need(rows):
        nonlocal r, page_start
        if r - page_start + rows > PAGE_ROWS:
            breaks.append(r - 1)
            page_start = r

    for head, items in SECTIONS:
        sizes = [(it, _estimate(it[3])) for it in items]
        # 구역 제목 + 첫 선반이 같은 페이지에 오도록
        first_h = 0
        x = 3
        for it, (h, w) in sizes:
            sw = max(w, _cap_w(it))
            if w > 61:
                first_h = max(first_h, h)
                break
            if x + sw - 1 > 63:
                break
            first_h = max(first_h, h)
            x += sw + 2
        need(2 + first_h + 1)
        cv.text(r, 2, head, sz=8.5, b=True, col=P["navy"])
        for c in range(2, 64):
            cv.add_border(r, c, bottom=side("thin", P["s300"]))
        r += 2
        x = 3
        shelf = []
        for it, (h, w) in sizes:
            sw = max(w, _cap_w(it))
            if x + sw - 1 > 63 and shelf:
                r = _draw_shelf(cv, r, shelf, need)
                x = 3
                shelf = []
            if w > 61:
                x = 2
            shelf.append((it, h, w, x))
            x += sw + 2
        if shelf:
            r = _draw_shelf(cv, r, shelf, need)
        r += 1
    last = r
    cv.finalize()
    cv.page_setup(63, last, title_rows="1:3", breaks=breaks)
    return cv


POSITIONS = {}


def _draw_shelf(cv, r, shelf, need):
    hmax = max(h for _, h, _, _ in shelf)
    need(1 + hmax + 1)
    for (code, name, size, fn), h, w, x in shelf:
        cell = cv.c(r, x)
        cell.value = CellRichText(TextBlock(InlineFont(rFont=FONT, sz=7, b=True, color=P["navy"]), code + " "),
                                  TextBlock(InlineFont(rFont=FONT, sz=7, color=P["ink"]), name + "  "),
                                  TextBlock(InlineFont(rFont=FONT, sz=6.5, color=P["s400"]), size))
        fn(cv, r + 1, x)
        POSITIONS[code] = (r + 1, x, h, w)
    return r + 1 + hmax + 1
