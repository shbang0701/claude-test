# -*- coding: utf-8 -*-
"""장비 시트 조립: 양식(빈칸)과 예시(작성 완료)를 같은 함수로 만든다."""
from xlkit import Canvas, P, font
from xlforms import title_block, view, table_section, finish, u_rows, C_LAST
from xlparts import note
import xldevices as D

TAB_FORM = "2F6DB5"
TAB_EX = "377F3F"


def device_sheet(wb, sheet_name, name, model, u, front, rear, n=None, info=None, rows=None, data=(),
                 nrows=None, half=False, tab=TAB_FORM, split=None, doc="장비 전·후면 구성도", foot_note=None,
                 labels=("전  면", "후  면"), back_link=False):
    """표준 장비 시트.
    front/rear: 그리기 함수 (cv, r, c, n) 또는 None
    split: None=자동, 'rear'=후면부터 새 페이지, 'table'=표부터 새 페이지"""
    n = n or {}
    cv = Canvas(wb, sheet_name, tab)
    r = title_block(cv, name, model, doc, info)
    h = rows or u_rows(u)
    breaks = []
    draw_first = r
    if front:
        r = view(cv, r, labels[0], rows=h, draw=lambda cv_, rr, cc: front(cv_, rr, cc, n))
        r += 1
    if rear:
        if split == "rear" or (split is None and front and h > 16):
            breaks.append(r - 1)
        r = view(cv, r, labels[1], rows=h, draw=lambda cv_, rr, cc: rear(cv_, rr, cc, n))
        r += 1
    draw_last = r - 1
    if foot_note:
        note(cv, r, 3, foot_note, sz=7)
        r += 1
    if split == "table" or (split is None and h > 12):
        breaks.append(r - 1)
    if nrows is None:
        nrows = max(len(data) + 2, 8)
    last = table_section(cv, r, list(data), nrows, (draw_first, draw_last), half=half)
    if back_link:
        add_back_link(cv)
    finish(cv, last, breaks=breaks)
    return cv


def add_back_link(cv):
    """인쇄 영역 밖(BM1)에 [랙] 시트로 돌아가는 링크."""
    cv.put(1, 65, '=HYPERLINK("#\'랙\'!A1","◀ 랙")', f=font(9, True, P["svc"]))


def tower_sheet(wb, sheet_name, name, model, n=None, info=None, data=(), nrows=16, tab=TAB_FORM,
                slots=None, filled=4, back_link=False):
    """타워: 전면·후면을 나란히(각 29칸 × 38행), 표는 다음 쪽."""
    n = n or {}
    cv = Canvas(wb, sheet_name, tab)
    r = title_block(cv, name, model, "장비 전·후면 구성도", info)
    h = 38
    view(cv, r, "전  면", rows=h, c=3, w=29, tab_c=2, draw=lambda cv_, rr, cc: D.tower_front(cv_, rr, cc, n, filled))
    view(cv, r, "후  면", rows=h, c=34, w=29, tab_c=33,
         draw=lambda cv_, rr, cc: D.tower_rear(cv_, rr, cc, n, {1: ("NIC 2P", 2), 3: ("GPU", 0)} if slots is None else slots))
    draw_first, draw_last = r, r + h - 1
    r += h + 1
    breaks = [r - 1]
    last = table_section(cv, r, list(data), nrows, (draw_first, draw_last))
    if back_link:
        add_back_link(cv)
    finish(cv, last, breaks=breaks)
    return cv


def free_sheet(wb, sheet_name="양식-자유형", tab=TAB_FORM):
    """생소한 장비용: 빈 틀(2U 앞·뒤) + 부품으로 채우기."""
    cv = Canvas(wb, sheet_name, tab)
    r = title_block(cv, "장비명", "제조사 모델명  (생소한 장비 — 빈 틀에 부품을 붙여 구성)", "장비 전·후면 구성도")
    top = r
    r = view(cv, r, "전  면", rows=10)
    r += 1
    r = view(cv, r, "후  면", rows=10)
    bottom = r - 1
    r += 1
    note(cv, r, 3, "틀 높이 바꾸기: 틀 안쪽 행에서 [행 삽입/삭제] (1U 5칸 · 2U 10칸 · 3U↑ U당 4칸). "
                   "틀 없이 그리려면 틀 범위를 선택해 [모두 지우기]", sz=7)
    r += 1
    last = table_section(cv, r, [], 10, (top, bottom))
    finish(cv, last)
    return cv
