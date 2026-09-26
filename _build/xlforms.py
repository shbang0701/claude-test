# -*- coding: utf-8 -*-
"""장비 시트(A4 가로) 구성: 제목 블록 · 전면/후면 도면 · 연결 표 · 인쇄 설정."""
from xlkit import P, font, align, side, Border, COL, fill
from xlparts import (side_tab, ear_frame, legend_colors, legend_ids, conn_table, TABLE_FULL,
                     TABLE_HALF, note)

C_TAB = 2        # B열: 세로 탭
C_FRAME = 3      # C열부터 장비 외곽
FRAME_W = 61     # 19인치 장비 = 61칸 (C~BK)
C_LAST = 63      # BK = 인쇄 영역 마지막 열
PAGE_ROWS = 40   # 제목 블록 아래 한 페이지에 들어가는 12pt 행 수(여유 포함)


def u_rows(u):
    """장비 높이 기준: 1~2U = U당 5칸, 3U 이상 = U당 4칸, 한 면 최대 38칸."""
    if u <= 2:
        return u * 5
    return min(u * 4, 38)


def title_block(cv, name, model, doc="장비 전·후면 구성도", info=None):
    cv.h(1, 24)
    cv.text(1, 2, name, sz=16, b=True, col=P["ink"])
    cv.text(1, 15, model, sz=11, col=P["s700"])
    cv.text(1, C_LAST, doc, sz=8, col=P["s500"], h="right")
    for c in range(2, C_LAST + 1):
        cv.add_border(1, c, bottom=side("medium", P["navy"]))
    cv.h(2, 16)
    info = info or {}
    spec = [("설치 위치", 4, 8), ("랙 · U", 3, 7), ("구분", 2, 5), ("시리얼", 3, 7), ("관리 IP", 3, 6),
            ("작성", 2, 7)]
    c = 2
    for lab, lw, vw in spec:
        cv.text(2, c, lab, sz=7.5, col=P["s500"])
        cv.text(2, c + lw, info.get(lab, ""), sz=8.5, b=True, col=P["ink"])
        c += lw + vw + 1
    cv.h(3, 5)
    return 4


def view(cv, r, label, u=None, rows=None, draw=None, c=C_FRAME, w=FRAME_W, tab_c=C_TAB, ears=True):
    """세로 탭 + 외곽 틀 + 내부 그리기. draw(cv, r, c) 는 틀 좌상단 기준."""
    h = rows or u_rows(u)
    side_tab(cv, r, tab_c, h, label)
    ear_frame(cv, r, c, h, w, ears)
    if draw:
        draw(cv, r, c)
    return r + h


def table_section(cv, r, data, nrows, draw_rows, half=False, legend=True):
    """연결 정보 제목 행 + 범례 + 표. draw_rows=(첫 행, 끝 행) 도면 영역."""
    cv.text(r, 2, "연결 정보", sz=9, b=True, col=P["navy"])
    if legend:
        used = legend_colors(cv, r, 8)
        legend_ids(cv, r, 8 + used + 1)
    r += 1
    rng = f"$C${draw_rows[0]}:$BK${draw_rows[1]}"
    if not half:
        first, last, _ = conn_table(cv, r, 2, TABLE_FULL, data, nrows, draw_range=rng)
        return last + 1
    n = (nrows + 1) // 2
    left, right = data[:n], data[n:]
    dups = [f"$B${r + 1}:$B${r + n}", f"$AH${r + 1}:$AH${r + n}"]
    f1, l1, _ = conn_table(cv, r, 2, TABLE_HALF, left, n, draw_range=rng, dup_ranges=dups)
    f2, l2, _ = conn_table(cv, r, 34, TABLE_HALF, right, n, draw_range=rng, dup_ranges=dups)
    return max(l1, l2) + 1


def finish(cv, last_row, breaks=(), doc_footer=True):
    cv.finalize()
    cv.page_setup(C_LAST, last_row, title_rows="1:3", breaks=breaks)
