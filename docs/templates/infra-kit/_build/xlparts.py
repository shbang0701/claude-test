# -*- coding: utf-8 -*-
"""셀 격자 도면 부품. 모든 함수는 (행 r, 열 c) 왼쪽 위 기준으로 그리고 (높이, 너비)를 돌려준다."""
import math
from xlkit import (P, USE, USE_NAME, USE_LIST, FONT, MONO, font, fill, align, side, Border, Side,
                   WHITE_MED, NOSIDE, add_cf, COL)
from openpyxl.worksheet.datavalidation import DataValidation

KSTYLE = {k: f"포트 {n}" for k, n in USE}
KSTYLE["none"] = "포트 미연결"

MOD_LINE = side("thin", P["s400"])
EMPTY_LINE = side("dashed", P["s300"])
FRAME_LINE = side("medium", P["s700"])


# ── 기본 요소 ───────────────────────────────────────────────
def port(cv, r, c, kind="none", num=None, w=2, label=None):
    """포트 1개 = w칸 × 1행. 번호(num)는 흰 굵은 숫자 = 도면 번호."""
    st = KSTYLE[kind]
    for cc in range(c, c + w):
        cv.put(r, cc, style=st)
        cv.cell_box(r, cc)
    cell = cv.c(r, c)
    if num is not None:
        cell.value = num
    elif label:
        cell.value = label
        cell.font = font(6, False, P["s500"] if kind == "none" else P["white"])
    cv.merge(r, c, r, c + w - 1)
    return 1, w


def pname(cv, r, c, text, w=2):
    """포트명(장비에 인쇄된 이름) = 회색 작은 글씨."""
    cv.put(r, c, text, style="포트명")
    cv.merge(r, c, r, c + w - 1)


def title(cv, r, c, text, empty=False, sz=7):
    if not text:
        return
    cv.text(r, c, text, sz=sz, b=not empty, col=P["s400"] if empty else P["ink"])


def module(cv, r, c, h, w, ttl=None, empty=False, fill_color=None):
    """모듈(카드·슬롯·PSU 등) 틀. empty=True 는 점선 빈 슬롯."""
    if empty:
        cv.outline(r, c, r + h - 1, c + w - 1, EMPTY_LINE)
    else:
        for rr in range(r, r + h):
            for cc in range(c, c + w):
                cv.put(rr, cc, style="모듈 채움")
                if fill_color:
                    cv.c(rr, cc).fill = fill(fill_color)
        cv.outline(r, c, r + h - 1, c + w - 1, MOD_LINE)
    title(cv, r, c, ttl, empty)
    return h, w


def ports_row(cv, r, c, specs, gap=1, names=True):
    """specs: [(kind, num, name[, w])] 왼→오 배치. names=True 면 다음 행에 포트명."""
    cc = c
    for s in specs:
        kind, num, name = s[0], s[1], s[2]
        w = s[3] if len(s) > 3 else 2
        lab = s[4] if len(s) > 4 else None
        port(cv, r, cc, kind, num, w, label=lab)
        if names and name:
            pname(cv, r + 1, cc, name, w)
        cc += w + gap
    return cc - gap - c  # 사용 너비


def specs_width(specs, gap=1):
    return sum((s[3] if len(s) > 3 else 2) for s in specs) + gap * (len(specs) - 1) if specs else 0


# ── 모듈 부품 ───────────────────────────────────────────────
def slot(cv, r, c, w, ttl, specs=(), empty=False, names=True, gap=1, h=2):
    """가로형 카드/슬롯 (2행): 1행 = 제목 + 포트(오른쪽 정렬), 2행 = 포트명."""
    module(cv, r, c, h, w, ttl, empty)
    if specs and not empty:
        pw = specs_width(specs, gap)
        ports_row(cv, r, c + w - 1 - pw, specs, gap, names)
    return h, w


def vmod(cv, r, c, w, ttl, specs=(), names=True, gap=1, empty=False, h=3):
    """세로형 소형 모듈 (3행): 1행 제목, 2행 포트, 3행 포트명. 포트는 가운데."""
    module(cv, r, c, h, w, ttl, empty)
    if specs:
        pw = specs_width(specs, gap)
        ports_row(cv, r + 1, c + (w - pw) // 2, specs, gap, names)
    return h, w


def strip(cv, r, c, w, ttl, specs=(), gap=1, empty=False):
    """1행 띠 모듈 (고밀도용): 제목 + 포트, 포트명 없음."""
    module(cv, r, c, 1, w, ttl, empty)
    if specs and not empty:
        pw = specs_width(specs, gap)
        ports_row(cv, r, c + w - 1 - pw, specs, gap, names=False)
    return 1, w


def psu(cv, r, c, w=8, ttl="PSU 1", kind="none", num=None, inlet="C14", iw=2, h=2, noport=False):
    """전원공급장치: 제목 + 전원 입력(인렛) 포트, 아래에 인렛 종류. h=3이면 제목 아래 줄에 포트."""
    module(cv, r, c, h, w, ttl)
    if noport:
        return h, w
    pr = r if h == 2 else r + 1
    pc = c + w - 1 - iw if h == 2 else c + (w - iw) // 2
    port(cv, pr, pc, kind, num, iw)
    if inlet and pr + 1 <= r + h - 1:
        pname(cv, pr + 1, pc, inlet, iw)
    return h, w


def fan(cv, r, c, w=5, h=3, ttl="FAN 1"):
    module(cv, r, c, h, w, ttl)
    cv.text(r + 1, c, "◎", sz=11 if h >= 3 else 8, col=P["s400"], h="center")
    cv.merge(r + 1, c, r + h - 1, c + w - 1)
    return h, w


RAID_LINE = side("thick", P["ink"])   # Excel [테두리 ▾ > 굵은 바깥쪽 테두리]와 같은 선


def _on(filled, i):
    """filled: 앞에서부터 장착된 개수(정수) 또는 장착된 순번 목록(0부터)."""
    return i < filled if isinstance(filled, int) else i in filled


def _cap(cap, i):
    if isinstance(cap, (list, tuple)):
        return cap[i] if i < len(cap) else None
    return cap


def bay_no(cv, r, c, r2, c2, text):
    """베이 번호(장비에 인쇄된 번호) — 디스크 칸 밖, 작은 회색 글씨. 문자로 넣어 도면 번호와 섞이지 않게."""
    cell = cv.c(r, c)
    cell.value = str(text)
    cell.font = font(6, False, P["s500"])
    cell.alignment = align("center", "center")
    cv.merge(r, c, r2, c2)


def disk(cv, r, c, h, w, on=True, cap=None):
    """디스크 1개(칸 안 작은 글씨 = 용량·종류). 빈 베이는 흰 칸에 '빈'."""
    st = "디스크 장착" if on else "디스크 빈칸"
    for rr in range(r, r + h):
        for x in range(c, c + w):
            cv.put(rr, x, style=st)
            cv.cell_box(rr, x)
    txt = cap if on else "빈"
    if txt:
        cv.c(r, c).value = txt
    cv.merge(r, c, r + h - 1, c + w - 1)


def raid_box(cv, r1, c1, r2, c2, label=None, label_rc=None):
    """RAID 묶음 = 디스크 범위 굵은 바깥쪽 테두리 + 옆/아래 빈 칸에 'RAID1 · OS' 같은 표시."""
    cv.outline(r1, c1, r2, c2, RAID_LINE, record=False)
    if label and label_rc:
        cv.text(label_rc[0], label_rc[1], label, sz=6.5, b=True, col=P["ink"])


def sff_bays(cv, r, c, n, filled, start=1, bh=6, ttl=None, pad=True, cap=None, raid=()):
    """2.5인치(SFF) 세로 베이 n개. 베이 = 2칸 × bh행: 위 bh-1행 = 디스크(칸 안에 용량),
    맨 아래 1행 = 베이 번호. cap: 용량 글씨(하나 또는 베이별 목록).
    raid: [(첫 순번, 끝 순번, 표시)] 순번은 0부터 — 굵은 테두리 + 아래 줄에 표시."""
    w = n * 2 + (2 if pad else 0)
    h = bh + (2 if pad else 0)
    if pad:
        module(cv, r, c, h, w, ttl)
        r0, c0 = r + 1, c + 1
    else:
        r0, c0 = r, c
    for i in range(n):
        cc = c0 + i * 2
        on = _on(filled, i)
        disk(cv, r0, cc, bh - 1, 2, on, _cap(cap, i))
        bay_no(cv, r0 + bh - 1, cc, r0 + bh - 1, cc + 1, start + i)
    for a, b, label in raid:
        raid_box(cv, r0, c0 + a * 2, r0 + bh - 1, c0 + b * 2 + 1, label, (r0 + bh, c0 + a * 2))
    return h, w


def lff_bay(cv, r, c, bw, no, on=True, cap=None):
    """3.5인치(LFF) 가로 베이 1개 = bw칸 × 2행: 왼쪽 2칸 베이 번호 + 디스크(칸 안에 용량)."""
    bay_no(cv, r, c, r + 1, c + 1, no)
    disk(cv, r, c + 2, 2, bw - 2, on, cap)
    return 2, bw


def lff_bays(cv, r, c, cols, rows, filled, start=1, bw=7, ttl=None, order="col", cap=None, raid=()):
    """3.5인치(LFF) 가로 베이: cols × rows, 각 베이 bw칸 × 2행.
    raid: [(첫 순번, 끝 순번, 표시)] — 묶음을 감싸는 굵은 테두리 + 위쪽 여백 줄에 표시."""
    w = cols * bw + 2
    h = rows * 2 + 2
    module(cv, r, c, h, w, ttl)
    pos = {}
    for j in range(cols):
        for i in range(rows):
            idx = (j * rows + i) if order == "col" else (i * cols + j)
            rr, cc = r + 1 + i * 2, c + 1 + j * bw
            lff_bay(cv, rr, cc, bw, start + idx, _on(filled, idx), _cap(cap, idx))
            pos[idx] = (rr, cc)
    for a, b, label in raid:
        cells = [pos[k] for k in range(a, b + 1) if k in pos]
        r1 = min(y for y, _ in cells)
        r2 = max(y for y, _ in cells) + 1
        c1 = min(x for _, x in cells)
        c2 = max(x for _, x in cells) + bw - 1
        # 표시는 묶음 바로 위(맨 위 줄 묶음) 또는 바로 아래(맨 아래 줄 묶음)의 여백 줄에
        at = (r, c1 + 2) if r1 == r + 1 else (r + h - 1, c1 + 2) if r2 == r + h - 2 else None
        raid_box(cv, r1, c1, r2, c2, label, at)
    return h, w


def port_block(cv, r, c, per_row, rows=2, start=1, order="oddeven", conn=None, gap_every=0, w=2):
    """스위치/패치패널 포트 블록. 칸 안 숫자 = 장비 표기 포트 번호.
    conn: {포트번호: kind} 연결된 포트만 색 표시. order: oddeven(위 홀수/아래 짝수) | rowwise."""
    conn = conn or {}
    cc = c
    width = 0
    for j in range(per_row):
        if gap_every and j and j % gap_every == 0:
            cc += 1
        for i in range(rows):
            if order == "oddeven" and rows == 2:
                n = start + j * 2 + i
            else:
                n = start + i * per_row + j
            kind = conn.get(n, "none")
            port(cv, r + i, cc, kind, n, w)
        cc += w
    return rows, cc - c


def badge(cv, r, c, num):
    cv.put(r, c, num, style="번호 배지")
    return 1, 1


def ear_frame(cv, r, c, h, w, ears=True):
    """장비 외곽(굵은 선). 좌우 1칸은 랙 귀(연회색)."""
    if ears:
        for rr in range(r, r + h):
            cv.c(rr, c).fill = fill(P["s100"])
            cv.c(rr, c + w - 1).fill = fill(P["s100"])
    cv.outline(r, c, r + h - 1, c + w - 1, FRAME_LINE)
    return h, w


def side_tab(cv, r, c, h, text):  # noqa
    """도면 왼쪽 세로 탭(전면/후면)."""
    for rr in range(r, r + h):
        cv.c(rr, c).fill = fill(P["navy"])
    cell = cv.c(r, c)
    cell.value = text
    cell.font = font(8, True, P["white"])
    cell.alignment = align("center", "center", rot=255)
    cv.merge(r, c, r + h - 1, c)


def note(cv, r, c, text, col=None, sz=7, b=False, h="left"):
    cv.text(r, c, text, sz=sz, col=col or P["s500"], b=b, h=h)


# ── 범례 ─────────────────────────────────────────────────
def _tw(s, sz=7):
    """대략적인 글자 폭(칸). 한글 1자 ≈ sz pt, 영문/숫자 ≈ 0.55 sz."""
    wpt = sum(sz if ord(ch) > 0x2E80 else sz * 0.58 for ch in s)
    return wpt / 12.0  # 1칸 ≈ 12pt


def legend_colors(cv, r, c, items=None, sz=7):
    """용도 색 범례. 칩 1칸 + 이름. 시작 열 c, 사용 너비 반환."""
    items = items or [(k, n) for k, n in USE] + [("none", "미연결")]
    cc = c
    for k, n in items:
        cv.put(r, cc, style=KSTYLE[k])
        if k == "none":
            cv.add_border(r, cc, left=side("thin", P["s400"]), right=side("thin", P["s400"]),
                          top=side("thin", P["s400"]), bottom=side("thin", P["s400"]))
        cv.text(r, cc + 1, n, sz=sz, col=P["s700"])
        cc += 1 + math.ceil(_tw(n, sz) + 0.6)
    return cc - c


def legend_ids(cv, r, c, sz=7):
    """번호·포트명·케이블 라벨 표기 구분 범례."""
    cc = c
    cv.put(r, cc, 7, style="번호 배지")
    cv.text(r, cc + 1, "도면 번호", sz=sz, col=P["s700"])
    cc += 1 + math.ceil(_tw("도면 번호", sz) + 0.8)
    cv.text(r, cc, "LOM1", sz=6.5, col=P["s500"])
    cc += 2
    cv.text(r, cc, "포트명(장비 표기)", sz=sz, col=P["s700"])
    cc += math.ceil(_tw("포트명(장비 표기)", sz) + 0.8)
    for x in range(cc, cc + 4):
        cv.put(r, x, style="케이블 라벨")
    cv.c(r, cc).value = "A-0001"
    cv.c(r, cc).font = font(7, False, P["ink"], MONO)
    cv.merge(r, cc, r, cc + 3)
    cv.text(r, cc + 4, "케이블 라벨", sz=sz, col=P["s700"])
    cc += 4 + math.ceil(_tw("케이블 라벨", sz) + 0.2)
    return cc - c


# ── 연결 표 ───────────────────────────────────────────────
TABLE_FULL = [("번호", 2), ("면", 2), ("포트명", 6), ("용도", 5), ("케이블 · 커넥터", 7),
              ("상대 장비", 8), ("상대 위치", 6), ("상대 포트", 6), ("케이블 라벨", 7), ("비고", 13)]
TABLE_HALF = [("번호", 2), ("포트명", 5), ("용도", 4), ("상대 장비", 7), ("상대 포트", 5), ("케이블 라벨", 7)]
CENTER_COLS = ("번호", "면")


def conn_table(cv, r, c, cols, data, nrows, draw_range=None, row_h=13.5, header=True, dup_ranges=None):
    """연결 정보 표. data: dict 목록(키 = 열 이름). nrows: 전체 행 수(빈 행 포함).
    draw_range: 도면 영역(예 'C4:BK30') — 번호 확인(분홍) 조건부서식에 사용."""
    ws = cv.ws
    rr = r
    if header:
        cc = c
        for name, w in cols:
            for x in range(cc, cc + w):
                cv.put(rr, x, style="표 머리글")
            cv.c(rr, cc).value = name
            if name in CENTER_COLS:
                cv.c(rr, cc).alignment = align("center")
            cv.merge(rr, cc, rr, cc + w - 1)
            cc += w
        cv.h(rr, 15)
        rr += 1
    first = rr
    colpos = {}
    cc = c
    for name, w in cols:
        colpos[name] = (cc, w)
        cc += w
    for i in range(nrows):
        d = data[i] if i < len(data) else {}
        cv.h(rr, row_h)
        for name, w in cols:
            x0, _ = colpos[name]
            st = {"번호": "표 번호", "케이블 라벨": "케이블 라벨"}.get(name, "표 본문")
            for x in range(x0, x0 + w):
                cv.put(rr, x, style=st)
            v = d.get(name)
            if v is not None and v != "":
                cv.c(rr, x0).value = v
            if name in CENTER_COLS:
                cv.c(rr, x0).alignment = align("center")
            cv.merge(rr, x0, rr, x0 + w - 1)
        rr += 1
    last = rr - 1
    # 드롭다운(데이터 유효성 검사)
    if "용도" in colpos:
        x0, w = colpos["용도"]
        dv = DataValidation(type="list", formula1='"' + ",".join(USE_LIST) + '"', allow_blank=True,
                            showErrorMessage=False)
        dv.add(f"{COL(x0)}{first}:{COL(x0)}{last}")
        ws.add_data_validation(dv)
    if "면" in colpos:
        x0, w = colpos["면"]
        dv = DataValidation(type="list", formula1='"전면,후면"', allow_blank=True, showErrorMessage=False)
        dv.add(f"{COL(x0)}{first}:{COL(x0)}{last}")
        ws.add_data_validation(dv)
    # 조건부 서식: 번호 칸 = 용도 색, 번호 확인(분홍)
    nx, nw = colpos["번호"]
    ncol = COL(nx)
    rng = f"{ncol}{first}:{COL(nx + nw - 1)}{last}"
    ref = f"${ncol}{first}"
    if draw_range:
        add_cf(ws, rng, f'AND(ISNUMBER({ref}),SUMPRODUCT(({draw_range}={ref})*({draw_range}<>""))<>1)',
               fill_color=P["warn_bg"], font_color=P["warn_fg"], stop=True)
    dups = dup_ranges or [f"${ncol}${first}:${ncol}${last}"]
    cnt = "+".join(f"COUNTIF({d},{ref})" for d in dups)
    add_cf(ws, rng, f"AND(ISNUMBER({ref}),{cnt}>1)",
           fill_color=P["warn_bg"], font_color=P["warn_fg"], stop=True)
    if "용도" in colpos:
        ux = COL(colpos["용도"][0])
        for k, n in USE:
            add_cf(ws, rng, f'${ux}{first}="{n}"', fill_color=P[k], font_color=P["white"], stop=True)
    return first, last, colpos
