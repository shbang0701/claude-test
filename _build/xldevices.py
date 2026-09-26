# -*- coding: utf-8 -*-
"""장비 전면/후면 그리기. n = {포트키: (용도키, 도면번호)} — 빈 양식은 n={} (모두 미연결)."""
from xlkit import P, font, align, fill
from xlparts import (port, pname, module, slot, vmod, strip, psu, fan, sff_bays, lff_bays, port_block,
                     ports_row, specs_width, note, title)


def g(n, key):
    v = n.get(key)
    if v is None:
        return ("none", None)
    if isinstance(v, str):
        return (v, None)
    return v


def sp(n, key, name, w=2, label=None):
    k, num = g(n, key)
    return (k, num, name, w, label)


# ═══ 1U 랙 서버 ═══════════════════════════════════════════════
def _raid_in(raid, first, count):
    """베이 번호 기준 RAID 묶음 → 이 묶음(첫 번호 first, count개) 안의 순번(0부터)."""
    out = []
    for a, b, lab in raid or ():
        if a <= first + count - 1 and b >= first:
            out.append((max(a, first) - first, min(b, first + count - 1) - first, lab))
    return out


def _caps(cap, first, count):
    if isinstance(cap, dict):
        return [cap.get(first + i) for i in range(count)]
    return cap


def srv1u_front(cv, r, c, n, bays=10, filled=4, ttl="SFF 1–10", cap=None, raid=()):
    """cap: 용량 글씨(하나 또는 {베이 번호: 글씨}). raid: [(첫 베이, 끝 베이, 표시)] — 표시는 틀 아래 줄."""
    module(cv, r + 1, c + 1, 3, 5, "UID")
    cv.text(r + 2, c + 1, "● 상태", sz=6.5, col=P["s500"])
    cv.text(r + 3, c + 1, "○ 전원", sz=6.5, col=P["s500"])
    module(cv, r + 1, c + 7, 3, bays * 2 + 2, None)
    sff_bays(cv, r + 1, c + 8, bays, filled, 1, bh=3, pad=False, cap=_caps(cap, 1, bays), raid=_raid_in(raid, 1, bays))
    module(cv, r + 1, c + 51, 3, 9, "전면 I/O")
    ports_row(cv, r + 2, c + 52, [sp(n, "f_usb", "", 2, "USB"), sp(n, "f_svc", "", 2, "SVC"), sp(n, "f_vga", "", 3, "VGA")],
              gap=0, names=False)
    return 5


def srv1u_rear(cv, r, c, n, slots=None):
    slots = slots or {1: ("NIC 2P", 2), 2: None, 3: None}
    xs = [c + 1, c + 21, c + 41]
    for i, x in enumerate(xs):
        s = slots.get(i + 1)
        if s:
            nm, cnt = s[0], s[1]
            specs = [sp(n, f"s{i+1}p{j+1}", "") for j in range(cnt)]
            strip(cv, r + 1, x, 19, f"S{i+1} · {nm}", specs)
        else:
            strip(cv, r + 1, x, 19, f"S{i+1}  빈 슬롯", empty=True)
    y = r + 2
    slot(cv, y, c + 1, 9, "OCP", [sp(n, "ocp1", "1"), sp(n, "ocp2", "2")])
    slot(cv, y, c + 11, 16, "I/O", [sp(n, "usb", "", 2, "USB"), sp(n, "vga", "", 3, "VGA"), sp(n, "bmc", "iDRAC")])
    slot(cv, y, c + 28, 12, "LOM", [sp(n, f"lom{i}", str(i)) for i in range(1, 5)], gap=0)
    psu(cv, y, c + 41, 9, "PSU 1", *g(n, "psu1"))
    psu(cv, y, c + 51, 9, "PSU 2", *g(n, "psu2"))
    return 5


# ═══ 2U 랙 서버 ═══════════════════════════════════════════════
def srv2u_front(cv, r, c, n, boxes=(8, 8, 0), lff=False, filled_lff=0, cap=None, raid=()):
    """boxes: 박스별 장착 수. cap: 용량 글씨(하나 또는 {베이 번호: 글씨}). raid: [(첫 베이, 끝 베이, 표시)]."""
    if lff:
        lff_bays(cv, r + 1, c + 1, 4, 3, filled_lff, bw=14, ttl=None, cap=_caps(cap, 1, 12), raid=_raid_in(raid, 1, 12))
        return 10
    x = c + 1
    for bi, filled in enumerate(boxes):
        first = bi * 8 + 1
        sff_bays(cv, r + 1, x, 8, filled, start=first, bh=6, ttl=f"BOX {bi + 1}", cap=_caps(cap, first, 8),
                 raid=_raid_in(raid, first, 8))
        x += 19
    module(cv, r + 1, c + 58, 8, 2, None)
    cv.text(r + 1, c + 58, "I/O", sz=6.5, b=True)
    port(cv, r + 3, c + 58, *g(n, "f_usb"), label="USB")
    port(cv, r + 5, c + 58, *g(n, "f_svc"), label="SVC")
    return 10


def srv2u_rear(cv, r, c, n, slots=None, lom="LOM 1GbE", ocp="FlexLOM", bmc="iLO"):
    """slots: {번호: (이름, 포트수[, 넓은포트])} — 없는 번호는 빈 슬롯."""
    slots = slots or {}
    grid = [(1, 0, 0), (2, 0, 1), (3, 0, 2), (4, 1, 0), (5, 1, 1), (6, 1, 2), (7, 2, 0), (8, 2, 1)]
    for no, col, row in grid:
        x, y = c + 1 + col * 20, r + 1 + row * 2
        s = slots.get(no)
        if s:
            wport = 3 if (len(s) > 2 and s[2]) else 2
            specs = [sp(n, f"s{no}p{j+1}", f"P{j+1}", wport) for j in range(s[1])]
            slot(cv, y, x, 19, f"S{no} · {s[0]}", specs)
        else:
            slot(cv, y, x, 19, f"S{no}  빈 슬롯", empty=True)
    note(cv, r + 5, c + 41, "라이저 3 : S7–S8", sz=6.5, col=P["s400"])
    y = r + 7
    slot(cv, y, c + 1, 9, ocp, [sp(n, "ocp1", "1"), sp(n, "ocp2", "2")])
    slot(cv, y, c + 11, 16, "I/O", [sp(n, "usb", "", 2, "USB"), sp(n, "vga", "", 3, "VGA"), sp(n, "bmc", bmc)])
    slot(cv, y, c + 28, 12, lom[:3], [sp(n, f"lom{i}", str(i)) for i in range(1, 5)], gap=0)
    cv.text(y, c + 28, lom, sz=7, b=True)
    psu(cv, y, c + 41, 9, "PSU 2", *g(n, "psu2"))
    psu(cv, y, c + 51, 9, "PSU 1", *g(n, "psu1"))
    return 10


# ═══ 4U 랙 서버 (GPU/스토리지형) ═════════════════════════════════
def srv4u_front(cv, r, c, n, filled=12, cap=None, raid=()):
    lff_bays(cv, r + 1, c + 1, 4, 6, filled, bw=14, ttl=None, cap=_caps(cap, 1, 24), raid=_raid_in(raid, 1, 24))
    return 16


def srv4u_rear(cv, r, c, n, slots=None):
    slots = slots or {}
    for i in range(8):
        col, row = i % 2, i // 2
        x, y = c + 1 + col * 30, r + 1 + row * 2
        s = slots.get(i + 1)
        w = 29
        if s:
            wport = 3 if (len(s) > 2 and s[2]) else 2
            specs = [sp(n, f"s{i+1}p{j+1}", f"P{j+1}", wport) for j in range(s[1])]
            slot(cv, y, x, w, f"S{i+1} · {s[0]}", specs)
        else:
            slot(cv, y, x, w, f"S{i+1}  빈 슬롯", empty=True)
    y = r + 10
    slot(cv, y, c + 1, 9, "OCP", [sp(n, "ocp1", "1"), sp(n, "ocp2", "2")])
    slot(cv, y, c + 11, 16, "I/O", [sp(n, "usb", "", 2, "USB"), sp(n, "vga", "", 3, "VGA"), sp(n, "bmc", "BMC")])
    slot(cv, y, c + 28, 12, "LOM", [sp(n, f"lom{i}", str(i)) for i in range(1, 5)], gap=0)
    y = r + 12
    for i in range(4):
        psu(cv, y, c + 1 + i * 15, 14, f"PSU {i+1}", *g(n, f"psu{i+1}"), inlet="C20", iw=3)
    return 16


# ═══ 1U 스위치 (48 + 업링크) ═══════════════════════════════════
def sw48_front(cv, r, c, n, media="SFP28 25G", up=6, up_media="QSFP28 100G", conn=None, mg=True):
    conn = conn or {}
    module(cv, r + 1, c + 1, 3, 48, f"1–48 · {media}")
    port_block(cv, r + 2, c + 1, 24, 2, 1, conn=conn)
    ux = c + 50
    module(cv, r + 1, ux, 3, up, f"49–{48 + up}")
    port_block(cv, r + 2, ux, up // 2, 2, 49, conn=conn)
    if mg:
        mx = c + 57
        module(cv, r + 1, mx, 3, 3, None)
        cv.text(r + 1, mx, "관리", sz=6.5, b=True)
        port(cv, r + 2, mx + 1, *g(n, "mgmt"), label=None if g(n, "mgmt")[1] else "M")
        port(cv, r + 3, mx + 1, *g(n, "con"), label=None if g(n, "con")[1] else "C")
    return 5


def sw_rear(cv, r, c, n, fans=4, psu_inlet="C14"):
    x = c + 1
    for i in range(fans):
        fan(cv, r + 1, x, 6, 3, f"FAN {i+1}")
        x += 7
    psu(cv, r + 1, c + 39, 10, "PSU 1", *g(n, "psu1"), inlet=psu_inlet, h=3)
    psu(cv, r + 1, c + 50, 10, "PSU 2", *g(n, "psu2"), inlet=psu_inlet, h=3)
    return 5


def rj48_front(cv, r, c, n, conn=None):
    """48 × RJ45 1G + 4 × SFP 업링크 (관리 스위치 등)."""
    conn = conn or {}
    module(cv, r + 1, c + 1, 3, 51, "1–48 · RJ45 1G")
    # 12포트(6열)마다 1칸 간격
    port_block(cv, r + 2, c + 1, 24, 2, 1, conn=conn, gap_every=6)
    ux = c + 53
    module(cv, r + 1, ux, 3, 4, "49–52")
    port_block(cv, r + 2, ux, 2, 2, 49, conn=conn)
    module(cv, r + 1, c + 58, 3, 2, None)
    port(cv, r + 2, c + 58, *g(n, "mgmt"), label=None if g(n, "mgmt")[1] else "M")
    port(cv, r + 3, c + 58, *g(n, "con"), label=None if g(n, "con")[1] else "C")
    return 5


# ═══ SAN 스위치 (8포트 그룹) ════════════════════════════════════
def san48_front(cv, r, c, n, conn=None, groups=6, qsfp=4):
    conn = conn or {}
    x = c + 1
    for gi in range(groups):
        module(cv, r + 1, x, 3, 8, f"{gi*8}–{gi*8+7}")
        port_block(cv, r + 2, x, 4, 2, gi * 8, order="rowwise", conn=conn)
        x += 9
    module(cv, r + 1, x, 3, qsfp, "QSFP")
    port_block(cv, r + 2, x, qsfp // 2, 2, 48, order="rowwise", conn=conn)
    return 5


def san_rear(cv, r, c, n):
    vmod(cv, r + 1, c + 1, 10, "관리", [sp(n, "mgmt", "MGMT"), sp(n, "con", "CON")])
    module(cv, r + 1, c + 12, 3, 6, "USB")
    for i in range(3):
        fan(cv, r + 1, c + 19 + i * 7, 6, 3, f"FAN {i+1}")
    psu(cv, r + 1, c + 40, 9, "PSU 1", *g(n, "psu1"), h=3)
    psu(cv, r + 1, c + 50, 10, "PSU 2", *g(n, "psu2"), h=3)
    return 5


# ═══ 스토리지 (2U 컨트롤러 인클로저) ════════════════════════════
def stg_front(cv, r, c, n, filled=12, bays=25, cap=None, raid=()):
    sff_bays(cv, r + 1, c + 1, bays, filled, 0, bh=6, ttl=f"드라이브 0–{bays-1}", cap=_caps(cap, 0, bays),
             raid=_raid_in(raid, 0, bays))
    module(cv, r + 1, c + 54, 8, 6, "상태")
    cv.text(r + 3, c + 54, "● 전원", sz=6.5, col=P["s500"])
    cv.text(r + 4, c + 54, "● 장애", sz=6.5, col=P["s500"])
    return 10


def stg_rear(cv, r, c, n):
    for k, sp_, y in (("b", "SP B", r + 1), ("a", "SP A", r + 5)):
        module(cv, y, c + 1, 4, 4, sp_)
        emb = [sp(n, f"{k}_mgmt", "MGMT"), sp(n, f"{k}_svc", "SVC")] + \
              [sp(n, f"{k}_eth{i}", f"ETH{i}") for i in range(4)] + \
              [sp(n, f"{k}_sas{i}", f"SAS{i}", 3) for i in range(2)]
        slot(cv, y, c + 6, 34, "임베디드", emb)
        slot(cv, y, c + 41, 19, "I/O 0 · FC 4P", [sp(n, f"{k}_io0p{i}", f"P{i}") for i in range(4)], gap=0)
        slot(cv, y + 2, c + 6, 17, "I/O 1  빈 슬롯", empty=True)
        slot(cv, y + 2, c + 24, 16, "USB", [sp(n, f"{k}_usb", "", 2, "USB")])
        psu(cv, y + 2, c + 41, 19, f"PSU {sp_[-1]}", *g(n, f"{k}_psu"))
    return 10


# ═══ 블레이드 섀시 (10U) ═══════════════════════════════════════
def blade_front(cv, r, c, n, bays=None):
    """bays: {번호: 모델명} — 없는 번호는 빈 베이. 반높이 16베이(위 1–8, 아래 9–16)."""
    bays = bays or {}
    for i in range(16):
        row, col = i // 8, i % 8
        x, y = c + 2 + col * 7, r + 1 + row * 14
        name = bays.get(i + 1)
        if name:
            module(cv, y, x, 14, 7, f"{i+1}")
            cell = cv.c(y + 2, x + 1)
            cell.value = name
            cell.font = font(7, False, P["s700"])
            cell.alignment = align("center", "center", rot=90)
            cv.merge(y + 2, x + 1, y + 10, x + 5)
            k, num = g(n, f"bay{i+1}")
            if num is not None:
                port(cv, y + 12, x + 2, k, num, 3)
        else:
            module(cv, y, x, 14, 7, f"{i+1}", empty=True)
            cv.text(y + 6, x, "빈 베이", sz=6.5, col=P["s400"], h="center")
            cv.merge(y + 6, x, y + 6, x + 6)
    y = r + 30
    for i in range(6):
        psu(cv, y, c + 2 + i * 9 + (1 if i >= 3 else 0), 9, f"PSU {i+1}", noport=True, h=4)
        cv.text(y + 2, c + 2 + i * 9 + (1 if i >= 3 else 0), "2650W", sz=6.5, col=P["s500"])
    module(cv, r + 29, c + 28, 1, 5, None)
    cv.text(r + 29, c + 28, "디스플레이", sz=6, col=P["s500"])
    return 38


def blade_rear(cv, r, c, n, ic=None):
    """ic: {베이: (이름, [(포트키, 포트명)...])}"""
    ic = ic or {}
    for i in range(5):
        fan(cv, r + 1, c + 1 + i * 12, 11, 4, f"FAN {i+1}")
    for b in range(8):
        col, row = b % 2, b // 2
        x, y = c + 1 + col * 30, r + 6 + row * 3
        spec = ic.get(b + 1)
        if spec:
            name, plist = spec
            specs = [sp(n, key, pn) for key, pn in plist]
            module(cv, y, x, 3, 29, f"BAY {b+1} · {name}")
            ports_row(cv, y + 1, x + 1, specs, gap=1)
        else:
            module(cv, y, x, 3, 29, f"BAY {b+1}  빈 베이", empty=True)
    y = r + 19
    for k, (nm, x) in enumerate((("OA 1", c + 1), ("OA 2", c + 31))):
        module(cv, y, x, 3, 29, nm)
        ports_row(cv, y + 1, x + 1, [sp(n, f"oa{k+1}", "MGMT"), sp(n, f"oa{k+1}_usb", "", 2, "USB"),
                                      sp(n, f"oa{k+1}_con", "", 3, "CON")], gap=2)
    for i in range(5):
        fan(cv, r + 23, c + 1 + i * 12, 11, 4, f"FAN {i+6}")
    y = r + 28
    for i in range(6):
        x = c + 1 + i * 10
        psu(cv, y, x, 9, f"AC {6 - i}", *g(n, f"ac{6-i}"), inlet="C20", iw=3, h=3)
    note(cv, r + 32, c + 1, "※ 섀시 뒤에서 보면 AC 입력 번호가 오른쪽→왼쪽 순서", sz=6.5, col=P["s500"])
    return 38


# ═══ 모듈형 섀시 스위치 (라인카드) ═══════════════════════════════
def chassis_front(cv, r, c, n, cards=None, conn=None):
    """cards: {슬롯: 이름} (라인카드 48포트). 슬롯 5·6 = 슈퍼바이저."""
    cards = cards or {}
    conn = conn or {}
    y = r + 1
    for s in (1, 2, 3, 4):
        spec = cards.get(s)
        if spec:
            name, nport = spec if isinstance(spec, tuple) else (spec, 48)
            module(cv, y, c + 1, 4, 59, None)
            title(cv, y, c + 1, f"SLOT {s} · {name}")
            port_block(cv, y + 1, c + 10, nport // 2, 2, s * 100 + 1, conn=conn)
            cv.text(y + 3, c + 1, f"번호 {s}01–{s}{nport:02d}", sz=6, col=P["s500"])
        else:
            module(cv, y, c + 1, 4, 59, f"SLOT {s}  빈 슬롯", empty=True)
        y += 5
    for s, nm in ((5, "SUP A"), (6, "SUP B")):
        x = c + 1 if s == 5 else c + 31
        module(cv, y, x, 3, 29, f"SLOT {s} · {nm}")
        ports_row(cv, y + 1, x + 1, [sp(n, f"sup{s}_mgmt", "MGMT"), sp(n, f"sup{s}_con", "CON"),
                                      sp(n, f"sup{s}_usb", "", 2, "USB")], gap=2)
    y += 4
    module(cv, y, c + 1, 3, 59, "팬 트레이 (전면)")
    return 32


def chassis_rear(cv, r, c, n):
    for i in range(3):
        fan(cv, r + 1, c + 1 + i * 20, 19, 12, f"FAN TRAY {i+1}")
    y = r + 14
    for i in range(4):
        psu(cv, y, c + 1 + i * 15, 14, f"PSU {i+1}", *g(n, f"psu{i+1}"), inlet="C20", iw=3, h=3)
    module(cv, r + 18, c + 1, 12, 59, "패브릭 모듈 영역 (후면)", empty=True)
    return 32


# ═══ 1U 소형 장비 (패치패널·콘솔서버 등) ═══════════════════════
def pp24_front(cv, r, c, n, conn=None, ports=24, ttl="1–24 · RJ45 Cat6"):
    conn = conn or {}
    title(cv, r + 1, c + 1, ttl)
    port_block(cv, r + 2, c + 6, ports, 1, 1, order="rowwise", conn=conn, gap_every=6)
    return 5


# ═══ PDU (0U 세로형 → 가로로 눕혀 표시) ═══════════════════════
def pdu_face(cv, r, c, n, conn=None):
    """콘센트 칸 안 숫자 = PDU에 인쇄된 콘센트 번호. 사용 중인 콘센트만 색 표시."""
    conn = conn or {}
    module(cv, r + 1, c + 1, 4, 7, "입력")
    k, num = g(n, "input")
    port(cv, r + 2, c + 3, k, num, w=3, label=None if num is not None else "IN")
    pname(cv, r + 3, c + 3, "IEC309", 3)
    banks = [("BANK 1 · C13", 1, 12, 2, c + 9, 14), ("BANK 2 · C13", 13, 12, 2, c + 24, 14),
             ("BANK 3 · C19", 25, 6, 3, c + 39, 11)]
    for ttl, first, cnt, w, x, mw in banks:
        module(cv, r + 1, x, 4, mw, ttl)
        per = cnt // 2
        for j in range(cnt):
            no = first + j
            rr = r + 2 + (j // per)
            cc = x + 1 + (j % per) * w
            k = conn.get(no, "none")
            port(cv, rr, cc, k, no, w)
    module(cv, r + 1, c + 51, 4, 9, "표시부")
    cv.text(r + 2, c + 51, " 전류  A", sz=6.5, col=P["s500"])
    cv.text(r + 3, c + 51, " 차단기 3", sz=6.5, col=P["s500"])
    return 6


# ═══ 타워 서버 (앞·뒤 나란히, 각 29칸 × 38행) ═══════════════════════
def tower_front(cv, r, c, n, filled=4, cap=None, raid=()):
    module(cv, r + 1, c + 2, 2, 25, "5.25\" 베이 1 · 광학 드라이브")
    module(cv, r + 4, c + 2, 2, 25, "5.25\" 베이 2  빈 베이", empty=True)
    vmod(cv, r + 7, c + 2, 25, "전면 I/O", [sp(n, "f_usb1", "", 2, "USB"), sp(n, "f_usb2", "", 2, "USB"),
                                           sp(n, "f_svc", "", 2, "SVC")], names=False)
    lff_bays(cv, r + 11, c + 2, 1, 8, filled, bw=23, ttl=None, cap=_caps(cap, 1, 8), raid=_raid_in(raid, 1, 8))
    module(cv, r + 30, c + 2, 6, 25, "전면 베젤 · 흡기구")
    cv.text(r + 32, c + 2, "  ○ 전원 버튼   ● 상태 LED", sz=7, col=P["s500"])
    return 38


def tower_rear(cv, r, c, n, slots=None):
    slots = {} if slots is None else slots
    module(cv, r + 1, c + 2, 21, 9, "I/O")
    items = [("bmc", "iLO"), ("lom1", "LOM1"), ("lom2", "LOM2"), ("lom3", "LOM3"), ("lom4", "LOM4"),
             ("usb", "USB"), ("vga", "VGA"), ("ser", "시리얼")]
    y = r + 3
    for key, nm in items:
        k, num = g(n, key)
        lab = None if (num is not None or key in ("bmc",) or key.startswith("lom")) else nm
        port(cv, y, c + 3, k, num, 2, label=lab if key in ("usb", "vga", "ser") else None)
        cv.text(y, c + 6, nm, sz=6.5, col=P["s500"])
        y += 2
    for i in range(5):
        x = c + 12 + i * 3
        s = slots.get(i + 1)
        if s:
            module(cv, r + 1, x, 16, 3, f"{i+1}")
            for j in range(s[1]):
                port(cv, r + 3 + j * 2, x + 1, *g(n, f"s{i+1}p{j+1}"))
            cv.text(r + 14, x, s[0], sz=6, col=P["s500"])
        else:
            module(cv, r + 1, x, 16, 3, f"{i+1}", empty=True)
    fan(cv, r + 18, c + 12, 15, 5, "후면 팬")
    psu(cv, r + 30, c + 2, 12, "PSU 1", *g(n, "psu1"), h=3)
    psu(cv, r + 30, c + 15, 12, "PSU 2", *g(n, "psu2"), h=3)
    note(cv, r + 25, c + 2, "PCIe 슬롯은 세로형 (번호 = 슬롯)", sz=6.5, col=P["s500"])
    return 38


# ═══ 5U GPU 서버 (예: HPE Cray XD670) — 공개 자료 요약 기준 초안 ═══════════════
def gpu5u_front(cv, r, c, n, filled=0, cap=None, slots=None):
    """전면: NVMe 8 · 전면 I/O(관리 MLAN 포함) · PCIe x16 4 · GPU 트레이. slots: {번호: (이름, 포트수[, 넓은포트])}."""
    slots = slots or {}
    sff_bays(cv, r + 1, c + 1, 8, filled, 0, bh=6, ttl="NVMe U.2 · 0–7", cap=_caps(cap, 0, 8))
    module(cv, r + 1, c + 20, 3, 40, "전면 I/O · 관리")
    ports_row(cv, r + 2, c + 21, [sp(n, "mlan", "MLAN"), sp(n, "lan1", "10G-1"), sp(n, "lan2", "10G-2"),
                                  sp(n, "usb", "", 2, "USB"), sp(n, "usb2", "", 2, "USB"), sp(n, "vga", "", 3, "VGA")])
    cv.text(r + 2, c + 50, "○ 전원  ● ID", sz=6.5, col=P["s500"])
    for i in range(4):
        x, y = c + 20 + (i % 2) * 20, r + 5 + (i // 2) * 2
        s = slots.get(i + 1)
        if s:
            wport = 3 if (len(s) > 2 and s[2]) else 2
            slot(cv, y, x, 19, f"S{i+1} · {s[0]}", [sp(n, f"s{i+1}p{j+1}", f"P{j+1}", wport) for j in range(s[1])])
        else:
            slot(cv, y, x, 19, f"S{i+1}  빈 슬롯", empty=True)
    module(cv, r + 10, c + 1, 9, 59, "GPU 트레이 · 8-GPU 보드 (흡기면)")
    cv.text(r + 14, c + 1, "연결 포트 없음", sz=7, col=P["s400"], h="center")
    cv.merge(r + 14, c + 1, r + 14, c + 59)
    return 20


def gpu5u_rear(cv, r, c, n, gnet=None):
    """후면: PSU 6 · 팬 6 · GPU 네트워크 슬레드 2 × 슬롯 4. gnet: {슬롯: 이름} (포트 1개, 넓은 포트)."""
    gnet = gnet or {}
    for i in range(6):
        psu(cv, r + 1, c + 1 + i * 10, 9, f"PSU {i+1}", *g(n, f"psu{i+1}"), inlet="C20", iw=3, h=3)
    for i in range(6):
        fan(cv, r + 5, c + 1 + i * 10, 9, 5, f"FAN {i+1}")
    for k, x in enumerate((c + 1, c + 31)):
        module(cv, r + 11, x, 5, 29, f"슬레드 {k+1} · GPU 네트워크 S{k*4+1}–S{k*4+4}")
        for j in range(4):
            no = k * 4 + j + 1
            nm = gnet.get(no)
            vmod(cv, r + 12, x + 1 + j * 7, 6, f"S{no}" + (f" {nm}" if nm else ""), [sp(n, f"g{no}", "P1", 3)])
    return 20


# ═══ 대형 블레이드 섀시 18U (예: HPE Superdome 2) — 구성 요소 수 = 공개 자료 요약, 배치는 확인 ═══
def sd2_front(cv, r, c, n, blades=None):
    """전면: 서버 블레이드 8 (세로) · 전원 12 (전원 입력은 후면)."""
    blades = blades or {}
    for i in range(8):
        x, y = c + 2 + i * 7, r + 1
        name = blades.get(i + 1)
        if name:
            module(cv, y, x, 26, 7, f"{i+1}")
            cell = cv.c(y + 2, x + 1)
            cell.value = name
            cell.font = font(7, False, P["s700"])
            cell.alignment = align("center", "center", rot=90)
            cv.merge(y + 2, x + 1, y + 23, x + 5)
        else:
            module(cv, y, x, 26, 7, f"{i+1}", empty=True)
            cv.text(y + 12, x, "빈 베이", sz=6.5, col=P["s400"], h="center")
            cv.merge(y + 12, x, y + 12, x + 6)
    for row, y in enumerate((r + 28, r + 32)):
        for i in range(6):
            psu(cv, y, c + 2 + i * 9 + (1 if i >= 3 else 0), 9, f"PSU {row * 6 + i + 1}", noport=True, h=3)
    note(cv, r + 36, c + 2, "서버 블레이드 8 · 전원 12 (전원 입력은 후면) — 수량은 공개 자료 기준, 배치는 현장 확인", sz=6.5)
    return 38


def sd2_rear(cv, r, c, n, ic=None):
    """후면: 팬 15 · 인터커넥트 베이 · XFM 4 · OA 2 · GPSM 2 · AC 입력. ic: {베이: (이름, [(키, 포트명)...])}"""
    ic = ic or {}
    for i in range(5):
        fan(cv, r + 1, c + 1 + i * 12, 11, 4, f"FAN {i+1}")
    for b in range(8):
        col, row = b % 2, b // 2
        x, y = c + 1 + col * 30, r + 6 + row * 3
        spec = ic.get(b + 1)
        if spec:
            name, plist = spec
            module(cv, y, x, 3, 29, f"IC {b+1} · {name}")
            ports_row(cv, y + 1, x + 1, [sp(n, key, pn) for key, pn in plist], gap=1)
        else:
            module(cv, y, x, 3, 29, f"IC {b+1}  빈 베이", empty=True)
    for i in range(4):
        module(cv, r + 19, c + 1 + i * 15, 3, 14, f"XFM {i+1}")
        cv.text(r + 20, c + 1 + i * 15, " 외부 링크 → IOX", sz=6, col=P["s400"])
    vmod(cv, r + 23, c + 1, 14, "OA 1", [sp(n, "oa1", "MGMT")])
    vmod(cv, r + 23, c + 16, 14, "OA 2", [sp(n, "oa2", "MGMT")])
    for k, x in enumerate((c + 31, c + 46)):
        module(cv, r + 23, x, 3, 14, f"GPSM {k+1}")
        cv.text(r + 24, x, " CAMnet → IOX", sz=6, col=P["s400"])
    for row, y in enumerate((r + 27, r + 31)):
        for i in range(5):
            fan(cv, y, c + 1 + i * 12, 11, 3, f"FAN {row * 5 + i + 6}")
    for k, x in enumerate((c + 1, c + 31)):
        first = k * 6 + 1
        slot(cv, r + 35, x, 29, f"AC {first}–{first + 5}" + (" (위)" if k == 0 else " (아래)"),
             [sp(n, f"ac{first + j}", str(first + j), 3) for j in range(6)], gap=1)
    return 38


# ═══ I/O 확장 박스 4U (IBM EMX0 · HPE Superdome 2 IOX 등) ═══════════════════════
def iox_front(cv, r, c, n, hint=True):
    module(cv, r + 1, c + 1, 14, 59, "전면 — 팬 · 전원 · 상태 표시 (연결 없음)")
    if hint:
        cv.text(r + 6, c + 1, "전면 모델 표기가 본체와 같게 보일 수 있음", sz=8, b=True, col=P["s600"], h="center")
        cv.merge(r + 6, c + 1, r + 6, c + 59)
        cv.text(r + 8, c + 1, "후면에 광케이블 포트(T1·T2)가 달린 I/O 모듈 2개 + 슬롯만 있고 HMC 포트가 없으면 I/O 드로어",
                sz=7, col=P["s500"], h="center")
        cv.merge(r + 8, c + 1, r + 8, c + 59)
    return 16


def iox_rear(cv, r, c, n, cards=None, mods=("P1", "P2"), cable=("T1", "T2")):
    """후면: I/O 모듈 2 (각 광케이블 포트 2 + 세로 슬롯 C1–C6) · PSU 2.
    cards: {(모듈 0/1, 슬롯 1–6): (이름, 포트수)}"""
    cards = cards or {}
    for m, x in enumerate((c + 1, c + 31)):
        mn = mods[m]
        module(cv, r + 1, x, 12, 29, f"{mn} · I/O 모듈")
        for t, y in enumerate((r + 3, r + 7)):
            port(cv, y, x + 1, *g(n, f"m{m+1}t{t+1}"), w=3)
            pname(cv, y + 1, x + 1, f"{mn}-{cable[t]}", 3)
        cv.text(r + 10, x + 1, "→ 본체", sz=6, col=P["s400"])
        for j in range(6):
            sx = x + 6 + j * 4
            spec = cards.get((m, j + 1))
            if spec:
                module(cv, r + 2, sx, 10, 3, f"C{j+1}")
                for p in range(spec[1]):
                    port(cv, r + 4 + p * 2, sx + 1, *g(n, f"m{m+1}c{j+1}p{p+1}"))
                cv.text(r + 11, sx, spec[0], sz=5.5, col=P["s600"])
            else:
                module(cv, r + 2, sx, 10, 3, f"C{j+1}", empty=True)
                cv.text(r + 6, sx, "빈", sz=6, col=P["s400"], h="center")
                cv.merge(r + 6, sx, r + 6, sx + 2)
    psu(cv, r + 13, c + 1, 14, "PSU 1", *g(n, "psu1"))
    psu(cv, r + 13, c + 46, 14, "PSU 2", *g(n, "psu2"))
    return 16


# ═══ IBM Power 본체 4U (예시 구성 — 슬롯 수·배치는 모델별 확인) ═══════════════════
def power4u_front(cv, r, c, n, filled=0, cap=None, raid=()):
    module(cv, r + 1, c + 1, 6, 12, "조작 패널 (LCD)")
    cv.text(r + 3, c + 2, "▭ 표시창", sz=6.5, col=P["s500"])
    cv.text(r + 4, c + 2, "● 전원 ○ ID", sz=6.5, col=P["s500"])
    cv.text(r + 6, c + 2, "← 본체에만 있음", sz=6, b=True, col=P["s600"])
    sff_bays(cv, r + 1, c + 15, 10, filled, 0, bh=6, ttl="NVMe U.2 · 0–9", cap=_caps(cap, 0, 10),
             raid=_raid_in(raid, 0, 10))
    module(cv, r + 10, c + 1, 5, 59, "팬 (베젤 안쪽 · 연결 없음)")
    return 16


def power4u_rear(cv, r, c, n, slots=None):
    """slots: {번호: (이름, 포트수[, 넓은포트[, 포트 접두 'P'|'T']])} — C1–C8, 없는 번호는 빈 슬롯."""
    slots = slots or {}
    for i in range(8):
        col, row = i % 2, i // 2
        x, y = c + 1 + col * 30, r + 1 + row * 2
        s = slots.get(i + 1)
        if s:
            wport = 3 if (len(s) > 2 and s[2]) else 2
            pre = s[3] if len(s) > 3 else "P"
            slot(cv, y, x, 29, f"C{i+1} · {s[0]}", [sp(n, f"c{i+1}p{j+1}", f"{pre}{j+1}", wport) for j in range(s[1])])
        else:
            slot(cv, y, x, 29, f"C{i+1}  빈 슬롯", empty=True)
    slot(cv, r + 10, c + 1, 29, "eBMC · 서비스 프로세서", [sp(n, "hmc1", "HMC1"), sp(n, "hmc2", "HMC2"),
                                                         sp(n, "usb", "", 2, "USB")])
    cv.text(r + 10, c + 31, "HMC 포트 = 본체 표시", sz=6.5, b=True, col=P["s600"])
    for i in range(4):
        psu(cv, r + 13, c + 1 + i * 15, 14, f"PSU {i+1}", *g(n, f"psu{i+1}"), inlet="C20", iw=3)
    return 16


# ═══ 2U 4노드 섀시 (노드 = 모듈 → 도면 번호 노드×100+포트) ═══════════════════════
def node4_front(cv, r, c, n, filled=(0, 0, 0, 0), cap=None):
    for k in range(4):
        sff_bays(cv, r + 1, c + 1 + k * 15, 6, filled[k], 0, bh=6, ttl=f"노드 {k+1} · 0–5", cap=cap)
    return 10


def node4_rear(cv, r, c, n):
    for k, (x, y) in enumerate(((c + 1, r + 5), (c + 34, r + 5), (c + 1, r + 1), (c + 34, r + 1))):
        no = k + 1
        slot(cv, y, x, 26, f"노드 {no}", [sp(n, f"n{no}_bmc", "BMC"), sp(n, f"n{no}_p1", "P1"), sp(n, f"n{no}_p2", "P2"),
                                         sp(n, f"n{no}_usb", "", 2, "USB"), sp(n, f"n{no}_vga", "", 3, "VGA")])
        slot(cv, y + 2, x, 26, f"노드 {no} PCIe  빈 슬롯", empty=True)
    psu(cv, r + 1, c + 28, 5, "PSU1", *g(n, "psu1"), h=4)
    psu(cv, r + 5, c + 28, 5, "PSU2", *g(n, "psu2"), h=4)
    return 10


# ═══ 2U 디스크 확장 선반 (DAE · JBOD) ═══════════════════════════════════════
def dae_front(cv, r, c, n, filled=0, cap=None, raid=()):
    return stg_front(cv, r, c, n, filled=filled, bays=25, cap=cap, raid=raid)


def dae_rear(cv, r, c, n):
    for k, y in (("b", r + 1), ("a", r + 4)):
        up = k.upper()
        slot(cv, y, c + 1, 29, f"LCC {up}", [sp(n, f"lcc{k}_a", "A (입력)", 3), sp(n, f"lcc{k}_b", "B (확장)", 3)], gap=3)
        psu(cv, y, c + 31, 29, f"PSU {up}", *g(n, f"psu{k}"))
    note(cv, r + 7, c + 1, "A = 컨트롤러(또는 앞 선반) 쪽 입력 · B = 다음 선반으로 확장", sz=6.5)
    return 10


# ═══ 반폭 1U 스위치 (트레이에 2대 나란히) ════════════════════════════════════
def half_sw_front(cv, r, c, n, conn=None):
    conn = conn or {}
    module(cv, r + 1, c + 1, 3, 12, "1–12 · SFP+")
    port_block(cv, r + 2, c + 1, 6, 2, 1, conn=conn)
    module(cv, r + 1, c + 14, 3, 6, "13–15")
    port_block(cv, r + 2, c + 14, 3, 1, 13, order="rowwise", conn=conn)
    module(cv, r + 1, c + 21, 3, 3, None)
    cv.text(r + 1, c + 21, "관리", sz=6.5, b=True)
    port(cv, r + 2, c + 22, *g(n, "mgmt"), label=None if g(n, "mgmt")[1] else "M")
    port(cv, r + 3, c + 22, *g(n, "con"), label=None if g(n, "con")[1] else "C")
    module(cv, r + 1, c + 31, 3, 29, "옆 자리 — 다른 반폭 장비 또는 빈 칸", empty=True)
    return 5


def half_sw_rear(cv, r, c, n):
    fan(cv, r + 1, c + 1, 8, 3, "FAN")
    psu(cv, r + 1, c + 10, 10, "PSU 1", *g(n, "psu1"), h=3)
    psu(cv, r + 1, c + 21, 8, "PSU 2", *g(n, "psu2"), h=3)
    module(cv, r + 1, c + 31, 3, 29, "옆 자리 (뒤에서 보면 좌우 반대)", empty=True)
    return 5


# ═══ 소형 · 데스크톱 장비 (선반 위) — 앞·뒤 나란히, 각 29칸 ═════════════════════
def small_front(cv, r, c, n, conn=None, ports=16):
    conn = conn or {}
    module(cv, r + 1, c + 1, 3, ports, f"1–{ports} · RJ45 1G")
    port_block(cv, r + 2, c + 1, ports // 2, 2, 1, conn=conn)
    module(cv, r + 1, c + ports + 2, 3, 26 - ports, "상태")
    cv.text(r + 2, c + ports + 2, " ● 전원", sz=6.5, col=P["s500"])
    return 6


def small_rear(cv, r, c, n, psu_n=1, dc=False):
    if dc:
        module(cv, r + 1, c + 1, 3, 8, "전원")
        port(cv, r + 2, c + 3, *g(n, "psu1"), label=None if g(n, "psu1")[1] else "DC")
        pname(cv, r + 3, c + 3, "12V", 2)
        module(cv, r + 1, c + 10, 3, 17, "잠금 슬롯 · 접지", empty=True)
        return 6
    module(cv, r + 1, c + 1, 3, 20, "후면 (단자·케이블 관리)")
    if psu_n:
        psu(cv, r + 1, c + 50, 10, "전원", *g(n, "psu1"), h=3)
    return 5
