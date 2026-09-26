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
def srv1u_front(cv, r, c, n, bays=10, filled=4, ttl="SFF 1–10"):
    module(cv, r + 1, c + 1, 3, 5, "UID")
    cv.text(r + 2, c + 1, "● 상태", sz=6.5, col=P["s500"])
    cv.text(r + 3, c + 1, "○ 전원", sz=6.5, col=P["s500"])
    module(cv, r + 1, c + 7, 3, bays * 2 + 2, None)
    sff_bays(cv, r + 1, c + 8, bays, filled, 1, bh=3, pad=False)
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
def srv2u_front(cv, r, c, n, boxes=(8, 8, 0), lff=False, filled_lff=0):
    if lff:
        lff_bays(cv, r + 1, c + 1, 4, 3, filled_lff, bw=14, ttl=None)
        return 10
    x = c + 1
    for bi, filled in enumerate(boxes):
        sff_bays(cv, r + 1, x, 8, filled, start=bi * 8 + 1, bh=6, ttl=f"BOX {bi + 1}")
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
def srv4u_front(cv, r, c, n, filled=12):
    lff_bays(cv, r + 1, c + 1, 4, 6, filled, bw=14, ttl=None)
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
def stg_front(cv, r, c, n, filled=12, bays=25):
    sff_bays(cv, r + 1, c + 1, bays, filled, 0, bh=6, ttl=f"드라이브 0–{bays-1}")
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


def small_rear(cv, r, c, n, psu_n=1):
    module(cv, r + 1, c + 1, 3, 20, "후면 (단자·케이블 관리)")
    if psu_n:
        psu(cv, r + 1, c + 50, 10, "전원", *g(n, "psu1"), h=3)
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
def tower_front(cv, r, c, n, filled=4):
    module(cv, r + 1, c + 2, 2, 25, "5.25\" 베이 1 · 광학 드라이브")
    module(cv, r + 4, c + 2, 2, 25, "5.25\" 베이 2  빈 베이", empty=True)
    vmod(cv, r + 7, c + 2, 25, "전면 I/O", [sp(n, "f_usb1", "", 2, "USB"), sp(n, "f_usb2", "", 2, "USB"),
                                           sp(n, "f_svc", "", 2, "SVC")], names=False)
    lff_bays(cv, r + 11, c + 2, 1, 8, filled, bw=23, ttl=None)
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
