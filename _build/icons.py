# -*- coding: utf-8 -*-
"""IT 인프라 아이콘 정의 (48×48, 단색). 파일명 키 · 한글 이름 · 분류."""
from icondsl import (Icon, rect, circle, ellipse, poly, ring, ring_sector, bar, arrow, darrow, gear, rot,
                     thick_line, _arc_segs, pt_on, Contour)
import math

ICONS = []
CATS = [("compute", "컴퓨팅 · 플랫폼"), ("network", "네트워크 · 보안"), ("storage", "스토리지 · 백업"),
        ("facility", "설비 · 물리"), ("people", "사람 · 장소"), ("ops", "운영 · 상태")]


def icon(key, name, cat):
    def deco(fn):
        ic = Icon(key, name, cat)
        fn(ic)
        ICONS.append(ic)
        return fn
    return deco


def chevron(cx, cy, w, h, t, left=True):
    if left:
        pts = [(cx + w / 2, cy - h / 2), (cx - w / 2, cy), (cx + w / 2, cy + h / 2)]
    else:
        pts = [(cx - w / 2, cy - h / 2), (cx + w / 2, cy), (cx - w / 2, cy + h / 2)]
    return thick_line(pts, t)


def db_band(cx, top_c, off0, off1, rx, ry):
    """원통 몸통의 곡선 띠(뚜껑 타원 아래)."""
    y0, y1 = top_c + off0, top_c + off1
    segs = [("M", cx - rx, y0)] + _arc_segs(cx, y0, 0, 180, 0, rx, ry)
    segs.append(("L", cx + rx, y1))
    segs += _arc_segs(cx, y1, 0, 0, 180, rx, ry)
    segs.append(("Z",))
    return Contour(segs)


# ═══ 컴퓨팅 · 플랫폼 ═══════════════════════════════════════════
@icon("server", "서버", "compute")
def _(i):
    for y in (6.5, 18.5, 30.5):
        i.fill(rect(7, y, 34, 10.5, 2.5))
        i.cut(rect(11, y + 4.1, 13, 2.4, 1.2), circle(30.5, y + 5.25, 1.7), circle(35.5, y + 5.25, 1.7))


@icon("server-tower", "타워 서버", "compute")
def _(i):
    i.fill(rect(13, 4, 22, 40, 3))
    i.cut(rect(17, 9, 14, 2.6, 1.3), rect(17, 14.5, 14, 2.6, 1.3), rect(17, 20, 14, 2.6, 1.3),
          circle(24, 36.5, 2.2))


@icon("blade", "블레이드 섀시", "compute")
def _(i):
    i.fill(rect(5, 8, 38, 32, 3))
    for k in range(6):
        i.cut(rect(9 + k * 5.25, 12, 3.4, 18.5, 1.2))
    i.cut(rect(9, 34, 30, 2.4, 1.2))


@icon("virtual-host", "가상화 호스트", "compute")
def _(i):
    i.fill(rect(5, 6, 38, 36, 4))
    i.cut(rect(8.5, 9.5, 31, 29, 2))
    for x in (11.5, 25):
        for y in (12.5, 25.5):
            i.fill(rect(x, y, 11.5, 10, 1.6))


@icon("vm", "가상 머신", "compute")
def _(i):
    L = 11
    t = 3.2
    for sx, sy in ((1, 1), (-1, 1), (1, -1), (-1, -1)):
        x0 = 5 if sx == 1 else 43
        y0 = 5 if sy == 1 else 43
        i.fill(poly([(x0, y0), (x0 + sx * L, y0), (x0 + sx * L, y0 + sy * t), (x0 + sx * t, y0 + sy * t),
                     (x0 + sx * t, y0 + sy * L), (x0, y0 + sy * L)]))
    i.fill(rect(14, 14, 20, 20, 3))
    i.cut(rect(17.5, 18, 13, 2.6, 1.3))


@icon("container", "컨테이너", "compute")
def _(i):
    i.fill(poly([(24, 5), (40.5, 13.2), (24, 21.4), (7.5, 13.2)]))
    i.fill(poly([(6.5, 16.2), (22.5, 24.2), (22.5, 43.5), (6.5, 35.5)]))
    i.fill(poly([(41.5, 16.2), (41.5, 35.5), (25.5, 43.5), (25.5, 24.2)]))
    for x in (11.2, 16.6):
        i.cut(poly([(x, 20.2 + (x - 6.5) * 0.5 + 1.5), (x + 1.8, 21.1 + (x - 6.5) * 0.5 + 1.5),
                    (x + 1.8, 34.8 + (x - 6.5) * 0.5 - 1.5), (x, 33.9 + (x - 6.5) * 0.5 - 1.5)]))
    for x in (35.0, 29.6):
        i.cut(poly([(x, 23.3 - (x - 25.5) * 0.5 + 1.5), (x + 1.8, 22.4 - (x - 25.5) * 0.5 + 1.5),
                    (x + 1.8, 36.1 - (x - 25.5) * 0.5 - 1.5), (x, 37.0 - (x - 25.5) * 0.5 - 1.5)]))


@icon("cluster", "클러스터", "compute")
def _(i):
    for x, y in ((18, 4.5), (5, 30.5), (31, 30.5)):
        i.fill(rect(x, y, 12, 11, 2.2))
        i.cut(rect(x + 3, y + 4.3, 6, 2.4, 1.2))
    i.new()
    i.fill(bar(21, 16.5, 13.5, 29.5, 2.6, "butt"))
    i.new()
    i.fill(bar(27, 16.5, 34.5, 29.5, 2.6, "butt"))
    i.new()
    i.fill(bar(18, 36, 30, 36, 2.6, "butt"))


@icon("database", "데이터베이스", "compute")
def _(i):
    cx, top, rx, ry = 24, 10, 16, 5.5
    i.fill(ellipse(cx, top, rx, ry))
    i.fill(db_band(cx, top, 2.6, 14.2, rx, ry))
    i.fill(db_band(cx, top, 16.8, 28.2, rx, ry))


@icon("application", "애플리케이션", "compute")
def _(i):
    i.fill(rect(4.5, 7, 39, 34, 3.5))
    for x in (9.5, 14, 18.5):
        i.cut(circle(x, 12, 1.4))
    i.cut(rect(8, 16.5, 32, 21, 1.5))
    i.fill(chevron(17.5, 27, 5, 10, 2.6, True))
    i.fill(chevron(30.5, 27, 5, 10, 2.6, False))
    i.fill(thick_line([(22.2, 32.5), (25.8, 21.5)], 2.4))


@icon("pc", "PC · 워크스테이션", "compute")
def _(i):
    i.fill(rect(4.5, 6, 39, 27, 3))
    i.cut(rect(8, 9.5, 32, 20, 1.5))
    i.new()
    i.fill(poly([(20.5, 33), (27.5, 33), (28.5, 38.5), (19.5, 38.5)]))
    i.new()
    i.fill(rect(13, 38, 22, 3.4, 1.7))


@icon("laptop", "노트북", "compute")
def _(i):
    i.fill(rect(9, 8, 30, 22, 2.5))
    i.cut(rect(12.2, 11.2, 23.6, 15.6, 1))
    i.fill(rect(3.5, 32, 41, 5.5, 2.5))
    i.cut(rect(19.5, 33.2, 9, 1.8, 0.9))


@icon("terminal", "콘솔 · 터미널", "compute")
def _(i):
    i.fill(rect(4.5, 7, 39, 34, 3.5))
    for x in (9.5, 14, 18.5):
        i.cut(circle(x, 12, 1.4))
    i.cut(rect(8, 16.5, 32, 21, 1.5))
    i.fill(chevron(15.5, 25, 5, 9, 2.6, False))
    i.fill(rect(21, 30.2, 11, 2.6, 1.3))


# ═══ 네트워크 · 보안 ═══════════════════════════════════════════
@icon("switch", "스위치 (L2)", "network")
def _(i):
    i.fill(rect(4, 13, 40, 22, 4))
    i.cut(arrow(9.5, 20.2, 38.5, 20.2, 3, 7, 7.4))
    i.cut(arrow(38.5, 27.8, 9.5, 27.8, 3, 7, 7.4))


@icon("switch-l3", "L3 · 코어 스위치", "network")
def _(i):
    i.fill(rect(5, 5, 38, 38, 6))
    i.cut(arrow(24, 20.5, 24, 9.5, 3, 6.5, 7.6))
    i.cut(arrow(24, 27.5, 24, 38.5, 3, 6.5, 7.6))
    i.cut(arrow(20.5, 24, 9.5, 24, 3, 6.5, 7.6))
    i.cut(arrow(27.5, 24, 38.5, 24, 3, 6.5, 7.6))


@icon("router", "라우터", "network")
def _(i):
    i.fill(circle(24, 24, 20))
    i.cut(arrow(8, 24, 19.5, 24, 3, 6, 7.4))
    i.cut(arrow(40, 24, 28.5, 24, 3, 6, 7.4))
    i.cut(arrow(24, 20.5, 24, 8.5, 3, 6, 7.4))
    i.cut(arrow(24, 27.5, 24, 39.5, 3, 6, 7.4))


@icon("firewall", "방화벽", "network")
def _(i):
    g, h = 1.8, 6.4
    rows = [7.5, 7.5 + h + g, 7.5 + 2 * (h + g), 7.5 + 3 * (h + g)]
    full = (38 - 2 * g) / 3
    for k, y in enumerate(rows):
        if k % 2 == 0:
            xs = [(5, full), (5 + full + g, full), (5 + 2 * (full + g), full)]
        else:
            half = (38 - 3 * g - 2 * full) / 2
            xs = [(5, half), (5 + half + g, full), (5 + half + full + 2 * g, full), (43 - half, half)]
        for x, w in xs:
            i.fill(rect(x, y, w, h, 1.2))


@icon("load-balancer", "로드밸런서 (L4)", "network")
def _(i):
    i.fill(rect(4, 9, 40, 30, 4))
    cx, cy = 20, 24
    i.cut(circle(cx, cy, 3.4))
    i.cut(bar(8.5, 24, 14.5, 24, 3, "butt"))
    for a, end in ((-35, (37, 13.5)), (0, (38.5, 24)), (35, (37, 34.5))):
        sx, sy = pt_on(cx, cy, 7.2, a)
        i.cut(arrow(sx, sy, end[0], end[1], 3, 6, 7))


@icon("security", "보안 장비 (IPS·WAF)", "network")
def _(i):
    i.fill(Contour([("M", 24, 4), ("L", 40.5, 9.5), ("L", 40.5, 22), ("C", 40.5, 32, 33, 39.5, 24, 44),
                    ("C", 15, 39.5, 7.5, 32, 7.5, 22), ("L", 7.5, 9.5), ("Z",)]))
    i.cut(thick_line([(16, 23.5), (21.5, 29), (32.5, 17.5)], 3.8))


@icon("vpn", "VPN", "network")
def _(i):
    i.fill(rect(2.5, 12, 43, 24, 12))
    i.cut(rect(6.5, 16, 35, 16, 8))
    # 자물쇠 (몸통 + 고리 외곽, 고리 안쪽은 구멍)
    R, r, cy = 4.3, 2.4, 21.2
    segs = [("M", 19.3, 30.2), ("L", 28.7, 30.2), ("L", 28.7, 23.2), ("L", 24 + R, 23.2), ("L", 24 + R, cy)]
    segs += _arc_segs(24, cy, R, 0, -180)
    segs += [("L", 24 - R, 23.2), ("L", 19.3, 23.2), ("Z",)]
    i.fill(Contour(segs))
    segs = [("M", 24 - r, 23.2), ("L", 24 - r, cy)] + _arc_segs(24, cy, r, 180, 360) + [("L", 24 + r, 23.2), ("Z",)]
    i.cut(Contour(segs))


@icon("wireless", "무선 AP", "network")
def _(i):
    i.fill(rect(7, 31, 34, 9, 4.5))
    i.cut(rect(19.5, 34.3, 9, 2.4, 1.2))
    cx, cy = 24, 29
    i.fill(circle(cx, 25.5, 2.3))
    for ro in (9, 15, 21):
        i.fill(ring_sector(cx, cy, ro, ro - 3, 222, 318))


@icon("internet", "인터넷", "network")
def _(i):
    i.fill(circle(24, 24, 20))
    i.cut(circle(24, 24, 17))
    i.new()
    i.fill(ellipse(24, 24, 8.5, 18.2))
    i.cut(ellipse(24, 24, 5.5, 18.2 - 3))
    i.new()
    i.fill(rect(5, 22.5, 38, 3, 1.5))
    i.new()
    i.fill(rect(9, 13, 30, 2.6, 1.3))
    i.new()
    i.fill(rect(9, 32.4, 30, 2.6, 1.3))


@icon("cloud", "클라우드", "network")
def _(i):
    i.fill(rect(5, 23, 38, 17, 8.5))
    i.new()
    i.fill(circle(17.5, 24, 8.5))
    i.new()
    i.fill(circle(28.5, 19.5, 11.5))


@icon("wan", "WAN · 전용회선", "network")
def _(i):
    i.fill(rect(3.5, 17, 11, 14, 2.5))
    i.cut(rect(6.3, 21.5, 5.4, 2.2, 1.1))
    i.fill(rect(33.5, 17, 11, 14, 2.5))
    i.cut(rect(36.3, 21.5, 5.4, 2.2, 1.1))
    i.new()
    i.fill(thick_line([(16, 27), (21.5, 18.5), (26.5, 29.5), (32, 21)], 2.8))


@icon("patch-panel", "패치패널", "network")
def _(i):
    i.fill(rect(3, 14, 42, 20, 3))
    for row in (18, 25.5):
        for k in range(6):
            i.cut(rect(7 + k * 6, row, 4.2, 4.6, 0.8))


@icon("cable-utp", "UTP 케이블", "network")
def _(i):
    i.fill(rect(14.5, 5, 19, 19, 2.5))
    for k in range(4):
        i.cut(rect(18 + k * 3.4, 8, 1.6, 6, 0.6))
    i.cut(rect(19, 17.5, 10, 2.6, 1.3))
    i.new()
    i.fill(rect(19.5, 23, 9, 5, 1))
    i.new()
    i.fill(bar(24, 27, 24, 43, 4.6))


@icon("cable-fiber", "광케이블", "network")
def _(i):
    i.fill(rect(16.5, 20.5, 15, 14, 2.5))
    i.cut(circle(24, 27, 2.6))
    i.fill(rect(20.5, 13.5, 7, 5.5, 1.2))
    i.new()
    i.fill(bar(24, 35.5, 24, 41.8, 4.4))
    for a in (-55, -90, -125):
        i.new()
        x1, y1 = pt_on(24, 12.5, 4.2, a)
        x2, y2 = pt_on(24, 12.5, 9, a)
        i.fill(bar(x1, y1, x2, y2, 2.2))


# ═══ 스토리지 · 백업 ═══════════════════════════════════════════
@icon("storage", "스토리지", "storage")
def _(i):
    i.fill(rect(4, 8, 40, 32, 3.5))
    for y in (12.5, 25.5):
        for k in range(5):
            i.cut(rect(8.3 + k * 6.6, y, 5.2, 9.5, 1.2))


@icon("san-switch", "SAN 스위치", "storage")
def _(i):
    i.fill(rect(4, 13, 40, 22, 4))
    i.cut(arrow(15, 20.2, 38.5, 20.2, 3, 6.5, 7.2))
    i.cut(arrow(38.5, 27.8, 15, 27.8, 3, 6.5, 7.2))
    i.cut(circle(9.5, 24, 2.6))


@icon("nas", "NAS", "storage")
def _(i):
    i.fill(rect(5, 7, 38, 34, 3.5))
    i.cut(poly([(13, 14.5), (20, 14.5), (22.5, 17.5), (35, 17.5), (35, 30.5), (13, 30.5)]))
    i.cut(rect(13, 34.5, 22, 2.4, 1.2))


@icon("hdd", "하드디스크", "storage")
def _(i):
    i.fill(rect(9, 4, 30, 40, 4))
    i.cut(circle(24, 19.5, 10.5))
    i.fill(circle(24, 19.5, 2.8))
    i.cut(circle(14, 38.5, 1.5), circle(34, 38.5, 1.5))
    i.new()
    i.fill(bar(33, 33, 27.5, 24.5, 2.8))


@icon("ssd", "SSD · 플래시", "storage")
def _(i):
    i.fill(rect(6, 9, 36, 30, 3))
    for x in (10.5, 20.5, 30.5):
        i.cut(rect(x, 13.5, 7, 13, 1.2))
    for k in range(5):
        i.cut(rect(11 + k * 5.6, 31, 3, 4, 0.7))


@icon("disk-shelf", "디스크 쉘프 (JBOD)", "storage")
def _(i):
    i.fill(rect(4, 11, 40, 26, 3.5))
    for k in range(8):
        i.cut(rect(7.8 + k * 4.2, 15, 2.8, 14.5, 1.1))
    i.cut(rect(7.8, 32, 32.2, 1.8, 0.9))


@icon("tape", "테이프", "storage")
def _(i):
    i.fill(rect(4, 9, 40, 30, 4))
    for x in (15, 33):
        i.cut(circle(x, 21.5, 6))
        i.fill(circle(x, 21.5, 2.4))
    i.cut(rect(15, 31.5, 18, 3.4, 1.7))


@icon("backup", "백업 · 복구", "storage")
def _(i):
    cx, cy = 24, 24
    i.fill(ring_sector(cx, cy, 19, 15.6, 30, 318))
    i.new()
    hx, hy = pt_on(cx, cy, 17.3, 318)
    tang = math.radians(318 + 90)
    ux, uy = math.cos(tang), math.sin(tang)
    nx, ny = -uy, ux
    i.fill(poly([(hx - ux * 1 + nx * 5.8, hy - uy * 1 + ny * 5.8), (hx + ux * 7.5, hy + uy * 7.5),
                 (hx - ux * 1 - nx * 5.8, hy - uy * 1 - ny * 5.8)]))
    i.new()
    i.fill(bar(24, 24, 24, 13, 3))
    i.new()
    i.fill(bar(24, 24, 31, 28.5, 3))


@icon("object-storage", "오브젝트 스토리지", "storage")
def _(i):
    i.fill(poly([(8.5, 14), (39.5, 14), (35, 41), (13, 41)]))
    for x, y in ((16, 19), (25, 19), (20.5, 28)):
        i.cut(rect(x, y, 7, 7, 1.2))
    i.new()
    i.fill(rect(5.5, 7.5, 37, 5, 2.5))


# ═══ 설비 · 물리 ═══════════════════════════════════════════════
@icon("rack", "랙", "facility")
def _(i):
    i.fill(rect(10, 3.5, 28, 41, 3))
    for k in range(5):
        i.cut(rect(14, 8 + k * 6.3, 20, 4.2, 1))
    i.cut(circle(24, 40, 1.5))


@icon("pdu", "PDU", "facility")
def _(i):
    i.fill(rect(16.5, 3, 15, 42, 3))
    for k in range(6):
        i.cut(rect(20, 7 + k * 6, 8, 4, 1.3))


@icon("ups", "UPS", "facility")
def _(i):
    i.fill(rect(8, 8.5, 32, 34.5, 4))
    i.cut(poly([(26.5, 13), (17, 27.5), (23.3, 27.5), (21, 38.5), (31, 23), (24.8, 23), (28, 13)]))
    i.new()
    i.fill(rect(13.5, 4.5, 6.5, 4.5, 1.2))
    i.new()
    i.fill(rect(28, 4.5, 6.5, 4.5, 1.2))


@icon("cooling", "항온항습기 · 냉방", "facility")
def _(i):
    i.fill(rect(5, 5, 38, 38, 5))
    i.cut(circle(24, 24, 14))
    for a in (0, 120, 240):
        e = rot(ellipse(24 + 7.2, 24, 5.4, 3.2), 24, 24, a - 90)
        i.fill(e)


@icon("power", "전원", "facility")
def _(i):
    i.fill(poly([(27.5, 4), (12, 26.5), (22, 26.5), (19.5, 44), (36, 20), (26, 20), (29.5, 4)]))


@icon("sensor", "온습도 센서", "facility")
def _(i):
    R = 8.5
    dy = math.sqrt(R * R - 25)
    a = math.degrees(math.atan2(-dy, 5))
    segs = [("M", 19, 10)] + _arc_segs(24, 10, 5, 180, 360) + [("L", 29, 36 - dy)]
    segs += _arc_segs(24, 36, R, a, 180 - a)
    segs += [("L", 19, 36 - dy), ("Z",)]
    i.fill(Contour(segs))
    i.cut(rect(22.2, 9, 3.6, 17, 1.8))
    i.cut(circle(24, 36, 4.4))


@icon("datacenter", "데이터센터", "facility")
def _(i):
    i.fill(rect(5, 6, 22, 37, 2.5))
    for r_ in range(4):
        for c_ in range(2):
            i.cut(rect(9.5 + c_ * 7.5, 11 + r_ * 7, 5, 3.6, 0.8))
    i.new()
    i.fill(rect(29, 17, 14, 26, 2.5))
    for r_ in range(3):
        i.cut(rect(32.5, 22 + r_ * 6.5, 7, 3.4, 0.8))


@icon("office", "사무실 · 지점", "facility")
def _(i):
    i.fill(poly([(24, 5), (44, 21), (39.5, 21), (39.5, 43), (8.5, 43), (8.5, 21), (4, 21)]))
    i.cut(rect(20, 29.5, 8, 10, 1.2))
    i.cut(rect(12.5, 25, 5, 5, 0.8), rect(30.5, 25, 5, 5, 0.8))


# ═══ 사람 · 장소 ═══════════════════════════════════════════════
def person(i, cx, top, s=1.0):
    i.fill(circle(cx, top + 8 * s, 7.6 * s))
    y0 = top + 19 * s
    w = 16 * s
    i.fill(Contour([("M", cx - w, top + 35 * s), ("L", cx - w, y0 + 8 * s),
                    ("C", cx - w, y0 + 3 * s, cx - 8 * s, y0, cx, y0),
                    ("C", cx + 8 * s, y0, cx + w, y0 + 3 * s, cx + w, y0 + 8 * s),
                    ("L", cx + w, top + 35 * s), ("Z",)]))


@icon("user", "사용자", "people")
def _(i):
    person(i, 24, 6)


@icon("users", "사용자 그룹", "people")
def _(i):
    person(i, 13.5, 13, 0.62)
    i.new()
    person(i, 34.5, 13, 0.62)


@icon("operator", "운영자 · 관리자", "people")
def _(i):
    person(i, 24, 7)
    i.new()
    i.fill(ring_sector(24, 15, 12, 9.6, 180, 360))
    i.new()
    i.fill(rect(9.6, 13, 4.2, 8, 2))
    i.new()
    i.fill(rect(34.2, 13, 4.2, 8, 2))


# ═══ 운영 · 상태 ═══════════════════════════════════════════════
@icon("monitoring", "모니터링", "ops")
def _(i):
    i.fill(rect(4.5, 6, 39, 27, 3))
    i.cut(rect(8, 9.5, 32, 20, 1.5))
    i.new()
    i.fill(thick_line([(10.5, 21), (16.5, 21), (19.5, 14), (24, 26), (27.5, 18), (30, 21), (37.5, 21)], 2.4))
    i.new()
    i.fill(poly([(20.5, 33), (27.5, 33), (28.5, 38.5), (19.5, 38.5)]))
    i.new()
    i.fill(rect(13, 38, 22, 3.4, 1.7))


@icon("document", "문서 · 로그", "ops")
def _(i):
    i.fill(poly([(9.5, 4), (29.5, 4), (38.5, 13), (38.5, 44), (9.5, 44)]))
    i.cut(poly([(28.5, 6.5), (36, 14), (28.5, 14)]))
    for y, w in ((20, 19), (26.5, 19), (33, 13)):
        i.cut(rect(14.5, y, w, 2.8, 1.4))


@icon("alarm", "알람 · 알림", "ops")
def _(i):
    i.fill(Contour([("M", 24, 7), ("C", 15.5, 7, 12, 13.5, 12, 20.5), ("L", 12, 29), ("L", 7.5, 34.5),
                    ("L", 40.5, 34.5), ("L", 36, 29), ("L", 36, 20.5), ("C", 36, 13.5, 32.5, 7, 24, 7), ("Z",)]))
    i.new()
    i.fill(circle(24, 5.5, 2.4))
    i.new()
    i.fill(Contour([("M", 19.5, 37)] + _arc_segs(24, 37, 4.5, 180, 0) + [("Z",)]))


@icon("warning", "경고", "ops")
def _(i):
    i.fill(Contour([("M", 21.4, 6.5), ("C", 22.6, 4.5, 25.4, 4.5, 26.6, 6.5), ("L", 43.3, 36.5),
                    ("C", 44.4, 38.6, 43, 41, 40.6, 41), ("L", 7.4, 41), ("C", 5, 41, 3.6, 38.6, 4.7, 36.5), ("Z",)]))
    i.cut(rect(22, 15, 4, 13.5, 2), circle(24, 34, 2.3))


@icon("key", "인증 · 키", "ops")
def _(i):
    i.fill(circle(15, 24, 10.5))
    i.cut(circle(15, 24, 4.2))
    i.new()
    i.fill(rect(23.5, 21.5, 20, 5, 1.2))
    i.new()
    i.fill(rect(33.5, 24, 3.6, 8, 0.8))
    i.new()
    i.fill(rect(39.4, 24, 3.6, 6, 0.8))


@icon("lock", "보안 · 잠금", "ops")
def _(i):
    i.fill(rect(10, 21, 28, 22, 3.5))
    i.cut(circle(24, 29.5, 2.7))
    i.cut(rect(22.9, 32.6, 2.2, 5.4, 1))
    i.new()
    i.fill(ring_sector(24, 19, 10, 6.6, 180, 360))
    i.new()
    i.fill(rect(14, 18.5, 3.4, 3.5, 0))
    i.new()
    i.fill(rect(30.6, 18.5, 3.4, 3.5, 0))


@icon("sync", "복제 · 동기화", "ops")
def _(i):
    cx, cy = 24, 24
    for a0, a1 in ((200, 330), (20, 150)):
        i.new()
        i.fill(ring_sector(cx, cy, 17.5, 14, a0, a1))
        i.new()
        hx, hy = pt_on(cx, cy, 15.75, a1)
        tang = math.radians(a1 + 90)
        ux, uy = math.cos(tang), math.sin(tang)
        nx, ny = -uy, ux
        i.fill(poly([(hx - ux * 1 + nx * 5.8, hy - uy * 1 + ny * 5.8), (hx + ux * 7.5, hy + uy * 7.5),
                     (hx - ux * 1 - nx * 5.8, hy - uy * 1 - ny * 5.8)]))


@icon("settings", "설정", "ops")
def _(i):
    i.fill(gear(24, 24, 19, 14.5, 8))
    i.cut(circle(24, 24, 6.2))


@icon("clock", "시간 · NTP", "ops")
def _(i):
    i.fill(circle(24, 24, 20))
    i.cut(circle(24, 24, 16.6))
    i.new()
    i.fill(bar(24, 24, 24, 12, 3))
    i.new()
    i.fill(bar(24, 24, 32, 28.5, 3))


@icon("mail", "메일", "ops")
def _(i):
    i.fill(rect(4.5, 10, 39, 28, 3))
    i.cut(poly([(9, 14.5), (24, 25.5), (39, 14.5), (39, 18.2), (24, 29.2), (9, 18.2)]))


@icon("ok", "정상", "ops")
def _(i):
    i.fill(circle(24, 24, 20))
    i.cut(thick_line([(14, 24.5), (21, 31.5), (34.5, 17)], 4.4))


@icon("error", "장애", "ops")
def _(i):
    i.fill(circle(24, 24, 20))
    t, L = 2.3, 9.5
    cross = poly([(-t, -L), (t, -L), (t, -t), (L, -t), (L, t), (t, t), (t, L), (-t, L), (-t, t), (-L, t), (-L, -t), (-t, -t)])
    i.cut(rot(cross.transformed(lambda x, y: (x + 24, y + 24)), 24, 24, 45))


@icon("info", "정보", "ops")
def _(i):
    i.fill(circle(24, 24, 20))
    i.cut(circle(24, 14.5, 2.7))
    i.cut(rect(21.8, 20.5, 4.4, 15, 2.2))


@icon("calendar", "일정", "ops")
def _(i):
    i.fill(rect(5.5, 8.5, 37, 34, 3.5))
    i.cut(rect(9, 18, 30, 21, 1.5))
    for r_ in range(2):
        for c_ in range(4):
            i.fill(rect(11.3 + c_ * 6.9, 21 + r_ * 8, 4.6, 5, 0.8))
    i.new()
    i.fill(rect(13.5, 4.5, 3.6, 8.5, 1.8))
    i.new()
    i.fill(rect(30.9, 4.5, 3.6, 8.5, 1.8))
