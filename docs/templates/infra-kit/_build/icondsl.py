# -*- coding: utf-8 -*-
"""아이콘 정의 도구: 48×48 격자에서 윤곽선을 만들고 SVG / DrawingML(custGeom)로 내보낸다.

규칙: 한 그룹(path) 안에서
  - 채움 윤곽 = 시계 방향(CW), 구멍 = 반시계(CCW), 구멍 안 섬 = CW
  - 같은 그룹의 채움끼리는 겹치지 않게 (겹치면 다른 그룹으로)
→ evenodd / nonzero 어느 규칙으로 칠해도 결과가 같다 (Office·브라우저 공통).
"""
import math

K = 0.5522847498


class Contour:
    def __init__(self, segs):
        self.segs = segs  # [('M',x,y),('L',x,y),('C',x1,y1,x2,y2,x,y),('Z',)]

    def points(self):
        pts = []
        for s in self.segs:
            if s[0] in ("M", "L"):
                pts.append((s[1], s[2]))
            elif s[0] == "C":
                pts += [(s[1], s[2]), (s[3], s[4]), (s[5], s[6])]
        return pts

    def area(self):
        p = self.points()
        a = 0
        for i in range(len(p)):
            x1, y1 = p[i]
            x2, y2 = p[(i + 1) % len(p)]
            a += x1 * y2 - x2 * y1
        return a / 2  # y-down: 양수 = 시계 방향

    def reversed(self):
        # 선분 목록을 뒤집는다
        pts = []  # (type, ctrl..., end) 순서로 기록
        cur = None
        items = []
        for s in self.segs:
            if s[0] == "M":
                cur = (s[1], s[2])
                start = cur
            elif s[0] == "L":
                items.append(("L", cur, (s[1], s[2])))
                cur = (s[1], s[2])
            elif s[0] == "C":
                items.append(("C", cur, (s[1], s[2]), (s[3], s[4]), (s[5], s[6])))
                cur = (s[5], s[6])
        if cur != start:
            items.append(("L", cur, start))
        out = [("M",) + items[-1][-1]]
        for it in reversed(items):
            if it[0] == "L":
                out.append(("L",) + it[1])
            else:
                out.append(("C",) + it[3] + it[2] + it[1])
        out.append(("Z",))
        return Contour(out)

    def oriented(self, cw=True):
        a = self.area()
        if (a > 0) == cw:
            return self
        return self.reversed()

    def transformed(self, fn):
        out = []
        for s in self.segs:
            if s[0] in ("M", "L"):
                out.append((s[0],) + fn(s[1], s[2]))
            elif s[0] == "C":
                out.append(("C",) + fn(s[1], s[2]) + fn(s[3], s[4]) + fn(s[5], s[6]))
            else:
                out.append(s)
        return Contour(out)


# ── 기본 도형 ───────────────────────────────────────────────
def rect(x, y, w, h, r=0):
    r = min(r, w / 2, h / 2)
    if r <= 0:
        return Contour([("M", x, y), ("L", x + w, y), ("L", x + w, y + h), ("L", x, y + h), ("Z",)])
    k = r * K
    return Contour([
        ("M", x + r, y), ("L", x + w - r, y),
        ("C", x + w - r + k, y, x + w, y + r - k, x + w, y + r),
        ("L", x + w, y + h - r),
        ("C", x + w, y + h - r + k, x + w - r + k, y + h, x + w - r, y + h),
        ("L", x + r, y + h),
        ("C", x + r - k, y + h, x, y + h - r + k, x, y + h - r),
        ("L", x, y + r),
        ("C", x, y + r - k, x + r - k, y, x + r, y), ("Z",)])


def ellipse(cx, cy, rx, ry):
    kx, ky = rx * K, ry * K
    return Contour([
        ("M", cx, cy - ry),
        ("C", cx + kx, cy - ry, cx + rx, cy - ky, cx + rx, cy),
        ("C", cx + rx, cy + ky, cx + kx, cy + ry, cx, cy + ry),
        ("C", cx - kx, cy + ry, cx - rx, cy + ky, cx - rx, cy),
        ("C", cx - rx, cy - ky, cx - kx, cy - ry, cx, cy - ry), ("Z",)])


def circle(cx, cy, r):
    return ellipse(cx, cy, r, r)


def poly(pts):
    segs = [("M",) + tuple(pts[0])] + [("L",) + tuple(p) for p in pts[1:]] + [("Z",)]
    return Contour(segs)


def _arc_segs(cx, cy, r, a0, a1, rx=None, ry=None):
    """a0→a1(도, 시계방향 증가 = y-down) 호를 베지어로. 시작점 제외 세그먼트 반환."""
    rx = rx or r
    ry = ry or r
    segs = []
    n = max(1, int(math.ceil(abs(a1 - a0) / 90.0)))
    da = (a1 - a0) / n
    for i in range(n):
        t0 = math.radians(a0 + da * i)
        t1 = math.radians(a0 + da * (i + 1))
        k = 4 / 3 * math.tan((t1 - t0) / 4)
        x0, y0 = cx + rx * math.cos(t0), cy + ry * math.sin(t0)
        x3, y3 = cx + rx * math.cos(t1), cy + ry * math.sin(t1)
        x1 = x0 - k * rx * math.sin(t0)
        y1 = y0 + k * ry * math.cos(t0)
        x2 = x3 + k * rx * math.sin(t1)
        y2 = y3 - k * ry * math.cos(t1)
        segs.append(("C", x1, y1, x2, y2, x3, y3))
    return segs


def pt_on(cx, cy, r, a, rx=None, ry=None):
    rx = rx or r
    ry = ry or r
    return (cx + rx * math.cos(math.radians(a)), cy + ry * math.sin(math.radians(a)))


def ring_sector(cx, cy, r_out, r_in, a0, a1):
    """두꺼운 호(고리 조각)."""
    p0 = pt_on(cx, cy, r_out, a0)
    segs = [("M",) + p0] + _arc_segs(cx, cy, r_out, a0, a1)
    segs.append(("L",) + pt_on(cx, cy, r_in, a1))
    segs += _arc_segs(cx, cy, r_in, a1, a0)
    segs.append(("Z",))
    return Contour(segs)


def ring(cx, cy, r_out, r_in):
    return [circle(cx, cy, r_out).oriented(True), circle(cx, cy, r_in).oriented(False)]


def bar(x1, y1, x2, y2, t, cap="round"):
    """두께 t 선분 (둥근/각진 끝)."""
    dx, dy = x2 - x1, y2 - y1
    L = math.hypot(dx, dy)
    ux, uy = dx / L, dy / L
    nx, ny = -uy * t / 2, ux * t / 2
    if cap == "butt":
        return poly([(x1 + nx, y1 + ny), (x2 + nx, y2 + ny), (x2 - nx, y2 - ny), (x1 - nx, y1 - ny)])
    a = math.degrees(math.atan2(dy, dx))
    r = t / 2
    segs = [("M", x1 + nx, y1 + ny), ("L", x2 + nx, y2 + ny)]
    segs += _arc_segs(x2, y2, r, a + 90, a - 90)
    segs.append(("L", x1 - nx, y1 - ny))
    segs += _arc_segs(x1, y1, r, a - 90, a - 270)
    segs.append(("Z",))
    return Contour(segs)


def thick_line(pts, t):
    """꺾인 선(마이터 연결, 각진 끝)을 다각형으로."""
    h = t / 2
    n = len(pts)
    left, right = [], []
    def nrm(p, q):
        dx, dy = q[0] - p[0], q[1] - p[1]
        L = math.hypot(dx, dy)
        return (-dy / L, dx / L)
    for i in range(n):
        if i == 0:
            nx, ny = nrm(pts[0], pts[1])
            left.append((pts[0][0] + nx * h, pts[0][1] + ny * h))
            right.append((pts[0][0] - nx * h, pts[0][1] - ny * h))
        elif i == n - 1:
            nx, ny = nrm(pts[-2], pts[-1])
            left.append((pts[-1][0] + nx * h, pts[-1][1] + ny * h))
            right.append((pts[-1][0] - nx * h, pts[-1][1] - ny * h))
        else:
            n1 = nrm(pts[i - 1], pts[i])
            n2 = nrm(pts[i], pts[i + 1])
            mx, my = n1[0] + n2[0], n1[1] + n2[1]
            ml = math.hypot(mx, my)
            mx, my = mx / ml, my / ml
            d = h / max(0.25, (mx * n1[0] + my * n1[1]))
            left.append((pts[i][0] + mx * d, pts[i][1] + my * d))
            right.append((pts[i][0] - mx * d, pts[i][1] - my * d))
    return poly(left + right[::-1])


def arrow(x1, y1, x2, y2, t=3, head=7, hw=None):
    """x1,y1 → x2,y2 화살표 (몸통 두께 t, 머리 길이 head, 머리 폭 hw)."""
    hw = hw or head * 1.1
    dx, dy = x2 - x1, y2 - y1
    L = math.hypot(dx, dy)
    ux, uy = dx / L, dy / L
    nx, ny = -uy, ux
    bx, by = x2 - ux * head, y2 - uy * head
    pts = [(x1 + nx * t / 2, y1 + ny * t / 2), (bx + nx * t / 2, by + ny * t / 2), (bx + nx * hw / 2, by + ny * hw / 2),
           (x2, y2), (bx - nx * hw / 2, by - ny * hw / 2), (bx - nx * t / 2, by - ny * t / 2),
           (x1 - nx * t / 2, y1 - ny * t / 2)]
    return poly(pts)


def darrow(x1, y1, x2, y2, t=3, head=6, hw=None):
    """양방향 화살표."""
    hw = hw or head * 1.2
    dx, dy = x2 - x1, y2 - y1
    L = math.hypot(dx, dy)
    ux, uy = dx / L, dy / L
    nx, ny = -uy, ux
    ax, ay = x1 + ux * head, y1 + uy * head
    bx, by = x2 - ux * head, y2 - uy * head
    pts = [(x1, y1), (ax + nx * hw / 2, ay + ny * hw / 2), (ax + nx * t / 2, ay + ny * t / 2),
           (bx + nx * t / 2, by + ny * t / 2), (bx + nx * hw / 2, by + ny * hw / 2), (x2, y2),
           (bx - nx * hw / 2, by - ny * hw / 2), (bx - nx * t / 2, by - ny * t / 2),
           (ax - nx * t / 2, ay - ny * t / 2), (ax - nx * hw / 2, ay - ny * hw / 2)]
    return poly(pts)


def gear(cx, cy, r_out, r_in, teeth=8, tooth=0.45):
    pts = []
    n = teeth * 4
    for i in range(n):
        a = 2 * math.pi * i / n - math.pi / 2
        phase = i % 4
        r = r_out if phase in (1, 2) else r_in
        # 톱니 옆면을 살짝 기울임
        off = (tooth * 2 * math.pi / teeth / 4) * (1 if phase == 1 else -1 if phase == 2 else 0) * 0.3
        pts.append((cx + r * math.cos(a + off), cy + r * math.sin(a + off)))
    return poly(pts)


def rot(c, cx, cy, deg):
    t = math.radians(deg)
    return c.transformed(lambda x, y: (cx + (x - cx) * math.cos(t) - (y - cy) * math.sin(t),
                                       cy + (x - cx) * math.sin(t) + (y - cy) * math.cos(t)))


# ── 아이콘 ────────────────────────────────────────────────
class Icon:
    def __init__(self, key, name, cat):
        self.key, self.name, self.cat = key, name, cat
        self.groups = [[]]

    def new(self):
        if self.groups[-1]:
            self.groups.append([])
        return self

    def fill(self, *cs):
        for c in cs:
            if isinstance(c, list):
                self.fill(*c)
            else:
                self.groups[-1].append(c.oriented(True))
        return self

    def cut(self, *cs):
        for c in cs:
            if isinstance(c, list):
                self.cut(*c)
            else:
                self.groups[-1].append(c.oriented(False))
        return self

    # SVG
    def svg_paths(self, fmt="{:.2f}"):
        out = []
        for g in self.groups:
            if not g:
                continue
            d = []
            for c in g:
                for s in c.segs:
                    if s[0] == "Z":
                        d.append("Z")
                    else:
                        d.append(s[0] + " ".join(fmt.format(v) for v in s[1:]))
            out.append(" ".join(d))
        return out

    def svg(self, color="#1F3A5F", size=48):
        paths = "".join(f'<path fill-rule="evenodd" d="{p}"/>' for p in self.svg_paths())
        return (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 48 48" width="{size}" height="{size}">'
                f'<title>{self.name}</title><g fill="{color}">{paths}</g></svg>')

    # DrawingML custGeom
    def custgeom(self, scale=1000):
        W = 48 * scale
        P = []
        for g in self.groups:
            if not g:
                continue
            seg = []
            for c in g:
                for s in c.segs:
                    if s[0] == "M":
                        seg.append(f'<a:moveTo><a:pt x="{round(s[1]*scale)}" y="{round(s[2]*scale)}"/></a:moveTo>')
                    elif s[0] == "L":
                        seg.append(f'<a:lnTo><a:pt x="{round(s[1]*scale)}" y="{round(s[2]*scale)}"/></a:lnTo>')
                    elif s[0] == "C":
                        seg.append('<a:cubicBezTo>' + "".join(
                            f'<a:pt x="{round(s[i]*scale)}" y="{round(s[i+1]*scale)}"/>' for i in (1, 3, 5)) + '</a:cubicBezTo>')
                    else:
                        seg.append('<a:close/>')
            P.append(f'<a:path w="{W}" h="{W}">' + "".join(seg) + '</a:path>')
        return ('<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/>'
                '<a:pathLst>' + "".join(P) + '</a:pathLst></a:custGeom>')
