# -*- coding: utf-8 -*-
"""재사용 요소: 각 함수는 도형들을 그리고 (그룹, 연결용 도형)을 돌려준다."""
from pptx.enum.shapes import MSO_SHAPE
from ppt_kit import (P, I, text, shape, rect, rrect, oval, line, connect, group, icon, set_font, line_fmt)

# 연결 종류: 이름, 색, 두께(pt), 점선, 설명
LINES = [
    ("svc", "서비스", P["svc"], 1.75, "solid", "업무 데이터 · 이더넷"),
    ("san", "스토리지 (SAN)", P["san"], 1.75, "solid", "FC · iSCSI · NFS"),
    ("mgmt", "관리 (OOB)", P["mgmt"], 1.25, "sysDash", "iLO · iDRAC · 콘솔"),
    ("bak", "백업", P["bak"], 1.25, "lgDashDot", "백업 전용망"),
    ("ha", "HA · 스택 · 피어링크", P["ic"], 1.25, "sysDot", "장비 간 동기·하트비트"),
    ("dr", "복제 · DR", P["ic"], 1.75, "lgDash", "센터 간 복제"),
    ("wan", "WAN · 인터넷", P["wan"], 2.25, "solid", "대외 · 전용회선"),
    ("pwr", "전원", P["pwr"], 1.25, "solid", "PDU · UPS 계통"),
]
LSTYLE = {k: (c, w, d) for k, _, c, w, d, _ in LINES}

ZONES = {
    "ext": ("인터넷 · 외부", P["t_ext"], P["s500"], "sysDash"),
    "dmz": ("DMZ", "FDF6E7", "C98A1A", "solid"),
    "svc": ("내부 서비스망", P["t_svc"], P["svc"], "solid"),
    "mgmt": ("관리망 (OOB)", P["t_mgmt"], P["mgmt"], "solid"),
    "bak": ("백업망", P["t_bak"], P["bak"], "solid"),
    "san": ("스토리지망 (SAN)", P["t_san"], P["san"], "solid"),
    "dr": ("DR 센터", P["t_ic"], P["ic"], "solid"),
    "cloud": ("클라우드 (VPC)", P["white"], P["svc"], "lgDash"),
    "phys": ("데이터센터 · 전산실", P["white"], P["s600"], "solid"),
    "group": ("클러스터 · 그룹", None, P["s400"], "sysDash"),
    "sec": ("보안 경계", None, P["crit"], "dash"),
}
KIND_COLOR = {"서버": P["svc"], "네트워크": P["mgmt"], "스토리지": P["san"], "보안": P["crit"], "전원": P["pwr"],
              "기타": P["s500"]}
KIND_TINT = {"서버": "DCE8F6", "네트워크": "D5EEEE", "스토리지": "FBE4D3", "보안": "F8DCDA", "전원": "EFE3F5",
             "기타": "E4E9EF", "빈칸": None}


# ── 장비 노드 ───────────────────────────────────────────────
def node_card(sl, x, y, key, name, sub="", color=None, w=2.1, h=0.62, style="outline", icon_color=None, gname=None):
    col = color or P["navy"]
    if style == "tint":
        bg = rrect(sl, x, y, w, h, 0.06, fill=P["s50"], line=P["s200"], lw=0.75)
    elif style == "solid":
        bg = rrect(sl, x, y, w, h, 0.06, fill=col, line=None)
    else:
        bg = rrect(sl, x, y, w, h, 0.06, fill=P["white"], line=P["s300"], lw=0.75)
    bg.name = "바탕"
    isz = min(0.42, h - 0.16)
    ic = icon(sl, key, x + 0.12, y + (h - isz) / 2, isz, P["white"] if style == "solid" else (icon_color or col))
    tc = P["white"] if style == "solid" else P["ink"]
    sc = "DCE6F2" if style == "solid" else P["s500"]
    tx = x + 0.12 + isz + 0.12
    if sub:
        t1 = text(sl, tx, y + h / 2 - 0.2, w - (tx - x) - 0.08, 0.2, name, 10.5, True, tc, anchor="b", wrap=False, name="이름")
        t2 = text(sl, tx, y + h / 2 + 0.02, w - (tx - x) - 0.08, 0.18, sub, 8.5, False, sc, wrap=False, name="설명")
        items = [bg, ic, t1, t2]
    else:
        t1 = text(sl, tx, y, w - (tx - x) - 0.08, h, name, 10.5, True, tc, anchor="m", wrap=False, name="이름")
        items = [bg, ic, t1]
    g = group(sl, items, gname or f"장비-{name}")
    return g, bg


def node_compact(sl, x, y, key, name, sub=None, size=0.56, color=None, w=1.3, gname=None):
    """아이콘 위 + 이름 아래 (좁은 구성도용). x,y = 아이콘 왼쪽 위."""
    ic = icon(sl, key, x, y, size, color or P["navy"])
    t1 = text(sl, x + size / 2 - w / 2, y + size + 0.04, w, 0.2, name, 9.5, True, P["ink"], align="c", name="이름")
    items = [ic, t1]
    if sub:
        items.append(text(sl, x + size / 2 - w / 2, y + size + 0.23, w, 0.18, sub, 8, False, P["s500"], align="c",
                          name="설명"))
    g = group(sl, items, gname or f"장비-{name}")
    return g, ic


def node_tile(sl, x, y, key, name, sub="", color=None, size=0.56, w=2.0, gname=None):
    col = color or P["svc"]
    t = rrect(sl, x, y, size, size, 0.1, fill=col, line=None, name="타일")
    ic = icon(sl, key, x + size * 0.17, y + size * 0.17, size * 0.66, P["white"])
    t1 = text(sl, x + size + 0.12, y + size / 2 - 0.2, w - size - 0.12, 0.2, name, 10.5, True, P["ink"], anchor="b",
              wrap=False, name="이름")
    items = [t, ic, t1]
    if sub:
        items.append(text(sl, x + size + 0.12, y + size / 2 + 0.02, w - size - 0.12, 0.18, sub, 8.5, False, P["s500"],
                          wrap=False, name="설명"))
    g = group(sl, items, gname or f"장비-{name}")
    return g, t


def node_box(sl, x, y, key, name, lines=(), color=None, w=1.55, h=1.05, gname=None):
    col = color or P["navy"]
    bg = rrect(sl, x, y, w, h, 0.07, fill=P["white"], line=P["s300"], lw=0.75, name="바탕")
    ic = icon(sl, key, x + 0.1, y + 0.1, 0.38, col)
    t1 = text(sl, x + 0.1, y + 0.52, w - 0.2, 0.2, name, 10, True, P["ink"], wrap=False, name="이름")
    items = [bg, ic, t1]
    for k, ln in enumerate(lines):
        items.append(text(sl, x + 0.1, y + 0.72 + k * 0.16, w - 0.2, 0.16, ln, 8, False, P["s500"], wrap=False,
                          name=f"줄{k+1}"))
    g = group(sl, items, gname or f"장비-{name}")
    return g, bg


def pill(sl, x, y, txt, color=None, fill=None, w=None, h=0.22, size=8, bold=True, tcolor=None, name=None):
    w = w or max(0.45, 0.075 * len(txt) * (size / 8) + 0.22)
    return rrect(sl, x, y, w, h, h / 2, fill=fill or P["white"], line=color or P["s400"], lw=0.75, txt=txt, size=size,
                 bold=bold, color=tcolor or color or P["s700"], margins=(0.05, 0.05, 0, 0), name=name or f"라벨-{txt}")


def badge(sl, x, y, n, d=0.26, fill=None, size=9.5, name=None):
    return oval(sl, x, y, d, d, fill=fill or P["navy"], line=None, txt=str(n), size=size, bold=True, color=P["white"],
                margins=(0, 0, 0, 0), name=name or f"번호-{n}")


def tag(sl, x, y, txt, kind="info", solid=False, w=None, h=0.24, size=8.5, name=None):
    col = {"ok": P["ok"], "warn": P["warn"], "crit": P["crit"], "info": P["info"], "mute": P["s500"],
           "new": P["ok"], "chg": P["warn"], "del": P["crit"], "keep": P["s500"]}[kind]
    tint = {"ok": P["t_ok"], "warn": P["t_warn"], "crit": P["t_crit"], "info": P["t_info"], "mute": P["s100"],
            "new": P["t_ok"], "chg": P["t_warn"], "del": P["t_crit"], "keep": P["s100"]}[kind]
    w = w or max(0.5, 0.11 * len(txt) + 0.2)
    if solid:
        return rrect(sl, x, y, w, h, 0.04, fill=col, line=None, txt=txt, size=size, bold=True, color=P["white"],
                     margins=(0.04, 0.04, 0, 0), name=name or f"태그-{txt}")
    return rrect(sl, x, y, w, h, 0.04, fill=tint, line=None, txt=txt, size=size, bold=True, color=col,
                 margins=(0.04, 0.04, 0, 0), name=name or f"태그-{txt}")


# ── 영역 ────────────────────────────────────────────────────
def zone(sl, x, y, w, h, kind="svc", label=None, tab=False, icon_key=None, gname=None):
    nm, f, c, d = ZONES[kind]
    label = label or nm
    lw = 2.0 if kind == "sec" else (1.25 if kind == "phys" else 1.0)
    bg = rrect(sl, x, y, w, h, 0.1, fill=f, line=c, lw=lw, dash=d, name="영역")
    items = [bg]
    if tab:
        tw = max(1.0, 0.13 * len(label) + 0.35)
        t = rrect(sl, x + 0.14, y - 0.14, tw, 0.28, 0.06, fill=c, line=None, txt=label, size=9, bold=True,
                  color=P["white"], name="영역 이름")
        items.append(t)
    else:
        lx = x + 0.14
        if icon_key:
            items.append(icon(sl, icon_key, lx, y + 0.1, 0.22, c))
            lx += 0.28
        items.append(text(sl, lx, y + 0.1, w - (lx - x) - 0.1, 0.22, label, 9.5, True, c, name="영역 이름"))
    g = group(sl, items, gname or f"영역-{label}")
    return g, bg


# ── 연결선 샘플 / 범례 ───────────────────────────────────────
def line_sample(sl, x, y, w, key, arrow=False):
    c, lw, d = LSTYLE[key]
    return line(sl, x, y, x + w, y, c, lw, d, tail="triangle" if arrow else None, name=f"선-{key}")


def line_legend(sl, x, y, keys=None, cols=1, colw=3.0, rowh=0.3, desc=True, gname="범례-연결선", title_txt="연결선"):
    keys = keys or [k for k, *_ in LINES]
    ws = list(colw) if isinstance(colw, (list, tuple)) else [colw] * cols     # 열마다 폭을 따로 줄 수 있음
    items = []
    if title_txt:
        items.append(text(sl, x, y, max(4.0, sum(ws)), 0.2, title_txt, 9, True, P["s700"], name="범례 제목"))
        y += 0.28
    for k, key in enumerate(keys):
        rec = [r for r in LINES if r[0] == key][0]
        c = k % cols
        cx = x + sum(ws[:c])
        cy = y + (k // cols) * rowh
        items.append(line_sample(sl, cx, cy + 0.1, 0.55, key))
        lab = rec[1] if not desc else [(rec[1], {"bold": True}), ("  " + rec[5], {"color": P["s500"], "size": 8})]
        items.append(text(sl, cx + 0.68, cy, ws[c] - 0.72, 0.2, [lab] if desc else lab, 9, False, P["ink"],
                          wrap=False, name=f"범례-{rec[1]}"))
    return group(sl, items, gname)


def color_legend(sl, x, y, items_, title_txt="구분", gname="범례-구분 색", sw=0.2, colw=1.25):
    items = []
    if title_txt:
        items.append(text(sl, x, y, 4, 0.2, title_txt, 9, True, P["s700"], name="범례 제목"))
    else:
        y -= 0.3
    for k, (nm, col) in enumerate(items_):
        cx = x + k * colw
        items.append(rrect(sl, cx, y + 0.3, sw, sw, 0.03, fill=col, line=None, name=f"칩-{nm}"))
        items.append(text(sl, cx + sw + 0.07, y + 0.3, colw - sw - 0.1, sw, nm, 9, False, P["ink"], anchor="m",
                          wrap=False, name=f"범례-{nm}"))
    return group(sl, items, gname)


# ── 현황 · 절차 ──────────────────────────────────────────────
def kpi(sl, x, y, label, value, unit="", delta=None, good=True, w=2.85, h=1.15, gname=None):
    bg = rrect(sl, x, y, w, h, 0.07, fill=P["white"], line=P["s200"], lw=0.75, name="바탕")
    t1 = text(sl, x + 0.18, y + 0.14, w - 0.3, 0.2, label, 9.5, False, P["s500"], name="항목")
    t2 = text(sl, x + 0.18, y + 0.38, w - 0.3, 0.45, [[(value, {"size": 26, "bold": True}),
                                                        (" " + unit, {"size": 11, "color": P["s500"]})]],
              26, True, P["ink"], anchor="b", name="값")
    items = [bg, t1, t2]
    if delta:
        items.append(text(sl, x + 0.18, y + 0.86, w - 0.3, 0.18, delta, 8.5, True, P["ok"] if good else P["crit"],
                          name="증감"))
    return group(sl, items, gname or f"KPI-{label}")


def usage_bar(sl, x, y, label, pct, w=5.2, color=None, warn=70, crit=85, gname=None):
    col = color or (P["crit"] if pct >= crit else P["warn"] if pct >= warn else P["svc"])
    lab = text(sl, x, y, 1.5, 0.22, label, 9.5, False, P["ink"], anchor="m", name="항목")
    tx = x + 1.55
    tw = w - 1.55 - 0.6
    track = rrect(sl, tx, y + 0.05, tw, 0.12, 0.06, fill=P["s150"], line=None, name="막대 바탕")
    fill_ = rrect(sl, tx, y + 0.05, max(0.12, tw * pct / 100), 0.12, 0.06, fill=col, line=None, name="막대 값")
    val = text(sl, tx + tw + 0.08, y, 0.52, 0.22, f"{pct}%", 9.5, True, col, anchor="m", align="r", name="값")
    return group(sl, [lab, track, fill_, val], gname or f"사용률-{label}")


def steps(sl, x, y, items_, w=12.3, h=0.55, color=None, gname="작업 단계"):
    n = len(items_)
    gap = 0.06
    sw = (w - gap * (n - 1)) / n
    shp = []
    for k, (ttl, desc) in enumerate(items_):
        sx = x + k * (sw + gap)
        prst = MSO_SHAPE.PENTAGON if k == 0 else MSO_SHAPE.CHEVRON
        s = shape(sl, prst, sx, y, sw, h, fill=color or P["navy"], line=None,
                  txt=[[(f"{k+1}  ", {"bold": True, "color": "AFC3DA"}), (ttl, {"bold": True})]], size=11,
                  color=P["white"], margins=(0.25 if k else 0.12, 0.2, 0, 0), name=f"단계{k+1}")
        s.adjustments[0] = 0.28
        shp.append(s)
        if desc:
            shp.append(text(sl, sx + (0.2 if k else 0.08), y + h + 0.12, sw - 0.3, 0.9, desc, 9, False, P["s700"],
                            name=f"단계{k+1} 설명", line_spacing=1.1))
    return group(sl, shp, gname)


def checkbox(sl, x, y, txt, checked=False, size=0.17, w=4.0, gname=None):
    b = rrect(sl, x, y + 0.02, size, size, 0.02, fill=P["svc"] if checked else P["white"],
              line=P["svc"] if checked else P["s400"], lw=1.0, name="체크 칸")
    items = [b]
    if checked:
        items.append(line(sl, x + size * 0.22, y + 0.02 + size * 0.52, x + size * 0.43, y + 0.02 + size * 0.74,
                          P["white"], 1.5, name="체크1"))
        items.append(line(sl, x + size * 0.43, y + 0.02 + size * 0.74, x + size * 0.8, y + 0.02 + size * 0.28,
                          P["white"], 1.5, name="체크2"))
    items.append(text(sl, x + size + 0.1, y, w - size - 0.1, 0.22, txt, 10, False, P["ink"], name="항목"))
    return group(sl, items, gname or f"체크-{txt[:10]}")


# ── 랙 ─────────────────────────────────────────────────────
def rack(sl, x, y, units=42, uh=0.118, w=2.0, devices=(), title_txt=None, gname=None, numbers=True):
    """devices: [(U하단, 높이, 이름, 구분)]"""
    H = units * uh
    items = []
    if title_txt:
        items.append(text(sl, x, y - 0.3, w + 0.4, 0.24, title_txt, 10, True, P["ink"], align="c", name="랙 이름"))
    rail = 0.22
    items.append(rect(sl, x, y, w + rail * 2, H + 0.06, fill=P["s100"], line=P["s700"], lw=1.25, name="랙 틀"))
    items.append(rect(sl, x + rail, y + 0.03, w, H, fill=P["white"], line=P["s400"], lw=0.5, name="랙 내부"))
    if numbers:
        for u in range(1, units + 1):
            uy = y + 0.03 + (units - u) * uh
            if u % 2 == 1 or units <= 24:
                items.append(text(sl, x, uy, rail - 0.03, uh, str(u), 5.5, False, P["s500"], align="r", anchor="m",
                                  wrap=False, name=f"U{u}"))
    for ub, hh, nm, kind in devices:
        dy = y + 0.03 + (units - (ub + hh - 1)) * uh
        items.append(rect(sl, x + rail, dy, w, hh * uh, fill=KIND_TINT.get(kind) or P["s150"], line=P["s600"], lw=0.5,
                          txt=nm, size=6.5 if hh == 1 else 8, bold=True, color=P["ink"], align="l",
                          margins=(0.06, 0.02, 0, 0), name=f"장비-{nm}"))
    return group(sl, items, gname or (f"랙-{title_txt}" if title_txt else "랙"))
