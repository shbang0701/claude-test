# -*- coding: utf-8 -*-
"""카탈로그 슬라이드: 표지 · 사용법 · 기준 · 아이콘 · 구분 페이지."""
from pptx.enum.shapes import MSO_SHAPE
from ppt_kit import P, I, SW, text, shape, rect, rrect, oval, line, connect, group, icon, slide, table, notes
from ppt_elements import (LINES, LSTYLE, node_card, node_tile, pill, badge, tag, line_sample)
from icons import ICONS, CATS

CATNAME = dict(CATS)


def cover(prs):
    sl = slide(prs, None, layout=6, bg=P["navy"])
    text(sl, 0.8, 1.55, 10, 0.3, "IT INFRASTRUCTURE DOCUMENT KIT", 11, True, "AFC3DA", name="분류")
    text(sl, 0.8, 1.95, 11.5, 0.8, "IT 인프라 문서 구성요소 라이브러리", 36, True, P["white"], name="제목")
    text(sl, 0.8, 2.85, 11.5, 0.4, "구성도 · 운영 보고서 · 작업 계획서에 바로 쓰는 PPT 요소, 슬라이드 샘플, 아이콘 60종",
         15, False, "DCE6F2", name="부제")
    keys = ["server", "switch", "router", "firewall", "storage", "san-switch", "database", "cloud", "rack", "ups",
            "backup", "monitoring"]
    for k, key in enumerate(keys):
        icon(sl, key, 0.8 + k * 0.98, 4.25, 0.52, "DCE6F2")
    y = 5.35
    items = [("요소", "장비 · 영역 · 연결선 · 라벨 · 범례 · 표 · 비교 · 현황 · 절차"),
             ("샘플", "시스템 구성도 · 랙 실장도 · 물리 연결 · 월간 보고 · 작업 계획 · 변경 전후 · 장애 보고 · 점검"),
             ("호환", "16:9 · 맑은 고딕 · Office 2016 · 모든 도형·글자 편집 가능 · 색은 RGB 고정(다른 양식에 붙여도 유지)")]
    for k, (a, b) in enumerate(items):
        text(sl, 0.8, y + k * 0.34, 1.0, 0.26, a, 10.5, True, "AFC3DA")
        text(sl, 1.6, y + k * 0.34, 11, 0.26, b, 10.5, False, P["white"])
    return sl


def divider(prs, no, title, desc):
    sl = slide(prs, None, layout=6, bg=P["navy"])
    text(sl, 0.8, 2.6, 2, 0.5, f"{no:02d}", 28, True, "AFC3DA", name="번호")
    text(sl, 0.8, 3.2, 11, 0.7, title, 32, True, P["white"], name="제목")
    text(sl, 0.8, 4.0, 11, 0.8, desc, 13, False, "DCE6F2", name="설명", line_spacing=1.2)
    return sl


def howto(prs, toc):
    sl = slide(prs, "사용법 · 목차", "시작하기")
    text(sl, 0.5, 1.3, 6, 0.3, "필요한 부분만 떼어 쓰기", 14, True, P["ink"])
    steps = [
        ("요소 선택", "요소를 한 번 클릭하면 그룹 전체가 선택됩니다 → Ctrl+C"),
        ("붙여넣기", "대상 PPT에서 Ctrl+V. 색·글꼴이 직접 지정되어 회사 양식에 붙여도 모양이 그대로입니다"),
        ("글자 수정", "그룹 안 글자를 한 번 더 클릭해 바로 입력 (표는 셀을 클릭해 입력)"),
        ("크기 조절", "모서리를 Shift+드래그 → 비율 유지. 아이콘은 비율이 고정되어 찌그러지지 않습니다"),
        ("색 바꾸기", "도형 선택 → [도형 채우기]. 기존 자료 색에 맞출 때는 [스포이트]"),
        ("연결선", "선 끝을 도형 가장자리의 연결점에 놓으면 붙어서, 도형을 옮겨도 선이 따라갑니다"),
        ("찾기", "[홈 > 정렬 > 선택 창] (Alt+F10): '장비-', '영역-', '범례-', '라벨-' 이름으로 찾기"),
    ]
    for k, (a, b) in enumerate(steps):
        y = 1.75 + k * 0.52
        badge(sl, 0.5, y, k + 1, 0.28, fill=P["svc"])
        text(sl, 0.92, y - 0.02, 1.3, 0.3, a, 11, True, P["ink"], anchor="m")
        text(sl, 2.2, y - 0.02, 4.3, 0.46, b, 9.5, False, P["s700"], line_spacing=1.1)
    box = rrect(sl, 0.5, 5.55, 5.95, 1.2, 0.06, fill=P["t_info"], line=None, name="참고")
    icon(sl, "info", 0.68, 5.72, 0.3, P["info"])
    text(sl, 1.1, 5.7, 5.2, 1.0, [[("Office 2016 참고", {"bold": True, "color": P["info"]})],
                                  "SVG 삽입은 Microsoft 365(구독형)에서만 됩니다. Office 2016은 아이콘 슬라이드의 도형을 복사하거나 "
                                  "03_아이콘 폴더의 PNG를 쓰세요. 도형 아이콘은 확대해도 선명하고 색도 바꿀 수 있습니다."],
         9.5, False, P["ink"], line_spacing=1.15)
    text(sl, 7.0, 1.3, 5.8, 0.3, "목차", 14, True, P["ink"])
    rows = [["구역", "쪽", "내용"]] + toc
    table(sl, 7.0, 1.72, [1.15, 0.75, 3.95], rows, row_h=0.305, size=9.5, align=["l", "c", "l"],
          name="목차 표", bold_cols=(0,))
    return sl


def design(prs):
    sl = slide(prs, "디자인 기준", "기준")
    # 색
    text(sl, 0.5, 1.3, 4, 0.3, "색 — 연결 · 용도", 12, True, P["ink"])
    x = 0.5
    for key, nm, col, lw, d, desc in LINES:
        rrect(sl, x, 1.7, 1.42, 0.42, 0.05, fill=col, line=None, txt=nm, size=8.5, bold=True, color=P["white"],
              name=f"색-{nm}")
        text(sl, x, 2.16, 1.42, 0.18, f"#{col}", 8, False, P["s500"], align="c")
        x += 1.53
    text(sl, 0.5, 2.55, 4, 0.3, "색 — 기본 · 상태 · 영역 바탕", 12, True, P["ink"])
    base = [("잉크", "ink"), ("남색", "navy"), ("진회색", "s700"), ("회색", "s500"), ("테두리", "s300"),
            ("바탕", "s150"), ("연바탕", "s50")]
    x = 0.5
    for nm, k in base:
        rrect(sl, x, 2.95, 0.72, 0.36, 0.04, fill=P[k], line=P["s200"] if k in ("s50", "s150") else None,
              name=f"색-{nm}")
        text(sl, x, 3.34, 0.72, 0.18, nm, 8, True, P["ink"], align="c")
        text(sl, x, 3.5, 0.72, 0.18, f"#{P[k]}", 7, False, P["s500"], align="c")
        x += 0.78
    x += 0.15
    for nm, k, t in (("정상", "ok", "t_ok"), ("주의", "warn", "t_warn"), ("장애", "crit", "t_crit"), ("정보", "info", "t_info")):
        rrect(sl, x, 2.95, 0.72, 0.36, 0.04, fill=P[t], line=None, txt=nm, size=9, bold=True, color=P[k],
              name=f"상태-{nm}")
        text(sl, x, 3.34, 0.72, 0.18, f"#{P[k]}", 7, False, P["s500"], align="c")
        x += 0.78
    x += 0.15
    for nm, t in (("서비스", "t_svc"), ("관리", "t_mgmt"), ("백업", "t_bak"), ("SAN", "t_san"), ("DR", "t_ic")):
        rrect(sl, x, 2.95, 0.62, 0.36, 0.04, fill=P[t], line=P["s200"], lw=0.5, name=f"영역색-{nm}")
        text(sl, x, 3.34, 0.62, 0.18, nm, 8, True, P["ink"], align="c")
        text(sl, x, 3.5, 0.62, 0.18, f"#{P[t]}", 7, False, P["s500"], align="c")
        x += 0.67
    # 글자
    text(sl, 0.5, 3.95, 4, 0.3, "글자 — 모두 맑은 고딕", 12, True, P["ink"])
    samples = [("슬라이드 제목", 24, True), ("소제목", 14, True), ("본문 11–12pt", 11, False),
               ("장비 이름 10.5pt", 10.5, True), ("설명 · 표 9–9.5pt", 9, False), ("각주 8pt", 8, False)]
    x = 0.5
    for s, sz, b in samples:
        text(sl, x, 4.35, 2.2, 0.45, s, sz, b, P["ink"], anchor="b", wrap=False)
        text(sl, x, 4.84, 2.0, 0.2, f"{sz}pt{' 굵게' if b else ''}", 8, False, P["s500"])
        x += {24: 2.75, 14: 1.3, 11: 1.85, 10.5: 2.0, 9: 1.9, 8: 1.2}[sz]
    # 선·모서리·간격
    text(sl, 0.5, 5.35, 4, 0.3, "선 · 모서리 · 간격 · 아이콘 크기", 12, True, P["ink"])
    rules = [("테두리", "0.75pt  #B4BFCB"), ("영역", "1pt  용도 색"), ("연결선", "1.25–2.25pt (연결선 쪽 참고)"),
             ("모서리", "카드 0.06in · 영역 0.1in"), ("간격", "요소 사이 0.3in · 여백 0.5in"),
             ("아이콘", "라벨 0.22 · 카드 0.42 · 단독 0.56 · 강조 0.9in")]
    for k, (a, b) in enumerate(rules):
        cx = 0.5 + (k % 3) * 4.15
        cy = 5.75 + (k // 3) * 0.42
        text(sl, cx, cy, 1.0, 0.3, a, 10, True, P["s700"], anchor="m")
        text(sl, cx + 0.95, cy, 3.1, 0.3, b, 10, False, P["ink"], anchor="m")
    for k, sz in enumerate((0.22, 0.42, 0.56, 0.9)):
        pass
    return sl


def icon_slide(prs, cats, title, eyebrow="아이콘", extra=None):
    sl = slide(prs, title, eyebrow)
    y = 1.3
    for cat in cats:
        ics = [ic for ic in ICONS if ic.cat == cat]
        text(sl, 0.5, y, 6, 0.26, CATNAME[cat], 11, True, P["navy"])
        text(sl, 4.0, y + 0.03, 8.8, 0.22, f"{len(ics)}종 · 도형 채우기로 색 변경 · 파일: 03_아이콘/svg, png",
             8.5, False, P["s500"], align="r")
        y += 0.34
        per = 10
        for k, ic in enumerate(ics):
            cx = 0.5 + (k % per) * 1.23
            cy = y + (k // per) * 1.08
            idx = ICONS.index(ic) + 1
            ish = icon(sl, ic.key, cx + 0.31, cy, 0.56, P["navy"])
            t = text(sl, cx, cy + 0.62, 1.18, 0.2, ic.name, 8.5, True, P["ink"], align="c", wrap=False)
            k2 = text(sl, cx, cy + 0.8, 1.18, 0.16, f"{idx:02d} · {ic.key}", 7, False, P["s400"], align="c", wrap=False)
            group(sl, [ish, t, k2], f"아이콘 세트-{ic.name}")
        rows = (len(ics) + per - 1) // per
        y += rows * 1.08 + 0.1
    if extra:
        extra(sl, y)
    return sl


def icon_variants(sl, y):
    text(sl, 0.5, y, 6, 0.26, "스타일 변형 — 같은 아이콘, 쓰임에 맞게", 11, True, P["navy"])
    y += 0.34
    x = 0.5
    # 기본 / 용도색 / 타일 / 원 배지 / 윤곽 카드
    icon(sl, "server", x, y, 0.56, P["navy"])
    text(sl, x - 0.2, y + 0.62, 1.0, 0.2, "기본 (남색)", 8.5, False, P["s500"], align="c")
    x += 1.25
    for key, col in (("server", P["svc"]), ("storage", P["san"]), ("switch", P["mgmt"])):
        icon(sl, key, x, y, 0.56, col)
        x += 0.75
    text(sl, 1.75, y + 0.62, 2.2, 0.2, "용도 색", 8.5, False, P["s500"], align="c")
    x += 0.4
    for key, col in (("server", P["svc"]), ("storage", P["san"]), ("firewall", P["crit"])):
        g = [rrect(sl, x, y, 0.56, 0.56, 0.1, fill=col, line=None, name="타일"), icon(sl, key, x + 0.1, y + 0.1, 0.36, P["white"])]
        group(sl, g, "아이콘 타일")
        x += 0.75
    text(sl, 4.35, y + 0.62, 2.2, 0.2, "타일 (흰 아이콘)", 8.5, False, P["s500"], align="c")
    x += 0.4
    for key, col, t in (("server", P["svc"], P["t_svc"]), ("database", P["san"], P["t_san"]), ("cloud", P["mgmt"], P["t_mgmt"])):
        g = [oval(sl, x, y, 0.56, 0.56, fill=t, line=None, name="원"), icon(sl, key, x + 0.11, y + 0.11, 0.34, col)]
        group(sl, g, "아이콘 원 배지")
        x += 0.75
    text(sl, 6.95, y + 0.62, 2.2, 0.2, "원 배지 (옅은 바탕)", 8.5, False, P["s500"], align="c")
    x += 0.4
    for sz in (0.22, 0.42, 0.56, 0.9):
        icon(sl, "router", x, y + 0.56 - sz, sz, P["navy"])
        x += sz + 0.2
    text(sl, 9.6, y + 0.62, 3.2, 0.2, "크기: 0.22 · 0.42 · 0.56 · 0.9in", 8.5, False, P["s500"], align="c")
