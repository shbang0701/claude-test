# -*- coding: utf-8 -*-
"""library/03_슬라이드_요소.pptx — 표현 유형별 슬라이드와 부품.
   모든 요소는 편집 가능한 도형·표·차트다(그림 아님)."""
import os, sys
HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
from pptx import Presentation
from pptx.util import Cm, Pt, Emu
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
import design as D
from pptx_lib import (W, H, M, CW, TOP, BOT, slide, box, rect, tbox, line_h, line_v, arrow,
                      node, zone, chip, pill, badge, dot, card, kpi, set_text, rich, rgb,
                      set_theme_fonts, _style, _apply_font, ptable, bar_chart, line_chart,
                      alpha, shp, conn, label, HEADS)

OUT = os.path.join(ROOT, "library", "03_슬라이드_요소.pptx")
ASSETS = os.path.join(ROOT, "assets")

def hdr(s, x, y, w, text, color=D.PRIMARY):
    tbox(s, x, y, w, 0.6, text, size=9.5, bold=True, color=color, anchor="t", margins=(0, 0, 0, 0))
    line_h(s, x, y + 0.62, w, D.LINE_SOFT, 1.0)

def lines_box(s, x, y, w, h, items, size=9.5, color=D.INK_SOFT, bullet="· ", gap=5, line=1.35):
    tb = tbox(s, x, y, w, h, "", anchor="t", margins=(0, 0, 0, 0))
    tf = tb.text_frame; tf.clear()
    for i, ln in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.line_spacing = line; p.space_after = Pt(gap)
        bold = ln.startswith("*")
        _apply_font(p.add_run(), size, bold, D.INK if bold else color)
        p.runs[0].text = (bullet if bullet else "") + ln.lstrip("*")
    return tb

def caption(s, x, y, w, text):
    return tbox(s, x, y, w, 0.7, text, size=9, color=D.INK_SOFT, anchor="t", margins=(0, 0, 0, 0))

def takeaway(s, y, text, label="핵심"):
    """장면의 결론 한 줄. 어느 슬라이드에나 붙여 쓴다."""
    b = box(s, M, y, CW, 1.5, D.PRIMARY_TINT, None, 0)
    rich(b, [(label + "   ", {"size": 10, "bold": True, "color": D.PRIMARY}),
             (text, {"size": 10.5, "color": D.INK})], align="l", anchor="ctr",
         margins=(0.6, 0.6, 0.1, 0.1))
    return b

# ══ 1. 사용 안내 ═══════════════════════════════════════════════════════════
def s_intro(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    rect(s, 0, 0, W, 0.26, D.PRIMARY, None)
    tbox(s, M, 2.5, 20, 0.8, "슬라이드 요소", size=9.5, bold=True, color=D.PRIMARY_MID,
         anchor="t", margins=(0, 0, 0, 0))
    tbox(s, M, 3.15, 26, 1.6, "내용만 바꿔 쓰는 장면 모음", size=29, bold=True,
         color=D.PRIMARY, anchor="t", margins=(0, 0, 0, 0))
    tbox(s, M, 5.05, 26, 0.9, "슬라이드를 통째로 복사하거나, 필요한 도형만 골라 기존 문서에 붙여 넣는다.",
         size=11.5, color=D.INK_SOFT, anchor="t", margins=(0, 0, 0, 0))
    line_h(s, M, 6.4, CW, D.LINE, 1.0)
    cols = [
        ("무엇이 들어 있나",
         ["*전하기   2 표지 · 3 섹션/목차 · 4 핵심과 근거 · 5 요약 카드",
          "*구성 보이기   6 구성도 · 7 계층과 관계",
          "*순서 보이기   8 단계 · 9 일정",
          "*견주기   10 전·후 · 11 선택지 · 12 두 축 비교",
          "*숫자 보이기   13 지표 · 14 막대 · 15 추세와 구성비",
          "*보여주기   16 화면·사진 주석 · 17 여러 장",
          "*정리하기   18 표 · 19 부품 · 20 마무리",
          "*부록   29 디자인 기준 · 30~31 아이콘 60종"]),
        ("어떻게 쓰나",
         ["*슬라이드 통째로: 왼쪽 목록에서 오른쪽 클릭 → 복사",
          "*도형만: Shift 클릭으로 여러 개 선택 후 복사",
          "*Word 에 붙일 때: 붙여넣기 옵션 → 원본 서식 유지",
          "글자는 도형을 두 번 눌러 바로 고친다",
          "묶인 도형은 Ctrl+Shift+G 로 해제한다",
          "차트는 오른쪽 클릭 → 데이터 편집 으로 숫자를 바꾼다",
          "회색 안내 글(아래쪽)은 복사한 뒤 지운다"]),
        ("맞춰 쓰는 규칙",
         ["한 장에 메시지는 하나만 담는다",
          "제목은 결론을 담은 문장으로 쓴다",
          "글꼴은 맑은 고딕 하나만 쓴다",
          "강조는 색보다 굵기·여백을 먼저 쓴다",
          "색은 주색 하나 + 상태색 세 가지 안에서",
          "도형 사이는 0.4cm 이상 띄운다",
          "글자는 9pt 아래로 줄이지 않는다"]),
    ]
    cw = (CW - 1.8) / 3
    for i, (t, items) in enumerate(cols):
        x = M + i * (cw + 0.9)
        hdr(s, x, 7.2, cw, t)
        lines_box(s, x, 8.05, cw, 6.2, items, size=9.5, gap=7)
    hdr(s, M, 14.3, CW, "색")
    sw = [("주색", D.PRIMARY), ("보조", D.PRIMARY_MID), ("정상", D.ST_OK), ("주의", D.ST_WARN),
          ("위험", D.ST_BAD), ("비활성", D.ST_IDLE), ("바탕", D.SURFACE), ("선", D.LINE)]
    for i, (nm, c) in enumerate(sw):
        x = M + i * 3.85
        rect(s, x, 15.2, 0.95, 0.95, c, D.LINE if c in (D.SURFACE, D.LINE) else None, 0.75)
        tbox(s, x + 1.2, 15.15, 2.5, 0.5, nm, size=9, bold=True, color=D.INK, anchor="t", margins=(0, 0, 0, 0))
        tbox(s, x + 1.2, 15.7, 2.5, 0.5, "#" + c, size=8, color=D.INK_FAINT, anchor="t", margins=(0, 0, 0, 0))
    tbox(s, M, 16.9, CW, 0.6, "슬라이드 16:9 · 맑은 고딕 · 본문 9.5~11pt · 제목 19pt",
         size=8.5, color=D.INK_FAINT, anchor="t", margins=(0, 0, 0, 0))
    return s

# ══ 2. 표지 ════════════════════════════════════════════════════════════════
def s_cover(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    rect(s, 0, 0, 1.0, H, D.PRIMARY, None)
    tbox(s, 3.2, 5.6, 24, 0.8, "〈문서 유형 — 보고 · 제안 · 안내 · 교육〉", size=10, bold=True,
         color=D.PRIMARY_MID, anchor="t", margins=(0, 0, 0, 0))
    tbox(s, 3.2, 6.5, 26, 2.4, "〈발표 제목을 한 줄로〉", size=32, bold=True, color=D.PRIMARY,
         anchor="t", margins=(0, 0, 0, 0))
    line_h(s, 3.2, 9.5, 8.0, D.PRIMARY_MID, 2.0)
    tbox(s, 3.2, 10.0, 24, 1.0, "〈부제 — 다루는 범위나 기간을 짧게〉", size=13, color=D.INK_SOFT,
         anchor="t", margins=(0, 0, 0, 0))
    tbox(s, 3.2, 13.4, 24, 1.6, "〈2026-01-01〉\n〈소속 · 발표자〉", size=11, color=D.INK_SOFT,
         anchor="t", line=1.5, margins=(0, 0, 0, 0))
    tbox(s, M, BOT + 0.35, CW, 0.6,
         "쓰는 법: 왼쪽 색 띠는 조직 색으로 바꿔 쓴다. 로고를 넣을 때는 오른쪽 아래를 비워 둔다.",
         size=8.5, color=D.INK_FAINT, anchor="t", margins=(0, 0, 0, 0))
    return s

# ══ 3. 섹션 · 목차 ═════════════════════════════════════════════════════════
def s_agenda(prs):
    s = slide(prs, "목차와 섹션 구분", "발표 흐름을 보여주고, 지금 어디인지 알려준다.",
              "쓰는 법: 진행 중인 항목만 진하게 둔다. 섹션 표지(오른쪽)는 장이 바뀔 때마다 복사해 번호와 제목만 바꾼다.")
    y = TOP + 0.6
    lw = CW * 0.54
    hdr(s, M, y, lw, "목차")
    items = [("01", "〈배경과 목적〉", True), ("02", "〈현황과 문제〉", False),
             ("03", "〈방안과 비교〉", False), ("04", "〈일정과 비용〉", False),
             ("05", "〈요청 사항〉", False)]
    for i, (num, t, cur) in enumerate(items):
        yy = y + 1.1 + i * 2.1
        c = D.PRIMARY if cur else D.INK_FAINT
        tbox(s, M, yy, 2.2, 1.2, num, size=17, bold=True, color=c, anchor="ctr", margins=(0, 0, 0, 0))
        tbox(s, M + 2.4, yy, lw - 2.4, 1.2, t, size=13, bold=cur, color=D.INK if cur else D.INK_SOFT,
             anchor="ctr", margins=(0, 0, 0, 0))
        line_h(s, M, yy + 1.35, lw, D.LINE_SOFT, 1.0)
        if cur: line_h(s, M, yy + 1.35, 2.2, D.PRIMARY, 2.0)
    RX = M + lw + 1.6; RW = CW - lw - 1.6
    hdr(s, RX, y, RW, "섹션 표지")
    sb = box(s, RX, y + 1.1, RW, 6.4, D.PRIMARY, None, 0)
    rich(sb, [("02\n", {"size": 22, "bold": True, "color": "8FB4CE"}),
              ("〈섹션 제목〉\n", {"size": 24, "bold": True, "color": D.WHITE}),
              ("〈이 장에서 답하려는 질문 한 줄〉", {"size": 11, "color": "C7D8E5"})],
         align="l", anchor="ctr", line=1.5, margins=(1.0, 0.8, 0.3, 0.3))
    hdr(s, RX, y + 8.2, RW, "마무리 장")
    eb = box(s, RX, y + 9.3, RW, 3.6, D.WHITE, D.LINE, 1.0)
    rich(eb, [("〈한 줄로 남길 결론〉\n", {"size": 15, "bold": True, "color": D.PRIMARY}),
              ("〈다음에 필요한 것 · 요청 사항〉", {"size": 10.5, "color": D.INK_SOFT})],
         align="l", anchor="ctr", line=1.5, margins=(0.8, 0.8, 0.2, 0.2))
    return s

# ══ 4. 핵심과 근거 ═════════════════════════════════════════════════════════
def s_message(prs):
    s = slide(prs, "〈여기에 결론을 문장으로 쓴다〉", "제목이 곧 메시지가 되게 한다. 아래는 그 근거 세 가지.",
              "쓰는 법: 근거는 세 개가 가장 읽기 좋다. 두 개면 카드를 지우고 남은 카드를 넓힌다.")
    y = TOP + 0.5
    b = box(s, M, y, CW, 3.0, D.PRIMARY_TINT, None, 0)
    rect(s, M, y, 0.18, 3.0, D.PRIMARY, None).line.fill.background()
    rich(b, [("〈가장 하고 싶은 말을 한 문장으로. 숫자가 있으면 문장 안에 넣는다.〉",
              {"size": 15, "bold": True, "color": D.PRIMARY})],
         align="l", anchor="ctr", margins=(0.9, 0.9, 0.2, 0.2))
    cw = (CW - 1.8) / 3
    for i, (t, lines) in enumerate([
            ("〈근거 1〉", ["〈사실이나 수치로 뒷받침한다〉", "〈출처나 기준을 덧붙인다〉",
                          "〈반론이 있다면 함께 적는다〉"]),
            ("〈근거 2〉", ["〈사실이나 수치〉", "〈출처나 기준〉", "〈덧붙일 설명〉"]),
            ("〈근거 3〉", ["〈사실이나 수치〉", "〈출처나 기준〉", "〈덧붙일 설명〉"])]):
        x = M + i * (cw + 0.9)
        card(s, x, y + 4.0, cw, 6.8, t, lines, accent=D.PRIMARY_MID)
        badge(s, x + cw - 1.5, y + 3.6, 0.9, str(i + 1))
    takeaway(s, y + 11.5, "〈그래서 무엇을 하자는 것인지. 요청이나 결정 사항을 적는다.〉", "그래서")
    return s

# ══ 5. 요약 카드 ═══════════════════════════════════════════════════════════
def s_summary(prs):
    s = slide(prs, "〈대상〉 한눈에 보기", "소개·개요·현황 요약에 쓴다.",
              "쓰는 법: 카드 수는 3~4개. 제목은 명사로, 설명은 한 줄로 맞춘다.")
    y = TOP + 0.5
    cw = (CW - 2.7) / 4
    data = [("〈무엇인가〉", ["〈한 문장 정의〉", "〈핵심 특징 하나〉", "〈비슷한 것과의 차이〉"], D.PRIMARY),
            ("〈왜 필요한가〉", ["〈해결하는 문제〉", "〈기대 효과〉", "〈근거 수치〉"], D.PRIMARY_MID),
            ("〈어떻게 쓰나〉", ["〈사용 방법 요약〉", "〈필요 조건〉", "〈걸리는 시간〉"], D.PRIMARY_MID),
            ("〈무엇을 조심하나〉", ["〈제약 사항〉", "〈예외 상황〉", "〈문의처〉"], D.ST_WARN)]
    for i, (t, lines, c) in enumerate(data):
        card(s, M + i * (cw + 0.9), y, cw, 7.2, t, lines, accent=c)
    hdr(s, M, y + 8.1, CW, "한 줄 정리")
    ptable(s, M, y + 9.0, [CW * 0.2, CW * 0.55, CW * 0.25],
           ["구분", "내용", "비고"],
           [["〈항목〉", "〈설명을 한 줄로〉", "〈덧붙임〉"],
            ["〈항목〉", "〈설명을 한 줄로〉", ""],
            ["〈항목〉", "〈설명을 한 줄로〉", ""],
            ["〈항목〉", "〈설명을 한 줄로〉", ""]], row_h=1.15, head_h=1.1)
    return s

# ══ 6. 구성 ════════════════════════════════════════════════════════════════
def s_structure(prs):
    s = slide(prs, "〈대상〉 구성", "무엇으로 이루어져 있고 어떻게 이어지는지 보여준다.",
              "쓰는 법: 영역 상자를 먼저 놓고 맨 뒤로 보낸 뒤 요소를 올린다. 요소를 늘릴 때는 같은 줄의 상자를 복사한다.")
    y = TOP + 0.4
    bands = [("〈영역 1〉", D.PRIMARY_MID, [("〈구성 요소〉", "〈역할 한 줄〉"), ("〈구성 요소〉", "〈역할 한 줄〉"),
                                            ("〈구성 요소〉", "〈역할 한 줄〉")]),
             ("〈영역 2〉", D.PRIMARY, [("〈구성 요소〉", "〈역할 한 줄〉"), ("〈구성 요소〉", "〈역할 한 줄〉"),
                                        ("〈구성 요소〉", "〈역할 한 줄〉")]),
             ("〈영역 3〉", D.PRIMARY_DEEP, [("〈구성 요소〉", "〈역할 한 줄〉"), ("〈구성 요소〉", "〈역할 한 줄〉")])]
    bh, gapb = 3.5, 1.2
    by = y
    for bi, (zname, zc, items) in enumerate(bands):
        zone(s, M, by, CW, bh, zname, color=zc, fill=None, label_pos="l")
        lx = M + 3.6; avail = CW - 4.1
        n = len(items)
        nw = (avail - 0.8 * (n - 1)) / n
        for i, (t, sub) in enumerate(items):
            node(s, lx + i * (nw + 0.8), by + 0.85, nw, 1.95, t, sub, accent=zc)
            if i < n - 1:
                arrow(s, lx + i * (nw + 0.8) + nw + 0.06, by + 1.82,
                      lx + (i + 1) * (nw + 0.8) - 0.06, by + 1.82, D.INK_FAINT, 1.25, head=False)
        if bi < len(bands) - 1:
            ax = lx + avail * 0.5
            arrow(s, ax, by + bh, ax, by + bh + gapb, D.PRIMARY_MID, 1.75)
            chip(s, ax + 0.4, by + bh + 0.22, 3.0, 0.62, "〈관계 · 조건〉", D.INK_FAINT)
        by += bh + gapb
    takeaway(s, by + 0.2, "〈이 구성에서 읽는 사람이 알아야 할 한 가지를 적는다.〉", "읽는 법")
    return s

# ══ 7. 계층과 관계 ═════════════════════════════════════════════════════════
def s_tree(prs):
    s = slide(prs, "계층과 관계", "위아래 관계는 왼쪽, 주고받는 관계는 오른쪽 모양으로 그린다.",
              "쓰는 법: 가지를 늘릴 때는 상자와 선을 함께 복사한다. 선은 ‘꺾인 연결선’을 쓰면 정리가 쉽다.")
    y = TOP + 0.5
    lw = CW * 0.52
    hdr(s, M, y, lw, "계층 — 조직 · 분류 · 구성")
    cx = M + lw / 2
    node(s, cx - 3.4, y + 1.3, 6.8, 1.8, "〈상위 항목〉", "〈전체를 아우르는 이름〉", accent=D.PRIMARY)
    for i, t in enumerate(["〈중간 항목 A〉", "〈중간 항목 B〉"]):
        x = M + 0.4 + i * (lw / 2 + 0.2)
        node(s, x, y + 4.7, lw / 2 - 0.8, 1.8, t, "〈역할〉", accent=D.PRIMARY_MID)
        arrow(s, cx, y + 3.1, x + (lw / 2 - 0.8) / 2, y + 4.7, D.LINE, 1.25, head=False,
              kind=MSO_CONNECTOR.ELBOW)
    for i in range(4):
        x = M + 0.4 + (i // 2) * (lw / 2 + 0.2) + (i % 2) * ((lw / 2 - 0.8) / 2 + 0.15)
        w = (lw / 2 - 0.8) / 2 - 0.15
        node(s, x, y + 8.3, w, 1.7, "〈하위 항목〉", "〈역할〉", accent=D.INK_FAINT, tsize=9.5, ssize=8)
        px = M + 0.4 + (i // 2) * (lw / 2 + 0.2) + (lw / 2 - 0.8) / 2
        arrow(s, px, y + 6.5, x + w / 2, y + 8.3, D.LINE, 1.25, head=False, kind=MSO_CONNECTOR.ELBOW)
    lines_box(s, M, y + 10.6, lw, 2.6,
              ["같은 층은 높이를 맞춘다", "한 층에 여섯 개가 넘으면 묶어서 단계를 하나 더 만든다",
               "설명은 상자 안이 아니라 아래 표로 뺀다"], size=9.5)
    RX = M + lw + 1.4; RW = CW - lw - 1.4
    hdr(s, RX, y, RW, "관계 — 연결 · 의존 · 주고받음")
    ccx, ccy = RX + RW / 2, y + 5.3
    hub = box(s, ccx - 2.8, ccy - 1.1, 5.6, 2.2, D.PRIMARY_TINT, D.PRIMARY, 1.5)
    set_text(hub, "〈가운데 대상〉", size=11.5, bold=True, color=D.PRIMARY, align="c")
    for dx, dy, t, sub in [(-1, -1, "〈상대 A〉", "〈주고받는 것〉"), (1, -1, "〈상대 B〉", "〈주고받는 것〉"),
                           (-1, 1, "〈상대 C〉", "〈주고받는 것〉"), (1, 1, "〈상대 D〉", "〈주고받는 것〉")]:
        w, h = 4.8, 1.6
        x = ccx + dx * (RW / 2 - w / 2 - 0.2) - w / 2
        yy = ccy + dy * 4.1 - h / 2
        node(s, x, yy, w, h, t, sub, accent=D.INK_FAINT, tsize=9.5, ssize=8)
        arrow(s, ccx + dx * 1.7, ccy + dy * 1.1, x + w / 2 - dx * 0.4, yy + (h if dy < 0 else 0),
              D.INK_FAINT, 1.25, head=True, tail=True)
    lines_box(s, RX, ccy + 5.6, RW, 2.6,
              ["선 위에 주고받는 내용과 주기를 적는다", "한쪽으로만 흐르면 화살표를 한 방향으로 그린다",
               "상대가 여섯을 넘으면 표로 정리한다"], size=9.5)
    return s

# ══ 8. 단계 ════════════════════════════════════════════════════════════════
def s_steps(prs):
    s = slide(prs, "〈과정 이름〉 진행 단계", "단계별로 무엇을 하고 무엇이 나오는지 보여준다.",
              "쓰는 법: 화살표 도형을 복사해 이어 붙인다. 단계는 3~5개일 때 가장 읽기 좋다.")
    y = TOP + 0.6
    steps = [("〈준비〉", "〈이 단계에서 하는 일〉", "〈나오는 것〉"),
             ("〈진행〉", "〈이 단계에서 하는 일〉", "〈나오는 것〉"),
             ("〈확인〉", "〈이 단계에서 하는 일〉", "〈나오는 것〉"),
             ("〈마무리〉", "〈이 단계에서 하는 일〉", "〈나오는 것〉")]
    n = len(steps); ov = 0.8
    bw = (CW + ov * (n - 1)) / n
    for i, (t, d, out) in enumerate(steps):
        x = M + i * (bw - ov)
        sh = s.shapes.add_shape(MSO_SHAPE.CHEVRON, Cm(x), Cm(y), Cm(bw), Cm(2.4))
        done = i < 2
        _style(sh, D.PRIMARY if i == 0 else (D.PRIMARY_MID if i == 1 else D.WHITE),
               None if done else D.PRIMARY_MID, 1.0)
        rich(sh, [(f"{i+1}. {t}", {"size": 13, "bold": True, "color": D.WHITE if done else D.PRIMARY})],
             align="c", anchor="ctr", margins=(0.7, 0.45, 0.05, 0.05))
        lines_box(s, x + 1.0, y + 2.9, bw - 1.6, 3.0, ["*" + d, "결과물: " + out], size=9.5, bullet="", gap=6)
    hdr(s, M, y + 6.3, CW, "단계별 기간과 담당")
    ptable(s, M, y + 7.2, [CW * 0.16, CW * 0.2, CW * 0.18, CW * 0.46],
           ["단계", "기간", "담당", "끝났다고 보는 기준"],
           [[f"{i}. 〈단계 이름〉", "〈MM/DD ~ MM/DD〉", "〈이름〉", "〈이 상태가 되면 완료〉"]
            for i in range(1, 5)], row_h=1.2, head_h=1.1)
    return s

# ══ 9. 일정 ════════════════════════════════════════════════════════════════
def s_timeline(prs):
    s = slide(prs, "〈일정 이름〉", "기간과 겹침, 중요한 시점을 한 장에 보여준다.",
              "쓰는 법: 막대는 직사각형이다. 좌우 끝을 끌어 기간을 맞추고 눈금 글자만 바꾼다.")
    y = TOP + 0.9
    ticks = ["〈1월〉", "〈2월〉", "〈3월〉", "〈4월〉", "〈5월〉", "〈6월〉"]
    x0 = M + 5.4; span = CW - 6.8
    line_h(s, x0, y + 0.95, span, D.LINE, 1.0)
    for i, t in enumerate(ticks):
        x = x0 + span * i / (len(ticks) - 1)
        line_v(s, x, y + 0.75, 0.4, D.LINE, 1.0)
        tbox(s, x - 1.1, y, 2.2, 0.6, t, size=8.5, color=D.INK_FAINT, align="c", anchor="t",
             margins=(0, 0, 0, 0))
    bars = [("〈준비 단계〉", "〈담당 A〉", 0.0, 0.22, D.PRIMARY_MID),
            ("〈본 작업〉", "〈담당 B〉", 0.18, 0.62, D.PRIMARY),
            ("〈검토·확인〉", "〈담당 A〉", 0.55, 0.8, D.PRIMARY_MID),
            ("〈정리·보고〉", "〈담당 C〉", 0.78, 1.0, D.INK_FAINT)]
    by = y + 1.9
    for nm, who, a, bb, c in bars:
        rich(tbox(s, M, by - 0.05, 5.2, 1.0, "", anchor="ctr", margins=(0, 0, 0, 0)),
             [(nm + "\n", {"size": 9.5, "bold": True, "color": D.INK}),
              (who, {"size": 8.5, "color": D.INK_FAINT})], align="l", anchor="ctr", line=1.2,
             margins=(0, 0, 0, 0))
        bar = rect(s, x0 + span * a, by + 0.1, span * (bb - a), 0.8, c, None)
        bar.line.fill.background()
        by += 1.45
    for pos, lb in ((0.18, "〈착수〉"), (0.62, "〈중간 점검〉"), (1.0, "〈완료〉")):
        x = x0 + span * pos
        line_v(s, x, y + 1.6, by - y - 1.5, D.ST_WARN, 1.0, dash="dash")
        chip(s, x - 1.5, by + 0.05, 3.0, 0.66, lb, D.ST_WARN)
    hdr(s, M, by + 1.3, CW, "요약")
    items = [("전체 기간", "〈6개월〉", D.PRIMARY), ("가장 긴 구간", "〈본 작업 3개월〉", D.PRIMARY_MID),
             ("겹치는 구간", "〈2~3월〉", D.ST_WARN), ("중요한 시점", "〈3월 중간 점검〉", D.ST_BAD)]
    kw = (CW - 2.4) / 4
    for i, (t, v, c) in enumerate(items):
        kpi(s, M + i * (kw + 0.8), by + 2.15, kw, 2.3, t, v, color=c, vsize=15)
    return s

# ══ 10. 전 · 후 ════════════════════════════════════════════════════════════
def s_before_after(prs):
    s = slide(prs, "〈무엇이〉 어떻게 달라지는가", "지금과 바뀐 뒤를 나란히 둔다.",
              "쓰는 법: 좌우 항목 수를 맞춘다. 화면이나 사진을 넣을 때는 16쪽 주석 요소를 함께 쓴다.")
    cw = (CW - 3.2) / 2
    y = TOP + 0.4
    for i, (t, color, fill, items) in enumerate([
        ("지금", D.INK_FAINT, D.SURFACE,
         ["〈현재 방식이나 상태〉", "〈그래서 생기는 문제〉", "〈영향을 받는 사람〉", "〈드는 시간·비용〉"]),
        ("바뀐 뒤", D.PRIMARY, D.PRIMARY_TINT,
         ["〈바뀌는 방식이나 상태〉", "〈해결되는 것〉", "〈달라지는 사용 방법〉", "〈줄어드는 시간·비용〉"])]):
        x = M + i * (cw + 3.2)
        box(s, x, y, cw, 9.6, D.WHITE, D.LINE, 1.0)
        head = rect(s, x, y, cw, 1.25, fill, None); head.line.fill.background()
        set_text(head, t, size=12.5, bold=True, color=color, align="c")
        lines_box(s, x + 0.8, y + 1.9, cw - 1.6, 7.2, items, size=10.5, color=D.INK, gap=14)
    ax = M + cw + 0.6
    a = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Cm(ax), Cm(y + 4.1), Cm(2.0), Cm(1.4))
    _style(a, D.PRIMARY_MID, None, 0)
    kw = (CW - 2.4) / 4
    for i, (t, v) in enumerate([("무엇을 바꾸나", "〈한 문장으로〉"), ("왜 바꾸나", "〈근거·배경〉"),
                                ("언제부터", "〈적용 시점〉"), ("해야 할 일", "〈사용자가 할 일〉")]):
        bx = M + i * (kw + 0.8)
        b = box(s, bx, y + 10.4, kw, 2.3, D.WHITE, D.LINE, 1.0)
        rich(b, [(t + "\n", {"size": 9, "color": D.INK_FAINT}),
                 (v, {"size": 10.5, "bold": True, "color": D.INK})],
             align="l", anchor="ctr", margins=(0.4, 0.3, 0.1, 0.1), line=1.4)
    return s

# ══ 11. 선택지 비교 ════════════════════════════════════════════════════════
def s_options(prs):
    s = slide(prs, "〈무엇을〉 고를 것인가", "같은 기준으로 견주고 고른 이유를 남긴다.",
              "쓰는 법: 고른 안만 진한 테두리로 둔다. 기준이 늘면 아래 표에 행을 추가한다.")
    y = TOP + 0.4
    cw = (CW - 1.8) / 3
    opts = [("〈선택지 A〉", "〈한 줄 성격〉", False, ["〈좋은 점〉", "〈걱정되는 점〉"]),
            ("〈선택지 B〉", "〈한 줄 성격〉", True, ["〈좋은 점〉", "〈걱정되는 점〉"]),
            ("〈선택지 C〉", "〈한 줄 성격〉", False, ["〈좋은 점〉", "〈걱정되는 점〉"])]
    for i, (t, sub, pick, pros) in enumerate(opts):
        x = M + i * (cw + 0.9)
        box(s, x, y, cw, 6.0, D.WHITE, D.PRIMARY if pick else D.LINE, 1.75 if pick else 1.0)
        top = rect(s, x, y, cw, 0.14, D.PRIMARY if pick else D.LINE, None); top.line.fill.background()
        rich(tbox(s, x + 0.7, y + 0.6, cw - 1.4, 1.5, "", anchor="t", margins=(0, 0, 0, 0)),
             [(t + "\n", {"size": 13, "bold": True, "color": D.INK}),
              (sub, {"size": 10, "color": D.INK_SOFT})], align="l", anchor="t", line=1.4,
             margins=(0, 0, 0, 0))
        if pick: chip(s, x + cw - 2.9, y + 0.6, 2.2, 0.66, "선택", D.PRIMARY)
        lines_box(s, x + 0.7, y + 2.5, cw - 1.4, 2.4, pros, size=9.5, gap=8)
        line_h(s, x + 0.7, y + 4.7, cw - 1.4, D.LINE_SOFT, 1.0)
        rich(tbox(s, x + 0.7, y + 4.9, cw - 1.4, 0.9, "", anchor="ctr", margins=(0, 0, 0, 0)),
             [("〈비용〉", {"size": 9.5, "color": D.INK_SOFT}),
              ("     〈기간〉", {"size": 9.5, "color": D.INK_SOFT})], align="l", anchor="ctr",
             margins=(0, 0, 0, 0))
    hdr(s, M, y + 6.6, CW, "같은 기준으로 보기")
    ptable(s, M, y + 7.5, [CW * 0.22, CW * 0.26, CW * 0.26, CW * 0.26],
           ["기준", "선택지 A", "선택지 B", "선택지 C"],
           [["〈기준 1〉", "〈값〉", "〈값〉", "〈값〉"],
            ["〈기준 2〉", "〈값〉", "〈값〉", "〈값〉"],
            ["〈기준 3〉", "〈값〉", "〈값〉", "〈값〉"],
            ["〈기준 4〉", "〈값〉", "〈값〉", "〈값〉"]], row_h=1.0, head_h=1.05)
    takeaway(s, y + 12.6, "〈무엇을 가장 중요하게 봤고, 그래서 무엇을 골랐는지 한 문장으로.〉", "고른 이유")
    return s

# ══ 12. 두 축 비교(매트릭스) ═══════════════════════════════════════════════
def s_matrix(prs):
    s = slide(prs, "〈두 기준〉으로 나눠 보기", "우선순위·난이도·중요도처럼 기준이 둘일 때.",
              "쓰는 법: 축 이름을 먼저 정한다. 항목은 사각형 안에 글머리로 적거나 작은 상자로 올린다.")
    y = TOP + 0.5
    size = 11.4
    x0 = M + 2.6
    quad = [("〈먼저 한다〉", D.PRIMARY_TINT, D.PRIMARY, ["〈항목〉", "〈항목〉"]),
            ("〈계획해서 한다〉", D.WHITE, D.PRIMARY_MID, ["〈항목〉", "〈항목〉"]),
            ("〈맡기거나 줄인다〉", D.WHITE, D.INK_FAINT, ["〈항목〉"]),
            ("〈하지 않는다〉", D.SURFACE, D.INK_FAINT, ["〈항목〉"])]
    hw = size / 2
    for i, (t, fl, c, items) in enumerate(quad):
        qx = x0 + (i % 2) * hw
        qy = y + (i // 2) * hw
        b = rect(s, qx, qy, hw, hw, fl, D.LINE, 1.0)
        tbox(s, qx + 0.5, qy + 0.45, hw - 1.0, 0.8, t, size=11, bold=True, color=c, anchor="t",
             margins=(0, 0, 0, 0))
        lines_box(s, qx + 0.5, qy + 1.5, hw - 1.0, hw - 2.0, items, size=9.5, gap=6)
    tbox(s, x0 - 2.5, y, 2.2, size, "〈세로축 —\n높음 ↑\n낮음 ↓〉", size=9.5, bold=True, color=D.INK_SOFT,
         align="r", anchor="ctr", line=1.6, margins=(0, 0, 0, 0))
    tbox(s, x0, y + size + 0.2, size, 0.8, "〈가로축 — 낮음  ←                   →  높음〉",
         size=9.5, bold=True, color=D.INK_SOFT, align="c", anchor="t", margins=(0, 0, 0, 0))
    RX = x0 + size + 1.6; RW = W - M - RX
    hdr(s, RX, y, RW, "이 그림을 읽는 법")
    lines_box(s, RX, y + 0.95, RW, 4.2,
              ["두 기준은 서로 독립적인 것으로 고른다",
               "항목은 한 칸에 셋을 넘기지 않는다",
               "칸 이름은 ‘무엇을 한다’처럼 행동으로 적는다"], size=10, gap=8)
    hdr(s, RX, y + 5.4, RW, "칸별 정리")
    ptable(s, RX, y + 6.3, [RW * 0.4, RW * 0.6],
           ["구분", "무엇을 할 것인가"],
           [["〈먼저 한다〉", "〈지금 착수〉"], ["〈계획해서 한다〉", "〈일정에 반영〉"],
            ["〈맡기거나 줄인다〉", "〈담당 조정〉"], ["〈하지 않는다〉", "〈보류·제외〉"]],
           row_h=1.1, head_h=1.05, size=9.5)
    takeaway(s, y + 12.0, "〈이 분류로 결정한 것을 한 문장으로 적는다.〉", "결론")
    return s

# ══ 13. 지표 ═══════════════════════════════════════════════════════════════
def s_kpi(prs):
    s = slide(prs, "〈대상〉 주요 수치", "〈2026-01-01 기준〉 숫자와 그 뜻을 함께 보여준다.",
              "쓰는 법: 숫자에는 반드시 비교 기준(목표·전월·작년)을 붙인다. 화살표는 좋아짐·나빠짐을 뜻한다.")
    y = TOP + 0.4
    kw = (CW - 2.4) / 4
    data = [("〈지표 이름〉", "〈1,240〉", "건", D.PRIMARY, "▲ 8%", D.ST_OK, "〈전월 1,148건〉"),
            ("〈지표 이름〉", "〈98.2〉", "%", D.ST_OK, "▲ 1.2%p", D.ST_OK, "〈목표 95%〉"),
            ("〈지표 이름〉", "〈2.4〉", "일", D.PRIMARY, "▼ 0.5일", D.ST_OK, "〈전월 2.9일〉"),
            ("〈지표 이름〉", "〈12〉", "건", D.ST_BAD, "▲ 4건", D.ST_BAD, "〈목표 10건 이하〉")]
    for i, (t, v, u, c, delta, dc, sub) in enumerate(data):
        x = M + i * (kw + 0.8)
        kpi(s, x, y, kw, 3.4, t, v, u, color=c, sub=sub, vsize=26)
        pill(s, x + kw - 3.0, y + 0.45, 2.6, 0.7, delta, dc, D.OK_BG if dc == D.ST_OK else D.DANGER_BG)
    hdr(s, M, y + 4.2, CW * 0.48, "수치의 뜻")
    lines_box(s, M, y + 5.1, CW * 0.48, 5.0,
              ["*〈지표 이름〉  〈무엇을 세는 숫자인지 한 줄로〉",
               "*〈지표 이름〉  〈어떻게 계산하는지〉",
               "*〈지표 이름〉  〈제외하는 대상이 있다면〉",
               "*〈지표 이름〉  〈기준일과 출처〉"], size=9.5, bullet="", gap=9)
    RX = M + CW * 0.52; RW = CW - CW * 0.52
    hdr(s, RX, y + 4.2, RW, "이번에 달라진 것")
    for i, (t, d, c) in enumerate([("〈늘어난 것〉", "〈왜 늘었는지 한 줄〉", D.ST_OK),
                                   ("〈줄어든 것〉", "〈왜 줄었는지 한 줄〉", D.PRIMARY_MID),
                                   ("〈지켜볼 것〉", "〈무엇을 언제까지 볼 것인지〉", D.ST_WARN)]):
        yy = y + 5.1 + i * 2.2
        b = box(s, RX, yy, RW, 1.9, D.WHITE, D.LINE, 1.0)
        rect(s, RX, yy, 0.13, 1.9, c, None).line.fill.background()
        rich(b, [(t + "\n", {"size": 10.5, "bold": True, "color": D.INK}),
                 (d, {"size": 9.5, "color": D.INK_SOFT})],
             align="l", anchor="ctr", margins=(0.45, 0.3, 0.1, 0.1), line=1.35)
    takeaway(s, y + 12.0, "〈숫자에서 읽어야 할 한 가지. 다음에 무엇을 할지까지.〉", "읽는 법")
    return s

# ══ 14. 막대 ═══════════════════════════════════════════════════════════════
def s_bar(prs):
    s = slide(prs, "〈무엇이〉 얼마나 되는가", "항목별 크기 비교. 값이 바뀌면 차트에서 숫자만 고친다.",
              "쓰는 법: 차트를 누르고 오른쪽 클릭 → 데이터 편집. 항목이 많으면 큰 것부터 정렬한다.")
    y = TOP + 0.5
    bar_chart(s, M, y, CW * 0.62, 11.0,
              ["〈항목 A〉", "〈항목 B〉", "〈항목 C〉", "〈항목 D〉", "〈항목 E〉"],
              [("〈올해〉", (42, 35, 28, 19, 11)), ("〈지난해〉", (36, 33, 30, 15, 9))],
              colors=[D.PRIMARY, "A8C3D8"], gap=70)
    RX = M + CW * 0.64; RW = CW - CW * 0.64
    hdr(s, RX, y, RW, "무엇을 보여주는 그림인가")
    lines_box(s, RX, y + 0.95, RW, 4.4,
              ["〈어느 항목이 가장 큰지〉", "〈작년과 견주어 달라진 항목〉", "〈예상과 달랐던 부분〉"],
              size=10, gap=9)
    hdr(s, RX, y + 5.6, RW, "숫자로 확인")
    ptable(s, RX, y + 6.5, [RW * 0.44, RW * 0.28, RW * 0.28],
           ["항목", "올해", "증감"],
           [["〈항목 A〉", "〈42〉", "〈+6〉"], ["〈항목 B〉", "〈35〉", "〈+2〉"],
            ["〈항목 C〉", "〈28〉", "〈-2〉"]], row_h=1.05, head_h=1.05, size=9.5,
           align=["l", "r", "r"])
    takeaway(s, y + 11.5, "〈이 그림에서 말하려는 한 가지. 숫자를 반복하지 말고 뜻을 적는다.〉", "읽는 법")
    return s

# ══ 15. 추세와 구성비 ══════════════════════════════════════════════════════
def s_trend(prs):
    s = slide(prs, "〈무엇이〉 어떻게 변했는가", "시간에 따른 흐름은 선으로, 안을 이루는 비율은 띠로 본다.",
              "쓰는 법: 두 차트 모두 오른쪽 클릭 → 데이터 편집으로 값을 바꾼다. 축 단위는 제목이나 축 이름에 밝힌다.")
    y = TOP + 0.5
    hw = (CW - 1.6) / 2
    hdr(s, M, y, hw, "흐름 — 시간에 따른 변화")
    line_chart(s, M, y + 0.9, hw, 8.6,
               ["〈1월〉", "〈2월〉", "〈3월〉", "〈4월〉", "〈5월〉", "〈6월〉"],
               [("〈올해〉", (120, 135, 128, 150, 162, 171)), ("〈지난해〉", (110, 118, 121, 125, 130, 128))],
               colors=[D.PRIMARY, "B9C4CE"])
    lines_box(s, M, y + 9.8, hw, 2.6,
              ["〈언제부터 달라졌는지, 그 이유는 무엇인지〉", "〈다음 달에 예상되는 흐름〉"], size=9.5, gap=8)
    RX = M + hw + 1.6
    hdr(s, RX, y, hw, "구성 — 무엇으로 이루어져 있는가")
    bar_chart(s, RX, y + 0.9, hw, 8.6, ["〈1분기〉", "〈2분기〉", "〈3분기〉"],
              [("〈구성 A〉", (55, 48, 42)), ("〈구성 B〉", (30, 34, 38)), ("〈구성 C〉", (15, 18, 20))],
              colors=[D.PRIMARY, D.PRIMARY_MID, "B9C4CE"], stacked=True, gap=90, overlap=100)
    lines_box(s, RX, y + 9.8, hw, 2.6,
              ["〈비중이 커지는 것과 작아지는 것〉", "〈그래서 무엇을 조정할 것인지〉"], size=9.5, gap=8)
    return s

# ══ 16. 화면 · 사진 주석 ═══════════════════════════════════════════════════
def s_annotate(prs):
    s = slide(prs, "화면과 사진 설명", "어디를 보고 무엇을 하는지 그림 위에 표시한다.",
              "쓰는 법: 그림을 넣고 그 위에 주석 도형을 올린다. 그림 교체는 그림 선택 → 오른쪽 클릭 → 그림 바꾸기.")
    y = TOP + 0.4
    iw = CW * 0.64
    s.shapes.add_picture(os.path.join(ASSETS, "screen_placeholder.png"), Cm(M), Cm(y), width=Cm(iw))
    ih = iw * 9 / 16.0
    hl = rect(s, M + iw * 0.10, y + ih * 0.16, iw * 0.34, ih * 0.16, None, D.ST_BAD, 1.75)
    hl.fill.background()
    badge(s, M + iw * 0.06, y + ih * 0.14, 0.8, "1", D.ST_BAD)
    badge(s, M + iw * 0.06, y + ih * 0.46, 0.8, "2", D.ST_BAD)
    badge(s, M + iw * 0.84, y + ih * 0.70, 0.8, "3", D.ST_BAD)
    arrow(s, M + iw * 0.70, y + ih * 0.62, M + iw * 0.84, y + ih * 0.70, D.ST_BAD, 1.75)
    cal = s.shapes.add_shape(MSO_SHAPE.RECTANGULAR_CALLOUT, Cm(M + iw * 0.52), Cm(y + ih * 0.05),
                             Cm(6.0), Cm(1.5))
    _style(cal, D.WHITE, D.ST_BAD, 1.25)
    set_text(cal, "〈여기를 본다〉", size=9.5, bold=True, color=D.ST_BAD, align="c")
    mask = rect(s, M + iw * 0.06, y + ih * 0.80, iw * 0.26, ih * 0.10, D.INK_FAINT, None)
    mask.line.fill.background()
    set_text(mask, "가림", size=8.5, bold=True, color=D.WHITE, align="c")
    caption(s, M, y + ih + 0.25, iw, "그림 〈1〉  〈무엇을 보여주는 그림인지〉")
    RX = M + iw + 1.2; RW = CW - iw - 1.2
    hdr(s, RX, y, RW, "그림 설명")
    for i, t in enumerate(["〈먼저 볼 곳과 그 뜻〉", "〈다음에 할 일〉", "〈결과를 확인하는 곳〉"]):
        yy = y + 0.95 + i * 1.6
        badge(s, RX, yy, 0.75, str(i + 1))
        tbox(s, RX + 1.05, yy - 0.1, RW - 1.05, 1.3, t, size=10, color=D.INK, anchor="t",
             margins=(0, 0, 0, 0))
    hdr(s, RX, y + 6.0, RW, "주석 부품")
    lines_box(s, RX, y + 6.9, RW, 5.2,
              ["*번호 배지  순서와 지점", "*빨간 테두리  볼 곳(채우기 없음)", "*말풍선  짧은 설명",
               "*화살표  이동과 연결", "*회색 상자  개인정보 가리기"], size=9.5, bullet="", gap=9)
    label(s, RX, y + 12.3, RW, "사진·화면에 이름·연락처가 보이면 가린 뒤 배포한다.", color=D.INK_FAINT)
    return s

# ══ 17. 그림 여러 장 ═══════════════════════════════════════════════════════
def s_gallery(prs):
    s = slide(prs, "여러 장면 늘어놓기", "사례·종류·단계별 모습을 나란히 보여준다.",
              "쓰는 법: 그림 틀(회색 상자)을 먼저 맞춰 놓고 그림을 넣으면 크기가 흐트러지지 않는다.")
    y = TOP + 0.5
    n = 3; gap = 1.0
    cw = (CW - gap * (n - 1)) / n
    ih = cw * 0.62
    for i in range(n):
        x = M + i * (cw + gap)
        s.shapes.add_picture(os.path.join(ASSETS, "photo_placeholder.png"), Cm(x), Cm(y),
                             width=Cm(cw), height=Cm(ih))
        rect(s, x, y, cw, ih, None, D.LINE, 1.0).fill.background()
        rich(tbox(s, x, y + ih + 0.3, cw, 2.0, "", anchor="t", margins=(0, 0, 0, 0)),
             [("〈제목 한 줄〉\n", {"size": 11, "bold": True, "color": D.INK}),
              ("〈무엇을 보여주는 장면인지 한두 줄로 적는다.〉", {"size": 9.5, "color": D.INK_SOFT})],
             align="l", anchor="t", line=1.35, margins=(0, 0, 0, 0))
    y2 = y + ih + 3.0
    hdr(s, M, y2, CW, "가로로 긴 그림 한 장")
    x = M
    s.shapes.add_picture(os.path.join(ASSETS, "diagram_placeholder.png"), Cm(M), Cm(y2 + 0.9),
                         width=Cm(CW * 0.62))
    lines_box(s, M + CW * 0.64, y2 + 0.9, CW * 0.36, 4.0,
              ["그림은 왼쪽, 설명은 오른쪽에 두면 시선이 자연스럽다",
               "그림에 글씨를 그려 넣지 말고 옆에 적는다",
               "여러 장을 넣을 때는 세로 크기를 똑같이 맞춘다"], size=9.5, gap=8)
    return s

# ══ 18. 표 ═════════════════════════════════════════════════════════════════
def s_table(prs):
    s = slide(prs, "〈표 제목〉", "숫자는 오른쪽, 글자는 왼쪽. 강조할 행은 한 줄만.",
              "쓰는 법: 행 추가는 마지막 칸에서 Tab. 8행이 넘으면 두 장으로 나눈다. 표 안 글자 크기는 10pt 아래로 줄이지 않는다.")
    y = TOP + 0.5
    ptable(s, M, y, [CW * 0.18, CW * 0.26, CW * 0.14, CW * 0.14, CW * 0.14, CW * 0.14],
           ["구분", "항목", "〈올해〉", "〈지난해〉", "〈증감〉", "〈판단〉"],
           [["〈구분 A〉", "〈항목 이름〉", "〈120〉", "〈100〉", "〈+20〉", "〈좋음〉"],
            ["〈구분 A〉", "〈항목 이름〉", "〈98〉", "〈102〉", "〈-4〉", "〈보통〉"],
            ["〈구분 B〉", "〈항목 이름〉", "〈45〉", "〈38〉", "〈+7〉", "〈좋음〉"],
            ["〈구분 B〉", "〈항목 이름〉", "〈12〉", "〈17〉", "〈-5〉", "〈확인〉"],
            ["합계", "", "〈275〉", "〈257〉", "〈+18〉", ""]],
           row_h=1.25, head_h=1.2, size=10.5,
           align=["l", "l", "r", "r", "r", "c"])
    takeaway(s, y + 8.6, "〈표에서 말하려는 결론을 한 문장으로. 표만 두면 읽는 사람이 각자 해석한다.〉", "읽는 법")
    label(s, M, y + 10.5, CW, "※ 〈기준일·산출 방법·제외 대상 같은 단서는 표 아래 작은 글씨로 적는다.〉",
          size=9, color=D.INK_FAINT)
    return s

# ══ 19. 마무리 ═════════════════════════════════════════════════════════════
def s_closing(prs):
    s = slide(prs, "정리와 다음 단계", "무엇을 결정했고 누가 무엇을 언제까지 하는지로 끝낸다.",
              "쓰는 법: 요청 사항은 상대가 바로 답할 수 있게 하나씩 적는다.")
    y = TOP + 0.5
    b = box(s, M, y, CW, 2.8, D.PRIMARY, None, 0)
    rich(b, [("〈오늘 남길 한 문장〉", {"size": 17, "bold": True, "color": D.WHITE})],
         align="l", anchor="ctr", margins=(1.0, 1.0, 0.2, 0.2))
    hdr(s, M, y + 3.6, CW * 0.62, "다음 단계")
    ptable(s, M, y + 4.5, [CW * 0.62 * 0.44, CW * 0.62 * 0.2, CW * 0.62 * 0.2, CW * 0.62 * 0.16],
           ["할 일", "담당", "기한", "상태"],
           [["〈해야 할 일〉", "〈이름〉", "〈MM/DD〉", "〈예정〉"],
            ["〈해야 할 일〉", "〈이름〉", "〈MM/DD〉", "〈예정〉"],
            ["〈해야 할 일〉", "〈이름〉", "〈MM/DD〉", "〈예정〉"]], row_h=1.2, head_h=1.1)
    RX = M + CW * 0.66; RW = CW - CW * 0.66
    hdr(s, RX, y + 3.6, RW, "요청 사항")
    for i, t in enumerate(["〈무엇을 결정해 주셔야 하는지〉", "〈언제까지 필요한지〉",
                           "〈필요한 지원이나 자료〉"]):
        yy = y + 4.5 + i * 1.9
        b2 = box(s, RX, yy, RW, 1.6, D.WHITE, D.LINE, 1.0)
        rect(s, RX, yy, 0.13, 1.6, D.PRIMARY_MID, None).line.fill.background()
        set_text(b2, t, size=10, color=D.INK, margins=(0.45, 0.3, 0.1, 0.1))
    hdr(s, M, y + 9.4, CW, "문의")
    lines_box(s, M, y + 10.3, CW, 2.0,
              ["〈소속 · 이름 · 연락처〉     〈다른 담당 · 연락처〉"], size=10.5, bullet="", gap=6)
    return s

# ════════════════════════════════════════════════════════════════════════════
#  도구상자 — 복사해서 쓰는 도형 부품
# ════════════════════════════════════════════════════════════════════════════
def how(s, x, y, w, items, title="직접 하는 법"):
    hdr(s, x, y, w, title)
    lines_box(s, x, y + 0.95, w, 14.0, items, size=9.5, gap=9)

# ══ 20. 선과 화살표 ════════════════════════════════════════════════════════
def s_lines(prs):
    s = slide(prs, "선과 화살표", "굵기 · 점선 · 화살촉 · 꺾인 선. 필요한 선을 골라 복사한다.",
              "쓰는 법: 선을 복사한 뒤 끝점을 도형 가장자리에 대면 연결점(●)에 붙는다. 붙여 두면 도형을 옮겨도 선이 따라온다.")
    y = TOP + 0.5
    LW = CW * 0.66
    hdr(s, M, y, LW, "굵기와 점선")
    specs = [("가는 선 0.75pt", D.INK_FAINT, 0.75, None), ("기본 1.5pt", D.PRIMARY_MID, 1.5, None),
             ("굵은 선 2.25pt", D.PRIMARY, 2.25, None), ("점선", D.INK_FAINT, 1.5, "dash"),
             ("일점쇄선", D.INK_FAINT, 1.5, "dashDot")]
    for i, (nm, c, lw, dash) in enumerate(specs):
        yy = y + 1.1 + i * 0.95
        conn(s, (M, yy, M + 6.4, yy), c, lw, dash, head=None)
        label(s, M + 6.8, yy - 0.32, 5.0, nm, size=9, color=D.INK)
    hdr(s, M + 12.4, y, LW - 12.4, "화살촉")
    for i, (nm, t) in enumerate(HEADS.items()):
        yy = y + 1.1 + i * 0.95
        conn(s, (M + 12.4, yy, M + 17.4, yy), D.PRIMARY_MID, 1.5, head=t)
        label(s, M + 17.8, yy - 0.32, 4.0, nm, size=9, color=D.INK)
    hdr(s, M, y + 6.2, LW, "꺾인 선 · 곡선 · 양방향")
    yb = y + 7.4
    conn(s, (M, yb, M + 5.2, yb + 2.6), D.PRIMARY_MID, 1.5, kind=MSO_CONNECTOR.ELBOW)
    label(s, M, yb + 2.9, 5.2, "꺾인 연결선", size=9, color=D.INK)
    conn(s, (M + 6.4, yb + 2.6, M + 11.6, yb), D.PRIMARY_MID, 1.5, kind=MSO_CONNECTOR.ELBOW)
    label(s, M + 6.4, yb + 2.9, 5.2, "꺾인 연결선(반대 방향)", size=9, color=D.INK)
    conn(s, (M + 12.8, yb, M + 18.0, yb + 2.6), D.PRIMARY_MID, 1.5, kind=MSO_CONNECTOR.CURVE)
    label(s, M + 12.8, yb + 2.9, 5.2, "곡선 연결선", size=9, color=D.INK)
    conn(s, (M, yb + 4.0, M + 5.2, yb + 4.0), D.PRIMARY_MID, 1.5, head="triangle", tail="triangle")
    label(s, M, yb + 4.3, 5.2, "양방향", size=9, color=D.INK)
    conn(s, (M + 6.4, yb + 4.0, M + 11.6, yb + 4.0), D.INK_FAINT, 1.25, dash="dash", head=None)
    label(s, M + 6.4, yb + 4.3, 5.2, "참조·보조 관계", size=9, color=D.INK)
    a = shp(s, MSO_SHAPE.RIGHT_ARROW, M + 12.8, yb + 3.65, 5.2, 0.8, D.PRIMARY_TINT, D.PRIMARY_MID, 1.0)
    label(s, M + 12.8, yb + 4.6, 5.2, "굵은 화살표 도형", size=9, color=D.INK)
    hdr(s, M, y + 12.5, LW, "선이 뜻하는 것 — 한 문서 안에서 규칙을 지킨다")
    rules = [("현재 연결 · 주경로", D.PRIMARY_MID, 1.5, None),
             ("예정 · 백업 · 논리", D.INK_FAINT, 1.25, "dash"),
             ("특히 중요한 경로", D.PRIMARY, 2.25, None),
             ("상태 표시", D.ST_BAD, 1.5, None)]
    rw = LW / 4
    for i, (nm, c, lwt, dash) in enumerate(rules):
        rx = M + i * rw
        conn(s, (rx, y + 13.5, rx + 1.2, y + 13.5), c, lwt, dash, head=None)
        label(s, rx + 1.45, y + 13.2, rw - 1.55, nm, size=8.5, color=D.INK)
    RX = M + LW + 1.2; RW = CW - LW - 1.2
    how(s, RX, y, RW, [
        "*선 그리기  삽입 › 도형 › 선. Shift 를 누르면 수평·수직·45°로 고정된다",
        "*연결점에 붙이기  선 끝을 도형 가장자리에 가져가면 ● 표시가 생긴다. 그 상태로 놓는다",
        "*꺾인 선  삽입 › 도형 › ‘꺾인 연결선’. 가운데 노란 점을 끌면 꺾이는 위치가 바뀐다",
        "*화살촉 바꾸기  선 선택 › 도형 서식 › 선 › 화살표 머리/꼬리 유형",
        "*굵기는 세 가지만  0.75 / 1.5 / 2.25pt 안에서 쓰면 문서가 정돈되어 보인다",
        "*복제  Ctrl+D, 또는 Ctrl 을 누른 채 끌기",
        "*여러 선 맞추기  모두 선택 › 도형 서식 › 맞춤 › 가운데 맞춤 / 간격을 동일하게"])
    return s

# ══ 21. 상자와 테두리 ══════════════════════════════════════════════════════
def s_boxes(prs):
    s = slide(prs, "상자와 테두리", "모서리 · 두께 · 채우기 단계. 그대로 복사해 크기만 바꾼다.",
              "쓰는 법: 자주 쓰는 모양은 오른쪽 클릭 › ‘기본 도형으로 설정’ 해 두면 다음부터 그 서식으로 그려진다.")
    y = TOP + 0.5
    LW = CW * 0.66
    hdr(s, M, y, LW, "모서리")
    row = [(MSO_SHAPE.RECTANGLE, "각진", None), (MSO_SHAPE.ROUNDED_RECTANGLE, "둥근(작게)", 0.06),
           (MSO_SHAPE.ROUNDED_RECTANGLE, "둥근(크게)", 0.22), (MSO_SHAPE.SNIP_1_RECTANGLE, "한쪽 자른", 0.18),
           (MSO_SHAPE.PLAQUE, "안쪽 둥근", 0.12)]
    bw = (LW - 4 * 0.6) / 5
    for i, (k, nm, adj) in enumerate(row):
        x = M + i * (bw + 0.6)
        shp(s, k, x, y + 1.05, bw, 1.9, D.WHITE, D.PRIMARY_MID, 1.25, adj=adj, text="")
        label(s, x, y + 3.05, bw, nm, size=9, color=D.INK, align="l")
    hdr(s, M, y + 4.0, LW, "채우기와 테두리 두께")
    fills = [(D.WHITE, D.LINE, 1.0, "흰 바탕 · 가는 선"), (D.SURFACE, D.LINE, 1.0, "연한 바탕"),
             (D.PRIMARY_TINT, D.PRIMARY_MID, 1.25, "강조 바탕"), (D.PRIMARY, None, 0, "진한 바탕 · 흰 글자")]
    for i, (fl, ln, lw, nm) in enumerate(fills):
        x = M + i * (bw + 0.6)
        b = shp(s, MSO_SHAPE.RECTANGLE, x, y + 5.05, bw, 1.9, fl, ln, lw)
        set_text(b, "〈내용〉", size=10, bold=True, color=D.WHITE if fl == D.PRIMARY else D.INK, align="c")
        label(s, x, y + 7.05, bw, nm, size=9, color=D.INK, align="l")
    x4 = M + 4 * (bw + 0.6)
    b = shp(s, MSO_SHAPE.RECTANGLE, x4, y + 5.05, bw, 1.9, D.WHITE, D.PRIMARY_MID, 1.0)
    b.line._get_or_add_ln().append(b.line._get_or_add_ln().makeelement(qn_dash(), {'val': 'dash'}))
    label(s, x4, y + 7.05, bw, "점선 테두리", size=9, color=D.INK, align="l")
    hdr(s, M, y + 8.0, LW, "강조 막대가 붙은 상자 · 구분선")
    b1 = rect(s, M, y + 9.05, LW * 0.31, 2.2, D.WHITE, D.LINE, 1.0)
    rect(s, M, y + 9.05, 0.15, 2.2, D.PRIMARY, None).line.fill.background()
    set_text(b1, "  〈왼쪽 강조〉", size=10, color=D.INK, align="l")
    b2 = rect(s, M + LW * 0.345, y + 9.05, LW * 0.31, 2.2, D.WHITE, D.LINE, 1.0)
    rect(s, M + LW * 0.345, y + 9.05, LW * 0.31, 0.15, D.PRIMARY, None).line.fill.background()
    set_text(b2, "〈위쪽 강조〉", size=10, color=D.INK, align="c")
    b3 = rect(s, M + LW * 0.69, y + 9.05, LW * 0.31, 2.2, D.PRIMARY_TINT, None, 0)
    set_text(b3, "〈테두리 없는 강조〉", size=10, color=D.PRIMARY, align="c", bold=True)
    for i, (c, lw, dash, nm) in enumerate([(D.LINE, 1.0, None, "얇은 구분선"),
                                           (D.PRIMARY_MID, 2.0, None, "굵은 구분선"),
                                           (D.LINE, 1.0, "dash", "점선 구분선")]):
        yy = y + 12.0 + i * 0.95
        conn(s, (M, yy, M + LW * 0.45, yy), c, lw, dash, head=None)
        label(s, M + LW * 0.48, yy - 0.32, 6.0, nm, size=9, color=D.INK)
    RX = M + LW + 1.2; RW = CW - LW - 1.2
    how(s, RX, y, RW, [
        "*모서리 둥글기  도형을 선택하면 보이는 노란 점을 안쪽으로 끌면 더 둥글어진다",
        "*모양만 바꾸기  도형 서식 › 도형 편집 › 도형 모양 변경 (글자와 크기는 그대로)",
        "*크기 똑같이  여러 개 선택 › 도형 서식 › 크기 칸에 숫자를 직접 입력",
        "*그림자 끄기  도형 효과 › 그림자 › 없음. 그림자는 대부분 없는 쪽이 깔끔하다",
        "*기본값으로 저장  마음에 드는 도형 › 오른쪽 클릭 › 기본 도형으로 설정",
        "*테두리 없이 쓰기  채우기만 연하게 두면 표·사진과 잘 어울린다",
        "*구분선  선 하나를 슬라이드 폭에 맞추고 Ctrl+D 로 복제해 쓴다"])
    return s

def qn_dash():
    from pptx.oxml.ns import qn as _q
    return _q('a:prstDash')

# ══ 22. 말풍선 · 괄호 · 지시선 ═════════════════════════════════════════════
def s_callouts(prs):
    s = slide(prs, "말풍선 · 괄호 · 지시선", "짧은 설명을 그림이나 표 옆에 붙일 때 쓴다.",
              "쓰는 법: 말풍선 꼬리는 노란 점을 끌어 가리키는 곳에 맞춘다. 꼬리가 대상에 닿게 두는 것이 중요하다.")
    y = TOP + 0.5
    LW = CW * 0.66
    hdr(s, M, y, LW, "말풍선")
    cal = [(MSO_SHAPE.RECTANGULAR_CALLOUT, "사각 말풍선"), (MSO_SHAPE.ROUNDED_RECTANGULAR_CALLOUT, "둥근 말풍선"),
           (MSO_SHAPE.OVAL_CALLOUT, "타원 말풍선"), (MSO_SHAPE.CLOUD_CALLOUT, "구름 말풍선")]
    bw = (LW - 3 * 0.7) / 4
    for i, (k, nm) in enumerate(cal):
        x = M + i * (bw + 0.7)
        c = shp(s, k, x, y + 1.05, bw, 2.3, D.WHITE, D.PRIMARY_MID, 1.25)
        set_text(c, "〈한 줄 설명〉", size=9.5, color=D.PRIMARY, align="c", margins=(0.2, 0.2, 0.05, 0.05))
        label(s, x, y + 3.9, bw, nm, size=9, color=D.INK)
    hdr(s, M, y + 4.8, LW, "괄호 · 묶음")
    bx = M
    shp(s, MSO_SHAPE.LEFT_BRACE, bx, y + 5.85, 0.7, 3.4, None, D.INK_FAINT, 1.25)
    lines_box(s, bx + 1.0, y + 6.0, 6.0, 3.2, ["〈묶이는 항목 1〉", "〈묶이는 항목 2〉", "〈묶이는 항목 3〉"],
              size=9.5, gap=6)
    label(s, bx, y + 9.4, 7.0, "세로 중괄호로 묶기", size=9, color=D.INK)
    bx2 = M + 8.4
    shp(s, MSO_SHAPE.LEFT_BRACE, bx2, y + 6.3, 0.7, 2.6, None, D.INK_FAINT, 1.25, rot=90)
    label(s, bx2 - 0.6, y + 8.4, 8.0, "가로로 눕힌 중괄호(90° 회전)", size=9, color=D.INK)
    bx3 = M + 12.6
    shp(s, MSO_SHAPE.LEFT_BRACKET, bx3, y + 5.85, 0.6, 3.4, None, D.INK_FAINT, 1.25)
    shp(s, MSO_SHAPE.RIGHT_BRACKET, bx3 + 5.6, y + 5.85, 0.6, 3.4, None, D.INK_FAINT, 1.25)
    tbox(s, bx3 + 0.8, y + 6.6, 4.8, 1.8, "〈대괄호로 감싸기〉", size=9.5, color=D.INK_SOFT, align="c",
         anchor="ctr", margins=(0, 0, 0, 0))
    label(s, bx3, y + 9.4, 7.0, "대괄호 두 개를 마주 보게 둔다", size=9, color=D.INK)
    hdr(s, M, y + 10.3, LW, "지시선 · 강조 원")
    rect(s, M, y + 11.4, 6.4, 2.4, D.SURFACE, D.LINE, 1.0)
    tbox(s, M, y + 11.4, 6.4, 2.4, "〈가리킬 대상〉", size=10, color=D.INK_SOFT, align="c", anchor="ctr",
         margins=(0, 0, 0, 0))
    conn(s, (M + 6.6, y + 12.6, M + 9.4, y + 11.9), D.ST_BAD, 1.25, head="triangle", tail="oval")
    label(s, M + 9.6, y + 11.6, 6.0, "지시선(끝을 점으로)", size=9, color=D.INK)
    o = s.shapes.add_shape(MSO_SHAPE.OVAL, Cm(M + 16.0), Cm(y + 11.2), Cm(3.4), Cm(1.8))
    _style(o, None, D.ST_BAD, 1.75); o.fill.background()
    label(s, M + 16.0, y + 13.2, 6.0, "강조 원(채우기 없음)", size=9, color=D.INK)
    RX = M + LW + 1.2; RW = CW - LW - 1.2
    how(s, RX, y, RW, [
        "*꼬리 맞추기  말풍선을 선택하면 꼬리 끝에 노란 점이 생긴다. 가리킬 곳으로 끌어 놓는다",
        "*글자 넣기  도형을 고르고 바로 입력하면 된다. 별도 텍스트 상자를 얹지 않는다",
        "*괄호 눕히기  도형 선택 › 회전 › 90° 회전. 또는 위쪽 회전 손잡이를 Shift 누르고 돌린다",
        "*강조 원  타원을 그리고 채우기 ‘없음’, 선은 2pt 안팎이 보기 좋다",
        "*지시선  선 하나 + 텍스트 상자. 두 개를 함께 선택해 Ctrl+G 로 묶어 두면 옮기기 편하다",
        "*색  주석은 본문 색과 다른 한 가지 색(빨강)만 쓴다"])
    return s

# ══ 23. 순서 · 묶음 부품 ═══════════════════════════════════════════════════
def s_process_parts(prs):
    s = slide(prs, "순서를 보여주는 부품", "화살표 단계 · 번호 원 · 타임라인 · 순환.",
              "쓰는 법: 한 개를 만든 뒤 Ctrl+D 로 복제하고, 여러 개를 선택해 ‘간격을 동일하게’로 맞춘다.")
    y = TOP + 0.5
    LW = CW * 0.66
    hdr(s, M, y, LW, "화살표 단계")
    n = 4; ov = 0.6
    bw = (LW + ov * (n - 1)) / n
    for i in range(n):
        x = M + i * (bw - ov)
        c = shp(s, MSO_SHAPE.CHEVRON, x, y + 1.05, bw, 1.7,
                D.PRIMARY if i == 0 else (D.PRIMARY_TINT if i < 2 else D.WHITE),
                None if i == 0 else D.PRIMARY_MID, 1.0)
        set_text(c, f"{i+1}. 〈단계〉", size=10, bold=True,
                 color=D.WHITE if i == 0 else D.PRIMARY, align="c", margins=(0.5, 0.35, 0.02, 0.02))
    hdr(s, M, y + 3.4, LW, "번호 원 단계")
    cy = y + 4.9
    for i in range(4):
        x = M + i * (LW / 4)
        badge(s, x, cy, 1.5, str(i + 1), D.PRIMARY if i < 2 else D.INK_FAINT, size=13)
        tbox(s, x - 0.6, cy + 1.7, LW / 4 - 0.4, 1.4, "〈단계 이름〉", size=9.5, bold=True, color=D.INK,
             align="c", anchor="t", margins=(0, 0, 0, 0))
        if i < 3:
            conn(s, (x + 1.6, cy + 0.75, x + LW / 4 - 0.1, cy + 0.75), D.LINE, 1.25, head=None)
    hdr(s, M, y + 8.2, LW, "타임라인")
    ty = y + 9.9
    conn(s, (M + 0.4, ty, M + LW - 0.4, ty), D.LINE, 1.5, head=None)
    for i in range(5):
        x = M + 0.8 + i * ((LW - 1.6) / 4)
        dot(s, x - 0.18, ty - 0.18, 0.36, D.PRIMARY if i in (0, 2, 4) else D.INK_FAINT)
        tbox(s, x - 1.6, ty - 1.3, 3.2, 0.8, "〈시점〉", size=9, bold=True, color=D.INK, align="c",
             anchor="t", margins=(0, 0, 0, 0))
        tbox(s, x - 1.8, ty + 0.4, 3.6, 1.2, "〈일어난 일〉", size=9, color=D.INK_SOFT, align="c",
             anchor="t", margins=(0, 0, 0, 0))
    hdr(s, M, y + 11.9, LW, "순환")
    ccx, ccy = M + LW / 2, y + 13.9
    for i, (dx, dy) in enumerate([(-1, -1), (1, -1), (1, 1), (-1, 1)]):
        w, h = 4.6, 1.5
        x = ccx + dx * 4.2 - w / 2
        yy = ccy + dy * 1.1 - h / 2
        shp(s, MSO_SHAPE.ROUNDED_RECTANGLE, x, yy, w, h, D.WHITE, D.PRIMARY_MID, 1.0, adj=0.15,
            text=f"〈{i+1}단계〉", size=9.5, color=D.INK)
    ca = shp(s, MSO_SHAPE.CIRCULAR_ARROW, ccx - 1.6, ccy - 1.6, 3.2, 3.2, D.PRIMARY_TINT, D.PRIMARY_MID, 1.0)
    RX = M + LW + 1.2; RW = CW - LW - 1.2
    how(s, RX, y, RW, [
        "*겹쳐 붙이기  화살표 단계는 조금씩 겹쳐 놓는다. 맞춤 › 간격을 동일하게 로 마무리",
        "*번호 원  타원을 그리고 Shift 를 누른 채 크기를 조절하면 정원이 된다",
        "*선 위에 점  선을 먼저 긋고 그 위에 작은 원을 올린 뒤 모두 선택해 ‘가운데 맞춤’",
        "*순환  가운데 원형 화살표는 삽입 › 도형 › 원형 화살표. 방향은 좌우 대칭으로 뒤집는다",
        "*글자  단계 이름은 명사로, 설명은 도형 밖에 둔다",
        "*묶기  완성한 조합은 Ctrl+G 로 묶어 두면 옮길 때 흐트러지지 않는다"])
    return s

# ══ 24. 강조 · 상태 부품 ═══════════════════════════════════════════════════
def s_accent_parts(prs):
    s = slide(prs, "강조와 상태 부품", "태그 · 배지 · 상태 표시 · 진행 막대 · 형광펜.",
              "쓰는 법: 상태 색은 정상(초록) · 주의(노랑) · 위험(빨강) 세 가지만 쓴다. 색을 늘리면 뜻이 흐려진다.")
    y = TOP + 0.5
    LW = CW * 0.66
    hdr(s, M, y, LW, "태그 · 배지")
    for i, (t, c, bg) in enumerate([("〈태그〉", D.PRIMARY, D.PRIMARY_TINT), ("정상", D.ST_OK, D.OK_BG),
                                    ("주의", D.ST_WARN, D.WARN_BG), ("위험", D.ST_BAD, D.DANGER_BG),
                                    ("〈비활성〉", D.ST_IDLE, D.SURFACE)]):
        pill(s, M + i * 3.2, y + 1.1, 2.8, 0.8, t, c, bg)
    for i in range(5):
        badge(s, M + 16.6 + i * 1.1, y + 1.05, 0.9, str(i + 1),
              D.PRIMARY_MID if i < 3 else D.INK_FAINT)
    hdr(s, M, y + 2.6, LW, "상태 점 · O · X")
    for i, (c, nm) in enumerate([(D.ST_OK, "정상"), (D.ST_WARN, "주의"), (D.ST_BAD, "위험"),
                                 (D.ST_IDLE, "해당 없음")]):
        x = M + i * 3.6
        dot(s, x, y + 3.75, 0.5, c)
        label(s, x + 0.8, y + 3.72, 2.6, nm, size=9.5, color=D.INK)
    for i, (mark, c) in enumerate([("○", D.ST_OK), ("△", D.ST_WARN), ("×", D.ST_BAD)]):
        x = M + 15.2 + i * 1.8
        o = s.shapes.add_shape(MSO_SHAPE.OVAL, Cm(x), Cm(y + 3.5), Cm(1.1), Cm(1.1))
        _style(o, D.WHITE, c, 1.25)
        set_text(o, mark, size=12, bold=True, color=c, align="c", margins=(0, 0, 0, 0))
    hdr(s, M, y + 5.2, LW, "진행 막대")
    for i, (nm, pct, c) in enumerate([("〈항목〉", 0.8, D.ST_OK), ("〈항목〉", 0.45, D.ST_WARN),
                                      ("〈항목〉", 0.15, D.INK_FAINT)]):
        yy = y + 6.3 + i * 1.2
        label(s, M, yy - 0.05, 4.0, nm, size=9.5, color=D.INK)
        rect(s, M + 4.2, yy, 12.0, 0.6, D.SURFACE, None).line.fill.background()
        rect(s, M + 4.2, yy, 12.0 * pct, 0.6, c, None).line.fill.background()
        label(s, M + 16.6, yy - 0.05, 3.0, f"〈{int(pct*100)}%〉", size=9.5, color=D.INK_SOFT)
    hdr(s, M, y + 10.0, LW, "형광펜 · 밑줄 · 리본 태그")
    hi = rect(s, M, y + 11.1, 8.6, 1.0, "FDF3C8", None)
    hi.line.fill.background()
    tbox(s, M, y + 11.1, 8.6, 1.0, "〈뒤에 깔아 강조하는 상자〉", size=10, color=D.INK, align="c",
         anchor="ctr", margins=(0, 0, 0, 0))
    tbox(s, M + 9.4, y + 11.05, 7.0, 1.0, "〈밑줄로 강조〉", size=11, bold=True, color=D.INK, align="c",
         anchor="ctr", margins=(0, 0, 0, 0))
    conn(s, (M + 10.6, y + 12.05, M + 15.2, y + 12.05), D.PRIMARY, 2.25, head=None)
    shp(s, MSO_SHAPE.PENTAGON, M + 17.2, y + 11.1, 4.0, 1.0, D.PRIMARY, None, 0,
        text="〈리본 태그〉", size=9.5, color=D.WHITE)
    RX = M + LW + 1.2; RW = CW - LW - 1.2
    how(s, RX, y, RW, [
        "*형광펜 효과  노란 사각형을 글자 뒤에 놓고 오른쪽 클릭 › 맨 뒤로 보내기",
        "*반투명  도형 서식 › 채우기 › 투명도 30~50%. 사진 위에 글자를 얹을 때 쓴다",
        "*진행 막대  회색 막대 위에 색 막대를 겹치고 왼쪽 끝을 맞춘다",
        "*상태 점  정원(Shift+타원)으로 그리고 색만 바꿔 쓴다",
        "*O · △ · × 표시  원 안에 글자를 넣는다. 표 안에서도 그대로 쓸 수 있다",
        "*색 규칙  강조색은 한 장에 한 가지. 상태색은 초록·노랑·빨강만"])
    return s

# ══ 25. 그림 배치 부품 ═════════════════════════════════════════════════════
def s_picture_parts(prs):
    s = slide(prs, "그림을 다루는 부품", "틀 · 가림 · 반투명 글상자 · 나란히 놓기.",
              "쓰는 법: 그림을 넣기 전에 회색 틀을 먼저 만들어 두면 크기와 위치가 흐트러지지 않는다.")
    y = TOP + 0.5
    LW = CW * 0.66
    hdr(s, M, y, LW, "그림 틀과 캡션")
    fw = 8.4; fh = fw * 0.56
    s.shapes.add_picture(os.path.join(ASSETS, "photo_placeholder.png"), Cm(M), Cm(y + 1.05),
                         width=Cm(fw), height=Cm(fh))
    rect(s, M, y + 1.05, fw, fh, None, D.LINE, 1.0).fill.background()
    caption(s, M, y + 1.05 + fh + 0.2, fw, "〈그림 설명 — 무엇을 보여주는지〉")
    x2 = M + fw + 1.2
    s.shapes.add_picture(os.path.join(ASSETS, "photo_placeholder.png"), Cm(x2), Cm(y + 1.05),
                         width=Cm(fw), height=Cm(fh))
    ov = rect(s, x2, y + 1.05 + fh - 2.0, fw, 2.0, D.INK, None)
    ov.line.fill.background(); alpha(ov, D.INK, 55)
    tbox(s, x2 + 0.5, y + 1.05 + fh - 1.7, fw - 1.0, 1.4, "〈사진 위에 얹는 글자〉", size=12, bold=True,
         color=D.WHITE, anchor="t", margins=(0, 0, 0, 0))
    caption(s, x2, y + 1.05 + fh + 0.2, fw, "반투명 상자를 깔고 흰 글자를 올린다")
    hdr(s, M, y + 7.4, LW, "가림 · 강조")
    s.shapes.add_picture(os.path.join(ASSETS, "screen_placeholder.png"), Cm(M), Cm(y + 8.45),
                         width=Cm(fw), height=Cm(fw * 9 / 16))
    mk = rect(s, M + 0.8, y + 9.3, 3.2, 0.8, D.INK_FAINT, None); mk.line.fill.background()
    set_text(mk, "가림", size=8.5, bold=True, color=D.WHITE, align="c")
    hl = rect(s, M + 4.6, y + 10.4, 3.2, 1.0, None, D.ST_BAD, 1.75); hl.fill.background()
    caption(s, M, y + 8.45 + fw * 9 / 16 + 0.2, fw, "가릴 곳은 회색 상자, 볼 곳은 빨간 테두리")
    x3 = M + fw + 1.2
    lines_box(s, x3, y + 8.6, LW - fw - 1.2, 5.0,
              ["*그림 바꾸기  그림 선택 › 오른쪽 클릭 › 그림 바꾸기 › 파일에서. 크기와 위치가 유지된다",
               "*비율 고정  크기 창에서 ‘가로 세로 비율 고정’을 켜고 한쪽만 입력",
               "*자르기  그림 › 자르기. 여러 장을 같은 비율로 자르면 나란히 두기 좋다",
               "*정렬  여러 장 선택 › 맞춤 › 위쪽 맞춤 + 가로 간격을 동일하게"], size=9.5, bullet="", gap=9)
    RX = M + LW + 1.2; RW = CW - LW - 1.2
    how(s, RX, y, RW, [
        "*틀 먼저  빈 사각형으로 자리를 잡고 그 크기에 맞춰 그림을 넣는다",
        "*테두리  사진에는 가는 회색 테두리(0.75pt) 하나면 충분하다",
        "*그림자  쓰지 않는다. 사진이 지저분해 보인다",
        "*반투명 상자  도형 서식 › 채우기 › 투명도 40~60%",
        "*캡션  그림 아래 9pt 회색. 그림 안에는 글씨를 넣지 않는다",
        "*개인정보  이름 · 연락처 · 사번이 보이면 회색 상자로 가린 뒤 배포한다"])
    return s

# ══ 26. 깔끔하게 ① 정렬과 간격 ════════════════════════════════════════════
def s_tidy_align(prs):
    s = slide(prs, "깔끔해 보이는 법 ① 정렬과 간격", "같은 내용도 줄만 맞추면 달라 보인다.",
              "쓰는 법: 도형을 여러 개 선택하고 [도형 서식] 탭 › 맞춤 에서 ‘위쪽 맞춤’과 ‘가로 간격을 동일하게’를 차례로 누른다.")
    y = TOP + 0.5
    hw = (CW - 1.6) / 2
    hdr(s, M, y, hw, "이렇게 하면 어수선하다")
    bad = [(0.0, 0.0, 5.2, 2.0), (5.9, 0.35, 4.4, 2.3), (11.0, 0.15, 5.0, 1.8)]
    for i, (dx, dy, w, h) in enumerate(bad):
        b = rect(s, M + dx, y + 1.2 + dy, w, h, D.WHITE, D.LINE, 1.0)
        set_text(b, f"〈항목 {i+1}〉", size=10 + i, color=D.INK, align="l" if i != 1 else "c",
                 margins=(0.3, 0.2, 0.05, 0.05))
    lines_box(s, M, y + 4.3, hw, 3.4,
              ["높이와 너비가 제각각이다", "위아래가 어긋나 있다", "사이 간격이 다르다",
               "글자 크기와 정렬이 섞여 있다"], size=9.5, gap=7, color=D.ST_BAD)
    RX = M + hw + 1.6
    hdr(s, RX, y, hw, "이렇게 하면 정돈된다")
    for i in range(3):
        bw = (hw - 1.2) / 3
        b = rect(s, RX + i * (bw + 0.6), y + 1.2, bw, 2.1, D.WHITE, D.LINE, 1.0)
        set_text(b, f"〈항목 {i+1}〉", size=10.5, color=D.INK, align="l", margins=(0.3, 0.2, 0.05, 0.05))
    lines_box(s, RX, y + 4.3, hw, 3.4,
              ["크기를 같게 맞췄다", "위쪽을 한 줄에 맞췄다", "간격을 같게 했다",
               "글자 크기와 정렬을 통일했다"], size=9.5, gap=7, color=D.ST_OK)
    hdr(s, M, y + 8.2, CW, "순서대로 하면 된다")
    steps = [("① 크기 맞추기", "여러 개 선택 › 도형 서식 › 크기 칸에 같은 숫자 입력"),
             ("② 줄 맞추기", "선택 › 맞춤 › 위쪽 맞춤(가로 배치) 또는 왼쪽 맞춤(세로 배치)"),
             ("③ 간격 맞추기", "선택 › 맞춤 › 가로 간격을 동일하게"),
             ("④ 여백 지키기", "슬라이드 가장자리에서 1.5cm 안쪽으로 내용을 둔다")]
    cw = (CW - 2.7) / 4
    for i, (t, d) in enumerate(steps):
        x = M + i * (cw + 0.9)
        b = box(s, x, y + 9.1, cw, 2.6, D.WHITE, D.LINE, 1.0)
        rich(b, [(t + "\n", {"size": 10.5, "bold": True, "color": D.PRIMARY}),
                 (d, {"size": 9, "color": D.INK_SOFT})], align="l", anchor="ctr",
             margins=(0.45, 0.35, 0.1, 0.1), line=1.35)
    takeaway(s, y + 12.3, "눈으로 맞추지 말고 맞춤 기능을 쓴다. 1분이면 끝나고 결과가 확실히 다르다.", "요령")
    return s

# ══ 27. 깔끔하게 ② 색과 글자 ══════════════════════════════════════════════
def s_tidy_color(prs):
    s = slide(prs, "깔끔해 보이는 법 ② 색과 글자", "색을 줄이고 굵기로 구분하면 전문가가 만든 것처럼 보인다.",
              "쓰는 법: 색은 [도형 서식] › 채우기 › 다른 채우기 색 › 사용자 지정 에서 RGB 값을 입력해 맞춘다.")
    y = TOP + 0.5
    hw = (CW - 1.6) / 2
    hdr(s, M, y, hw, "이렇게 하면 산만하다")
    bad = [("〈제목〉", "E74C3C", 16), ("〈소제목〉", "8E44AD", 13), ("〈본문〉", "16A085", 11)]
    for i, (t, c, sz) in enumerate(bad):
        b = rect(s, M, y + 1.2 + i * 1.7, hw, 1.4, "FFF2CC" if i == 1 else D.WHITE, c, 2.25)
        set_text(b, t, size=sz, bold=True, color=c, align="c")
    lines_box(s, M, y + 6.6, hw, 3.6,
              ["색이 네 가지 넘게 쓰였다", "테두리가 굵어 내용보다 눈에 띈다",
               "글자 크기가 제각각이다", "강조가 많아 무엇이 중요한지 알 수 없다"],
              size=9.5, gap=7, color=D.ST_BAD)
    RX = M + hw + 1.6
    hdr(s, RX, y, hw, "이렇게 하면 정돈된다")
    good = [("〈제목〉", D.PRIMARY, 14, True, D.WHITE, D.LINE),
            ("〈소제목〉", D.INK, 11.5, True, D.SURFACE, None),
            ("〈본문〉", D.INK_SOFT, 10.5, False, D.WHITE, D.LINE)]
    for i, (t, c, sz, bold, fl, ln) in enumerate(good):
        b = rect(s, RX, y + 1.2 + i * 1.7, hw, 1.4, fl, ln, 1.0)
        set_text(b, t, size=sz, bold=bold, color=c, align="c")
    lines_box(s, RX, y + 6.6, hw, 3.6,
              ["색은 주색 한 가지 + 회색 단계만 썼다", "선은 1pt 이하로 얇게 했다",
               "크기 차이로 위계를 만들었다", "강조는 한 곳만 했다"],
              size=9.5, gap=7, color=D.ST_OK)
    hdr(s, M, y + 10.6, CW, "지켜야 할 값")
    items = [("글자 크기", "제목 19 · 소제목 12 · 본문 10~11 · 주석 9pt"),
             ("색", "주색 1 + 회색 3단계 + 상태색 3"),
             ("선 굵기", "0.75 / 1.0 / 2.25pt"),
             ("여백", "가장자리 1.5cm · 도형 사이 0.4cm 이상")]
    cw = (CW - 2.7) / 4
    for i, (t, d) in enumerate(items):
        x = M + i * (cw + 0.9)
        b = box(s, x, y + 11.5, cw, 2.4, D.PRIMARY_TINT, None, 0)
        rich(b, [(t + "\n", {"size": 10, "bold": True, "color": D.PRIMARY}),
                 (d, {"size": 9, "color": D.INK})], align="l", anchor="ctr",
             margins=(0.45, 0.35, 0.1, 0.1), line=1.35)
    return s

# ══ 28. 빠르게 하는 법 ═════════════════════════════════════════════════════
def s_shortcuts(prs):
    s = slide(prs, "빠르게 하는 법", "자주 쓰는 것만 모았다. 이 장은 인쇄해 옆에 두고 써도 된다.",
              "쓰는 법: 단축키는 한글·영문 입력 상태와 상관없이 동작한다.")
    y = TOP + 0.5
    hw = (CW - 1.6) / 2
    hdr(s, M, y, hw, "도형 다루기")
    ptable(s, M, y + 0.9, [hw * 0.34, hw * 0.66],
           ["단축키", "하는 일"],
           [["Ctrl + D", "선택한 도형 복제"],
            ["Ctrl 끌기", "끌면서 복사"],
            ["Shift 끌기", "수평·수직으로만 이동"],
            ["Shift + 크기 조절", "비율 유지하며 크기 변경"],
            ["Ctrl + G / Ctrl + Shift + G", "묶기 / 묶음 풀기"],
            ["Alt + F10", "선택 창 열기(겹친 도형 고르기)"],
            ["F4", "직전 동작 반복"]], row_h=1.0, head_h=1.0, size=9.5)
    RX = M + hw + 1.6
    hdr(s, RX, y, hw, "리본에서 찾기")
    ptable(s, RX, y + 0.9, [hw * 0.34, hw * 0.66],
           ["하고 싶은 것", "어디에 있나"],
           [["줄 맞추기·간격", "도형 서식 › 맞춤"],
            ["도형 모양 바꾸기", "도형 서식 › 도형 편집 › 도형 모양 변경"],
            ["앞뒤 순서", "도형 서식 › 앞으로 가져오기 / 뒤로 보내기"],
            ["안내선 켜기", "보기 › 안내선 · 눈금선"],
            ["기본 서식 저장", "도형 오른쪽 클릭 › 기본 도형으로 설정"],
            ["글꼴 한 번에 바꾸기", "홈 › 바꾸기 › 글꼴 바꾸기"],
            ["슬라이드 크기", "디자인 › 슬라이드 크기"]], row_h=1.0, head_h=1.0, size=9.5)
    hdr(s, M, y + 8.8, CW, "마무리 전 3분 점검")
    items = ["가장자리 여백이 일정한가", "도형 줄과 간격이 맞는가", "글꼴이 한 가지인가",
             "색이 네 가지를 넘지 않는가", "글자가 9pt 아래로 내려간 곳은 없는가",
             "한 장에 메시지가 하나인가"]
    cw = (CW - 1.8) / 3
    for i, t in enumerate(items):
        x = M + (i % 3) * (cw + 0.9)
        yy = y + 9.7 + (i // 3) * 1.5
        b = box(s, x, yy, cw, 1.2, D.WHITE, D.LINE, 1.0)
        rich(b, [("□  ", {"size": 11, "color": D.PRIMARY_MID}), (t, {"size": 10, "color": D.INK})],
             align="l", anchor="ctr", margins=(0.4, 0.3, 0.05, 0.05))
    return s

# ════════════════════════════════════════════════════════════════════════════
#  부록 — 디자인 기준 · 아이콘
# ════════════════════════════════════════════════════════════════════════════
ICON_SRC = os.path.join(ASSETS, "아이콘_원본.pptx")
# 원본 색 → 이 라이브러리 색
ICON_DARK = {"1F3A5F": D.PRIMARY, "1E2A36": D.INK, "8C99A8": D.INK_FAINT}
ICON_LIGHT = {"1F3A5F": D.WHITE, "1E2A36": D.WHITE, "8C99A8": "9DB6CA"}

def _icon_grid(s, icons, mapping, label_color):
    cols, iw = 10, 2.85
    px, py = 3.06, 2.42
    x0, y0 = M + 0.05, TOP + 0.35
    for i, (nm, el) in enumerate(icons):
        r, c = divmod(i, cols)
        from pptx_lib import place_group
        place_group(s, el, x0 + c * px, y0 + r * py, w=iw, mapping=mapping, name="아이콘-" + nm)

def s_design_spec(prs):
    s = slide(prs, "디자인 기준", "이 라이브러리가 쓰는 값. 새로 만들 때도 이 안에서 고른다.",
              "쓰는 법: 색은 [도형 채우기] → 다른 채우기 색 → 사용자 지정 에서 RGB 값을 입력한다.")
    y = TOP + 0.4
    hdr(s, M, y, CW, "색")
    sw = [("주색", D.PRIMARY, "제목 · 표 머리글 글자 · 강조"), ("보조", D.PRIMARY_MID, "선 · 화살표 · 번호"),
          ("주색 바탕", D.PRIMARY_TINT, "표 머리글 · 강조 상자"), ("진한 주색", D.PRIMARY_DEEP, "아주 진한 강조"),
          ("본문", D.INK, "글자"), ("보조 글자", D.INK_SOFT, "설명 · 캡션"),
          ("흐린 글자", D.INK_FAINT, "라벨 · 주석"), ("선", D.LINE, "표 선 · 테두리"),
          ("연한 바탕", D.SURFACE, "구분 · 비활성"), ("정상", D.ST_OK, "상태"),
          ("주의", D.ST_WARN, "상태"), ("위험", D.ST_BAD, "상태")]
    cw = (CW - 3 * 0.6) / 4
    for i, (nm, c, use) in enumerate(sw):
        r, col = divmod(i, 4)
        x = M + col * (cw + 0.6); yy = y + 1.0 + r * 1.55
        rect(s, x, yy, 1.15, 1.15, c, D.LINE if c in (D.SURFACE, D.LINE, D.WHITE) else None, 0.75)
        rich(tbox(s, x + 1.4, yy - 0.05, cw - 1.4, 1.25, "", anchor="ctr", margins=(0, 0, 0, 0)),
             [(nm + "   ", {"size": 9.5, "bold": True, "color": D.INK}),
              ("#" + c + "\n", {"size": 8, "color": D.INK_FAINT}),
              (use, {"size": 8.5, "color": D.INK_SOFT})], align="l", anchor="ctr", line=1.3,
             margins=(0, 0, 0, 0))
    y2 = y + 5.9
    lw = (CW - 1.8) / 3
    hdr(s, M, y2, lw, "글자")
    ptable(s, M, y2 + 0.9, [lw * 0.46, lw * 0.28, lw * 0.26],
           ["쓰는 곳", "크기", "굵기"],
           [["슬라이드 제목", "19pt", "굵게"], ["구역 제목", "9.5pt", "굵게"],
            ["본문", "10 ~ 11pt", "보통"], ["표 안", "9.5pt", "보통"],
            ["주석 · 라벨", "8.5 ~ 9pt", "보통"]], row_h=0.95, head_h=1.0, size=9)
    x2 = M + lw + 0.9
    hdr(s, x2, y2, lw, "선")
    for i, (nm, lwt) in enumerate([("0.75pt  가는 선", 0.75), ("1.0pt  기본", 1.0),
                                   ("1.5pt  연결선", 1.5), ("2.25pt  강조", 2.25)]):
        yy = y2 + 1.3 + i * 1.0
        conn(s, (x2, yy, x2 + 3.6, yy), D.PRIMARY_MID, lwt, head=None)
        label(s, x2 + 4.0, yy - 0.32, lw - 4.0, nm, size=9, color=D.INK)
    label(s, x2, y2 + 5.5, lw, "네 가지 안에서만 쓴다", size=8.5, color=D.INK_FAINT)
    x3 = M + 2 * (lw + 0.9)
    hdr(s, x3, y2, lw, "간격 · 모서리")
    lines_box(s, x3, y2 + 0.95, lw, 5.0,
              ["*슬라이드 여백  좌우 1.6cm", "*본문 영역  위 3.05 ~ 아래 17.6cm",
               "*도형 사이  0.4cm 이상", "*각진 모서리  장비 · 요소 · 표",
               "*둥근 모서리  사람 · 외부 · 태그", "*그림자  쓰지 않는다"], size=9.5, bullet="", gap=7)
    return s

def s_icons_dark(prs):
    s = slide(prs, "아이콘 60종", "구성도 · 흐름도에 쓰는 아이콘. 도형이라 색과 크기를 바꿀 수 있다.",
              "쓰는 법: 아이콘을 클릭해 복사한 뒤 붙여 넣는다. 색은 [도형 채우기], 크기는 모서리를 Shift 로 끈다. 글자는 지워도 된다.")
    from pptx_lib import load_groups
    icons = load_groups(ICON_SRC, (5, 6, 7), "아이콘 세트-")
    _icon_grid(s, icons, ICON_DARK, D.INK)
    return s

def s_icons_light(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    rect(s, 0, 0, W, H, D.PRIMARY_DEEP, None)
    tbox(s, M, 1.05, CW - 8, 1.1, "아이콘 60종 — 밝은 색", size=19, bold=True, color=D.WHITE,
         anchor="t", margins=(0, 0, 0, 0))
    tbox(s, M, 2.05, CW - 8, 0.7, "어두운 배경 · 표지 · 섹션 장에 쓴다.", size=10, color="9DB6CA",
         anchor="t", margins=(0, 0, 0, 0))
    line_h(s, M, 2.72, CW, "2C4B66", 1.0)
    from pptx_lib import load_groups
    icons = load_groups(ICON_SRC, (5, 6, 7), "아이콘 세트-")
    _icon_grid(s, icons, ICON_LIGHT, D.WHITE)
    tbox(s, M, BOT + 0.35, CW, 0.6,
         "쓰는 법: 어두운 배경 위에 올려 쓴다. 배경이 밝으면 앞 장(진한 색)을 쓴다.",
         size=8.5, color="7C97AE", anchor="t", margins=(0, 0, 0, 0))
    return s

SLIDES = [s_intro, s_cover, s_agenda, s_message, s_summary, s_structure, s_tree, s_steps,
          s_timeline, s_before_after, s_options, s_matrix, s_kpi, s_bar, s_trend,
          s_annotate, s_gallery, s_table, s_closing,
          s_lines, s_boxes, s_callouts, s_process_parts, s_accent_parts, s_picture_parts,
          s_tidy_align, s_tidy_color, s_shortcuts,
          s_design_spec, s_icons_dark, s_icons_light]

def build():
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(12192000), Emu(6858000)
    set_theme_fonts(prs)
    for f in SLIDES: f(prs)
    prs.core_properties.title = "슬라이드 요소와 도형 도구상자"
    prs.core_properties.author = "문서 작성 라이브러리"
    prs.save(OUT)
    return OUT

if __name__ == "__main__":
    print("PPT 생성:", build())
