# -*- coding: utf-8 -*-
"""library/03_인프라_시각요소.pptx 생성 — 구성도·흐름·비교·상태·화면주석 재사용 요소."""
import os, sys
HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
from pptx import Presentation
from pptx.util import Cm, Pt, Emu
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
import design as D
from pptx_lib import (W, H, M, CW, TOP, BOT, slide, box, rect, tbox, line_h, line_v, arrow,
                      node, zone, chip, pill, badge, dot, card, kpi, set_text, rich, rgb,
                      set_theme_fonts, _style, _apply_font, ptable)

OUT = os.path.join(ROOT, "library", "03_인프라_시각요소.pptx")
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

# ══ 1. 사용 안내 ═══════════════════════════════════════════════════════════
def s_intro(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    rect(s, 0, 0, W, 0.26, D.PRIMARY, None)
    tbox(s, M, 2.6, 20, 0.8, "인프라 시각 요소", size=9.5, bold=True, color=D.PRIMARY_MID,
         anchor="t", margins=(0, 0, 0, 0))
    tbox(s, M, 3.25, 26, 1.6, "구성도 · 흐름 · 비교 · 상태 · 화면 주석", size=29, bold=True,
         color=D.PRIMARY, anchor="t", margins=(0, 0, 0, 0))
    tbox(s, M, 5.15, 26, 0.9, "필요한 슬라이드를 통째로 복사하거나, 도형만 골라 기존 문서에 붙여 넣는다.",
         size=11.5, color=D.INK_SOFT, anchor="t", margins=(0, 0, 0, 0))
    line_h(s, M, 6.5, CW, D.LINE, 1.0)
    cols = [
        ("무엇이 들어 있나",
         ["2쪽 구성도 부품 — 노드·영역·연결선", "3쪽 시스템 구성도(완성)", "4쪽 처리 흐름",
          "5쪽 진행 단계", "6쪽 작업 시간 계획", "7쪽 변경 전·후", "8쪽 방안 비교",
          "9쪽 상태 현황", "10쪽 장애 보고", "11쪽 역할·인수인계",
          "12쪽 화면 주석", "13쪽 구조도·관계도", "14쪽 부품 모음", "15쪽 표"]),
        ("어떻게 쓰나",
         ["*슬라이드 통째로: 왼쪽 목록에서 오른쪽 클릭 → 복사",
          "*도형만: Shift 클릭으로 여러 개 선택 후 복사",
          "*Word 에 붙일 때: 붙여넣기 옵션 → 원본 서식 유지",
          "글자는 도형을 두 번 눌러 바로 고친다",
          "묶인 도형은 Ctrl+Shift+G 로 해제한다",
          "선을 노드 가장자리에 대면 연결점에 붙는다",
          "회색 안내 글(아래쪽)은 복사한 뒤 지운다"]),
        ("맞춰 쓰는 규칙",
         ["글꼴은 맑은 고딕 하나만 쓴다",
          "강조는 색보다 굵기·여백을 먼저 쓴다",
          "상태 색은 정상·주의·위험 세 가지만 쓴다",
          "도형 사이는 0.4cm 이상 띄운다",
          "한 장에 메시지는 하나만 담는다",
          "글자는 9pt 아래로 줄이지 않는다"]),
    ]
    cw = (CW - 1.8) / 3
    for i, (t, items) in enumerate(cols):
        x = M + i * (cw + 0.9)
        hdr(s, x, 7.3, cw, t)
        lines_box(s, x, 8.15, cw, 6.0, items, size=9.5)
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

# ══ 2. 구성도 부품 ═════════════════════════════════════════════════════════
def s_parts(prs):
    s = slide(prs, "구성도 부품", "영역 · 노드 · 연결선 · 라벨. 필요한 것만 골라 복사한다.",
              "노드는 5.9 × 2.0cm 기준이다. 여러 개를 선택하고 [도형 서식] → 맞춤 → ‘가로 간격을 동일하게’로 줄을 맞춘다.")
    LW = 18.6
    hdr(s, M, TOP, LW, "장비 · 서비스 노드")
    nw, gp = 5.9, 0.45
    r1 = [("〈서버명〉", "〈역할〉", D.PRIMARY_MID, D.WHITE, None, "WEB"),
          ("〈서버명〉", "〈역할〉", D.PRIMARY, D.PRIMARY_TINT, None, "WAS"),
          ("〈DB명〉", "〈주 / 대기〉", D.PRIMARY_DEEP, D.WHITE, None, "DB")]
    r2 = [("〈외부 서비스〉", "〈클라우드 · 관리형〉", D.INK_FAINT, D.WHITE, "dash", None),
          ("〈장애 장비〉", "〈조치 필요〉", D.ST_BAD, D.DANGER_BG, None, None),
          ("〈점검 대상〉", "〈주의 상태〉", D.ST_WARN, D.WARN_BG, None, None)]
    for row, y in ((r1, TOP + 0.95), (r2, TOP + 3.35)):
        for i, (t, sub, ac, fl, dash, tag) in enumerate(row):
            node(s, M + i * (nw + gp), y, nw, 2.0, t, sub, accent=ac, fill=fl, dash=dash, tag=tag)
    u = box(s, M, TOP + 5.75, nw, 1.5, D.WHITE, D.LINE, 1.0, radius=0.5)
    set_text(u, "〈사용자〉", size=10.5, align="c")
    u2 = box(s, M + nw + gp, TOP + 5.75, nw, 1.5, D.WHITE, D.LINE, 1.0, radius=0.5)
    set_text(u2, "〈외부 연동처〉", size=10.5, align="c")
    tbox(s, M + 2 * (nw + gp), TOP + 5.75, nw, 1.5,
         "둥근 모서리는 사람·외부,\n각진 모서리는 장비·서비스에 쓴다.",
         size=9, color=D.INK_SOFT, anchor="ctr", margins=(0, 0, 0, 0))

    hdr(s, M, TOP + 7.75, LW, "영역(존)")
    zone(s, M, TOP + 8.7, 9.0, 3.9, "〈DMZ〉")
    zone(s, M + 9.6, TOP + 8.7, 9.0, 3.9, "〈내부망〉", color=D.INK_FAINT, fill=D.SURFACE)
    tbox(s, M + 0.3, TOP + 9.8, 8.4, 2.4, "영역 박스를 먼저 그리고\n오른쪽 클릭 → 맨 뒤로 보내기.\n"
         "노드는 그 위에 올린다.", size=9, color=D.INK_SOFT, anchor="t", margins=(0, 0, 0, 0))
    tbox(s, M + 9.9, TOP + 9.8, 8.4, 2.4, "영역 이름은 왼쪽 위에 둔다.\n"
         "망 분리·보안 구간이 다르면\n색을 달리한다.", size=9, color=D.INK_SOFT, anchor="t", margins=(0, 0, 0, 0))

    RX = M + 19.4; RW = W - M - RX
    hdr(s, RX, TOP, RW, "연결선")
    ly = TOP + 1.1
    for nm, c, lw, dash, head in [("일반 연결", D.INK_FAINT, 1.25, None, True),
                                  ("주요 경로", D.PRIMARY_MID, 2.25, None, True),
                                  ("이중화 · 대기", D.INK_FAINT, 1.25, "dash", False),
                                  ("양방향", D.PRIMARY_MID, 1.25, None, "both")]:
        arrow(s, RX, ly, RX + 3.9, ly, c, lw, dash, head=bool(head), tail=(head == "both"))
        tbox(s, RX + 4.3, ly - 0.42, RW - 4.3, 0.85, nm, size=9.5, color=D.INK, anchor="ctr",
             margins=(0, 0, 0, 0))
        ly += 1.2
    hdr(s, RX, TOP + 5.5, RW, "라벨 · 태그")
    chip(s, RX, TOP + 6.4, 2.6, 0.7, "〈구간〉")
    chip(s, RX + 2.9, TOP + 6.4, 2.6, 0.7, "〈VLAN〉", D.INK_FAINT)
    pill(s, RX, TOP + 7.35, 2.6, 0.7, "정상", D.ST_OK, D.OK_BG)
    pill(s, RX + 2.9, TOP + 7.35, 2.6, 0.7, "주의", D.ST_WARN, D.WARN_BG)
    pill(s, RX + 5.8, TOP + 7.35, 2.6, 0.7, "위험", D.ST_BAD, D.DANGER_BG)
    badge(s, RX + 5.8, TOP + 6.4, 0.7, "1")
    badge(s, RX + 6.7, TOP + 6.4, 0.7, "2", D.INK_FAINT)
    hdr(s, RX, TOP + 8.6, RW, "노드 안에 무엇을 쓸까")
    lines_box(s, RX, TOP + 9.5, RW, 4.0,
              ["이름(호스트명·서비스명)", "역할 한 줄", "이중화·상태는 태그로",
               "IP·포트는 선 옆 라벨로", "제품명·버전은 표로 뺀다"], size=9.5)
    return s

# ══ 3. 시스템 구성도 ═══════════════════════════════════════════════════════
def s_arch(prs):
    s = slide(prs, "〈시스템명〉 구성도", "구간별 계층 구성. 이름·역할만 바꿔 쓴다.",
              "쓰는 법: 노드를 추가하려면 같은 줄의 노드를 복사해 오른쪽에 붙이고 간격만 맞춘다. 구간이 없으면 줄째로 지운다.")
    ux = (W - 5.6) / 2
    up = box(s, ux, TOP + 0.05, 5.6, 1.2, D.WHITE, D.LINE, 1.0, radius=0.5)
    set_text(up, "〈사용자 · 외부 요청〉", size=10.5, align="c")
    bands = [("〈DMZ〉", D.PRIMARY_MID, [("〈방화벽〉", "〈이중화〉", D.PRIMARY_DEEP),
                                        ("〈L4 / LB〉", "〈부하 분산〉", D.PRIMARY_MID),
                                        ("〈WEB-01〉", "〈웹 서버〉", D.PRIMARY_MID),
                                        ("〈WEB-02〉", "〈웹 서버〉", D.PRIMARY_MID)]),
             ("〈서비스망〉", D.PRIMARY, [("〈WAS-01〉", "〈업무 처리〉", D.PRIMARY),
                                         ("〈WAS-02〉", "〈업무 처리〉", D.PRIMARY),
                                         ("〈배치〉", "〈야간 작업〉", D.INK_FAINT),
                                         ("〈모니터링〉", "〈수집·알람〉", D.INK_FAINT)]),
             ("〈데이터망〉", D.PRIMARY_DEEP, [("〈DB-01〉", "〈주〉", D.PRIMARY_DEEP),
                                              ("〈DB-02〉", "〈대기〉", D.PRIMARY_DEEP),
                                              ("〈스토리지〉", "〈백업〉", D.INK_FAINT)])]
    by = TOP + 1.95
    bh, gapb = 3.45, 1.15
    labels = ["〈443〉", "〈8080〉"]
    links = [[(0, 1, "seq"), (1, 2, "seq"), (2, 3, "pair", "〈이중화〉")],
             [(0, 1, "pair", "〈이중화〉")],
             [(0, 1, "pair", "〈동기 복제〉")]]
    for bi, (zname, zc, nodes) in enumerate(bands):
        zone(s, M, by, CW, bh, zname, color=zc, fill=None, label_pos="l")
        lx = M + 3.6
        avail = CW - 4.1
        n = len(nodes)
        nw = (avail - 0.7 * (n - 1)) / n
        cy = by + 1.82
        for i, (t, sub, ac) in enumerate(nodes):
            node(s, lx + i * (nw + 0.7), by + 0.85, nw, 1.95, t, sub, accent=ac)
        for lk in links[bi]:
            i, j, kind = lk[0], lk[1], lk[2]
            x1 = lx + i * (nw + 0.7) + nw + 0.06
            x2 = lx + j * (nw + 0.7) - 0.06
            if kind == "seq":
                arrow(s, x1, cy, x2, cy, D.PRIMARY_MID, 1.5)
            else:
                arrow(s, x1, cy, x2, cy, D.INK_FAINT, 1.25, dash="dash", head=False)
                chip(s, (x1 + x2) / 2 - 1.3, by + 0.12, 2.6, 0.6, lk[3], D.INK_FAINT)
        if bi < 2:
            ax = lx + avail * 0.70
            arrow(s, ax, by + bh, ax, by + bh + gapb, D.PRIMARY_MID, 1.75)
            chip(s, ax + 0.35, by + bh + 0.2, 2.4, 0.62, labels[bi], D.INK_FAINT)
        by += bh + gapb
    arrow(s, W / 2, TOP + 1.25, W / 2, TOP + 1.95, D.PRIMARY_MID, 1.75)
    return s

# ══ 4. 처리 흐름 ═══════════════════════════════════════════════════════════
def s_flow(prs):
    s = slide(prs, "〈업무명〉 처리 흐름", "요청이 지나가는 순서와 구간별로 확인할 것.",
              "쓰는 법: 구간이 늘면 노드와 아래 카드를 함께 복사한다. 구간은 5~7개를 넘기지 않는다.")
    y = TOP + 1.0
    names = [("〈사용자〉", "〈웹 브라우저〉"), ("〈방화벽〉", "〈접근 제어〉"), ("〈L4/LB〉", "〈부하 분산〉"),
             ("〈WEB〉", "〈정적 처리〉"), ("〈WAS〉", "〈업무 로직〉"), ("〈DB〉", "〈데이터 저장〉")]
    n = len(names); gap = 1.0
    bw = (CW - gap * (n - 1)) / n
    for i, (t, sub) in enumerate(names):
        x = M + i * (bw + gap)
        node(s, x, y, bw, 2.1, t, sub, accent=D.PRIMARY if i == 4 else D.PRIMARY_MID)
        if i < n - 1:
            arrow(s, x + bw + 0.08, y + 1.05, x + bw + gap - 0.08, y + 1.05, D.PRIMARY_MID, 1.5)
    for i, lb in enumerate(["〈HTTPS〉", "〈정책 통과〉", "〈서버 선택〉", "〈요청 전달〉", "〈조회·갱신〉"]):
        chip(s, M + i * (bw + gap) + bw - 0.55, y + 2.45, 2.7, 0.62, lb, D.INK_FAINT)
    hdr(s, M, y + 3.6, CW, "구간별 확인 사항")
    cols = [("〈접속 구간〉", ["확인: 〈방화벽 정책, 포트〉", "증상: 〈접속 불가〉", "확인 방법: 〈telnet · 정책표〉",
                              "담당: 〈네트워크〉"]),
            ("〈처리 구간〉", ["확인: 〈서비스 기동, 응답 시간〉", "증상: 〈지연 · 오류〉",
                              "확인 방법: 〈상태 명령 · 로그〉", "담당: 〈운영〉"]),
            ("〈데이터 구간〉", ["확인: 〈세션 수, 잠금, 용량〉", "증상: 〈타임아웃〉",
                                "확인 방법: 〈DB 모니터링 화면〉", "담당: 〈DBA〉"])]
    cwid = (CW - 1.8) / 3
    for i, (t, lines) in enumerate(cols):
        card(s, M + i * (cwid + 0.9), y + 4.45, cwid, 5.4, t, lines, accent=D.PRIMARY_MID)
    return s

# ══ 5. 진행 단계 ═══════════════════════════════════════════════════════════
def s_steps(prs):
    s = slide(prs, "〈작업명〉 진행 단계", "단계별 목적과 산출물. 계획서·보고서 요약용.",
              "쓰는 법: 화살표 도형은 복사해 이어 붙인다. 단계는 3~5개일 때 가장 읽기 좋다.")
    y = TOP + 0.7
    steps = [("준비", "〈사전 점검 · 공지〉", "〈점검표〉"), ("작업", "〈설정 변경 · 적용〉", "〈작업 로그〉"),
             ("검증", "〈기능 · 성능 확인〉", "〈확인 결과〉"), ("마무리", "〈결과 보고 · 인계〉", "〈결과 보고서〉")]
    n = len(steps); ov = 0.8
    bw = (CW + ov * (n - 1)) / n
    for i, (t, d, out) in enumerate(steps):
        x = M + i * (bw - ov)
        sh = s.shapes.add_shape(MSO_SHAPE.CHEVRON, Cm(x), Cm(y), Cm(bw), Cm(2.3))
        done = i < 2
        _style(sh, D.PRIMARY if i == 0 else (D.PRIMARY_MID if i == 1 else D.WHITE),
               None if done else D.PRIMARY_MID, 1.0)
        rich(sh, [(f"{i+1}. {t}", {"size": 13, "bold": True, "color": D.WHITE if done else D.PRIMARY})],
             align="c", anchor="ctr", margins=(0.7, 0.45, 0.05, 0.05))
        lines_box(s, x + 1.0, y + 2.75, bw - 1.6, 2.2,
                  ["*" + d, "산출물: " + out], size=9.5, bullet="", gap=5)
    hdr(s, M, y + 5.1, CW, "단계별 일정과 담당")
    ptable(s, M, y + 6.0, [CW * 0.16, CW * 0.18, CW * 0.18, CW * 0.48],
           ["단계", "기간", "담당", "완료 기준"],
           [[f"{i}. 〈단계 이름〉", "〈MM/DD ~ MM/DD〉", "〈담당자〉", "〈이 상태가 되면 완료〉"]
            for i in range(1, 5)], row_h=1.15, head_h=1.05)
    return s

# ══ 6. 작업 시간 계획 ══════════════════════════════════════════════════════
def s_timeline(prs):
    s = slide(prs, "〈작업명〉 시간 계획", "작업 시간대와 중단 구간을 한 장으로 보여준다.",
              "쓰는 법: 막대는 직사각형이다. 좌우 끝을 끌어 시간을 맞추고 눈금 글자만 바꾼다.")
    y = TOP + 0.9
    ticks = ["22:00", "22:30", "23:00", "23:30", "24:00", "00:30"]
    x0 = M + 5.2; span = CW - 6.6
    line_h(s, x0, y + 0.95, span, D.LINE, 1.0)
    for i, t in enumerate(ticks):
        x = x0 + span * i / (len(ticks) - 1)
        line_v(s, x, y + 0.75, 0.4, D.LINE, 1.0)
        tbox(s, x - 1.1, y, 2.2, 0.6, t, size=8.5, color=D.INK_FAINT, align="c", anchor="t",
             margins=(0, 0, 0, 0))
    bars = [("〈사전 점검〉", "〈담당 A〉", 0.0, 0.18, D.PRIMARY_MID),
            ("〈설정 변경〉", "〈담당 B〉", 0.18, 0.5, D.PRIMARY),
            ("〈서비스 중단〉", "〈영향 구간〉", 0.42, 0.5, D.ST_BAD),
            ("〈검증〉", "〈담당 A〉", 0.5, 0.78, D.PRIMARY_MID),
            ("〈결과 보고〉", "〈작업 책임자〉", 0.78, 1.0, D.INK_FAINT)]
    by = y + 1.85
    for nm, who, a, bb, c in bars:
        rich(tbox(s, M, by - 0.05, 5.0, 1.0, "", anchor="ctr", margins=(0, 0, 0, 0)),
             [(nm + "\n", {"size": 9.5, "bold": True, "color": D.INK}),
              (who, {"size": 8.5, "color": D.INK_FAINT})], align="l", anchor="ctr", line=1.2,
             margins=(0, 0, 0, 0))
        bar = rect(s, x0 + span * a, by + 0.1, span * (bb - a), 0.78, c, None)
        bar.line.fill.background()
        by += 1.35
    line_v(s, x0 + span * 0.42, y + 1.6, by - y - 1.5, D.ST_BAD, 1.0, dash="dash")
    chip(s, x0 + span * 0.42 - 1.75, by + 0.05, 3.5, 0.68, "중단 시작 〈23:00〉", D.ST_BAD)
    hdr(s, M, by + 1.3, CW, "구간 요약")
    items = [("총 작업 시간", "〈2시간 30분〉", D.PRIMARY), ("서비스 중단", "〈15분〉", D.ST_BAD),
             ("되돌리기 판단", "〈23:30까지〉", D.ST_WARN), ("영향 범위", "〈전체 사용자〉", D.INK)]
    kw = (CW - 2.4) / 4
    for i, (t, v, c) in enumerate(items):
        kpi(s, M + i * (kw + 0.8), by + 2.15, kw, 2.3, t, v, color=c, vsize=15)
    return s

# ══ 7. 변경 전·후 ══════════════════════════════════════════════════════════
def s_before_after(prs):
    s = slide(prs, "〈작업명〉 변경 전 · 후", "무엇이 어떻게 달라지는지 한 장으로 보여준다.",
              "쓰는 법: 가운데 화살표를 기준으로 좌우 항목 수를 맞춘다. 화면 캡처를 넣을 때는 12쪽 화면 주석을 함께 쓴다.")
    cw = (CW - 3.2) / 2
    y = TOP + 0.3
    for i, (t, color, fill, items) in enumerate([
        ("변경 전", D.INK_FAINT, D.SURFACE,
         ["〈구성 · 설정 상태〉", "〈문제점 또는 제약〉", "〈영향 받는 대상〉", "〈운영 부담〉"]),
        ("변경 후", D.PRIMARY, D.PRIMARY_TINT,
         ["〈바뀐 구성 · 설정〉", "〈해결되는 점〉", "〈달라지는 사용 방법〉", "〈기대 효과〉"])]):
        x = M + i * (cw + 3.2)
        b = box(s, x, y, cw, 9.2, D.WHITE, D.LINE, 1.0)
        head = rect(s, x, y, cw, 1.15, fill, None); head.line.fill.background()
        set_text(head, t, size=12, bold=True, color=color, align="c")
        lines_box(s, x + 0.8, y + 1.7, cw - 1.6, 7.0, items, size=10.5, color=D.INK, gap=12)
    ax = M + cw + 0.6
    a = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Cm(ax), Cm(y + 3.9), Cm(2.0), Cm(1.4))
    _style(a, D.PRIMARY_MID, None, 0)
    hdr(s, M, y + 9.9, CW, "요약")
    cols = [("무엇을 바꾸나", "〈한 문장으로〉"), ("왜 바꾸나", "〈근거 · 배경〉"),
            ("영향", "〈중단 시간 · 대상〉"), ("되돌리기", "〈롤백 방법〉")]
    kw = (CW - 2.4) / 4
    for i, (t, v) in enumerate(cols):
        bx = M + i * (kw + 0.8)
        b = box(s, bx, y + 10.75, kw, 2.2, D.WHITE, D.LINE, 1.0)
        rich(b, [(t + "\n", {"size": 9, "color": D.INK_FAINT}), (v, {"size": 10.5, "bold": True, "color": D.INK})],
             align="l", anchor="ctr", margins=(0.4, 0.3, 0.1, 0.1), line=1.4)
    return s

# ══ 8. 방안 비교 ═══════════════════════════════════════════════════════════
def s_compare(prs):
    s = slide(prs, "〈검토 항목〉 방안 비교", "선택지를 같은 기준으로 비교하고 결론을 먼저 적는다.",
              "쓰는 법: 선정안 카드만 진한 테두리로 둔다. 기준이 늘면 아래 표에 행을 추가한다.")
    y = TOP + 0.3
    cw = (CW - 1.8) / 3
    opts = [("방안 A", "〈현행 유지〉", False, ["〈비용 부담 없음〉", "〈문제 지속〉"]),
            ("방안 B", "〈부분 개선〉", True, ["〈효과 대비 비용 적정〉", "〈일부 중단 필요〉"]),
            ("방안 C", "〈전면 교체〉", False, ["〈근본 해결〉", "〈비용 · 기간 부담〉"])]
    for i, (t, sub, pick, pros) in enumerate(opts):
        x = M + i * (cw + 0.9)
        b = box(s, x, y, cw, 5.6, D.WHITE, D.PRIMARY if pick else D.LINE, 1.75 if pick else 1.0)
        top = rect(s, x, y, cw, 0.14, D.PRIMARY if pick else D.LINE, None); top.line.fill.background()
        rich(tbox(s, x + 0.7, y + 0.6, cw - 1.4, 1.4, "", anchor="t", margins=(0, 0, 0, 0)),
             [(t + "   ", {"size": 13, "bold": True, "color": D.INK}),
              (sub, {"size": 10, "color": D.INK_SOFT})], align="l", anchor="t", margins=(0, 0, 0, 0))
        if pick: chip(s, x + cw - 3.0, y + 0.55, 2.3, 0.66, "선정", D.PRIMARY)
        lines_box(s, x + 0.7, y + 2.1, cw - 1.4, 2.4, pros, size=9.5, gap=8)
        line_h(s, x + 0.7, y + 4.35, cw - 1.4, D.LINE_SOFT, 1.0)
        rich(tbox(s, x + 0.7, y + 4.55, cw - 1.4, 0.9, "", anchor="ctr", margins=(0, 0, 0, 0)),
             [("비용 〈000〉만원", {"size": 9.5, "color": D.INK_SOFT}),
              ("   기간 〈0〉주", {"size": 9.5, "color": D.INK_SOFT})], align="l", anchor="ctr",
             margins=(0, 0, 0, 0))
    hdr(s, M, y + 6.2, CW, "비교 기준")
    ptable(s, M, y + 7.1, [CW * 0.22, CW * 0.26, CW * 0.26, CW * 0.26],
           ["기준", "방안 A", "방안 B", "방안 C"],
           [["〈비용〉", "〈없음〉", "〈중〉", "〈높음〉"],
            ["〈작업 시간〉", "〈-〉", "〈2시간〉", "〈8시간〉"],
            ["〈서비스 영향〉", "〈없음〉", "〈15분 중단〉", "〈4시간 중단〉"],
            ["〈위험〉", "〈문제 지속〉", "〈낮음〉", "〈중〉"]], row_h=0.95, head_h=1.0)
    b = box(s, M, y + 12.2, CW, 1.5, D.PRIMARY_TINT, None, 0)
    rich(b, [("선정 결과   ", {"size": 10, "bold": True, "color": D.PRIMARY}),
             ("〈방안 B — 선택 이유를 한 문장으로 적는다. 비교 기준 중 무엇을 가장 크게 본 것인지 밝힌다.〉",
              {"size": 10.5, "color": D.INK})], align="l", anchor="ctr", margins=(0.6, 0.6, 0.1, 0.1))
    return s

# ══ 9. 상태 현황 ═══════════════════════════════════════════════════════════
def s_status(prs):
    s = slide(prs, "〈대상〉 운영 현황", "〈2026-01-01 기준〉 지표와 상태를 함께 보여준다.",
              "쓰는 법: 상태는 정상·주의·위험 세 가지만 쓴다. 숫자는 굵게, 단위는 작게.")
    y = TOP + 0.3
    kw = (CW - 2.4) / 4
    for i, (t, v, u, c, sub) in enumerate([
            ("가동률", "〈99.9〉", "%", D.ST_OK, "〈최근 30일〉"),
            ("장애 건수", "〈2〉", "건", D.ST_WARN, "〈전월 3건〉"),
            ("평균 복구 시간", "〈28〉", "분", D.PRIMARY, "〈목표 30분〉"),
            ("점검 완료율", "〈100〉", "%", D.ST_OK, "〈12/12 항목〉")]):
        kpi(s, M + i * (kw + 0.8), y, kw, 3.2, t, v, u, color=c, sub=sub, vsize=24)
    hdr(s, M, y + 3.85, CW * 0.58, "대상별 상태")
    rows = [("〈WEB-01〉", "정상", D.ST_OK, D.OK_BG, "〈01-01 09:00〉", "〈-〉"),
            ("〈WAS-01〉", "주의", D.ST_WARN, D.WARN_BG, "〈01-01 09:00〉", "〈메모리 82%〉"),
            ("〈DB-01〉", "정상", D.ST_OK, D.OK_BG, "〈01-01 09:00〉", "〈-〉"),
            ("〈백업 서버〉", "위험", D.ST_BAD, D.DANGER_BG, "〈12-31 22:00〉", "〈백업 실패 1건〉"),
            ("〈모니터링〉", "정상", D.ST_OK, D.OK_BG, "〈01-01 09:00〉", "〈-〉")]
    ry = y + 4.75
    tw = CW * 0.58
    hrow = rect(s, M, ry, tw, 1.0, D.PRIMARY_TINT, None); hrow.line.fill.background()
    for lbl, fx in (("대상", 0.02), ("상태", 0.33), ("최근 확인", 0.53), ("비고", 0.76)):
        tbox(s, M + tw * fx + 0.35, ry, tw * 0.25, 1.0, lbl, size=9.5, bold=True, color=D.PRIMARY,
             anchor="ctr", margins=(0, 0, 0, 0))
    for i, (nm, st, c, bg, when, memo) in enumerate(rows):
        yy = ry + 1.0 + i * 1.35
        rect(s, M, yy, tw, 1.35, D.WHITE if i % 2 else D.SURFACE, None).line.fill.background()
        tbox(s, M + 0.35, yy, tw * 0.3, 1.35, nm, size=10, bold=True, color=D.INK, anchor="ctr",
             margins=(0, 0, 0, 0))
        pill(s, M + tw * 0.33 + 0.35, yy + 0.35, 1.9, 0.64, st, c, bg)
        tbox(s, M + tw * 0.53 + 0.35, yy, tw * 0.22, 1.35, when, size=9, color=D.INK_SOFT, anchor="ctr",
             margins=(0, 0, 0, 0))
        tbox(s, M + tw * 0.76 + 0.35, yy, tw * 0.23, 1.35, memo, size=9, color=D.INK_SOFT, anchor="ctr",
             margins=(0, 0, 0, 0))
    line_h(s, M, ry + 1.0 + len(rows) * 1.35, tw, D.PRIMARY_MID, 1.0)
    tbox(s, M, ry + 1.0 + len(rows) * 1.35 + 0.35, tw, 0.8,
         "※ 〈점검 기준과 판정 방법은 운영 매뉴얼 4장 참고〉", size=8.5, color=D.INK_FAINT,
         anchor="t", margins=(0, 0, 0, 0))
    RX = M + CW * 0.62; RW = CW - CW * 0.62
    hdr(s, RX, y + 3.85, RW, "조치가 필요한 항목")
    items = [("〈백업 실패〉", "〈원인 확인 후 재수행〉", "〈01-02까지〉", D.ST_BAD),
             ("〈메모리 사용률 상승〉", "〈추세 관찰, 증설 검토〉", "〈01-15까지〉", D.ST_WARN),
             ("〈로그 용량〉", "〈보관 주기 조정〉", "〈01-31까지〉", D.ST_WARN)]
    for i, (t, act, due, c) in enumerate(items):
        yy = y + 4.75 + i * 2.15
        b = box(s, RX, yy, RW, 1.9, D.WHITE, D.LINE, 1.0)
        rect(s, RX, yy, 0.13, 1.9, c, None).line.fill.background()
        rich(b, [(t + "\n", {"size": 10.5, "bold": True, "color": D.INK}),
                 (act + "   ·   " + due, {"size": 9, "color": D.INK_SOFT})],
             align="l", anchor="ctr", margins=(0.45, 0.3, 0.1, 0.1), line=1.35)
    hdr(s, RX, y + 11.3, RW, "예정된 작업")
    lines_box(s, RX, y + 12.2, RW, 3.0,
              ["〈01-10  정기 점검(월)〉", "〈01-15  백업 정책 변경 작업〉", "〈01-20  인수인계 2차〉"],
              size=9.5, gap=8)
    return s

# ══ 10. 장애 보고 ══════════════════════════════════════════════════════════
def s_incident(prs):
    s = slide(prs, "〈장애명〉 장애 보고", "〈발생일〉 · 〈대상 시스템〉 — 경과 · 원인 · 조치 요약.",
              "쓰는 법: 시간은 왼쪽부터 순서대로. 상세 내용은 문서로 넘기고 이 장에는 결론만 남긴다.")
    y = TOP + 0.3
    kw = (CW - 3.2) / 5
    for i, (t, v, c) in enumerate([("발생", "〈01-01 14:20〉", D.INK), ("인지", "〈14:25〉", D.INK),
                                   ("조치 시작", "〈14:31〉", D.PRIMARY), ("복구", "〈15:05〉", D.ST_OK),
                                   ("총 영향 시간", "〈45분〉", D.ST_BAD)]):
        b = box(s, M + i * (kw + 0.8), y, kw, 2.0, D.WHITE, D.LINE, 1.0)
        rich(b, [(t + "\n", {"size": 9, "color": D.INK_FAINT}), (v, {"size": 12.5, "bold": True, "color": c})],
             align="l", anchor="ctr", margins=(0.4, 0.3, 0.1, 0.1), line=1.4)
    ty = y + 3.2
    hdr(s, M, ty, CW, "경과")
    lx = M + 0.6; lw = CW - 1.2
    line_h(s, lx, ty + 1.9, lw, D.LINE, 1.25)
    events = [("14:20", "〈장애 발생〉", D.ST_BAD), ("14:25", "〈알람 · 신고 접수〉", D.ST_WARN),
              ("14:31", "〈1차 조치 시작〉", D.PRIMARY_MID), ("14:52", "〈원인 확인〉", D.PRIMARY_MID),
              ("15:05", "〈서비스 정상화〉", D.ST_OK)]
    for i, (t, e, c) in enumerate(events):
        x = lx + lw * i / (len(events) - 1)
        dot(s, x - 0.16, ty + 1.74, 0.32, c)
        tbox(s, x - 2.0, ty + 1.05, 4.0, 0.6, t, size=9.5, bold=True, color=D.INK, align="c",
             anchor="t", margins=(0, 0, 0, 0))
        tbox(s, x - 2.4, ty + 2.2, 4.8, 1.0, e, size=9, color=D.INK_SOFT, align="c", anchor="t",
             margins=(0, 0, 0, 0))
    cy = ty + 3.9
    cw = (CW - 1.8) / 3
    for i, (t, lines, c) in enumerate([
            ("원인", ["〈직접 원인 한 줄〉", "〈확인 근거 — 로그 · 지표〉", "〈배경 요인〉"], D.ST_BAD),
            ("조치", ["〈응급 조치 내용〉", "〈정상 확인 방법〉", "〈서비스 영향 범위〉"], D.PRIMARY),
            ("재발 방지", ["〈조치 항목〉", "〈담당 · 기한〉", "〈확인 방법〉"], D.ST_OK)]):
        x = M + i * (cw + 0.9)
        card(s, x, cy, cw, 4.6, t, lines, accent=c)
    return s

# ══ 11. 역할 · 인수인계 ════════════════════════════════════════════════════
def s_rnr(prs):
    s = slide(prs, "역할 분담과 인수인계 현황", "누가 무엇을 맡는지, 인계가 어디까지 되었는지.",
              "쓰는 법: ● 주담당 ○ 협조 △ 승인. 기호만 바꿔 쓰고 빈칸은 그대로 둔다.")
    y = TOP + 0.3
    tw = CW * 0.56
    hdr(s, M, y, tw, "역할 분담(R&R)")
    ptable(s, M, y + 0.9, [tw * 0.34, tw * 0.22, tw * 0.22, tw * 0.22],
           ["업무", "〈운영팀〉", "〈유지보수〉", "〈책임자〉"],
           [["〈일상 점검〉", "●", "○", ""], ["〈장애 1차 대응〉", "●", "○", ""],
            ["〈설정 변경〉", "○", "●", "△"], ["〈증설 · 구성 변경〉", "○", "●", "△"],
            ["〈백업 · 복구〉", "○", "●", "△"], ["〈보고 · 승인〉", "○", "", "●"]],
           row_h=1.0, head_h=1.0, align=["l", "c", "c", "c"])
    nb = box(s, M, y + 8.6, tw, 3.2, D.SURFACE, None, 0)
    rich(nb, [("합의 사항\n", {"size": 9.5, "bold": True, "color": D.PRIMARY}),
              ("〈연락 가능 시간, 대응 시작 기준, 승인 없이 수행할 수 있는 범위처럼 "
               "역할 표만으로 정해지지 않는 내용을 적는다.〉", {"size": 9.5, "color": D.INK_SOFT})],
         align="l", anchor="ctr", margins=(0.6, 0.6, 0.2, 0.2), line=1.4)
    RX = M + tw + 1.2; RW = CW - tw - 1.2
    hdr(s, RX, y, RW, "인수인계 현황")
    items = [("〈계정 · 권한〉", 100, "완료", D.ST_OK, D.OK_BG),
             ("〈운영 문서〉", 80, "진행", D.ST_WARN, D.WARN_BG),
             ("〈장애 이력〉", 100, "완료", D.ST_OK, D.OK_BG),
             ("〈모니터링 인계〉", 40, "진행", D.ST_WARN, D.WARN_BG),
             ("〈미결 과제〉", 0, "예정", D.ST_IDLE, D.SURFACE)]
    for i, (nm, pct, st, c, bg) in enumerate(items):
        yy = y + 1.15 + i * 2.28
        tbox(s, RX, yy, RW - 3.0, 0.7, nm, size=10.5, bold=True, color=D.INK, anchor="t",
             margins=(0, 0, 0, 0))
        pill(s, RX + RW - 2.4, yy - 0.05, 2.4, 0.66, st, c, bg)
        rect(s, RX, yy + 0.85, RW, 0.42, D.SURFACE, None).line.fill.background()
        if pct:
            rect(s, RX, yy + 0.85, RW * pct / 100.0, 0.42, c, None).line.fill.background()
        tbox(s, RX, yy + 1.35, RW, 0.6, f"〈{pct}%〉 〈남은 일 · 기한〉", size=8.5, color=D.INK_FAINT,
             anchor="t", margins=(0, 0, 0, 0))
    return s

# ══ 12. 화면 주석 ══════════════════════════════════════════════════════════
def s_screen(prs):
    s = slide(prs, "화면 캡처와 주석", "어디를 보고 무엇을 누르는지 화면 위에 표시한다.",
              "쓰는 법: 캡처를 넣고 그림 위에 주석 도형을 올린다. 그림 교체는 그림 선택 → 오른쪽 클릭 → 그림 바꾸기.")
    y = TOP + 0.3
    iw = CW * 0.63
    pic = s.shapes.add_picture(os.path.join(ASSETS, "screen_placeholder.png"),
                               Cm(M), Cm(y), width=Cm(iw))
    ih = iw * 9 / 16.0
    # 주석 도형
    hl = rect(s, M + iw * 0.10, y + ih * 0.16, iw * 0.34, ih * 0.16, None, D.ST_BAD, 1.75)
    hl.fill.background()
    badge(s, M + iw * 0.06, y + ih * 0.14, 0.78, "1", D.ST_BAD)
    badge(s, M + iw * 0.06, y + ih * 0.46, 0.78, "2", D.ST_BAD)
    badge(s, M + iw * 0.82, y + ih * 0.70, 0.78, "3", D.ST_BAD)
    arrow(s, M + iw * 0.70, y + ih * 0.62, M + iw * 0.82, y + ih * 0.70, D.ST_BAD, 1.75)
    cal = s.shapes.add_shape(MSO_SHAPE.RECTANGULAR_CALLOUT, Cm(M + iw * 0.55), Cm(y + ih * 0.06),
                             Cm(5.6), Cm(1.5))
    _style(cal, D.WHITE, D.ST_BAD, 1.25)
    set_text(cal, "〈여기를 누른다〉", size=9.5, bold=True, color=D.ST_BAD, align="c")
    mask = rect(s, M + iw * 0.06, y + ih * 0.78, iw * 0.26, ih * 0.10, D.INK_FAINT, None)
    mask.line.fill.background()
    set_text(mask, "가림", size=8.5, bold=True, color=D.WHITE, align="c")
    tbox(s, M, y + ih + 0.25, iw, 0.7, "그림 〈1〉  〈화면 이름 — 무엇을 보여주는 화면인지〉",
         size=9, color=D.INK_SOFT, anchor="t", margins=(0, 0, 0, 0))

    RX = M + iw + 1.2; RW = CW - iw - 1.2
    hdr(s, RX, y, RW, "화면 설명")
    for i, t in enumerate(["〈왼쪽 위 메뉴에서 ○○을 고른다〉", "〈목록에서 대상 항목을 찾는다〉",
                           "〈오른쪽 아래 [적용]을 누른다〉"]):
        yy = y + 0.95 + i * 1.5
        badge(s, RX, yy, 0.72, str(i + 1))
        tbox(s, RX + 1.0, yy - 0.1, RW - 1.0, 1.2, t, size=10, color=D.INK, anchor="t",
             margins=(0, 0, 0, 0))
    hdr(s, RX, y + 5.7, RW, "주석 부품")
    parts = [("번호 배지", "순서·지점 표시"), ("빨간 테두리", "볼 곳 강조(채우기 없음)"),
             ("말풍선", "짧은 설명"), ("화살표", "이동·연결"), ("회색 가림 상자", "이름·IP 가리기")]
    for i, (t, d) in enumerate(parts):
        yy = y + 6.6 + i * 1.15
        rich(tbox(s, RX, yy, RW, 1.0, "", anchor="ctr", margins=(0, 0, 0, 0)),
             [(t + "  ", {"size": 9.5, "bold": True, "color": D.INK}),
              (d, {"size": 9, "color": D.INK_SOFT})], align="l", anchor="ctr", margins=(0, 0, 0, 0))
    tbox(s, RX, y + 12.5, RW, 1.6, "캡처에 개인정보·실제 IP가 보이면 회색 상자로 가린 뒤 배포한다.",
         size=9, color=D.INK_FAINT, anchor="t", margins=(0, 0, 0, 0))
    return s

# ══ 13. 구조도 · 관계도 ════════════════════════════════════════════════════
def s_tree(prs):
    s = slide(prs, "구조도와 관계도", "계층(위에서 아래) · 관계(가운데에서 바깥)로 나눠 그린다.",
              "쓰는 법: 가지를 늘릴 때는 노드와 선을 함께 복사한다. 선은 ‘꺾인 연결선’을 쓰면 정리가 쉽다.")
    y = TOP + 0.5
    lw = CW * 0.52
    hdr(s, M, y, lw, "계층 구조 — 조직 · 시스템 · 서비스")
    cx = M + lw / 2
    root = node(s, cx - 3.4, y + 1.3, 6.8, 1.8, "〈상위 시스템〉", "〈전체 서비스〉", accent=D.PRIMARY)
    mid = [("〈하위 시스템 A〉", "〈역할〉"), ("〈하위 시스템 B〉", "〈역할〉")]
    for i, (t, sub) in enumerate(mid):
        x = M + 0.4 + i * (lw / 2 + 0.2)
        node(s, x, y + 4.7, lw / 2 - 0.8, 1.8, t, sub, accent=D.PRIMARY_MID)
        arrow(s, cx, y + 3.1, x + (lw / 2 - 0.8) / 2, y + 4.7, D.LINE, 1.25, head=False,
              kind=MSO_CONNECTOR.ELBOW)
    lines_box(s, M, y + 10.6, lw, 2.6,
              ["같은 층은 높이를 맞춘다", "한 층에 6개를 넘기면 묶어서 한 단계를 더 만든다",
               "가지가 한쪽으로 몰리면 좌우 순서를 바꿔 균형을 맞춘다"], size=9.5)
    for i in range(4):
        x = M + 0.4 + (i // 2) * (lw / 2 + 0.2) + (i % 2) * ((lw / 2 - 0.8) / 2 + 0.15)
        w = (lw / 2 - 0.8) / 2 - 0.15
        node(s, x, y + 8.3, w, 1.7, "〈구성 요소〉", "〈역할〉", accent=D.INK_FAINT, tsize=9.5, ssize=8)
        px = M + 0.4 + (i // 2) * (lw / 2 + 0.2) + (lw / 2 - 0.8) / 2
        arrow(s, px, y + 6.5, x + w / 2, y + 8.3, D.LINE, 1.25, head=False, kind=MSO_CONNECTOR.ELBOW)

    RX = M + lw + 1.4; RW = CW - lw - 1.4
    hdr(s, RX, y, RW, "관계도 — 연동 · 의존")
    ccx, ccy = RX + RW / 2, y + 5.3
    hub = box(s, ccx - 2.6, ccy - 1.0, 5.2, 2.0, D.PRIMARY_TINT, D.PRIMARY, 1.5)
    set_text(hub, "〈우리 시스템〉", size=11, bold=True, color=D.PRIMARY, align="c")
    around = [(-1, -1, "〈연동 A〉", "〈전문 송수신〉"), (1, -1, "〈연동 B〉", "〈인증〉"),
              (-1, 1, "〈연동 C〉", "〈결제〉"), (1, 1, "〈연동 D〉", "〈알림〉")]
    for dx, dy, t, sub in around:
        w, h = 4.6, 1.5
        x = ccx + dx * (RW / 2 - w / 2 - 0.2) - w / 2
        yy = ccy + dy * 4.1 - h / 2
        node(s, x, yy, w, h, t, sub, accent=D.INK_FAINT, tsize=9.5, ssize=8)
        arrow(s, ccx + dx * 1.6, ccy + dy * 1.0, x + w / 2 - dx * 0.4, yy + (h if dy < 0 else 0),
              D.INK_FAINT, 1.25, head=True, tail=True)
    lines_box(s, RX, ccy + 5.6, RW, 2.6,
              ["선 위에 연동 방식(전문·API·파일)과 주기를 라벨로 적는다",
               "보내는 쪽과 받는 쪽이 다르면 화살표를 한 방향으로 그린다",
               "연동처가 6개를 넘으면 표로 정리하고 그림에는 주요 연동만 남긴다"], size=9.5)
    return s

# ══ 14. 부품 모음 ══════════════════════════════════════════════════════════
def s_kit(prs):
    s = slide(prs, "부품 모음", "자주 쓰는 작은 요소. 복사해서 색과 글자만 바꾼다.",
              "쓰는 법: 같은 종류끼리 간격을 맞춰 두면 문서 전체가 정돈되어 보인다.")
    y = TOP + 0.3
    cw = (CW - 1.8) / 3
    # 1열
    hdr(s, M, y, cw, "번호 · 상태")
    for i in range(5):
        badge(s, M + i * 1.2, y + 0.95, 0.85, str(i + 1), D.PRIMARY_MID if i < 3 else D.INK_FAINT)
    for i, (t, c, bg) in enumerate([("정상", D.ST_OK, D.OK_BG), ("주의", D.ST_WARN, D.WARN_BG),
                                    ("위험", D.ST_BAD, D.DANGER_BG), ("대기", D.ST_IDLE, D.SURFACE)]):
        pill(s, M + (i % 2) * 3.0, y + 2.2 + (i // 2) * 1.0, 2.7, 0.7, t, c, bg)
    for i, c in enumerate([D.ST_OK, D.ST_WARN, D.ST_BAD, D.ST_IDLE]):
        dot(s, M + 6.5 + i * 0.85, y + 2.35, 0.42, c)
    tbox(s, M + 6.5, y + 2.95, cw - 6.5, 0.6, "상태 점", size=8.5, color=D.INK_FAINT, anchor="t",
         margins=(0, 0, 0, 0))
    hdr(s, M, y + 4.6, cw, "강조 상자")
    b1 = box(s, M, y + 5.5, cw, 1.8, D.PRIMARY_TINT, None, 0)
    rich(b1, [("핵심   ", {"size": 9.5, "bold": True, "color": D.PRIMARY}),
              ("〈한 문장으로 결론을 적는다〉", {"size": 10, "color": D.INK})],
         align="l", anchor="ctr", margins=(0.5, 0.4, 0.1, 0.1))
    b2 = box(s, M, y + 7.6, cw, 1.8, D.WARN_BG, None, 0)
    rect(s, M, y + 7.6, 0.13, 1.8, D.ST_WARN, None).line.fill.background()
    rich(b2, [("주의   ", {"size": 9.5, "bold": True, "color": D.WARN}),
              ("〈놓치기 쉬운 점〉", {"size": 10, "color": D.INK})],
         align="l", anchor="ctr", margins=(0.5, 0.4, 0.1, 0.1))
    hdr(s, M, y + 10.0, cw, "범례")
    lg = box(s, M, y + 10.9, cw, 3.0, D.WHITE, D.LINE, 1.0)
    for i, (c, t) in enumerate([(D.PRIMARY_MID, "〈주 경로〉"), (D.INK_FAINT, "〈대기 · 백업〉"),
                                (D.ST_BAD, "〈장애 구간〉")]):
        line_h(s, M + 0.5, y + 11.6 + i * 0.8, 1.6, c, 1.75, "dash" if i == 1 else None)
        tbox(s, M + 2.4, y + 11.25 + i * 0.8, cw - 2.8, 0.7, t, size=9, color=D.INK_SOFT, anchor="ctr",
             margins=(0, 0, 0, 0))
    # 2열
    x2 = M + cw + 0.9
    hdr(s, x2, y, cw, "화살표 · 흐름")
    for i, (t, c, lw_, dash) in enumerate([("진행", D.PRIMARY_MID, 1.75, None),
                                           ("되돌림", D.INK_FAINT, 1.5, "dash"),
                                           ("양방향", D.PRIMARY_MID, 1.5, None)]):
        yy = y + 1.1 + i * 1.1
        arrow(s, x2, yy, x2 + 4.0, yy, c, lw_, dash, head=True, tail=(i == 2))
        tbox(s, x2 + 4.4, yy - 0.4, cw - 4.4, 0.8, t, size=9, color=D.INK_SOFT, anchor="ctr",
             margins=(0, 0, 0, 0))
    for i, sh in enumerate([MSO_SHAPE.RIGHT_ARROW, MSO_SHAPE.CHEVRON, MSO_SHAPE.DOWN_ARROW]):
        a = s.shapes.add_shape(sh, Cm(x2 + i * 2.6), Cm(y + 4.5), Cm(2.2), Cm(1.2))
        _style(a, D.PRIMARY_TINT, D.PRIMARY_MID, 1.0)
    hdr(s, x2, y + 6.2, cw, "설명 목록 조합")
    for i, (t, d) in enumerate([("〈항목 이름〉", "〈한 줄 설명을 적는다〉"),
                                ("〈항목 이름〉", "〈한 줄 설명을 적는다〉"),
                                ("〈항목 이름〉", "〈한 줄 설명을 적는다〉")]):
        yy = y + 7.1 + i * 1.9
        badge(s, x2, yy + 0.15, 0.8, str(i + 1))
        rich(tbox(s, x2 + 1.1, yy, cw - 1.1, 1.6, "", anchor="ctr", margins=(0, 0, 0, 0)),
             [(t + "\n", {"size": 10.5, "bold": True, "color": D.INK}),
              (d, {"size": 9.5, "color": D.INK_SOFT})], align="l", anchor="ctr", line=1.35,
             margins=(0, 0, 0, 0))
    # 3열
    x3 = M + 2 * (cw + 0.9)
    hdr(s, x3, y, cw, "지표 타일")
    kpi(s, x3, y + 0.95, cw / 2 - 0.4, 3.0, "〈지표〉", "〈99.9〉", "%", color=D.ST_OK, vsize=18)
    kpi(s, x3 + cw / 2 + 0.4, y + 0.95, cw / 2 - 0.4, 3.0, "〈지표〉", "〈2〉", "건", color=D.ST_BAD, vsize=18)
    hdr(s, x3, y + 4.4, cw, "카드")
    card(s, x3, y + 5.3, cw, 4.0, "〈카드 제목〉",
         ["〈첫 줄 설명〉", "〈둘째 줄 설명〉", "〈셋째 줄 설명〉"], accent=D.PRIMARY_MID)
    hdr(s, x3, y + 9.8, cw, "구분선 · 메모")
    line_h(s, x3, y + 10.8, cw, D.LINE, 1.0)
    line_h(s, x3, y + 11.4, cw, D.PRIMARY_MID, 1.75)
    tbox(s, x3, y + 11.9, cw, 2.0, "〈보충 설명은 회색 8.5~9pt 로 적고, 본문과 한 줄 띄운다.〉",
         size=9, color=D.INK_FAINT, anchor="t", margins=(0, 0, 0, 0))
    return s

# ══ 15. 표 ═════════════════════════════════════════════════════════════════
def s_table(prs):
    s = slide(prs, "〈표 제목〉", "발표·보고용 표. 숫자는 오른쪽, 글자는 왼쪽으로 맞춘다.",
              "쓰는 법: 행 추가는 마지막 칸에서 Tab. 열 너비는 경계선을 끌어 조정한다. 8행을 넘기면 두 장으로 나눈다.")
    y = TOP + 0.5
    ptable(s, M, y, [CW * 0.16, CW * 0.24, CW * 0.15, CW * 0.15, CW * 0.15, CW * 0.15],
           ["구분", "항목", "〈현재〉", "〈목표〉", "〈차이〉", "판정"],
           [["〈구분 A〉", "〈항목 이름〉", "〈100〉", "〈120〉", "〈-20〉", "〈미달〉"],
            ["〈구분 A〉", "〈항목 이름〉", "〈98〉", "〈95〉", "〈+3〉", "〈충족〉"],
            ["〈구분 B〉", "〈항목 이름〉", "〈45〉", "〈50〉", "〈-5〉", "〈미달〉"],
            ["〈구분 B〉", "〈항목 이름〉", "〈12〉", "〈10〉", "〈+2〉", "〈충족〉"],
            ["〈구분 C〉", "〈항목 이름〉", "〈7〉", "〈7〉", "〈0〉", "〈충족〉"]],
           row_h=1.15, head_h=1.15, size=10,
           align=["l", "l", "r", "r", "r", "c"])
    yy = y + 1.15 + 1.15 * 5 + 0.8
    b = box(s, M, yy, CW, 1.6, D.PRIMARY_TINT, None, 0)
    rich(b, [("읽는 법   ", {"size": 10, "bold": True, "color": D.PRIMARY}),
             ("〈표에서 말하려는 결론을 한 문장으로 적는다. 표만 두면 읽는 사람이 각자 해석한다.〉",
              {"size": 10.5, "color": D.INK})], align="l", anchor="ctr", margins=(0.6, 0.6, 0.1, 0.1))
    tbox(s, M, yy + 2.0, CW, 1.2, "※ 〈기준일·산출 방법·제외 대상 같은 단서는 표 아래 작은 글씨로 적는다.〉",
         size=9, color=D.INK_FAINT, anchor="t", margins=(0, 0, 0, 0))
    return s

SLIDES = [s_intro, s_parts, s_arch, s_flow, s_steps, s_timeline, s_before_after, s_compare,
          s_status, s_incident, s_rnr, s_screen, s_tree, s_kit, s_table]

def build():
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(12192000), Emu(6858000)
    set_theme_fonts(prs)
    for f in SLIDES: f(prs)
    prs.save(OUT)
    return OUT

if __name__ == "__main__":
    print("PPT 생성:", build())
