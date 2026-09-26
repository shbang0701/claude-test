# -*- coding: utf-8 -*-
"""슬라이드 샘플: 요소를 조합한 완성 슬라이드 8종."""
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.util import Pt
from ppt_kit import (P, I, text, shape, rect, rrect, oval, line, connect, group, icon, slide, table, rgb, line_fmt,
                     notes)
from ppt_elements import (LSTYLE, KIND_TINT, node_card, node_compact, node_tile, node_box, pill, badge, tag, zone,
                          line_legend, color_legend, kpi, usage_bar, steps, checkbox, rack)


def L(key):
    return LSTYLE[key]


# ── S1 시스템 구성도 ─────────────────────────────────────────
def s_system(prs):
    sl = slide(prs, "인터넷뱅킹 시스템 구성도", "샘플 · 시스템 구성도")
    zone(sl, 0.5, 1.35, 2.05, 2.95, "ext", icon_key="internet")
    zone(sl, 0.5, 4.5, 2.05, 1.75, "ext", "대외기관")
    zone(sl, 2.8, 1.35, 3.05, 4.9, "dmz")
    zone(sl, 6.1, 1.35, 4.1, 3.05, "svc")
    zone(sl, 6.1, 4.6, 4.1, 1.65, "san")
    zone(sl, 10.45, 1.35, 2.38, 4.9, "mgmt")
    g, inet = node_compact(sl, 1.25, 2.05, "internet", "인터넷", "ISP-A · ISP-B", 0.56, P["s600"])
    g, ext = node_compact(sl, 1.25, 4.95, "wan", "대외기관", "전용회선 2회선", 0.5, P["s600"])
    g, fw1 = node_card(sl, 3.05, 1.8, "firewall", "FW-01", "외부 방화벽 · Active", P["crit"], w=2.35, h=0.52)
    g, fw2 = node_card(sl, 3.05, 2.4, "firewall", "FW-02", "외부 방화벽 · Standby", P["crit"], w=2.35, h=0.52)
    g, l4 = node_card(sl, 3.05, 3.45, "load-balancer", "L4-01", "VIP 10.10.1.10", P["mgmt"], w=2.55, h=0.52)
    g, w1 = node_card(sl, 3.05, 4.55, "server", "WEB-01", "10.10.1.11", P["svc"], w=2.55, h=0.52)
    g, w2 = node_card(sl, 3.05, 5.15, "server", "WEB-02", "10.10.1.12", P["svc"], w=2.55, h=0.52)
    g, a1 = node_card(sl, 6.35, 1.85, "application", "WAS-01", "10.10.2.21", P["svc"], w=1.8, h=0.52)
    g, a2 = node_card(sl, 6.35, 2.45, "application", "WAS-02", "10.10.2.22", P["svc"], w=1.8, h=0.52)
    g, d1 = node_card(sl, 8.2, 1.85, "database", "DB-01", "Active", P["san"], w=1.8, h=0.52)
    g, d2 = node_card(sl, 8.2, 2.45, "database", "DB-02", "Standby", P["san"], w=1.8, h=0.52)
    g, bt = node_card(sl, 6.35, 3.45, "server", "BAT-01", "배치 서버", P["svc"], w=1.8, h=0.52)
    g, sa = node_card(sl, 8.2, 4.95, "san-switch", "SAN-A", "Fabric A", P["san"], w=1.8, h=0.46)
    g, sb = node_card(sl, 8.2, 5.55, "san-switch", "SAN-B", "Fabric B", P["san"], w=1.8, h=0.46)
    g, st = node_card(sl, 6.35, 5.12, "storage", "STG-01", "Unity XT 480", P["san"], w=1.6, h=0.52)
    g, ms = node_card(sl, 10.65, 1.85, "switch", "MGMT-SW", "OOB 10.10.0.0/24", P["mgmt"], w=2.0, h=0.52)
    g, nm = node_card(sl, 10.65, 2.55, "monitoring", "NMS-01", "통합 모니터링", P["mgmt"], w=2.0, h=0.52)
    g, bk = node_card(sl, 10.65, 3.25, "backup", "BAK-01", "백업 서버", P["bak"], w=2.0, h=0.52)
    g, tp = node_card(sl, 10.65, 3.95, "tape", "TL-01", "테이프 라이브러리", P["bak"], w=2.0, h=0.52)
    node_compact(sl, 11.4, 4.85, "operator", "운영자", "관제실", 0.46, P["s600"], w=1.2)
    connect(sl, inet, 3, fw1, 1, "elbow", *L("wan"))
    connect(sl, ext, 3, fw2, 1, "elbow", *L("wan"), adj=0.8)
    connect(sl, fw1, 3, fw2, 3, "elbow", *L("ha"), bulge=0.2)
    connect(sl, fw2, 2, l4, 0, "elbow", *L("svc"))
    connect(sl, l4, 2, w1, 0, "elbow", *L("svc"))
    connect(sl, l4, 1, w2, 1, "elbow", *L("svc"), bulge=0.14)
    connect(sl, w1, 3, a1, 1, "elbow", *L("svc"), adj=0.44)
    connect(sl, w2, 3, a2, 1, "elbow", *L("svc"), adj=0.57)
    connect(sl, a1, 3, d1, 1, "straight", *L("svc"))
    connect(sl, a2, 3, d2, 1, "straight", *L("svc"))
    connect(sl, d1, 3, d2, 3, "elbow", *L("ha"), bulge=0.1)
    connect(sl, d2, 2, sa, 0, "straight", *L("san"))
    connect(sl, d1, 3, sb, 3, "elbow", *L("san"), bulge=0.17)
    connect(sl, sa, 1, st, 3, "elbow", *L("san"))
    connect(sl, sb, 1, st, 3, "elbow", *L("san"))
    connect(sl, bk, 2, tp, 0, "straight", *L("bak"))
    connect(sl, ms, 1, d2, 3, "elbow", *L("mgmt"), adj=0.45)
    pill(sl, 1.95, 2.02, "1G ×2", P["wan"], w=0.55)
    pill(sl, 4.1, 3.05, "10G", P["svc"], w=0.45)
    pill(sl, 5.33, 2.2, "HA", P["ic"], w=0.4)
    tag(sl, 8.95, 1.6, "Active / Standby", "info", w=1.05, h=0.22, size=7.5)
    line_legend(sl, 0.5, 6.42, keys=["wan", "svc", "san", "mgmt", "bak", "ha"], cols=6, colw=2.05, rowh=0.3, desc=False,
                title_txt=None, gname="범례-연결선")
    notes(sl, "구성도 샘플: 영역(맨 뒤) → 장비 카드 → 연결선 순서로 쌓여 있습니다. 장비를 옮기면 연결선이 따라갑니다.")
    return sl


# ── S2 랙 실장도 ─────────────────────────────────────────────
def s_rack(prs):
    sl = slide(prs, "DC1 B열 랙 실장도 (2026년 10월 변경 반영)", "샘플 · 랙 실장도")
    r02 = [(42, 1, "PP-01 패치패널", "기타"), (41, 1, "MGMT-SW", "네트워크"), (39, 1, "TOR-A", "네트워크"),
           (38, 1, "TOR-B", "네트워크"), (36, 1, "SAN-A", "스토리지"), (35, 1, "SAN-B", "스토리지"),
           (30, 2, "STG-01  Unity XT", "스토리지"), (20, 2, "SRV-01  DL380", "서버"), (18, 2, "SRV-02  DL380", "서버"),
           (16, 1, "SRV-03  R640", "서버"), (1, 10, "ENC-01  c7000", "서버")]
    r03 = [(40, 2, "FW-01", "보안"), (38, 2, "FW-02", "보안"), (35, 1, "L4-01", "네트워크"), (34, 1, "L4-02", "네트워크"),
           (24, 4, "BAK-01  백업 서버", "기타"), (10, 6, "TL-01  테이프", "기타")]
    r04 = [(40, 1, "TOR-C", "네트워크"), (30, 2, "SRV-10  DL380 Gen11", "서버"), (28, 2, "SRV-11  DL380 Gen11", "서버"),
           (20, 2, "OLD-APP-01", "기타")]
    rack(sl, 0.75, 1.75, devices=r02, title_txt="R02", gname="랙-R02")
    rack(sl, 3.45, 1.75, devices=r03, title_txt="R03", gname="랙-R03")
    rack(sl, 6.15, 1.75, devices=r04, title_txt="R04", gname="랙-R04")
    uh = 0.118
    H = 42 * uh

    def mark(rx, ub, h, kind, label):
        y = 1.75 + 0.03 + (42 - (ub + h - 1)) * uh
        col = P["ok"] if kind == "new" else P["crit"]
        rrect(sl, rx + 0.18, y - 0.04, 2.08, h * uh + 0.08, 0.03, fill=None, line=col, lw=1.5, dash="dash",
              name=f"강조-{label}")
        tag(sl, rx + 1.82, y - 0.2, label, kind, solid=True, w=0.48, h=0.2, size=7.5)
    mark(6.15, 28, 4, "new", "신규")
    mark(6.15, 20, 2, "del", "철거")
    x0 = 9.0
    text(sl, x0, 1.4, 3.8, 0.3, "10월 변경 내역", 12, True, P["ink"])
    rows = [["랙", "U", "장비", "변경"], ["R04", "40", "TOR-C", "유지"], ["R04", "30–31", "SRV-10", "신규"],
            ["R04", "28–29", "SRV-11", "신규"], ["R04", "20–21", "OLD-APP-01", "철거"]]
    colors = {(2, 3): P["ok"], (3, 3): P["ok"], (4, 3): P["crit"], (1, 3): P["s500"]}
    table(sl, x0, 1.8, [0.6, 0.75, 1.45, 0.85], rows, row_h=0.3, name="표-변경 내역", colors=colors, bold_cols=(3,),
          align=["l", "l", "l", "c"])
    color_legend(sl, x0, 3.55, [("서버", KIND_TINT["서버"]), ("네트워크", KIND_TINT["네트워크"]), ("스토리지", KIND_TINT["스토리지"])],
                 title_txt="구분", colw=1.2)
    color_legend(sl, x0, 4.15, [("보안", KIND_TINT["보안"]), ("전원", KIND_TINT["전원"]), ("기타", KIND_TINT["기타"])],
                 title_txt=None, colw=1.2)
    text(sl, x0, 4.75, 3.8, 1.8, [
        [("요약", {"bold": True, "color": P["svc"]})],
        "· R04에 신규 서버 2대(4U) 입고, 전원 A/B 연결",
        "· OLD-APP-01 철거 후 U20–21 비움",
        "· R02 여유 U: 19U / R04 여유 U: 33U",
        [("※ 정확한 실장도는 Excel 랙 파일 [랙] 시트 기준", {"color": P["s500"], "size": 8.5})]],
         9.5, False, P["s700"], line_spacing=1.2)
    return sl


# ── S3 서버 물리 연결 ─────────────────────────────────────────
def s_cabling(prs):
    from xlexamples import device_model
    sl = slide(prs, "서버 물리 연결 구성 — SRV-01", "샘플 · 물리 연결")
    n, data = device_model("SRV-01")
    usekey = {"서비스": "svc", "관리": "mgmt", "백업": "bak", "스토리지": "san", "인터커넥트": "ha", "전원": "pwr"}
    colkey = {"서비스": P["svc"], "관리": P["mgmt"], "백업": P["bak"], "스토리지": P["san"], "인터커넥트": P["ic"], "전원": P["pwr"]}
    # 왼쪽: 서버 포트 목록
    sx, sy = 0.5, 1.35
    rrect(sl, sx, sy, 3.7, 5.45, 0.08, fill=P["s50"], line=P["s300"], lw=0.75, name="서버 바탕")
    icon(sl, "server", sx + 0.15, sy + 0.14, 0.4, P["svc"])
    text(sl, sx + 0.65, sy + 0.12, 3, 0.25, "SRV-01", 12, True, P["ink"])
    text(sl, sx + 0.65, sy + 0.37, 3, 0.2, "HPE DL380 Gen10 · R02 U20–21", 8.5, False, P["s500"])
    text(sl, sx + 0.15, sy + 0.72, 0.4, 0.2, "번호", 7.5, True, P["s500"], align="c")
    text(sl, sx + 0.6, sy + 0.72, 1.2, 0.2, "포트명", 7.5, True, P["s500"])
    text(sl, sx + 1.55, sy + 0.72, 1.3, 0.2, "케이블 라벨", 7.5, True, P["s500"])
    ports = []
    for k, d in enumerate(data):
        y = sy + 0.98 + k * 0.43
        col = colkey[d["용도"]]
        b = rrect(sl, sx + 0.15, y, 0.4, 0.3, 0.03, fill=col, line=None, txt=str(d["번호"]), size=9, bold=True,
                  color=P["white"], name=f"번호-{d['번호']}")
        text(sl, sx + 0.62, y, 0.9, 0.3, d["포트명"], 9, False, P["ink"], anchor="m")
        rrect(sl, sx + 1.55, y + 0.03, 1.05, 0.24, 0.02, fill="FFF1BF", line=None, txt=d["케이블 라벨"], size=8,
              color=P["ink"], name=f"라벨-{d['케이블 라벨']}")
        anchor = rect(sl, sx + 3.3, y + 0.05, 0.2, 0.2, fill=P["white"], line=col, lw=1.0, name=f"포트-{d['포트명']}")
        ports.append((anchor, d))
    # 오른쪽: 상대 장비
    order = []
    for _, d in ports:
        if d["상대 장비"] not in order:
            order.append(d["상대 장비"])
    icons_ = {"SAN-A": "san-switch", "SAN-B": "san-switch", "TOR-A": "switch", "TOR-B": "switch", "MGMT-SW": "switch",
              "PP-01": "patch-panel", "PDU-A": "pdu", "PDU-B": "pdu"}
    tgt = {}
    ty0 = 1.45
    step = (6.35 - 0.5 - ty0) / max(1, len(order) - 1)
    for k, dev in enumerate(order):
        loc = [d["상대 위치"] for _, d in ports if d["상대 장비"] == dev][0]
        g, bg = node_card(sl, 9.6, ty0 + k * step, icons_.get(dev, "server"), dev, loc, P["navy"], w=2.2, h=0.5)
        tgt[dev] = bg
    # 연결선 (꺾임 위치를 조금씩 달리해 겹치지 않게)
    n_ = len(ports)
    for k, (anchor, d) in enumerate(ports):
        c, w, dsh = L(usekey[d["용도"]])
        connect(sl, anchor, 3, tgt[d["상대 장비"]], 1, "elbow", c, w, dsh, adj=0.2 + 0.6 * k / max(1, n_ - 1),
                name=f"연결-{d['포트명']}")
        # 상대 포트 라벨
    for dev, bg in tgt.items():
        pp = [d["상대 포트"] for _, d in ports if d["상대 장비"] == dev]
        s = " / ".join(pp)
        w = 0.1 + 0.075 * len(s)
        rect(sl, 9.5 - w, bg.top / 914400 + 0.14, w, 0.22, fill=P["white"], line=None, txt=s, size=8, bold=True,
             color=P["s700"], margins=(0.03, 0.03, 0, 0), name=f"상대 포트-{dev}")
    line_legend(sl, 4.6, 6.6, keys=["svc", "san", "mgmt", "bak", "pwr"], cols=5, colw=[1.35, 1.8, 1.6, 1.2, 1.0],
                rowh=0.3, desc=False,
                title_txt=None, gname="범례-연결선")
    text(sl, 4.6, 1.4, 4.5, 0.5, [[("번호 · 라벨은 Excel 예시(2_랙예시 · U20 SRV-01)와 같습니다", {"color": P["s500"]})]],
         8.5, False, P["s500"])
    return sl


# ── S4 월간 운영 보고 ─────────────────────────────────────────
def s_monthly(prs):
    sl = slide(prs, "2026년 9월 인프라 운영 현황", "샘플 · 월간 운영 보고")
    for k, (lab, val, unit, delta, good) in enumerate((("서비스 가용률", "99.98", "%", "▲ 목표 99.95% 달성", True),
                                                       ("장애", "1", "건", "▼ 전월 3건", True),
                                                       ("변경 작업", "14", "건", "계획 16 · 연기 2", True),
                                                       ("스토리지 사용률", "86", "%", "▲ 증설 검토 필요", False))):
        kpi(sl, 0.5 + k * 3.1, 1.35, lab, val, unit, delta, good, h=1.05)
    text(sl, 0.5, 2.62, 4, 0.25, "월별 장애 · 변경", 11, True, P["ink"])
    cd = CategoryChartData()
    cd.categories = ["4월", "5월", "6월", "7월", "8월", "9월"]
    cd.add_series("장애", (3, 2, 4, 1, 3, 1))
    cd.add_series("변경", (11, 9, 15, 12, 16, 14))
    gf = sl.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, I(0.5), I(2.92), I(5.4), I(2.1), cd)
    gf.name = "차트-월별"
    ch = gf.chart
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.TOP
    ch.legend.include_in_layout = False
    ch.legend.font.size = Pt(8.5)
    ch.legend.font.name = "맑은 고딕"
    for s, col in zip(ch.series, (P["crit"], P["svc"])):
        s.format.fill.solid()
        s.format.fill.fore_color.rgb = rgb(col)
    pl = ch.plots[0]
    pl.gap_width = 70
    pl.has_data_labels = True
    pl.data_labels.font.size = Pt(8)
    pl.data_labels.font.name = "맑은 고딕"
    pl.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
    ch.category_axis.tick_labels.font.size = Pt(8.5)
    ch.category_axis.tick_labels.font.name = "맑은 고딕"
    ch.category_axis.format.line.color.rgb = rgb(P["s300"])
    ch.value_axis.visible = False
    ch.value_axis.has_major_gridlines = False
    text(sl, 6.3, 2.62, 4, 0.25, "주요 이슈", 11, True, P["ink"])
    rows = [["일자", "구분", "내용", "상태"],
            ["09/03", "변경", "TOR 펌웨어 업그레이드 (R02)", "완료"],
            ["09/14", "장애", "SAN-B 포트 12 오류 → SFP 교체", "완료"],
            ["09/21", "변경", "SRV-10·11 신규 입고 (R04)", "완료"],
            ["09/28", "점검", "3분기 정기 점검 (42대)", "진행"],
            ["10월", "계획", "스토리지 증설 검토 · 발주", "예정"]]
    colors = {(1, 3): P["ok"], (2, 3): P["ok"], (3, 3): P["ok"], (4, 3): P["info"], (5, 3): P["s500"], (2, 1): P["crit"]}
    table(sl, 6.3, 2.95, [0.7, 0.65, 4.1, 0.9], rows, row_h=0.33, name="표-주요 이슈", colors=colors, bold_cols=(3,),
          align=["l", "l", "l", "c"])
    text(sl, 0.5, 5.25, 4, 0.25, "자원 사용률", 11, True, P["ink"])
    for k, (lab, pct) in enumerate((("CPU", 38), ("메모리", 72), ("스토리지", 86))):
        usage_bar(sl, 0.5, 5.6 + k * 0.36, lab, pct, w=5.4)
    text(sl, 6.3, 5.25, 4, 0.25, "다음 달 계획", 11, True, P["ink"])
    text(sl, 6.3, 5.6, 6.5, 1.2, ["· 스토리지 증설 (Unity XT DAE 1식) — 10월 3주",
                                  "· OLD-APP-01 철거 및 R04 U20–21 정리 — 10월 2주",
                                  "· DR 전환 훈련 — 10월 4주 (토) 22:00"], 10, False, P["ink"], line_spacing=1.25)
    return sl


# ── S5 작업 계획서 ───────────────────────────────────────────
def s_workplan(prs):
    sl = slide(prs, "SAN 스위치 교체 작업 계획", "샘플 · 작업 계획서")
    rows = [["항목", "내용"], ["작업명", "SAN-B (Brocade G620) 교체 — 포트 오류 재발"], ["일시", "2026-10-17 (토) 22:00 – 24:00"],
            ["영향", "Fabric B 경로 중단 (Fabric A로 무중단 운영)"], ["담당", "인프라팀 홍길동 · 협력사 김철수"],
            ["승인", "CR-2026-1017 (변경관리위원회 10/10)"]]
    table(sl, 0.5, 1.35, [1.0, 4.9], rows, row_h=0.34, name="표-작업 개요", bold_cols=(0,), zebra=True)
    steps(sl, 6.7, 1.35, [("사전", "경로 이중화 확인"), ("분리", "Fabric B 포트 비활성"), ("교체", "장비 교체 · 설정 복원"),
                          ("확인", "경로 · 로그 확인")], w=6.13, h=0.5)
    text(sl, 6.7, 2.55, 3, 0.25, "영향 받는 시스템", 11, True, P["ink"])
    for k, (nm, st) in enumerate((("DB-01 · 02", "경로 1개로 운영"), ("SRV-01 · 02", "경로 1개로 운영"), ("ENC-01 (블레이드 6대)", "경로 1개로 운영"),
                                  ("STG-01", "SP A·B 모두 Fabric A 유지"))):
        y = 2.9 + k * 0.33
        tag(sl, 6.7, y, "경로 감소", "warn", w=0.85, h=0.24, size=8)
        text(sl, 7.65, y, 2.2, 0.24, nm, 9.5, True, P["ink"], anchor="m")
        text(sl, 9.85, y, 3.0, 0.24, st, 9, False, P["s500"], anchor="m")
    text(sl, 0.5, 3.6, 3, 0.25, "시간표", 11, True, P["ink"])
    y = 4.35
    line(sl, 0.7, y, 5.9, y, P["s300"], 2.0)
    for k, (t, d) in enumerate((("22:00", "공지 · 사전 점검"), ("22:20", "Fabric B 분리"), ("22:40", "교체 · 설정"),
                                ("23:30", "정상 확인"), ("24:00", "종료 보고"))):
        x = 0.85 + k * 1.25
        oval(sl, x - 0.08, y - 0.08, 0.16, 0.16, fill=P["svc"], line=P["white"], lw=1.5)
        text(sl, x - 0.55, y - 0.4, 1.1, 0.22, t, 10, True, P["ink"], align="c")
        text(sl, x - 0.62, y + 0.14, 1.24, 0.36, d, 8.5, False, P["s700"], align="c")
    text(sl, 6.7, 4.35, 3, 0.25, "롤백 계획", 11, True, P["crit"])
    g = [rrect(sl, 6.7, 4.7, 6.13, 1.05, 0.06, fill=P["t_crit"], line=None, name="바탕"),
         text(sl, 6.9, 4.8, 5.8, 0.9, ["· 기준: 교체 후 30분 내 포트 링크 미기동, 또는 경로 복구 실패",
                                       "· 방법: 기존 장비 재장착 → 설정 백업 복원 → Fabric B 포트 활성",
                                       "· 소요: 약 20분 (최종 결정: 인프라팀장)"], 9.5, False, P["ink"], line_spacing=1.2)]
    group(sl, g, "롤백 계획")
    text(sl, 0.5, 5.05, 3, 0.25, "사전 체크리스트", 11, True, P["ink"])
    for k, (t, c) in enumerate((("설정 백업 (configupload)", True), ("존 · 별칭 목록 출력", True), ("예비 SFP 4개", False),
                                ("협력사 출입 신청", True))):
        checkbox(sl, 0.5 + (k % 2) * 2.95, 5.4 + (k // 2) * 0.34, t, c, w=2.9)
    text(sl, 6.7, 5.95, 6.1, 0.8, [[("비상 연락", {"bold": True, "color": P["svc"]}),
                                   ("  인프라팀 홍길동 010-0000-0000 · 협력사 김철수 010-0000-0001 · 관제 02-000-0000", {})]],
         9, False, P["s700"])
    return sl


# ── S6 변경 전·후 ────────────────────────────────────────────
def s_beforeafter(prs):
    sl = slide(prs, "TOR 이중화 구성 변경 — 전 · 후", "샘플 · 변경 전후")
    for k, ttl in enumerate(("변경 전 (As-Is)", "변경 후 (To-Be)")):
        x = 0.5 + k * 6.25
        rrect(sl, x, 1.35, 6.05, 4.3, 0.08, fill=P["s50"] if k == 0 else P["t_info"], line=None, name="패널")
        text(sl, x + 0.2, 1.47, 4, 0.3, ttl, 13, True, P["ink"] if k == 0 else P["svc"])
    # 전
    g, up = node_card(sl, 2.05, 1.95, "switch-l3", "CORE-01", "", P["s500"], w=2.0, h=0.5)
    g, t1 = node_card(sl, 2.05, 3.05, "switch", "TOR-A", "단일", P["s500"], w=2.0, h=0.5)
    servers = []
    for k in range(3):
        g, s = node_card(sl, 0.8 + k * 1.9, 4.3, "server", f"SRV-0{k+1}", "", P["s500"], w=1.7, h=0.48)
        servers.append(s)
    connect(sl, up, 2, t1, 0, "straight", *L("svc"))
    for s in servers:
        connect(sl, t1, 2, s, 0, "elbow", *L("svc"))
    tag(sl, 4.25, 3.15, "단일 장애점", "crit", w=0.95, h=0.24, size=8)
    # 후
    ox = 6.25
    g, up2 = node_card(sl, 2.05 + ox, 1.95, "switch-l3", "CORE-01", "", P["navy"], w=2.0, h=0.5)
    g, ta = node_card(sl, 0.9 + ox, 3.05, "switch", "TOR-A", "vPC 1", P["mgmt"], w=1.9, h=0.5)
    g, tb = node_card(sl, 3.3 + ox, 3.05, "switch", "TOR-B", "vPC 1", P["mgmt"], w=1.9, h=0.5)
    tag(sl, 4.72 + ox, 2.87, "신규", "new", solid=True, w=0.45, h=0.2, size=7.5)
    line_fmt(tb, P["ok"], 1.25, "dash")
    servers2 = []
    for k in range(3):
        g, s = node_card(sl, 0.8 + ox + k * 1.9, 4.3, "server", f"SRV-0{k+1}", "", P["svc"], w=1.7, h=0.48)
        servers2.append(s)
    connect(sl, up2, 2, ta, 0, "elbow", *L("svc"))
    connect(sl, up2, 2, tb, 0, "elbow", *L("svc"))
    connect(sl, ta, 3, tb, 1, "straight", *L("ha"))
    pill(sl, 2.8 + ox + 0.02, 3.2, "피어링크", P["ic"], w=0.62, h=0.2, size=7.5)
    for s in servers2:
        connect(sl, ta, 2, s, 0, "elbow", *L("svc"), adj=0.4)
        connect(sl, tb, 2, s, 0, "elbow", *L("svc"), adj=0.6)
    shape(sl, MSO_SHAPE.RIGHT_ARROW, 6.02, 3.25, 0.45, 0.5, fill=P["svc"], line=None, name="화살표")
    text(sl, 0.5, 5.85, 3, 0.25, "변경 효과", 11, True, P["ink"])
    for k, (a, b) in enumerate((("가용성", "스위치 1대 장애에도 서비스 유지 (단일 장애점 제거)"), ("대역폭", "서버당 10G → 25G × 2 (LACP)"),
                                ("작업", "10/24 (토) 22:00 · 서버별 5분 이내 순단")),):
        text(sl, 0.5 + k * 4.15, 6.15, 1.0, 0.5, a, 10, True, P["svc"])
        text(sl, 1.35 + k * 4.15, 6.15, 3.1, 0.6, b, 9.5, False, P["ink"], line_spacing=1.1)
    return sl


# ── S7 장애 보고 ─────────────────────────────────────────────
def s_incident(prs):
    sl = slide(prs, "장애 보고 — SAN-B 포트 오류 (09/14)", "샘플 · 장애 보고서")
    g = [rrect(sl, 0.5, 1.35, 12.33, 0.8, 0.06, fill=P["t_crit"], line=None, name="바탕"),
         icon(sl, "error", 0.7, 1.5, 0.5, P["crit"])]
    for k, (a, b) in enumerate((("영향", "서비스 영향 없음 (Fabric A 경로 유지)"), ("등급", "3등급 (단일 경로 장애)"),
                                ("시간", "14:02 발생 → 14:48 복구 (46분)"), ("원인", "SFP 모듈 불량 (포트 12)"))):
        g.append(text(sl, 1.45 + k * 2.85, 1.47, 2.8, 0.22, a, 9, True, P["crit"]))
        g.append(text(sl, 1.45 + k * 2.85, 1.7, 2.8, 0.4, b, 9.5, False, P["ink"], line_spacing=1.05))
    group(sl, g, "장애 요약")
    text(sl, 0.5, 2.4, 3, 0.25, "경과", 11, True, P["ink"])
    y = 3.1
    line(sl, 0.7, y, 12.6, y, P["s300"], 2.0)
    ev = (("14:02", "발생", "SAN-B P12 링크 다운 · 경로 경고", "crit"), ("14:05", "인지", "NMS 알람 → 관제 전파", "warn"),
          ("14:15", "확인", "SFP 오류 카운터 증가 확인", "info"), ("14:40", "조치", "예비 SFP 교체", "info"),
          ("14:48", "복구", "링크 · 경로 정상 확인", "ok"), ("15:30", "종료", "관련 부서 공유", "ok"))
    for k, (t, a, d, st) in enumerate(ev):
        x = 0.9 + k * 2.3
        col = {"crit": P["crit"], "warn": P["warn"], "info": P["svc"], "ok": P["ok"]}[st]
        oval(sl, x - 0.1, y - 0.1, 0.2, 0.2, fill=col, line=P["white"], lw=1.5)
        text(sl, x - 0.6, y - 0.45, 1.2, 0.22, t, 10, True, P["ink"], align="c")
        text(sl, x - 0.9, y + 0.17, 1.8, 0.22, a, 9.5, True, col, align="c")
        text(sl, x - 1.0, y + 0.4, 2.0, 0.4, d, 8.5, False, P["s700"], align="c")
    text(sl, 0.5, 4.2, 3, 0.25, "원인 · 조치", 11, True, P["ink"])
    text(sl, 0.5, 4.55, 5.9, 1.9, ["· 원인: SFP(16G SW) 모듈 수광 레벨 저하 → CRC 오류 누적 후 링크 다운",
                                   "· 조치: 예비 SFP 교체, 포트 통계 초기화 후 1시간 모니터링",
                                   "· 확인: 모든 호스트 경로 2개 복구 (multipath 확인)"], 9.5, False, P["ink"], line_spacing=1.3)
    text(sl, 6.7, 4.2, 3, 0.25, "재발 방지", 11, True, P["ink"])
    rows = [["대책", "담당", "기한", "상태"], ["SFP 수광 레벨 임계치 알람 추가", "운영팀", "09/30", "완료"],
            ["동일 로트 SFP 전수 점검", "협력사", "10/15", "진행"], ["SAN-B 교체 (노후)", "인프라팀", "10/17", "예정"]]
    colors = {(1, 3): P["ok"], (2, 3): P["info"], (3, 3): P["s500"]}
    table(sl, 6.7, 4.55, [3.2, 0.95, 0.9, 1.08], rows, row_h=0.33, name="표-재발 방지", colors=colors, bold_cols=(3,),
          align=["l", "l", "l", "c"])
    return sl


# ── S8 점검 결과 ─────────────────────────────────────────────
def s_inspection(prs):
    sl = slide(prs, "3분기 정기 점검 결과", "샘플 · 점검 결과 보고")
    for k, (lab, val, unit, col) in enumerate((("점검 대상", "42", "대", P["ink"]), ("정상", "39", "대", P["ok"]),
                                               ("주의", "2", "대", P["warn"]), ("조치 필요", "1", "대", P["crit"]))):
        g = [rrect(sl, 0.5 + k * 3.1, 1.35, 2.85, 0.95, 0.07, fill=P["white"], line=P["s200"], lw=0.75, name="바탕"),
             text(sl, 0.7 + k * 3.1, 1.47, 2.5, 0.2, lab, 9.5, False, P["s500"]),
             text(sl, 0.7 + k * 3.1, 1.7, 2.5, 0.5, [[(val, {"size": 24, "bold": True, "color": col}),
                                                       (" " + unit, {"size": 11, "color": P["s500"]})]], 24, True, col,
                  anchor="b")]
        group(sl, g, f"요약-{lab}")
    text(sl, 0.5, 2.55, 4, 0.25, "주의 · 조치 대상", 11, True, P["ink"])
    rows = [["장비", "위치", "항목", "결과", "상태", "조치"],
            ["STG-01", "R02 U30", "디스크", "드라이브 7 예측 장애", "주의", "10/05 교체"],
            ["SAN-B", "R02 U35", "포트", "P12 CRC 오류 재발", "조치", "10/17 장비 교체"],
            ["SRV-03", "R02 U16", "펌웨어", "iDRAC 권고 버전 미적용", "주의", "10월 정기 작업"]]
    colors = {(1, 4): P["warn"], (2, 4): P["crit"], (3, 4): P["warn"]}
    table(sl, 0.5, 2.9, [1.1, 1.0, 0.9, 2.3, 0.8, 1.6], rows, row_h=0.36, name="표-점검 대상", colors=colors,
          bold_cols=(4,), align=["l", "l", "l", "l", "c", "l"])
    text(sl, 0.5, 4.55, 4, 0.25, "항목별 결과", 11, True, P["ink"])
    for k, (lab, pct) in enumerate((("전원 이중화", 100), ("팬 · 온도", 100), ("디스크", 98), ("펌웨어", 93), ("로그 · 알람", 95))):
        usage_bar(sl, 0.5, 4.9 + k * 0.36, lab, pct, w=7.7, color=P["ok"] if pct == 100 else (P["warn"] if pct >= 95 else P["crit"]))
    text(sl, 8.6, 2.55, 4, 0.25, "점검 기준", 11, True, P["ink"])
    for k, (t, c) in enumerate((("PSU 2/2 정상 · A/B 계통 분리", True), ("입구 온도 27℃ 이하", True), ("디스크 오류 · 예측 장애 없음", False),
                                ("펌웨어 권고 버전", False), ("시간 동기화 (NTP)", True), ("관리 포트 접속", True))):
        checkbox(sl, 8.6, 2.92 + k * 0.36, t, c, w=4.2)
    g = [rrect(sl, 8.6, 5.2, 4.23, 1.2, 0.06, fill=P["t_info"], line=None, name="바탕"),
         icon(sl, "info", 8.75, 5.35, 0.3, P["info"]),
         text(sl, 9.2, 5.3, 3.5, 1.05, ["종합 의견", "조치 필요 1건은 10/17 작업에 포함. 주의 2건은 10월 정기 작업으로 처리."],
              9.5, False, P["ink"], line_spacing=1.2)]
    group(sl, g, "종합 의견")
    return sl


SAMPLES = [("시스템 구성도", s_system), ("랙 실장도", s_rack), ("서버 물리 연결", s_cabling), ("월간 운영 보고", s_monthly),
           ("작업 계획서", s_workplan), ("변경 전 · 후 비교", s_beforeafter), ("장애 보고서", s_incident),
           ("점검 결과 보고", s_inspection)]
