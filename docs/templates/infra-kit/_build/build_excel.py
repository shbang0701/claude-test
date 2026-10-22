# -*- coding: utf-8 -*-
"""Excel 산출물 3종 생성.
  1_장비도면_라이브러리.xlsx  — 사용법·기준·부품·빈 양식·작성 예시
  2_랙예시_DC1-R02.xlsx      — 랙 1대 = 파일 1개 관리 예시
  3_랙_빈양식.xlsx           — 새 랙 파일 시작용
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from xlkit import new_workbook, postprocess
from xlguide import build_guide, build_standard, build_index
from xlpalette import build_palette
from xlrack import build_rack
from sheets import device_sheet, tower_sheet, free_sheet, small_sheet, TAB_FORM, TAB_EX
from xlcatalog import TYPES, example_sheet, TODAY
from xlexamples import DEV, RACK, loc
import xldevices as D

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "01_엑셀_장비도면")


NOTE_GPU = ("공개 자료(QuickSpecs 요약) 기준 초안 — 전면 NVMe 8 · 관리 MLAN 전면 · 후면 PSU 6 · 팬 6 · GPU 네트워크 슬롯 8"
            "(슬레드 2)  ·  PCIe x16 슬롯 위치(앞/뒤)는 현장 확인")
NOTE_SD2 = ("구성 요소 수 = 공개 자료 요약(블레이드 8 · PSU 12 · 팬 15 · XFM 4 · OA 2 · GPSM 2) · 배치 · IC 베이 수 · AC 입력 수는 현장 확인"
            "  ·  IOX는 [양식-IO확장박스]")
NOTE_IOX = ("IBM EMX0: I/O 모듈 = 팬아웃 모듈 P1 · P2, 본체와 T1 · T2 한 쌍씩  /  HPE Superdome 2 IOX: I/O 베이 1 · 2 (각 6슬롯), "
            "XFM 링크 케이블 + CAMnet(RJ45)")


def forms(wb):
    F = [
        ("양식-1U서버", "1U 랙 서버", 1, D.srv1u_front, lambda cv, r, c, n: D.srv1u_rear(cv, r, c, n, slots={}), {}),
        ("양식-2U서버", "2U 랙 서버", 2, lambda cv, r, c, n: D.srv2u_front(cv, r, c, n, boxes=(0, 0, 0)),
         lambda cv, r, c, n: D.srv2u_rear(cv, r, c, n, slots={}), {}),
        ("양식-4U서버", "4U 서버 (GPU·스토리지형)", 4, lambda cv, r, c, n: D.srv4u_front(cv, r, c, n, filled=0),
         D.srv4u_rear, {}),
        ("양식-GPU서버5U", "5U GPU 서버 (예: HPE Cray XD670)", 5, D.gpu5u_front, D.gpu5u_rear, dict(foot_note=NOTE_GPU)),
        ("양식-멀티노드", "2U 4노드 섀시 (노드마다 BMC · LAN)", 2, D.node4_front, D.node4_rear,
         dict(foot_note="도면 번호 = 노드 × 100 + 포트 (예: 노드 2의 P1 = 202) · 노드 번호 위치는 장비 표기 확인")),
        ("양식-Power본체", "IBM Power 4U 본체 (슬롯 배치는 모델별 확인)", 4, D.power4u_front, D.power4u_rear,
         dict(role="본체", foot_note="본체 표시: 전면 조작 패널(LCD) · 후면 HMC 포트  ·  I/O 드로어는 [양식-IO확장박스]")),
        ("양식-IO확장박스", "I/O 확장 박스 4U (IBM EMX0 · HPE Superdome 2 IOX 등)", 4, D.iox_front, D.iox_rear,
         dict(role="I/O 드로어", foot_note=NOTE_IOX)),
        ("양식-대형섀시", "대형 블레이드 섀시 18U (예: HPE Superdome 2)", 18, D.sd2_front, D.sd2_rear,
         dict(half=True, nrows=30, foot_note=NOTE_SD2)),
        ("양식-1U스위치", "1U 스위치 48 × SFP + 6 × QSFP", 1, TYPES["sw48"]["front"], D.sw_rear, dict(half=True, nrows=24)),
        ("양식-반폭스위치", "반폭 1U 스위치 12 × SFP+ + 3 × QSFP28 (트레이에 2대)", 1, _half_front, D.half_sw_rear,
         dict(half=True, nrows=16, foot_note="반폭 장비 2대 = 트레이 1개: 랙 목록에는 트레이를 구분 '선반'으로 한 줄, 스위치 2대는 면 = 선반")),
        ("양식-관리스위치", "1U 스위치 48 × RJ45 + 4 × SFP", 1, TYPES["rj48"]["front"], D.sw_rear, dict(half=True, nrows=24)),
        ("양식-SAN스위치", "1U SAN 스위치 48 × SFP (8포트 그룹)", 1, TYPES["san48"]["front"], D.san_rear, dict(half=True, nrows=24)),
        ("양식-스토리지", "2U 스토리지 컨트롤러 (SFF 25)", 2, lambda cv, r, c, n: D.stg_front(cv, r, c, n, filled=0),
         D.stg_rear, dict(role="컨트롤러")),
        ("양식-디스크선반", "2U 디스크 확장 선반 (DAE · JBOD, SFF 25)", 2, D.dae_front, D.dae_rear, dict(role="디스크 선반")),
        ("양식-블레이드섀시", "블레이드 섀시 10U", 10, lambda cv, r, c, n: D.blade_front(cv, r, c, n, bays={}),
         lambda cv, r, c, n: D.blade_rear(cv, r, c, n, ic={}), dict(half=True, nrows=30)),
        ("양식-모듈형섀시", "모듈형 섀시 스위치 8U (라인카드)", 8, D.chassis_front, D.chassis_rear, dict(half=True, nrows=30)),
        ("양식-패치패널", "1U 패치패널 · 소형 장비", 1, TYPES["pp24"]["front"], TYPES["pp24"]["rear"], {}),
    ]
    for sheet, model, u, fr, rr, kw in F:
        device_sheet(wb, sheet, "장비명", model, u, fr, rr, tab=TAB_FORM, **kw)
    tower_sheet(wb, "양식-타워", "장비명", "타워 서버", slots={}, filled=0)
    small_sheet(wb, "양식-소형장비", "장비명", "소형 · 데스크톱 장비 (예: 16포트 데스크톱 스위치 · 미니 PC)",
                lambda cv, r, c, n: D.small_front(cv, r, c, n, conn={}),
                lambda cv, r, c, n: D.small_rear(cv, r, c, n, dc=True),
                foot_note="틀 너비 = 실제 폭(mm) ÷ 8 칸 (예: 160mm → 20칸) · 선반 위 장비는 랙 목록에서 면 = 선반 · 포트 수는 부품으로 조정")
    device_sheet(wb, "양식-PDU", "장비명", "0U PDU (세로형 → 가로로 표시)", 1, TYPES["pdu"]["front"], None, rows=6,
                 half=True, nrows=24, labels=("콘센트", ""),
                 foot_note="0U 세로형 PDU를 가로로 눕혀 그림 · 칸 안 숫자 = PDU에 인쇄된 콘센트 번호 · 사용 중인 콘센트만 색 지정")
    free_sheet(wb)


def _half_front(cv, r, c, n):
    from xlexamples import conn_from_n
    return D.half_sw_front(cv, r, c, n, conn=conn_from_n(n))


def tower_example(wb, sheet="예시-WS-01", back=False):
    n = {"bmc": ("mgmt", 1), "lom1": ("svc", 2), "lom2": ("bak", 3), "s1p1": ("svc", 4), "s1p2": ("svc", 5),
         "psu1": ("pwr", 6), "psu2": ("pwr", 7)}
    rows = [
        (1, "후면", "iLO", "관리", "UTP Cat6 / RJ45", "OPS-SW", "2F 운영실 R1", "Gi1/0/3", "OP-D003", ""),
        (2, "후면", "LOM1", "서비스", "UTP Cat6 / RJ45", "OPS-SW", "2F 운영실 R1", "Gi1/0/4", "OP-D004", ""),
        (3, "후면", "LOM2", "백업", "UTP Cat6 / RJ45", "OPS-SW", "2F 운영실 R1", "Gi1/0/28", "OP-D005", "백업 VLAN"),
        (4, "후면", "S1-P1", "서비스", "OM4 LC-LC", "OPS-SW", "2F 운영실 R1", "Te1/1/1", "OP-D006", ""),
        (5, "후면", "S1-P2", "서비스", "OM4 LC-LC", "OPS-SW", "2F 운영실 R1", "Te1/1/2", "OP-D007", ""),
        (6, "후면", "PSU1", "전원", "C13-C14 1.8m", "UPS 콘센트", "2F 운영실 벽면", "A-3", "OP-P001", "A 계통"),
        (7, "후면", "PSU2", "전원", "C13-C14 1.8m", "UPS 콘센트", "2F 운영실 벽면", "B-3", "OP-P002", "B 계통"),
    ]
    keys = ["번호", "면", "포트명", "용도", "케이블 · 커넥터", "상대 장비", "상대 위치", "상대 포트", "케이블 라벨", "비고"]
    data = [dict(zip(keys, x)) for x in rows]
    info = {"설치 위치": "DC1 2F 운영실", "랙 · U": "바닥 (타워)", "구분": "서버", "시리얼": "CZ2A01BCDE",
            "관리 IP": "10.20.0.15", "작성": TODAY}
    tower_sheet(wb, sheet, "WS-01", "HPE ProLiant ML350 Gen10 (타워)", n=n, info=info, data=data, nrows=10, tab=TAB_EX,
                slots={1: ("NIC 2P", 2), 3: ("GPU", 0)}, filled=8, back_link=back, cap="4T SATA",
                raid=[(1, 2, "RAID1 · OS"), (3, 8, "RAID5 · DATA")])


def chassis_example(wb, sheet="예시-CORE-01"):
    conn = {}
    n = {}
    rows = []
    lab = 1
    for s, peers in ((1, ["TOR-A", "TOR-B", "TOR-C", "TOR-D"]), (2, ["TOR-A", "TOR-B", "TOR-C", "TOR-D"])):
        for i, peer in enumerate(peers, 1):
            num = s * 100 + i
            n[f"p{num}"] = ("svc", num)
            rows.append((num, "전면", f"Eth{s}/{i}", "서비스", "QSFP28 AOC 10m", peer.replace("TOR", f"R0{i+1}-TOR"),
                         f"R0{i+1} U39" if s == 1 else f"R0{i+1} U38", "Eth1/49" if s == 1 else "Eth1/50",
                         f"R10-D{lab:03d}", ""))
            lab += 1
    n["sup5_mgmt"] = ("mgmt", 501)
    n["sup6_mgmt"] = ("mgmt", 601)
    rows.append((501, "전면", "SUP A mgmt0", "관리", "UTP Cat6 / RJ45", "OOB-CORE", "R10 U42", "Gi1/0/1", f"R10-D{lab:03d}", ""))
    rows.append((601, "전면", "SUP B mgmt0", "관리", "UTP Cat6 / RJ45", "OOB-CORE", "R10 U42", "Gi1/0/2", f"R10-D{lab+1:03d}", ""))
    for i in range(1, 5):
        n[f"psu{i}"] = ("pwr", 900 + i)
        rows.append((900 + i, "후면", f"PSU{i}", "전원", "C19-C20 2m", "R10-PDU-" + ("A" if i <= 2 else "B"),
                     "R10 후면 " + ("좌" if i <= 2 else "우"), f"C19-{25 + (i - 1) % 2}", f"R10-P{i:03d}",
                     "A 계통" if i <= 2 else "B 계통"))
    keys = ["번호", "면", "포트명", "용도", "케이블 · 커넥터", "상대 장비", "상대 위치", "상대 포트", "케이블 라벨", "비고"]
    data = [dict(zip(keys, x)) for x in rows]
    conn = {int(k[1:]): v[0] for k, v in n.items() if k.startswith("p") and k[1:].isdigit()}
    info = {"설치 위치": "DC1 2F 전산실", "랙 · U": "R10 · U01–08", "구분": "네트워크", "시리얼": "FOX2412P0AB",
            "관리 IP": "10.10.0.10", "작성": TODAY}
    device_sheet(wb, sheet, "CORE-01", "모듈형 섀시 스위치 8U (예: Nexus 9504급)", 8,
                 lambda cv, r, c, n_: D.chassis_front(cv, r, c, n_, cards={1: ("36 × QSFP28 100G", 36), 2: ("36 × QSFP28 100G", 36)}, conn=conn),
                 D.chassis_rear, n=n, info=info, data=data, nrows=16, tab=TAB_EX,
                 foot_note="도면 번호 = 슬롯 × 100 + 포트 (예: 슬롯 2의 3번 = 203) · 슈퍼바이저 관리포트 = 501/601 · PSU = 901–904")


LIB_EXAMPLES = ["SRV-01", "SRV-03", "TOR-A", "MGMT-SW", "SAN-A", "STG-01", "STG-01-DAE1", "AIX-01", "AIX-01-IO1",
                "ENC-01", "PDU-A", "PP-01"]
RACK_ORDER = ["PP-01", "MGMT-SW", "TOR-A", "TOR-B", "SAN-A", "SAN-B", "STG-01-DAE1", "STG-01", "AIX-01-IO1", "AIX-01",
              "SRV-01", "SRV-02", "SRV-03", "ENC-01", "PDU-A", "PDU-B"]


def rack_sheet_name(dev):
    d = DEV[dev]
    return f"0U {dev}" if d["u"] is None else f"U{d['u']:02d} {dev}"


FORM_DESC = {"양식-1U서버": "1U 랙 서버 (SFF 10 · 띠형 슬롯 3)", "양식-2U서버": "2U 랙 서버 (SFF 24 · 라이저 8슬롯)",
             "양식-4U서버": "4U 서버 (LFF 24 · 슬롯 8 · PSU 4) · 2쪽",
             "양식-GPU서버5U": "5U GPU 서버 — XD670형 초안 (전면 관리포트 · NVMe 8) · 3쪽",
             "양식-멀티노드": "2U 4노드 (노드×100+포트)", "양식-Power본체": "IBM Power 4U 본체 (조작 패널 · HMC 포트) · 2쪽",
             "양식-IO확장박스": "I/O 드로어 4U — EMX0 · SD2 IOX · 2쪽",
             "양식-대형섀시": "18U 블레이드 섀시 — Superdome 2형 초안 · 3쪽",
             "양식-1U스위치": "48 × SFP + 6 × QSFP (포트 번호 인쇄형)", "양식-반폭스위치": "반폭 1U (트레이에 2대 · 면 = 선반)",
             "양식-관리스위치": "48 × RJ45 + 4 × SFP", "양식-SAN스위치": "48 × SFP, 8포트 그룹 (0번부터)",
             "양식-스토리지": "2U 컨트롤러 (SFF 25 · SP A/B)", "양식-디스크선반": "2U 확장 선반 DAE · JBOD (LCC A/B)",
             "양식-블레이드섀시": "10U 블레이드 (반높이 16베이) · 3쪽",
             "양식-모듈형섀시": "8U 라인카드형 스위치 (슬롯×100+포트) · 3쪽", "양식-패치패널": "1U 패치패널 · 소형 장비",
             "양식-타워": "타워 서버 (앞·뒤 나란히) · 2쪽", "양식-소형장비": "데스크톱 · 미니 PC (폭 8mm = 1칸, 선반 위)",
             "양식-PDU": "0U PDU (콘센트 번호 인쇄형)", "양식-자유형": "생소한 장비: 빈 틀 + 부품으로 구성"}
EX_DESC = {"예시-SRV-01": "HPE DL380 Gen10 (2U 서버 · 디스크 용량 · RAID 표시)", "예시-SRV-03": "Dell R640 (1U 서버)",
           "예시-TOR-A": "Nexus 93180YC-FX (1U 스위치)",
           "예시-MGMT-SW": "C9200L-48T-4X (관리 스위치)", "예시-SAN-A": "Brocade G620 (SAN 스위치)",
           "예시-STG-01": "Unity XT 480 (스토리지 컨트롤러)", "예시-STG-01-DAE1": "Unity DAE (확장 선반 · 시스템 STG-01)",
           "예시-AIX-01": "IBM Power 본체 (시스템 AIX-01)", "예시-AIX-01-IO1": "EMX0 I/O 드로어 (본체와 T1·T2 연결)",
           "예시-ENC-01": "HPE c7000 (블레이드 섀시)", "예시-PDU-A": "APC AP8868 (0U PDU)",
           "예시-PP-01": "패치패널 24P", "예시-WS-01": "HPE ML350 Gen10 (타워)", "예시-CORE-01": "모듈형 섀시 스위치 8U"}


def build_library(path):
    wb = new_workbook("IT 인프라 장비 도면 라이브러리")
    idx_cv = None
    build_guide(wb)
    build_standard(wb)
    build_palette(wb)
    forms(wb)
    for dev in LIB_EXAMPLES:
        example_sheet(wb, dev, f"예시-{dev}")
    tower_example(wb)
    chassis_example(wb)
    entries = [("안내", "사용법", "5분 요약: 복사 · 수정 · 번호 · 랙 관리 · 인쇄"), ("안내", "기준", "격자 · 크기 · 글자 · 선 · 색 · 쪽 구성"),
               ("안내", "부품", "복사해서 붙이는 셀 부품 60여 종")]
    entries += [("빈 양식", s, FORM_DESC.get(s, "")) for s in wb.sheetnames if s.startswith("양식-")]
    entries += [("작성 예시", s, EX_DESC.get(s, "")) for s in wb.sheetnames if s.startswith("예시-")]
    build_index(wb, entries)
    wb.move_sheet("목차", offset=-(len(wb.sheetnames) - 1))
    wb.active = 0
    wb.save(path)
    postprocess(path)


def rack_devices():
    out = []
    for dev in RACK_ORDER:
        d = DEV[dev]
        out.append({"U": d["u"], "높이": d["h"] if d["h"] else None, "면": d["side"], "장비명": dev, "시스템": d.get("sys"),
                    "모델": d.get("short", d["model"].split(" (")[0]), "구분": d["kind"], "시트": rack_sheet_name(dev)})
    return out


def build_rack_example(path):
    wb = new_workbook("랙 실장도 · 장비 도면 — DC1-R02 예시")
    build_rack(wb, "DC1-R02", info={"설치 위치": "DC1 2F 전산실 B열", "규격": "42U · 600×1200", "전원": "PDU-A(A 계통) / PDU-B(B 계통)",
                                     "관리 담당": "인프라팀", "작성": TODAY}, devices=rack_devices())
    for dev in RACK_ORDER:
        cv = example_sheet(wb, dev, rack_sheet_name(dev))
        cv.put(1, 65, '=HYPERLINK("#\'랙\'!A1","◀ 랙")')
        from xlkit import font, P
        cv.c(1, 65).font = font(9, True, P["svc"])
    build_guide(wb)
    wb.active = 0
    wb.save(path)
    postprocess(path)


def build_rack_blank(path):
    wb = new_workbook("랙 실장도 · 장비 도면 — 빈 양식")
    build_rack(wb, "DC?-R??", info={"설치 위치": "", "규격": "42U", "전원": "", "관리 담당": "", "작성": ""}, devices=[],
               title_note="입력 예: U 20 · 높이 2 · 면 양면 · 장비명 SRV-01 · 시스템 (한 대면 비움) · 모델 DL380 Gen10 · 구분 서버 · "
                          "시트 이름 U20 SRV-01")
    build_guide(wb)
    wb.active = 0
    wb.save(path)
    postprocess(path)


if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    targets = sys.argv[1:] or ["lib", "rack", "blank"]
    if "lib" in targets:
        build_library(os.path.join(OUT, "1_장비도면_라이브러리.xlsx"))
        import json, xlpalette
        with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "palette_positions.json"), "w") as f:
            json.dump(xlpalette.POSITIONS, f, ensure_ascii=False, indent=0)
    if "rack" in targets:
        build_rack_example(os.path.join(OUT, "2_랙예시_DC1-R02.xlsx"))
    if "blank" in targets:
        build_rack_blank(os.path.join(OUT, "3_랙_빈양식.xlsx"))
    print("done:", os.listdir(OUT))
