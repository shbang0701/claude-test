# -*- coding: utf-8 -*-
"""장비 종류별 도면 설정 + 양식/예시 시트 생성."""
import xldevices as D
from xlexamples import DEV, device_model, conn_from_n, RACK, loc, AIX_SLOTS, IOD_CARDS
from sheets import device_sheet, TAB_FORM, TAB_EX

DL380_SLOTS = {1: ("FC HBA 2P", 2), 4: ("NIC 2P SFP28", 2)}
# 디스크: 칸 안 용량 글씨 + RAID 묶음(굵은 테두리) 예시
DL380_CAP = {**{i: "480G SSD" for i in (1, 2)}, **{i: "1.2T SAS" for i in range(3, 9)},
             **{i: "1.92T SSD" for i in range(9, 13)}}
DL380_RAID = [(1, 2, "RAID1 · OS"), (3, 8, "RAID5 · DATA"), (9, 12, "RAID10 · DB")]


def _sw(fn):
    return lambda cv, r, c, n: fn(cv, r, c, n, conn=conn_from_n(n))


TYPES = {
    "dl380": dict(u=2, front=lambda cv, r, c, n: D.srv2u_front(cv, r, c, n, boxes=(8, 4, 0), cap=DL380_CAP,
                                                              raid=DL380_RAID),
                  rear=lambda cv, r, c, n: D.srv2u_rear(cv, r, c, n, slots=DL380_SLOTS)),
    "r640": dict(u=1, front=lambda cv, r, c, n: D.srv1u_front(cv, r, c, n, filled=4, cap="600G SAS",
                                                              raid=[(1, 2, "RAID1 · OS"), (3, 4, "RAID1 · DATA")]),
                 rear=lambda cv, r, c, n: D.srv1u_rear(cv, r, c, n, slots={1: ("NIC 2P", 2)})),
    "sw48": dict(u=1, front=_sw(D.sw48_front), rear=D.sw_rear, half=True, nrows=24),
    "rj48": dict(u=1, front=_sw(D.rj48_front), rear=D.sw_rear, half=True, nrows=24),
    "san48": dict(u=1, front=_sw(D.san48_front), rear=D.san_rear, half=True, nrows=24),
    "pp24": dict(u=1, front=_sw(D.pp24_front), rear=lambda cv, r, c, n: D.small_rear(cv, r, c, n, psu_n=0)),
    "unity": dict(u=2, front=lambda cv, r, c, n: D.stg_front(cv, r, c, n, filled=12, cap="1.92T SSD",
                                                             raid=[(0, 11, "Pool 1 · RAID5")]),
                  rear=D.stg_rear),
    "c7000": dict(u=10, front=lambda cv, r, c, n: D.blade_front(cv, r, c, n, bays={i: "BL460c Gen10" for i in range(1, 7)}),
                  rear=lambda cv, r, c, n: D.blade_rear(cv, r, c, n, ic={
                      1: ("VC FlexFabric-20/40 F8", [(f"b1x{i}", f"X{i}") for i in range(1, 9)]),
                      2: ("VC FlexFabric-20/40 F8", [(f"b2x{i}", f"X{i}") for i in range(1, 9)])}),
                  half=True, nrows=36),
    "pdu": dict(u=None, rows=6, front=_sw(D.pdu_face), rear=None, half=True, nrows=24),
    "power4u": dict(u=4, front=lambda cv, r, c, n: D.power4u_front(cv, r, c, n, filled=2, cap="1.6T NVMe",
                                                                   raid=[(0, 1, "rootvg 미러")]),
                    rear=lambda cv, r, c, n: D.power4u_rear(cv, r, c, n, slots=AIX_SLOTS)),
    "iodrawer": dict(u=4, front=D.iox_front, rear=lambda cv, r, c, n: D.iox_rear(cv, r, c, n, cards=IOD_CARDS)),
    "dae": dict(u=2, front=lambda cv, r, c, n: D.dae_front(cv, r, c, n, filled=12, cap="3.84T SSD",
                                                           raid=[(0, 11, "Pool 1 · RAID5")]),
                rear=D.dae_rear),
}

# 예시 시트 아래 한 줄 안내 (번호 규칙 · 구분법)
TYPE_NOTE = {
    "pdu": "0U 세로형 PDU를 가로로 눕혀 그림 · 칸 안 숫자 = PDU에 인쇄된 콘센트 번호 · 색 칸 = 사용 중",
    "sw48": "칸 안 숫자 = 장비에 인쇄된 포트 번호(= 도면 번호) · 색 칸 = 연결됨",
    "c7000": "도면 번호 = 베이×100 + 포트 (예: 베이 2의 X3 = 203) · OA = 901/902 · 전원 입력 = 1–6",
    "power4u": "본체 표시: 전면 조작 패널(LCD) · 후면 HMC 포트 · 광케이블 어댑터(T1·T2) → I/O 드로어  ·  "
               "슬롯 번호·배치는 모델별 확인",
    "iodrawer": "I/O 드로어: 전면 표기가 본체와 같아 보여도 후면에 HMC 포트가 없음 · 모듈마다 T1·T2 한 쌍이 "
                "본체 광케이블 어댑터로 연결",
    "dae": "확장 선반: 컨트롤러 SP A/B의 SAS → LCC A/B의 A 포트 · 다음 선반은 B 포트에서 이어짐",
}
for _t in ("rj48", "san48", "pp24"):
    TYPE_NOTE[_t] = TYPE_NOTE["sw48"]

TYPE_LABEL = {"dl380": "2U 서버", "r640": "1U 서버", "sw48": "1U 스위치", "rj48": "관리 스위치",
              "san48": "SAN 스위치", "pp24": "패치패널", "unity": "스토리지", "c7000": "블레이드 섀시",
              "pdu": "PDU", "power4u": "Power 본체", "iodrawer": "I/O 드로어", "dae": "디스크 선반"}

TODAY = "2026-09-26"


def example_sheet(wb, dev, sheet_name=None, tab=TAB_EX):
    d = DEV[dev]
    t = TYPES[d["type"]]
    n, data = device_model(dev)
    u_txt = loc(dev).replace(f"{RACK} ", "")
    info = {"설치 위치": "DC1 2F 전산실", "랙 · U": f"{RACK} · {u_txt}", "구분": d["kind"], "시리얼": d["sn"],
            "관리 IP": d["ip"], "작성": TODAY}
    nrows = t.get("nrows")
    if nrows is not None:
        nrows = max(nrows, len(data) + (2 if not t.get("half") else 4))
        if t.get("half") and nrows % 2:
            nrows += 1
    if d.get("sys"):
        info["구분"] = f'{d["kind"]} · {d["sys"]}'
    return device_sheet(wb, sheet_name or f"예시-{dev}", dev, d["model"], t.get("u") or 1, t["front"], t.get("rear"),
                        n=n, info=info, rows=t.get("rows"), data=data, nrows=nrows, half=t.get("half", False),
                        tab=tab, foot_note=TYPE_NOTE.get(d["type"]), role=d.get("role"))
