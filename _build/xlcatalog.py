# -*- coding: utf-8 -*-
"""장비 종류별 도면 설정 + 양식/예시 시트 생성."""
import xldevices as D
from xlexamples import DEV, device_model, conn_from_n, RACK, loc
from sheets import device_sheet, TAB_FORM, TAB_EX

DL380_SLOTS = {1: ("FC HBA 2P", 2), 4: ("NIC 2P SFP28", 2)}


def _sw(fn):
    return lambda cv, r, c, n: fn(cv, r, c, n, conn=conn_from_n(n))


TYPES = {
    "dl380": dict(u=2, front=lambda cv, r, c, n: D.srv2u_front(cv, r, c, n, boxes=(8, 4, 0)),
                  rear=lambda cv, r, c, n: D.srv2u_rear(cv, r, c, n, slots=DL380_SLOTS)),
    "r640": dict(u=1, front=lambda cv, r, c, n: D.srv1u_front(cv, r, c, n, filled=4),
                 rear=lambda cv, r, c, n: D.srv1u_rear(cv, r, c, n, slots={1: ("NIC 2P", 2)})),
    "sw48": dict(u=1, front=_sw(D.sw48_front), rear=D.sw_rear, half=True, nrows=24),
    "rj48": dict(u=1, front=_sw(D.rj48_front), rear=D.sw_rear, half=True, nrows=24),
    "san48": dict(u=1, front=_sw(D.san48_front), rear=D.san_rear, half=True, nrows=24),
    "pp24": dict(u=1, front=_sw(D.pp24_front), rear=lambda cv, r, c, n: D.small_rear(cv, r, c, n, psu_n=0)),
    "unity": dict(u=2, front=D.stg_front, rear=D.stg_rear),
    "c7000": dict(u=10, front=lambda cv, r, c, n: D.blade_front(cv, r, c, n, bays={i: "BL460c Gen10" for i in range(1, 7)}),
                  rear=lambda cv, r, c, n: D.blade_rear(cv, r, c, n, ic={
                      1: ("VC FlexFabric-20/40 F8", [(f"b1x{i}", f"X{i}") for i in range(1, 9)]),
                      2: ("VC FlexFabric-20/40 F8", [(f"b2x{i}", f"X{i}") for i in range(1, 9)])}),
                  half=True, nrows=36),
    "pdu": dict(u=None, rows=6, front=_sw(D.pdu_face), rear=None, half=True, nrows=24),
}

TYPE_LABEL = {"dl380": "2U 서버", "r640": "1U 서버", "sw48": "1U 스위치", "rj48": "관리 스위치",
              "san48": "SAN 스위치", "pp24": "패치패널", "unity": "스토리지", "c7000": "블레이드 섀시",
              "pdu": "PDU"}

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
    fnote = None
    if d["type"] == "pdu":
        fnote = "0U 세로형 PDU를 가로로 눕혀 그림 · 칸 안 숫자 = PDU에 인쇄된 콘센트 번호 · 색 칸 = 사용 중"
    if d["type"] in ("sw48", "rj48", "san48", "pp24"):
        fnote = "칸 안 숫자 = 장비에 인쇄된 포트 번호(= 도면 번호) · 색 칸 = 연결됨"
    if d["type"] == "c7000":
        fnote = "도면 번호 = 베이×100 + 포트 (예: 베이 2의 X3 = 203) · OA = 901/902 · 전원 입력 = 1–6"
    return device_sheet(wb, sheet_name or f"예시-{dev}", dev, d["model"], t.get("u") or 1, t["front"], t.get("rear"),
                        n=n, info=info, rows=t.get("rows"), data=data, nrows=nrows, half=t.get("half", False),
                        tab=tab, foot_note=fnote)
