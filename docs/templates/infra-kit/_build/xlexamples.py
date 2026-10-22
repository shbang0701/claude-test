# -*- coding: utf-8 -*-
"""작성 예시 데이터: 랙 DC1-R02 한 대분의 장비·케이블 모델.
케이블 목록 하나에서 양쪽 장비의 표를 만들므로, 두 장비 시트의 케이블 라벨이 항상 일치한다."""
import xldevices as D

RACK = "R02"

# ── 장비 정의 ────────────────────────────────────────────────
# type, U(하단), 높이, 모델, 면(양면/전면/후면/0U), 구분, 시트이름, 기타
DEV = {
    "PP-01":   dict(type="pp24", u=42, h=1, model="패치패널 24P Cat6", short="Cat6 24P", side="양면", kind="네트워크", sn="-", ip="-"),
    "MGMT-SW": dict(type="rj48", u=41, h=1, model="Cisco C9200L-48T-4X", short="C9200L-48T-4X", side="양면", kind="네트워크", sn="JAE2301ABCD", ip="10.10.0.41"),
    "TOR-A":   dict(type="sw48", u=39, h=1, model="Cisco Nexus 93180YC-FX", short="Nexus 93180YC-FX", side="양면", kind="네트워크", sn="FDO2315A1BC", ip="10.10.0.39"),
    "TOR-B":   dict(type="sw48", u=38, h=1, model="Cisco Nexus 93180YC-FX", short="Nexus 93180YC-FX", side="양면", kind="네트워크", sn="FDO2315A2BC", ip="10.10.0.38"),
    "SAN-A":   dict(type="san48", u=36, h=1, model="Brocade G620", short="Brocade G620", side="양면", kind="스토리지", sn="BRCAG1234A01", ip="10.10.0.36"),
    "SAN-B":   dict(type="san48", u=35, h=1, model="Brocade G620", short="Brocade G620", side="양면", kind="스토리지", sn="BRCAG1234A02", ip="10.10.0.35"),
    "STG-01":  dict(type="unity", u=30, h=2, model="Dell EMC Unity XT 480 (DPE)", short="Unity XT 480", side="양면", kind="스토리지", sn="CKM00201900123", ip="10.10.0.30", sys="STG-01", role="컨트롤러"),
    "STG-01-DAE1": dict(type="dae", u=32, h=2, model="Dell EMC Unity 25-드라이브 DAE (2U)", short="Unity DAE 25D", side="양면", kind="스토리지", sn="CKM00202000456", ip="-", sys="STG-01", role="디스크 선반"),
    "AIX-01":  dict(type="power4u", u=22, h=4, model="IBM Power 4U 본체 (예: Power E1050)", short="Power 4U 본체", side="양면", kind="서버", sn="78A1B2C", ip="10.10.1.40", sys="AIX-01", role="본체"),
    "AIX-01-IO1": dict(type="iodrawer", u=26, h=4, model="IBM EMX0 PCIe3 I/O 확장 드로어", short="EMX0 I/O 드로어", side="양면", kind="서버", sn="78C9D3E", ip="-", sys="AIX-01", role="I/O 드로어"),
    "SRV-01":  dict(type="dl380", u=20, h=2, model="HPE ProLiant DL380 Gen10", short="DL380 Gen10", side="양면", kind="서버", sn="SGH0123ABC", ip="10.10.1.21"),
    "SRV-02":  dict(type="dl380", u=18, h=2, model="HPE ProLiant DL380 Gen10", short="DL380 Gen10", side="양면", kind="서버", sn="SGH0123ABD", ip="10.10.1.22"),
    "SRV-03":  dict(type="r640", u=16, h=1, model="Dell PowerEdge R640", short="PowerEdge R640", side="양면", kind="서버", sn="7XK2Q53", ip="10.10.1.23"),
    "ENC-01":  dict(type="c7000", u=1, h=10, model="HPE BladeSystem c7000", short="c7000", side="양면", kind="서버", sn="CZJ12345AB", ip="10.10.1.10"),
    "PDU-A":   dict(type="pdu", u=None, h=0, model="APC AP8868 (24×C13 + 6×C19)", short="AP8868", side="0U", kind="전원", sn="5A1234E01234", ip="10.10.0.201", loc="R02 후면 좌"),
    "PDU-B":   dict(type="pdu", u=None, h=0, model="APC AP8868 (24×C13 + 6×C19)", short="AP8868", side="0U", kind="전원", sn="5A1234E05678", ip="10.10.0.202", loc="R02 후면 우"),
}
# 랙 밖 장비(상대 위치 표시용)
EXT = {"SPINE-01": "R10 U40", "SPINE-02": "R10 U39", "OOB-CORE": "R10 U42", "BAK-SW": "R05 U30"}

# ── 장비 종류별 포트 목록: (키, 포트명, 면) — 도면 읽는 순서(전면→후면, 위→아래, 왼→오) ──
def _dl380():
    L = []
    for row in ((1, 4, 7), (2, 5, 8), (3, 6)):
        for s in row:
            for p in (1, 2):
                L.append((f"s{s}p{p}", f"S{s}-P{p}", "후면"))
    L += [("ocp1", "FLR1", "후면"), ("ocp2", "FLR2", "후면"), ("bmc", "iLO", "후면")]
    L += [(f"lom{i}", f"LOM{i}", "후면") for i in range(1, 5)]
    L += [("psu2", "PSU2", "후면"), ("psu1", "PSU1", "후면")]
    return L


def _r640():
    L = [(f"s{s}p{p}", f"S{s}-P{p}", "후면") for s in (1, 2, 3) for p in (1, 2)]
    L += [("ocp1", "NDC1", "후면"), ("ocp2", "NDC2", "후면"), ("bmc", "iDRAC", "후면")]
    L += [(f"lom{i}", f"LOM{i}", "후면") for i in range(1, 5)]
    L += [("psu1", "PSU1", "후면"), ("psu2", "PSU2", "후면")]
    return L


def _unity():
    L = []
    for k in ("b", "a"):
        up = k.upper()
        L += [(f"{k}_mgmt", f"SP{up}-MGMT", "후면"), (f"{k}_svc", f"SP{up}-SVC", "후면")]
        L += [(f"{k}_eth{i}", f"SP{up}-ETH{i}", "후면") for i in range(4)]
        L += [(f"{k}_sas{i}", f"SP{up}-SAS{i}", "후면") for i in range(2)]
        L += [(f"{k}_io0p{i}", f"SP{up}-IO0-P{i}", "후면") for i in range(4)]
        L += [(f"{k}_psu", f"PSU {up}", "후면")]
    return L


def _c7000():
    L = []
    for b in range(1, 9):
        L += [(f"b{b}x{i}", f"B{b}-X{i}", "후면") for i in range(1, 9)]
    L += [("oa1", "OA1", "후면"), ("oa2", "OA2", "후면")]
    L += [(f"ac{i}", f"AC{i}", "후면") for i in range(6, 0, -1)]
    return L


# IBM Power 본체 슬롯 구성(예시)과 I/O 드로어 카드 — 도면과 표가 같은 정의를 쓴다
AIX_SLOTS = {1: ("NIC 4P 25G", 4), 2: ("FC 4P 32G", 4), 3: ("광케이블 어댑터 → IO1 P1", 2, True, "T"),
             4: ("광케이블 어댑터 → IO1 P2", 2, True, "T")}
IOD_CARDS = {(0, 1): ("FC 2P", 2)}


def _power4u():
    L = []
    for s in range(1, 9):
        spec = AIX_SLOTS.get(s)
        if spec:
            pre = spec[3] if len(spec) > 3 else "P"
            L += [(f"c{s}p{j}", f"C{s}-{pre}{j}", "후면") for j in range(1, spec[1] + 1)]
    L += [("hmc1", "HMC1", "후면"), ("hmc2", "HMC2", "후면")]
    L += [(f"psu{i}", f"PSU{i}", "후면") for i in range(1, 5)]
    return L


def _iodrawer():
    L = []
    for m in (1, 2):
        L += [(f"m{m}t1", f"P{m}-T1", "후면"), (f"m{m}t2", f"P{m}-T2", "후면")]
        for j in range(1, 7):
            spec = IOD_CARDS.get((m - 1, j))
            if spec:
                L += [(f"m{m}c{j}p{p}", f"P{m}-C{j}-T{p}", "후면") for p in range(1, spec[1] + 1)]
    L += [("psu1", "PSU1", "후면"), ("psu2", "PSU2", "후면")]
    return L


def _dae():
    L = []
    for k in ("b", "a"):
        up = k.upper()
        L += [(f"lcc{k}_a", f"LCC{up}-A", "후면"), (f"lcc{k}_b", f"LCC{up}-B", "후면"), (f"psu{k}", f"PSU {up}", "후면")]
    return L


PORTS = {"dl380": _dl380, "r640": _r640, "unity": _unity, "c7000": _c7000, "power4u": _power4u,
         "iodrawer": _iodrawer, "dae": _dae}

# 번호 규칙: seq = 연결된 포트만 1부터 차례로 / face = 장비에 인쇄된 번호 / slot = 슬롯×100+포트
NUMBERING = {"dl380": "seq", "r640": "seq", "unity": "seq", "c7000": "slot",
             "sw48": "face", "rj48": "face", "san48": "face", "pp24": "face", "pdu": "face",
             "power4u": "seq", "iodrawer": "seq", "dae": "seq"}


def face_name(dtype, key):
    """스위치·패치패널·PDU: 키 → (번호, 포트명, 면)."""
    if key in ("mgmt", "con"):
        nm = {"sw48": ("mgmt0", "console"), "rj48": ("Gi0/0", "console"), "san48": ("MGMT", "CON"),
              }.get(dtype, ("MGMT", "CON"))
        base = {"sw48": 55, "rj48": 53, "san48": 64}.get(dtype, 90)
        i = 0 if key == "mgmt" else 1
        face = "후면" if dtype == "san48" else "전면"
        return base + i, nm[i], face
    if key.startswith("psu"):
        i = int(key[3:])
        base = {"sw48": 57, "rj48": 55, "san48": 66, "pp24": 91}.get(dtype, 90)
        return base + i - 1, f"PSU{i}", "후면"
    if key == "input":
        return 31, "입력", "전면"
    no = int(key[1:])
    if dtype == "sw48":
        return no, f"Eth1/{no}", "전면"
    if dtype == "rj48":
        return no, (f"Gi1/0/{no}" if no <= 48 else f"Te1/1/{no-48}"), "전면"
    if dtype == "san48":
        return no, f"P{no}", "전면"
    if dtype == "pp24":
        return no, f"{no}", "전면"
    if dtype == "pdu":
        return no, (f"C13-{no}" if no <= 24 else f"C19-{no}"), "전면"
    return no, key, "전면"


# ── 케이블 목록: (장비A, 키A, 장비B, 키B, 용도, 케이블·커넥터, 비고) ─────────────
C = []
def cab(a, ka, b, kb, use, media, note=""):
    C.append((a, ka, b, kb, use, media, note))


# 관리망 (MGMT-SW)
mg = 1
for dev, key in (("SRV-01", "bmc"), ("SRV-02", "bmc"), ("SRV-03", "bmc"), ("ENC-01", "oa1"), ("ENC-01", "oa2"),
                 ("STG-01", "a_mgmt"), ("STG-01", "b_mgmt"), ("TOR-A", "mgmt"), ("TOR-B", "mgmt"),
                 ("SAN-A", "mgmt"), ("SAN-B", "mgmt")):
    cab(dev, key, "MGMT-SW", f"p{mg}", "관리", "UTP Cat6 / RJ45")
    mg += 1
cab("MGMT-SW", "p49", "OOB-CORE", "Te1/0/12", "관리", "OM4 LC-LC", "관리망 업링크")
# 서비스망 (TOR-A/B)
cab("SRV-01", "ocp1", "TOR-A", "p11", "서비스", "DAC 25G 3m")
cab("SRV-01", "ocp2", "TOR-B", "p11", "서비스", "DAC 25G 3m")
cab("SRV-01", "s4p1", "TOR-A", "p12", "서비스", "OM4 LC-LC", "vMotion")
cab("SRV-01", "s4p2", "TOR-B", "p12", "서비스", "OM4 LC-LC", "vMotion")
cab("SRV-02", "ocp1", "TOR-A", "p13", "서비스", "DAC 25G 3m")
cab("SRV-02", "ocp2", "TOR-B", "p13", "서비스", "DAC 25G 3m")
cab("SRV-02", "s4p1", "TOR-A", "p14", "서비스", "OM4 LC-LC", "vMotion")
cab("SRV-02", "s4p2", "TOR-B", "p14", "서비스", "OM4 LC-LC", "vMotion")
cab("SRV-03", "ocp1", "TOR-A", "p15", "서비스", "DAC 25G 3m")
cab("SRV-03", "ocp2", "TOR-B", "p15", "서비스", "DAC 25G 3m")
cab("ENC-01", "b1x3", "TOR-A", "p1", "서비스", "OM4 LC-LC")
cab("ENC-01", "b1x4", "TOR-A", "p2", "서비스", "OM4 LC-LC")
cab("ENC-01", "b2x3", "TOR-B", "p1", "서비스", "OM4 LC-LC")
cab("ENC-01", "b2x4", "TOR-B", "p2", "서비스", "OM4 LC-LC")
cab("STG-01", "a_eth0", "TOR-A", "p21", "스토리지", "DAC 25G 3m", "iSCSI")
cab("STG-01", "b_eth0", "TOR-B", "p21", "스토리지", "DAC 25G 3m", "iSCSI")
# 백업망 (다른 랙, 패치패널 경유)
cab("SRV-01", "lom1", "PP-01", "p1", "백업", "UTP Cat6 / RJ45", "→ R05 BAK-SW Gi1/0/3")
cab("SRV-02", "lom1", "PP-01", "p2", "백업", "UTP Cat6 / RJ45", "→ R05 BAK-SW Gi1/0/4")
cab("SRV-03", "lom1", "PP-01", "p3", "백업", "UTP Cat6 / RJ45", "→ R05 BAK-SW Gi1/0/5")
# TOR 업링크 / 피어링크
cab("TOR-A", "p49", "SPINE-01", "Eth1/1", "서비스", "QSFP28 AOC 10m", "업링크")
cab("TOR-A", "p50", "SPINE-02", "Eth1/1", "서비스", "QSFP28 AOC 10m", "업링크")
cab("TOR-B", "p49", "SPINE-01", "Eth1/2", "서비스", "QSFP28 AOC 10m", "업링크")
cab("TOR-B", "p50", "SPINE-02", "Eth1/2", "서비스", "QSFP28 AOC 10m", "업링크")
cab("TOR-A", "p53", "TOR-B", "p53", "인터커넥트", "QSFP28 DAC 1m", "vPC 피어링크")
cab("TOR-A", "p54", "TOR-B", "p54", "인터커넥트", "QSFP28 DAC 1m", "vPC 피어링크")
# SAN (FC)
fa = {"SAN-A": 0, "SAN-B": 0}
for dev, key, fab in (("SRV-01", "s1p1", "SAN-A"), ("SRV-01", "s1p2", "SAN-B"),
                      ("SRV-02", "s1p1", "SAN-A"), ("SRV-02", "s1p2", "SAN-B"),
                      ("ENC-01", "b1x1", "SAN-A"), ("ENC-01", "b1x2", "SAN-A"),
                      ("ENC-01", "b2x1", "SAN-B"), ("ENC-01", "b2x2", "SAN-B"),
                      ("STG-01", "a_io0p0", "SAN-A"), ("STG-01", "b_io0p0", "SAN-A"),
                      ("STG-01", "a_io0p1", "SAN-B"), ("STG-01", "b_io0p1", "SAN-B")):
    cab(dev, key, fab, f"p{fa[fab]}", "스토리지", "OM4 LC-LC", "Fabric " + fab[-1])
    fa[fab] += 1
# c7000 모듈 간 스태킹 (같은 섀시 안)
cab("ENC-01", "b1x7", "ENC-01", "b2x7", "인터커넥트", "DAC 10G 0.5m", "VC 스태킹")
cab("ENC-01", "b1x8", "ENC-01", "b2x8", "인터커넥트", "DAC 10G 0.5m", "VC 스태킹")
# 전원 (PDU-A = A 계통, PDU-B = B 계통)
pa = {"PDU-A": 1, "PDU-B": 1}
def pwr(dev, key, pdu, c19=False, note=""):
    if c19:
        no = pa[pdu + "19"] if (pdu + "19") in pa else 25
        pa[pdu + "19"] = no + 1
        cab(dev, key, pdu, f"p{no}", "전원", "C19-C20 2m", note)
    else:
        cab(dev, key, pdu, f"p{pa[pdu]}", "전원", "C13-C14 2m", note)
        pa[pdu] += 1
for dev, ka, kb in (("PP-01", None, None), ("MGMT-SW", "psu1", "psu2"), ("TOR-A", "psu1", "psu2"),
                    ("TOR-B", "psu1", "psu2"), ("SAN-A", "psu1", "psu2"), ("SAN-B", "psu1", "psu2"),
                    ("STG-01", "a_psu", "b_psu"), ("SRV-01", "psu1", "psu2"), ("SRV-02", "psu1", "psu2"),
                    ("SRV-03", "psu1", "psu2")):
    if ka:
        pwr(dev, ka, "PDU-A", note="A 계통")
        pwr(dev, kb, "PDU-B", note="B 계통")
for i in (1, 2, 3):
    pwr("ENC-01", f"ac{i}", "PDU-A", c19=True, note="A 계통")
for i in (4, 5, 6):
    pwr("ENC-01", f"ac{i}", "PDU-B", c19=True, note="B 계통")
cab("PDU-A", "input", "분전반 A", "A-12", "전원", "IEC309 32A", "A 계통 입력")
cab("PDU-B", "input", "분전반 B", "B-12", "전원", "IEC309 32A", "B 계통 입력")

# ── 여러 박스로 된 시스템 ── (기존 라벨이 바뀌지 않도록 목록 끝에 추가)
# IBM Power 본체 + I/O 드로어 (시스템 AIX-01): 드로어는 본체의 광케이블 어댑터와 T1·T2 한 쌍씩 연결
cab("AIX-01", "hmc1", "MGMT-SW", "p12", "관리", "UTP Cat6 / RJ45", "HMC 망")
cab("AIX-01", "hmc2", "MGMT-SW", "p13", "관리", "UTP Cat6 / RJ45", "HMC 망")
cab("AIX-01", "c1p1", "TOR-A", "p16", "서비스", "DAC 25G 3m")
cab("AIX-01", "c1p2", "TOR-B", "p16", "서비스", "DAC 25G 3m")
cab("AIX-01", "c2p1", "SAN-A", "p6", "스토리지", "OM4 LC-LC", "Fabric A")
cab("AIX-01", "c2p2", "SAN-B", "p6", "스토리지", "OM4 LC-LC", "Fabric B")
for _s, _m in ((3, 1), (4, 2)):
    for _t in (1, 2):
        cab("AIX-01", f"c{_s}p{_t}", "AIX-01-IO1", f"m{_m}t{_t}", "인터커넥트", "광케이블 CXP 3m",
            f"드로어 P{_m} · T{_t}–T{_t} (한 쌍)")
cab("AIX-01-IO1", "m1c1p1", "SAN-A", "p7", "스토리지", "OM4 LC-LC", "Fabric A")
cab("AIX-01-IO1", "m1c1p2", "SAN-B", "p7", "스토리지", "OM4 LC-LC", "Fabric B")
# 스토리지 컨트롤러 + 확장 선반 (시스템 STG-01)
cab("STG-01", "a_sas0", "STG-01-DAE1", "lcca_a", "스토리지", "Mini-SAS HD 1m", "SP A → LCC A")
cab("STG-01", "b_sas0", "STG-01-DAE1", "lccb_a", "스토리지", "Mini-SAS HD 1m", "SP B → LCC B")
for _i, _pdu, _no in ((1, "PDU-A", 28), (2, "PDU-B", 28), (3, "PDU-A", 29), (4, "PDU-B", 29)):
    cab("AIX-01", f"psu{_i}", _pdu, f"p{_no}", "전원", "C19-C20 2m", _pdu[-1] + " 계통")
cab("AIX-01-IO1", "psu1", "PDU-A", "p10", "전원", "C13-C14 2m", "A 계통")
cab("AIX-01-IO1", "psu2", "PDU-B", "p10", "전원", "C13-C14 2m", "B 계통")
cab("STG-01-DAE1", "psua", "PDU-A", "p11", "전원", "C13-C14 2m", "A 계통")
cab("STG-01-DAE1", "psub", "PDU-B", "p11", "전원", "C13-C14 2m", "B 계통")

# 라벨 부여: 데이터 = R02-D###, 전원 = R02-P###
_d = _p = 0
LABELS = []
for c in C:
    if c[4] == "전원":
        _p += 1
        LABELS.append(f"{RACK}-P{_p:03d}")
    else:
        _d += 1
        LABELS.append(f"{RACK}-D{_d:03d}")

USEKEY = {"서비스": "svc", "관리": "mgmt", "백업": "bak", "스토리지": "san", "인터커넥트": "ic", "콘솔": "con",
          "전원": "pwr"}


def loc(dev):
    if dev in DEV:
        d = DEV[dev]
        if d.get("loc"):
            return d["loc"]
        return f"{RACK} U{d['u']}" if d["h"] <= 1 else f"{RACK} U{d['u']}–{d['u'] + d['h'] - 1}"
    return EXT.get(dev, "")


def port_info(dev, key):
    """(번호후보, 포트명, 면)"""
    if dev not in DEV:
        return None, key, ""
    t = DEV[dev]["type"]
    if NUMBERING[t] == "face":
        return face_name(t, key)
    for i, (k, nm, face) in enumerate(PORTS[t]()):
        if k == key:
            if NUMBERING[t] == "slot":
                if k.startswith("b") and "x" in k:
                    b, x = k[1:].split("x")
                    return int(b) * 100 + int(x), nm, face
                if k.startswith("oa"):
                    return 900 + int(k[2:]), nm, face
                if k.startswith("ac"):
                    return int(k[2:]), nm, face
            return i, nm, face
    return None, key, ""


def peer_port_name(dev, key):
    if dev in DEV:
        return port_info(dev, key)[1]
    return key


def device_model(dev):
    """→ (n 딕셔너리, 표 데이터 목록)"""
    t = DEV[dev]["type"]
    ends = []
    for idx, (a, ka, b, kb, use, media, note) in enumerate(C):
        if a == dev:
            ends.append((ka, b, kb, use, media, note, LABELS[idx]))
        if b == dev:
            ends.append((kb, a, ka, use, media, note, LABELS[idx]))
    rows = []
    for key, peer, pkey, use, media, note, label in ends:
        order, pname, face = port_info(dev, key)
        rows.append(dict(key=key, order=order, pname=pname, face=face, peer=peer, pport=peer_port_name(peer, pkey),
                         ploc=loc(peer), use=use, media=media, note=note, label=label))
    rows.sort(key=lambda x: (x["order"] if x["order"] is not None else 9999))
    mode = NUMBERING[t]
    n = {}
    data = []
    for i, rw in enumerate(rows, 1):
        num = i if mode == "seq" else rw["order"]
        n[rw["key"]] = (USEKEY[rw["use"]], num)
        data.append({"번호": num, "면": rw["face"], "포트명": rw["pname"], "용도": rw["use"],
                     "케이블 · 커넥터": rw["media"], "상대 장비": rw["peer"], "상대 위치": rw["ploc"],
                     "상대 포트": rw["pport"], "케이블 라벨": rw["label"], "비고": rw["note"]})
    return n, data


def conn_from_n(n):
    """스위치류: {'p12': (kind, 12)} → {12: kind}"""
    out = {}
    for k, v in n.items():
        if k.startswith("p") and k[1:].isdigit():
            out[int(k[1:])] = v[0]
    return out
