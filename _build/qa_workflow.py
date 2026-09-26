# -*- coding: utf-8 -*-
"""작업 흐름 검수 (LibreOffice UNO로 실제 조작):
샘플 선택 → 낯선 장비에 맞게 부품 변경 → 번호·표 작성 → 다른 파일에 복사 → 랙에 장비 추가 → 인쇄(PDF)
실행: /usr/bin/python3 qa_workflow.py <출력폴더>   (python3-uno 필요)"""
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time

import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.table import CellAddress, CellRangeAddress
from com.sun.star.sheet.CellInsertMode import ROWS as INSERT_ROWS

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import lo  # noqa: E402
from xlrack import LC  # noqa: E402  랙 목록 열 위치 (U · 높이 · 면 · 장비명 · 시스템 · 모델 · 구분 · 시트 · 이동)

XL = os.path.join(os.path.dirname(HERE), "01_엑셀_장비도면")
LOG = []


def log(ok, msg):
    LOG.append(("PASS" if ok else "FAIL", msg))
    print(("PASS " if ok else "FAIL ") + msg, flush=True)


def pv(n, v):
    p = PropertyValue()
    p.Name, p.Value = n, v
    return p


def start():
    prof = tempfile.mkdtemp(prefix="lo_qa_")
    port = 2099
    proc = subprocess.Popen(["soffice", f"-env:UserInstallation=file://{prof}", "--headless", "--invisible",
                             "--norestore", "--nologo",
                             f"--accept=socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext"],
                            env=lo.env(), stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    local = uno.getComponentContext()
    res = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    for _ in range(90):
        try:
            ctx = res.resolve(f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext")
            break
        except Exception:
            time.sleep(1)
    else:
        raise SystemExit("soffice 연결 실패")
    desk = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    return proc, desk, ctx


def ui_paste(doc, ctx, sheet, src, dst):
    """화면에서 하는 [복사] → [붙여넣기]와 같게 (병합 칸 위에 붙여넣기는 UNO copyRange로는 안 됨).
    LibreOffice의 '덮어쓰기 확인' 창은 [예]로 간주하도록 끈다 (Excel은 확인 없이 덮어씀)."""
    cp = ctx.ServiceManager.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider", ctx)
    acc = cp.createInstanceWithArguments("com.sun.star.configuration.ConfigurationUpdateAccess",
                                         (pv("nodepath", "/org.openoffice.Office.Calc/Input"),))
    acc.setPropertyValue("ReplaceCellsWarning", False)
    acc.commitChanges()
    ctl = doc.getCurrentController()
    ctl.setActiveSheet(sheet)
    disp = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.DispatchHelper", ctx)
    ctl.select(src)
    disp.executeDispatch(ctl.getFrame(), ".uno:Copy", "", 0, ())
    ctl.select(dst)
    disp.executeDispatch(ctl.getFrame(), ".uno:Paste", "", 0, ())


def load(desk, path):
    return desk.loadComponentFromURL(uno.systemPathToFileUrl(path), "_blank", 0, (pv("Hidden", True),))


def cell(sh, r, c):
    return sh.getCellByPosition(c - 1, r - 1)  # 1-based 입력


def rng_addr(sheet_idx, r, c, h, w):
    a = CellRangeAddress()
    a.Sheet, a.StartRow, a.StartColumn, a.EndRow, a.EndColumn = sheet_idx, r - 1, c - 1, r - 1 + h - 1, c - 1 + w - 1
    return a


def addr(sheet_idx, r, c):
    a = CellAddress()
    a.Sheet, a.Row, a.Column = sheet_idx, r - 1, c - 1
    return a


def idx_of(doc, name):
    for i in range(doc.Sheets.Count):
        if doc.Sheets.getByIndex(i).Name == name:
            return i
    raise KeyError(name)


def main(out):
    os.makedirs(out, exist_ok=True)
    lib = os.path.join(out, "lib.xlsx")
    rack = os.path.join(out, "rack_new.xlsx")
    shutil.copy(os.path.join(XL, "1_장비도면_라이브러리.xlsx"), lib)
    shutil.copy(os.path.join(XL, "3_랙_빈양식.xlsx"), rack)
    pos = json.load(open(os.path.join(HERE, "palette_positions.json")))
    proc, desk, ctx = start()
    try:
        L = load(desk, lib)
        R = load(desk, rack)
        log(L is not None and R is not None, "라이브러리 · 빈 랙 파일 열기")
        # 1) 샘플 선택 + 다른 파일로 복사 (이동/복사와 같은 동작)
        n0 = R.Sheets.Count
        R.Sheets.importSheet(L, "양식-2U서버", n0)
        dev = R.Sheets.getByIndex(n0)
        dev.Name = "U12 GPU-01"
        log(dev.Name == "U12 GPU-01", "양식-2U서버 → 랙 파일로 시트 복사 · 이름 변경")
        ps = dev.PageStyle
        st = R.StyleFamilies.getByName("PageStyles").getByName(ps)
        log(bool(st.IsLandscape) and st.Width > st.Height, f"복사한 시트 인쇄 설정 유지 (가로, 스타일 {ps})")
        cs = R.StyleFamilies.getByName("CellStyles")
        need = ["포트 서비스", "포트 인터커넥트", "포트 전원", "포트 관리", "포트 미연결", "모듈 채움", "케이블 라벨"]
        log(all(cs.hasByName(n) for n in need), "셀 스타일이 랙 파일에도 있음 (포트·모듈·라벨)")
        # 2) 부품 가져오기: 부품 시트를 잠시 가져와 같은 크기 부품으로 교체
        R.Sheets.importSheet(L, "부품", R.Sheets.Count)
        pi = R.Sheets.Count - 1
        di = idx_of(R, "U12 GPU-01")
        # 후면 틀: 15행부터 (제목 3행 + 전면 10행 + 간격 1행)
        rear_top = 4 + 10 + 1
        r0, c0 = rear_top + 1, 3 + 1          # S1 자리 (틀 안 첫 칸)
        b03, b02 = pos["B03"], pos["B02"]
        dev.copyRange(addr(di, r0, c0), rng_addr(pi, b03[0], b03[1], b03[2], b03[3]))
        dev.copyRange(addr(di, r0, c0 + 20), rng_addr(pi, b02[0], b02[1], b02[2], b02[3]))
        cell(dev, r0, c0).setString("S1 · IB HCA 2P")
        cell(dev, r0, c0 + 20).setString("S4 · NIC 4P 25G")
        log(cell(dev, r0, c0).getString().startswith("S1 · IB"), "부품 붙여넣기: S1 = 넓은 포트 카드(B03), S4 = 4포트 카드(B02)")
        merged = dev.getCellRangeByPosition(c0 + 11 - 1, r0 - 1, c0 + 13 - 1, r0 - 1)
        log(merged.getIsMerged(), "붙여넣은 포트의 병합(3칸) 유지")
        # 3) 번호 + 셀 스타일로 색 지정
        marks = [((r0, c0 + 11), 1, "포트 인터커넥트"), ((r0, c0 + 15), 2, "포트 인터커넥트"),
                 ((r0, c0 + 20 + 7), 3, "포트 서비스"), ((r0, c0 + 20 + 10), 4, "포트 서비스"),
                 ((rear_top + 7, 3 + 11 + 6 + 7), 5, "포트 관리")]
        # I/O 모듈 관리 포트 위치: 모듈 시작 14열, 포트폭 합 9 → 20열 USB, 23열 VGA, 27열 관리
        marks[4] = ((rear_top + 7, 27), 5, "포트 관리")
        marks += [((rear_top + 7, 50), 6, "포트 전원"), ((rear_top + 7, 60), 7, "포트 전원")]
        for (r, c), num, sty in marks:
            x = cell(dev, r, c)
            x.setValue(num)
            x.CellStyle = sty
        bd = cell(dev, r0, c0 + 11).TopBorder
        log(bd.OuterLineWidth > 0 or bd.LineWidth > 0, "셀 스타일 적용 후에도 포트 테두리 유지 (스타일에 테두리 미포함)")
        # 3-1) 디스크 용량: 베이 1에 셀 스타일 + 용량 입력 → 그 디스크 칸을 복사해 베이 2–4 디스크 칸에 한 번에 붙여넣기
        #      (전면 틀 4행부터: BOX 1 모듈 5행, 디스크 칸 6–10행, 베이 번호 11행 · 베이 1 = 5–6열, 베이 2–4 = 7–12열)
        dev.getCellRangeByPosition(5 - 1, 6 - 1, 6 - 1, 10 - 1).CellStyle = "디스크 장착"
        cell(dev, 6, 5).setString("1.92T SSD")
        ui_paste(R, ctx, dev, dev.getCellRangeByPosition(5 - 1, 6 - 1, 6 - 1, 10 - 1),
                 dev.getCellRangeByPosition(7 - 1, 6 - 1, 12 - 1, 10 - 1))
        ok = all(cell(dev, 6, c).getString() == "1.92T SSD" and cell(dev, 11, c).getString() == str((c - 3) // 2)
                 and dev.getCellRangeByPosition(c - 1, 6 - 1, c, 10 - 1).getIsMerged() for c in (7, 9, 11))
        log(ok, "디스크 용량: 베이 1에 입력 → 베이 2–4에 한 번에 붙여넣기 (병합 · 베이 번호 유지)")
        # 4) 표 작성 (번호 2열, 면 4, 포트명 6, 용도 12, 케이블 17, 상대 장비 24, 위치 32, 포트 38, 라벨 44, 비고 51)
        t0 = rear_top + 10 + 1 + 2   # 간격 1 + 제목 행 1 + 머리글 1
        rows = [(1, "후면", "S1-P1", "인터커넥트", "QSFP56 DAC 2m", "IB-SW-01", "R05 U40", "P1", "R05-D101"),
                (2, "후면", "S1-P2", "인터커넥트", "QSFP56 DAC 2m", "IB-SW-02", "R05 U39", "P1", "R05-D102"),
                (3, "후면", "S4-P1", "서비스", "DAC 25G 3m", "TOR-A", "R05 U42", "Eth1/5", "R05-D103"),
                (4, "후면", "S4-P2", "서비스", "DAC 25G 3m", "TOR-B", "R05 U41", "Eth1/5", "R05-D104"),
                (5, "후면", "iLO", "관리", "UTP Cat6", "MGMT-SW", "R05 U43", "Gi1/0/9", "R05-D105"),
                (6, "후면", "PSU2", "전원", "C13-C14 2m", "PDU-B", "R05 후면 우", "C13-9", "R05-P106"),
                (7, "후면", "PSU1", "전원", "C13-C14 2m", "PDU-A", "R05 후면 좌", "C13-9", "R05-P107"),
                (99, "후면", "(오타 확인용)", "기타", "", "", "", "", "")]
        cols = [2, 4, 6, 12, 17, 24, 32, 38, 44]
        for i, row in enumerate(rows):
            for c, v in zip(cols, row):
                x = cell(dev, t0 + i, c)
                x.setValue(v) if isinstance(v, int) else x.setString(v)
        log(cell(dev, t0 + 2, 24).getString() == "TOR-A", "연결 표 입력 (병합 칸에 한 칸씩)")
        # 표 줄 추가: 빈 줄을 행째 복사해 삽입(병합 유지) — 안내서의 방법
        last = t0 + 7
        dev.insertCells(rng_addr(di, last + 1, 1, 1, 70), INSERT_ROWS)
        dev.copyRange(addr(di, last + 1, 1), rng_addr(di, last, 1, 1, 70))
        log(dev.getCellRangeByPosition(6 - 1, last, 11 - 1, last).getIsMerged(), "표 줄 추가(행 복사 삽입) 후 병합 유지")
        R.Sheets.removeByName("부품")
        # 5) 랙에 장비 추가: 목록 한 줄
        rk = R.Sheets.getByName("랙")
        col = {k: v[0] for k, v in LC.items()}
        vals = {col["U"]: 12, col["높이"]: 2, col["면"]: "양면", col["장비명"]: "GPU-01", col["모델"]: "DL380 Gen11 + IB",
                col["구분"]: "서버", col["시트"]: "U12 GPU-01"}
        for c, v in vals.items():
            x = cell(rk, 6, c)
            x.setValue(v) if isinstance(v, int) else x.setString(v)
        R.calculateAll()
        top = 6 + (42 - 13)
        got = cell(rk, top, 3).getString()
        log(got.startswith("GPU-01"), f"랙 실장도 자동 표시 (U13 행: '{got}')")
        log(cell(rk, top + 1, 3).getString() == "" and cell(rk, top + 1, 67).getValue() == 1,
            "2U 장비의 아래 U(U12)도 차지 표시(이름은 맨 위 U에만)")
        link = cell(rk, 6, col["이동"])
        log(link.getString() == "▶" and "HYPERLINK" in link.getFormula().upper(), "시트 이동 링크(▶) 자동 생성")
        summ = cell(rk, 49, 6).getString()
        log(summ.startswith("2 U"), f"전면 사용 U 합계 갱신 ('{summ}')")
        # 겹침 확인: 같은 U에 장비 하나 더
        for c, v in {col["U"]: 13, col["높이"]: 1, col["면"]: "양면", col["장비명"]: "TEST", col["구분"]: "기타"}.items():
            x = cell(rk, 7, c)
            x.setValue(v) if isinstance(v, int) else x.setString(v)
        R.calculateAll()
        log(cell(rk, top, 3).getString().startswith("⚠"), "U 겹침 입력 시 '⚠ U 겹침' 경고")
        for c in (col["U"], col["높이"], col["면"], col["장비명"], col["구분"]):
            cell(rk, 7, c).setString("")
        R.calculateAll()
        # 5-1) 선반: 선반 한 줄(구분 '선반') + 그 위 장비(면 '선반') → 겹침 경고 없음 · 장비 수에서 선반 제외
        rows2 = ((7, {col["U"]: 5, col["높이"]: 2, col["면"]: "양면", col["장비명"]: "선반-U05", col["모델"]: "NAS · 미니 PC",
                      col["구분"]: "선반"}),
                 (8, {col["U"]: 5, col["높이"]: 1, col["면"]: "선반", col["장비명"]: "NAS-01", col["모델"]: "RS2821RP+",
                      col["구분"]: "스토리지"}))
        for rr, vv in rows2:
            for c, v in vv.items():
                x = cell(rk, rr, c)
                x.setValue(v) if isinstance(v, int) else x.setString(v)
        R.calculateAll()
        u6 = cell(rk, 6 + (42 - 6), 3).getString()
        cnt = cell(rk, 49, 15).getString()
        log(u6.startswith("선반-U05") and not u6.startswith("⚠") and cnt.startswith("2"),
            f"선반 + 선반 위 장비: 겹침 경고 없음 · 장비 수에서 선반 제외 (U6 '{u6}', {cnt})")
        # 6) 저장 · 인쇄(PDF)
        R.storeToURL(uno.systemPathToFileUrl(os.path.join(out, "rack_after.xlsx")),
                     (pv("FilterName", "Calc MS Excel 2007 XML"),))
        R.storeToURL(uno.systemPathToFileUrl(os.path.join(out, "rack_after.pdf")), (pv("FilterName", "calc_pdf_Export"),))
        log(os.path.getsize(os.path.join(out, "rack_after.pdf")) > 10000, "랙 파일 전체 PDF 인쇄(내보내기)")
        # 7) 예시 시트를 예시 랙 파일에 추가하는 경우 (같은 모델 추가)
        rx = os.path.join(out, "rack_example.xlsx")
        shutil.copy(os.path.join(XL, "2_랙예시_DC1-R02.xlsx"), rx)
        RX = load(desk, rx)
        RX.Sheets.importSheet(L, "예시-SRV-01", RX.Sheets.Count)
        s2 = RX.Sheets.getByIndex(RX.Sheets.Count - 1)
        s2.Name = "U14 SRV-04"
        rk2 = RX.Sheets.getByName("랙")
        row = next(r for r in range(6, 46) if cell(rk2, r, col["장비명"]).getString() == "")   # 목록의 첫 빈 줄
        for c, v in {col["U"]: 14, col["높이"]: 2, col["면"]: "양면", col["장비명"]: "SRV-04", col["모델"]: "DL380 Gen10",
                     col["구분"]: "서버", col["시트"]: "U14 SRV-04"}.items():
            x = cell(rk2, row, c)
            x.setValue(v) if isinstance(v, int) else x.setString(v)
        RX.calculateAll()
        t = cell(rk2, 6 + (42 - 15), 3).getString()
        log(t.startswith("SRV-04"), f"예시 랙에 같은 모델 추가 → 실장도 U15에 표시 ('{t}')")
        RX.storeToURL(uno.systemPathToFileUrl(os.path.join(out, "rack_example_after.pdf")), (pv("FilterName", "calc_pdf_Export"),))
        RX.close(True)
        R.close(True)
        L.close(True)
    finally:
        try:
            desk.terminate()
        except Exception:
            pass
        proc.wait(timeout=30)
    with open(os.path.join(out, "qa_log.txt"), "w", encoding="utf-8") as f:
        for s, m in LOG:
            f.write(f"{s}  {m}\n")
    print(f"{sum(1 for s, _ in LOG if s == 'PASS')}/{len(LOG)} PASS")


if __name__ == "__main__":
    main(sys.argv[1])
