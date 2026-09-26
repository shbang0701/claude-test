# -*- coding: utf-8 -*-
"""실사용 상황 점검 — 내용 교체 / 긴 설명 / 단계 추가 / 이미지 교체 / 다른 양식에 복사 / 인쇄(PDF).
   LibreOffice 로 변환해 결과를 확인한다(실제 Office 2016 검증은 별도)."""
import os, sys, copy, subprocess, shutil
HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Cm
from PIL import Image

LIB = os.path.join(ROOT, "library")
WORK = "/tmp/selftest"
TPL = os.path.join(LIB, "01_문서_템플릿.docx")
BLK = os.path.join(LIB, "02_Word_블록.docx")

def pdf(path, tag):
    out = os.path.join(WORK, tag)
    os.makedirs(out, exist_ok=True)
    subprocess.run(["soffice", "--headless", "--norestore", "--convert-to", "pdf",
                    "--outdir", out, path], check=True, capture_output=True)
    f = [x for x in os.listdir(out) if x.endswith(".pdf")][0]
    p = os.path.join(out, f)
    n = int(subprocess.run(["pdfinfo", p], capture_output=True, text=True
                           ).stdout.split("Pages:")[1].split()[0])
    return p, n

def paras(doc): return doc.paragraphs

def find(doc, text):
    for p in doc.paragraphs:
        if text in p.text: return p
    return None

def set_text(p, s):
    for r in p.runs[1:]: r._r.getparent().remove(r._r)
    if p.runs: p.runs[0].text = s
    else: p.add_run(s)

def clone_after(p, style=None, text=None):
    new = copy.deepcopy(p._p)
    p._p.addnext(new)
    from docx.text.paragraph import Paragraph
    np = Paragraph(new, p._parent)
    if text is not None: set_text(np, text)
    if style: np.style = style
    return np

def t1_edit_and_add_steps():
    """내용 교체 + 긴 설명 + 단계 중간 삽입 → 번호가 자동으로 밀리는지."""
    doc = Document(TPL)
    # 내용 교체
    n = 0
    for p in doc.paragraphs:
        if "〈문서 제목〉" in p.text or "〈항목 1〉" in p.text:
            for r in p.runs:
                r.text = r.text.replace("〈문서 제목〉", "신입 사원 안내서").replace("〈항목 1〉", "출입 등록")
            n += 1
    # 긴 설명 입력
    tgt = find(doc, "〈보충 설명이 필요한 단계에만")
    if tgt:
        set_text(tgt, "신청서는 사내 포털의 서식 자료실에서 내려받는다. 부서장 결재가 필요한 항목은 "
                      "미리 담당자에게 확인해 두면 반려를 줄일 수 있다. 제출 후에는 수정이 어려우므로 "
                      "보내기 전에 첨부 파일과 기간을 한 번 더 확인한다.")
    # 단계 중간 삽입 (3.1 의 1번 단계 뒤)
    s1 = find(doc, "〈첫 번째로 할 일을")
    added = clone_after(s1, text="중간에 끼워 넣은 단계 — 뒤 번호가 자동으로 밀려야 한다")
    out = os.path.join(WORK, "t1_편집.docx"); doc.save(out)
    p, pages = pdf(out, "t1")
    return out, p, pages, n

def t2_replace_image():
    """이미지 교체 — 비율이 다른 새 캡처로 바꿨을 때."""
    doc = Document(TPL)
    new_img = os.path.join(WORK, "new_capture.png")
    Image.new("RGB", (1200, 1000), (222, 232, 240)).save(new_img)   # 4:3.3 세로 긴 이미지
    part = doc.part
    # 문서의 첫 번째 그림(화면 캡처)의 관계를 새 이미지로 교체
    rids = [r for r in part.rels.values() if "image" in r.reltype]
    target = [r for r in rids if r._target.partname.endswith("image2.png")] or rids
    rid, image = part.get_or_add_image(new_img)
    for blip in doc.element.body.iter(qn('a:blip')):
        blip.set(qn('r:embed'), rid)   # 모든 그림을 새 이미지로 교체
    out = os.path.join(WORK, "t2_이미지교체.docx"); doc.save(out)
    p, pages = pdf(out, "t2")
    return out, p, pages

def t3_paste_into_other_template():
    """다른 회사 양식에 블록 붙여넣기 — '원본 서식 유지'(스타일 동반)와
       '대상 서식 사용'(스타일 없이 본문만) 두 경우."""
    src = Document(BLK)
    # 대상: 다른 서식의 문서(바탕글 굴림 11pt, 여백 다름)
    dst = Document()
    sec = dst.sections[0]
    sec.left_margin = sec.right_margin = Cm(3.0)
    st = dst.styles["Normal"]
    st.font.name = "Gulim"; st.font.size = __import__("docx").shared.Pt(11)
    dst.add_paragraph("다른 회사 양식 문서입니다. 아래에 라이브러리 블록을 붙여 넣었습니다.")
    # 복사할 범위: 1.4 단계-상세형 블록 전체
    body = src.element.body
    items = list(body)
    start = None; picked = []
    for el in items:
        txt = "".join(el.itertext())
        if start is None and "단계 — 상세" in txt and el.tag.endswith('}p'):
            start = True; continue
        if start:
            if el.tag.endswith('}p') and "하위 단계와 분기" in txt: break
            picked.append(el)
    # (1) 원본 서식 유지 = 쓰이는 스타일 정의까지 함께 가져간다
    keep = Document(); ks = keep.sections[0]
    ks.left_margin = ks.right_margin = Cm(3.0)
    src_styles = {s.style_id: s for s in src.styles}
    have = {s.style_id for s in keep.styles}
    for sid, s in src_styles.items():
        if sid not in have:
            keep.styles.element.append(copy.deepcopy(s.element))
    # 번호 정의도 함께
    try:
        kn = keep.part.numbering_part.element
        for ch in list(kn): kn.remove(ch)
        for ch in list(src.part.numbering_part.element): kn.append(copy.deepcopy(ch))
    except Exception as e:
        print("   numbering copy:", e)
    keep.add_paragraph("다른 회사 양식 + ‘원본 서식 유지’로 붙여넣은 경우")
    for el in picked: keep.element.body.append(copy.deepcopy(el))
    out1 = os.path.join(WORK, "t3a_원본서식유지.docx"); keep.save(out1)
    # (2) 대상 서식 사용 = 스타일 없이 문단만
    for el in picked: dst.element.body.append(copy.deepcopy(el))
    out2 = os.path.join(WORK, "t3b_대상서식사용.docx"); dst.save(out2)
    a = pdf(out1, "t3a"); b = pdf(out2, "t3b")
    return (out1, a), (out2, b)

def t4_print_all():
    res = {}
    for f in sorted(os.listdir(LIB)):
        p = os.path.join(LIB, f)
        if not os.path.isfile(p) or os.path.splitext(f)[1].lower() not in (".docx", ".pptx", ".xlsx"):
            continue
        try:
            _, n = pdf(p, "print_" + f[:2])
            res[f] = n
        except Exception as e:
            res[f] = "실패: %s" % e
    return res

if __name__ == "__main__":
    shutil.rmtree(WORK, ignore_errors=True); os.makedirs(WORK, exist_ok=True)
    print("1) 내용 교체 · 긴 설명 · 단계 중간 삽입")
    o, p, n, cnt = t1_edit_and_add_steps()
    print(f"   치환 {cnt}곳, PDF {n}쪽 → {p}")
    print("2) 이미지 교체(세로로 긴 이미지)")
    o, p, n = t2_replace_image(); print(f"   PDF {n}쪽 → {p}")
    print("3) 다른 회사 양식에 블록 붙여넣기")
    (o1, (p1, n1)), (o2, (p2, n2)) = t3_paste_into_other_template()
    print(f"   원본 서식 유지 → {p1} ({n1}쪽)")
    print(f"   대상 서식 사용 → {p2} ({n2}쪽)")
    print("4) 전체 인쇄(PDF) 변환")
    for k, v in t4_print_all().items(): print(f"   {k}: {v}쪽")
