# -*- coding: utf-8 -*-
"""PPT 산출물: 02_PPT_구성요소/IT인프라_PPT_구성요소_라이브러리.pptx"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ppt_kit import new_presentation, notes
import ppt_slides_a as A
import ppt_slides_b as B

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "02_PPT_구성요소")


def build(path, with_samples=True):
    prs = new_presentation()
    A.cover(prs)
    # 목차 쪽 번호 계산용 계획
    plan = [("시작", "사용법 · 목차"), ("기준", "디자인 기준"), ("div", "아이콘"),
            ("아이콘", "컴퓨팅 · 스토리지"), ("아이콘", "네트워크 · 설비"), ("아이콘", "사람 · 운영 + 스타일 변형"),
            ("div", "구성 요소"), ("요소", "장비"), ("요소", "랙 · 하드웨어"), ("요소", "영역 박스"), ("요소", "연결선"),
            ("요소", "라벨 · 태그 · 번호"), ("요소", "범례 · 주석 · 강조"), ("요소", "표"), ("요소", "비교"), ("요소", "현황"),
            ("요소", "작업 절차")]
    samples = []
    if with_samples:
        import ppt_slides_c as C
        samples = C.SAMPLES
        plan += [("div", "슬라이드 샘플")] + [("샘플", t) for t, _ in samples]
    toc = []
    no = 2
    for kind, t in plan:
        if kind != "div":
            toc.append([kind, str(no), t])
        no += 1
    A.howto(prs, toc)
    A.design(prs)
    A.divider(prs, 1, "아이콘 60종", "IT 인프라 아이콘을 PowerPoint 도형으로 만들어 두었습니다.\n"
                                  "복사해서 쓰고, [도형 채우기]로 색을 바꾸세요. 같은 아이콘의 SVG · PNG는 03_아이콘 폴더에 있습니다.")
    A.icon_slide(prs, ["compute", "storage"], "아이콘 — 컴퓨팅 · 스토리지")
    A.icon_slide(prs, ["network", "facility"], "아이콘 — 네트워크 · 설비")
    A.icon_slide(prs, ["people", "ops"], "아이콘 — 사람 · 운영 상태", extra=A.icon_variants)
    A.divider(prs, 2, "구성 요소", "장비 · 영역 · 연결선 · 라벨 · 범례 · 표 · 비교 · 현황 · 절차.\n"
                               "필요한 요소만 골라 복사하세요. 모든 요소는 이름이 붙은 그룹입니다 (선택 창 Alt+F10).")
    for fn in (B.nodes, B.racks, B.zones, B.lines_slide, B.labels, B.legends, B.tables, B.compare, B.dashboard,
               B.procedure):
        fn(prs)
    if with_samples:
        A.divider(prs, 3, "슬라이드 샘플", "요소를 조합한 완성 슬라이드입니다. 통째로 복사해 내용만 바꾸거나,\n"
                                      "일부 영역만 떼어 쓰세요. 회사 양식에 붙일 때는 [홈 > 새 슬라이드 > 슬라이드 다시 사용]도 가능합니다.")
        for t, fn in samples:
            fn(prs)
    prs.save(path)
    return path


if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    p = build(os.path.join(OUT, "IT인프라_PPT_구성요소_라이브러리.pptx"), with_samples="--no-samples" not in sys.argv)
    print(p)
