---
name: infra-design
description: IT 인프라 문서 세트(docs/templates/infra-kit)의 디자인 기준 — 색 팔레트, 맑은 고딕 글꼴 크기, 선·모서리·간격, PPT 슬라이드 레이아웃, Excel 격자 도면 규칙, 아이콘 규칙. 이 세트의 PPT 슬라이드·구성도·요소, Excel 장비 도면·양식·부품, 아이콘, PDF 가이드를 새로 만들거나 고칠 때, 또는 그 색·글꼴·배치를 다룰 때 먼저 읽는다. 다른 Office 문서를 같은 스타일로 만들 때도 쓴다. WorkTool 앱 화면에는 쓰지 않는다.
---

# IT 인프라 문서 디자인 기준

이 문서의 경로는 모두 `docs/templates/infra-kit/` 기준이다.

목표는 **어느 회사 양식에 붙여도, 어느 장비를 그려도 같은 모양**이 나오는 것이다. 새 요소를 만들 때는 아래 기준 안에서만 고르고, 기준에 없는 값(새 색, 새 글꼴, 임의 크기)을 만들지 않는다.

## 먼저 지킬 원칙

1. **결과물은 직접 고치지 않는다.** `.xlsx` · `.pptx` · `.pdf` · 아이콘 파일은 `_build/`의 Python 소스로 다시 만든다. 디자인을 바꾸려면 소스를 고친다.
2. **기준값은 코드에 있다.** 팔레트는 `_build/xlkit.py`와 `_build/ppt_kit.py`의 `P` 딕셔너리(두 곳의 값이 같아야 함), 연결선은 `_build/ppt_elements.py`의 `LINES`, 영역 박스는 `ZONES`. 새 hex 값을 코드에 직접 쓰지 말고 `P["svc"]`처럼 키로 참조한다.
3. **색은 RGB 직접 지정, 테마 색은 쓰지 않는다.** 다른 양식에 붙여도 색이 바뀌지 않게 하기 위해서다.
4. **글꼴은 맑은 고딕만 쓴다.** 예외는 케이블 라벨과 hex 표기에 쓰는 `Consolas` 하나. Aptos, Calibri 같은 글꼴은 Office 2016에 없거나 양식마다 달라 쓰지 않는다.
5. **그림자, 그라데이션, 3D, 테마 효과는 쓰지 않는다.** 납작한 단색 + 가는 테두리만 쓴다.
6. **같은 의미 = 같은 색.** 연결 용도 색(서비스=파랑, 관리=청록 …)은 Excel 포트 칸, 표 번호 칸, PPT 연결선, 영역 박스에서 모두 같다.
7. **Office 2016에서 열린다.** 매크로, SVG 삽입, 외부 서비스에 기대지 않는다.

## 빠른 참조

| 항목 | 기준 |
|---|---|
| 기본 글자색 | 잉크 `#1E2A36` · 강조/제목 남색 `#1F3A5F` · 보조 회색 `#677789` |
| 테두리 | 0.75pt `#B4BFCB` (PPT) / 가는 선 `#8C99A8`, 표 줄 `#D5DCE4` (Excel) |
| PPT 슬라이드 | 16:9 (13.333 × 7.5in), 여백 0.5in, 요소 사이 0.3in |
| PPT 글자 | 제목 24 굵게 · 소제목 14 굵게 · 본문 11–12 · 장비 이름 10.5 굵게 · 표·설명 9–9.5 · 각주 8 |
| Excel 격자 | 1칸 ≈ 4.2mm (열 너비 2.0 · 행 높이 12pt), 모든 시트 동일 |
| Excel 인쇄 | A4 가로 · 100% · 바닥글 자동 (배율 줄이지 않음) |
| 아이콘 | 48×48 단색, 기본 남색 `#1F3A5F` / 어두운 바탕 흰색 |

## 작업별 상세 문서

필요한 것만 읽는다.

- 색 전체 표(기본 · 용도 · 상태 · 영역 바탕): [references/palette.md](references/palette.md)
- PPT 슬라이드, 요소, 연결선, 영역 박스: [references/ppt.md](references/ppt.md)
- Excel 장비 도면, 셀 스타일, 쪽 구성: [references/excel.md](references/excel.md)
- 아이콘 추가와 수정: [references/icons.md](references/icons.md)

## 작업 순서

1. 비슷한 기존 요소를 먼저 찾는다 (PPT: `ppt_elements.py` · `ppt_slides_b.py` · `ppt_slides_c.py`, Excel: `xlparts.py` · `xldevices.py` · `xlforms.py`). 새로 그리기보다 기존 헬퍼를 조합한다.
2. 위 기준값과 `P` 키만 써서 소스를 고친다.
3. 다시 만들고 눈으로 확인한다. 렌더링해서 PNG로 직접 본 뒤에만 "됐다"고 말한다.

   ```bash
   cd docs/templates/infra-kit/_build
   python3 build_ppt.py        # 또는 build_excel.py / build_icons.py / build_docs.py
   python3 -c "import lo; print(lo.convert('../02_PPT_구성요소/IT인프라_PPT_구성요소_라이브러리.pptx', 'pdf', '/tmp/render'))"
   pdftoppm -png -r 60 -f <쪽> -l <쪽> /tmp/render/*.pdf /tmp/render/p   # 바뀐 쪽만 이미지로 확인
   ```

   Excel을 고쳤다면 README의 "다시 만들기" 순서대로 `inject_cache.py`까지 실행하고, 가능하면 `qa_workflow.py`로 검수한다.
4. 확인할 것: 글자 넘침과 잘림, 요소 겹침, 정렬선(0.5in 여백 · 같은 x 좌표), 색이 용도 규칙과 맞는지, 테마 색이나 다른 글꼴이 섞이지 않았는지.
5. 기준 자체를 바꿨다면(새 색 · 새 크기) 코드의 기준값, PPT [디자인 기준] 슬라이드(`ppt_slides_a.design`), Excel [기준] 시트(`xlguide.build_standard`), README의 표, 이 스킬 문서를 함께 고친다.
