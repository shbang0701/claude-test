# AGENTS.md

IT 인프라 문서 템플릿 · 에셋 저장소입니다. Excel 장비 도면, PPT 구성요소, 아이콘 60종과 이 파일들을 다시 만드는 Python 소스(`_build/`)가 있습니다. 결과물은 폐쇄망의 Office 2016에서 매크로 없이 씁니다.

## 문서 지도

| 파일 | 내용 | 언제 읽나 |
|---|---|---|
| `README.md` | 구성 · 사용법 · 도면 기준 · 검수 결과 · 다시 만들기 | 작업을 시작할 때 |
| `design.md` | GUI 화면 디자인 기준 (시각 + 시각으로 인한 사용자 경험) | GUI 작업을 시작하기 전 |

`design.md`는 GUI 작업에만 필요해서 매번 불러오지 않습니다. 그래서 `@` 가져오기 대신 이름으로만 가리킵니다.

## GUI 작업

GUI 화면을 만들거나 고치는 작업이면 **시작하기 전에 `design.md`를 끝까지 읽고** 따릅니다.

- 해당하는 일: 화면 배치, 색, 글꼴, 간격, 모서리, 그림자, 아이콘, 움직임, 상태 표시(불러오는 중 · 비어 있음 · 오류), 버튼 · 메뉴 · 오류 문구
- 색 · 크기 · 간격 · 시간은 `design.md`의 토큰만 씁니다. 새 값이 필요하면 `design.md`에 먼저 추가하고 이유를 한 줄 남깁니다.
- `design.md`에 없는 것은 Windows 기본 컨트롤(Fluent) 동작을 따릅니다.
- 끝내기 전에 `design.md`의 「완료 전 점검」을 확인하고, 지키지 못한 항목은 결과 보고에 적습니다.
- Excel · PPT · 아이콘 같은 문서 산출물은 GUI가 아닙니다. 그 기준은 아래 「문서 작업」과 `README.md`에 있습니다.

## 문서 작업 (Excel · PPT · PDF · Word)

문서 파일을 만들거나 읽거나 고칠 때는 [anthropics/skills](https://github.com/anthropics/skills)의 문서 스킬을 씁니다: `.xlsx` → `xlsx`, `.pptx` → `pptx`, `.pdf` → `pdf`, `.docx` → `docx`.

- **Claude Code (터미널 · 데스크톱 · IDE)**: `.claude/settings.json`이 마켓플레이스 `anthropic-agent-skills`를 등록하고 `document-skills` 플러그인을 켭니다. 폴더를 신뢰하면 적용되며, `/plugin`에서 확인할 수 있습니다.
- **클라우드 세션 (claude.ai/code)**: 저장소 설정의 플러그인을 불러오지 않습니다. 세션에 같은 스킬(`anthropic-skills:xlsx` 등)이 있으면 그것을 씁니다.
- 문서 스킬은 Anthropic 서비스 안에서만 쓸 수 있습니다(라이선스상 복사 · 배포 금지). 이 저장소에 복사해 넣지 않습니다.

**이 저장소의 기준이 스킬보다 우선합니다.** 스킬은 읽기 · 검수 · PDF 처리에 쓰고, 만드는 방식은 아래를 따릅니다.

- 산출물은 `_build/` 스크립트로 만듭니다. 결과 파일만 고치면 다시 만들 때 사라지므로, 스크립트를 고친 뒤 다시 만듭니다(`README.md` 「다시 만들기」).
- PPT는 `_build/ppt_kit.py`(python-pptx)로 만듭니다. 스킬의 pptxgenjs 방식과 디자인 제안(색 팔레트 · 글꼴)은 쓰지 않습니다.
- 색은 `_build/ppt_kit.py` · `_build/xlkit.py`의 `P`를 RGB로 직접 지정하고, 글꼴은 맑은 고딕(케이블 라벨만 Consolas)입니다. 테마 색 · 그림자 효과는 쓰지 않습니다. 스킬의 기본 글꼴(Arial 등)은 따르지 않습니다.
- 수식 값은 `_build/inject_cache.py`로 넣습니다. 스킬의 `recalc.py`는 LibreOffice로 파일을 다시 저장해 인쇄 설정이 바뀔 수 있으므로, 원본에는 쓰지 않고 사본 검사에만 씁니다.
- 대상은 Office 2016 영구 라이선스입니다. Excel 함수는 2016에 있는 것만 씁니다(`TEXTJOIN` · `CONCAT` · `IFS` · `SWITCH` · `MAXIFS` · `MINIFS` · `XLOOKUP` · 동적 배열 함수 금지). 매크로 · 외부 서비스는 쓰지 않습니다. PPT에는 SVG 대신 도형 아이콘을 넣습니다.
