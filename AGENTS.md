# AGENTS.md

IT 인프라 문서 템플릿 · 에셋 저장소입니다. Excel 장비 도면, PPT 구성요소, 아이콘 60종과 이 파일들을 다시 만드는 Python 소스(`_build/`)가 있습니다. 결과물은 폐쇄망의 Office 2016에서 매크로 없이 씁니다.

## 작업별 스킬

작업 주제에 맞는 행을 골라 그 스킬만 씁니다. 한 작업에 주제가 섞이면 주제마다 해당 행을 따릅니다.

| 작업 | 먼저 읽을 것 | 쓰는 스킬 | 쓰지 않는 것 |
|---|---|---|---|
| GUI 화면 만들기 · 고치기 · 검토 (화면 안 차트 포함) | `design.md` | `ui-ux-pro-max` | 스킬의 스타일 · 색 · 글꼴 · 아이콘 추천, `frontend-design` · `theme-factory` · `brand-guidelines` |
| Excel · PPT · PDF · Word 파일 | `README.md` | `xlsx` · `pptx` · `pdf` · `docx` | 슬라이드 · 문서 아티팩트, pptxgenjs, 결과 파일만 직접 고치기 |
| 저장소에 둘 Markdown 문서 (README · design.md · AGENTS.md · 설계 메모 · 제안서) | 고칠 문서 | `doc-writing` (→ doc-coauthoring 흐름) | Claude Docs · 아티팩트 |
| 아이콘 SVG | `README.md` | 없음 (`_build/icons.py` · `_build/icondsl.py`) | 외부 아이콘 세트 · 이미지 생성 |

- 사용자가 형식이나 도구를 직접 지정하면(예: "Word로", "Claude Docs로") 그것을 따릅니다.
- 오타 · 한두 줄 수정 같은 작은 일에는 스킬을 부르지 않습니다.
- 위 주제(GUI · 문서 파일 · Markdown 문서 · 아이콘)에는 표의 스킬만 씁니다. 세션에 다른 스킬이 보여도 쓰지 않습니다.

| 스킬 | 로컬 Claude Code | 클라우드 세션 (claude.ai/code) |
|---|---|---|
| `ui-ux-pro-max` · `doc-writing` | 저장소 `.claude/skills/` (자동) | 같음 (자동) |
| `xlsx` · `pptx` · `pdf` · `docx` | `.claude/settings.json`의 `document-skills` 플러그인 (폴더를 신뢰하면 자동, `/plugin`에서 확인) | 계정에서 켠 스킬 `anthropic-skills:*` |

- 필요한 스킬이 세션에 없으면 흉내 내지 말고, 작업 전에 사용자에게 알립니다.

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

### ui-ux-pro-max 스킬

`design.md`를 보완하는 점검 도구로만 씁니다.

- 쓰는 것: `--domain ux`(UX 지침 · 안티패턴), `--domain chart`(차트 고르기), 정해진 기술 스택의 지침(`--stack winui` · `wpf` · `avalonia` · `uno` · `uwp` 등), `references/quick-reference.md`
- 쓰지 않는 것: `--design-system` · `--persist`(디자인 시스템은 `design.md`입니다). `style` · `color` · `typography` · `google-fonts` · `icons` · `gsap` · `landing` · `product` 검색의 추천(색 · 글꼴 · 아이콘 · 움직임은 `design.md`가 정하고, 외부 글꼴 · 아이콘 패키지는 폐쇄망에서 쓸 수 없습니다)
- 검색 결과의 수치가 `design.md`와 다르면 `design.md`를 따릅니다. 예: 스킬의 누르는 영역 44px · 본문 16px · 차트 색 → `design.md`의 32 · 14 · 용도 색
- 실행: 저장소 루트에서 `python3 .claude/skills/ui-ux-pro-max/scripts/search.py "<검색어>" --domain ux` (Python 3 표준 라이브러리만 쓰고 네트워크를 쓰지 않음)
- 출처: [nextlevelbuilder/ui-ux-pro-max-skill](https://github.com/nextlevelbuilder/ui-ux-pro-max-skill) v2.15.0(`623a9a9`), MIT. 바꾼 점: 설명문(이 저장소 GUI 화면 전용), 스크립트 경로, `SKILL.md` 맨 위 안내, `scripts/tests` 제외. 업데이트할 때는 원본의 `.claude/skills/ui-ux-pro-max/`를 다시 복사하고 이 네 가지를 다시 적용합니다.

## 문서 작업 (Excel · PPT · PDF · Word)

문서 파일을 만들거나 읽거나 고칠 때는 [anthropics/skills](https://github.com/anthropics/skills)의 문서 스킬을 씁니다: `.xlsx` → `xlsx`, `.pptx` → `pptx`, `.pdf` → `pdf`, `.docx` → `docx`. 이 스킬들은 Anthropic 서비스 안에서만 쓸 수 있어(라이선스상 복사 · 배포 금지) 저장소에 복사하지 않습니다.

**이 저장소의 기준이 스킬보다 우선합니다.** 스킬은 읽기 · 검수 · PDF 처리에 쓰고, 만드는 방식은 아래를 따릅니다.

- 산출물은 `_build/` 스크립트로 만듭니다. 결과 파일만 고치면 다시 만들 때 사라지므로, 스크립트를 고친 뒤 다시 만듭니다(`README.md` 「다시 만들기」).
- PPT는 `_build/ppt_kit.py`(python-pptx)로 만듭니다. 스킬의 pptxgenjs 방식과 디자인 제안(색 팔레트 · 글꼴)은 쓰지 않습니다.
- 색은 `_build/ppt_kit.py` · `_build/xlkit.py`의 `P`를 RGB로 직접 지정하고, 글꼴은 맑은 고딕(케이블 라벨만 Consolas)입니다. 테마 색 · 그림자 효과는 쓰지 않습니다. 스킬의 기본 글꼴(Arial 등)은 따르지 않습니다.
- 수식 값은 `_build/inject_cache.py`로 넣습니다. 스킬의 `recalc.py`는 LibreOffice로 파일을 다시 저장해 인쇄 설정이 바뀔 수 있으므로, 원본에는 쓰지 않고 사본 검사에만 씁니다.
- 대상은 Office 2016 영구 라이선스입니다. Excel 함수는 2016에 있는 것만 씁니다(`TEXTJOIN` · `CONCAT` · `IFS` · `SWITCH` · `MAXIFS` · `MINIFS` · `XLOOKUP` · 동적 배열 함수 금지). 매크로 · 외부 서비스는 쓰지 않습니다. PPT에는 SVG 대신 도형 아이콘을 넣습니다.

## 문서 글쓰기 (Markdown)

README · `design.md` · AGENTS.md · 설계 메모 · 제안서를 새로 쓰거나 크게 고칠 때는 `doc-writing` 스킬을 씁니다. 이 스킬이 [anthropics/skills](https://github.com/anthropics/skills)의 `doc-coauthoring` 흐름(맥락 모으기 → 구조 잡고 다듬기 → 새로 읽는 사람 테스트)을 원본에서 불러오고, 저장소 규칙을 덧붙입니다.

- 결과는 저장소 안의 Markdown 파일이고, 문체는 `README.md`와 같습니다. 자세한 규칙은 `.claude/skills/doc-writing/SKILL.md`에 있습니다.
- `doc-coauthoring`은 라이선스 표시가 없어 저장소에 복사하지 않습니다.
