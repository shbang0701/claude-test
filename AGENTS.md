# AGENTS.md

IT 인프라 문서 템플릿 · 에셋 저장소입니다. Excel 장비 도면, PPT 구성요소, 아이콘 60종과 이 파일들을 다시 만드는 Python 소스(`_build/`)가 있습니다. 결과물은 폐쇄망의 Office 2016에서 매크로 없이 씁니다.

## 문서 지도

| 파일 | 내용 | 언제 읽나 |
|---|---|---|
| `README.md` | 구성 · 사용법 · 도면 기준 · 검수 결과 · 다시 만들기 | 작업을 시작할 때 |
| `design.md` | GUI 화면 디자인 기준 (시각 + 시각으로 인한 사용자 경험) | GUI 작업을 시작하기 전 |
| `.claude/skills/ui-ux-pro-max/` | UI/UX 점검 스킬 (UX 지침 · 차트 · 기술 스택별 지침 검색) | GUI 작업 중 필요할 때 |

`design.md`는 GUI 작업에만 필요해서 매번 불러오지 않습니다. 그래서 `@` 가져오기 대신 이름으로만 가리킵니다.

## GUI 작업

GUI 화면을 만들거나 고치는 작업이면 **시작하기 전에 `design.md`를 끝까지 읽고** 따릅니다.

- 해당하는 일: 화면 배치, 색, 글꼴, 간격, 모서리, 그림자, 아이콘, 움직임, 상태 표시(불러오는 중 · 비어 있음 · 오류), 버튼 · 메뉴 · 오류 문구
- 색 · 크기 · 간격 · 시간은 `design.md`의 토큰만 씁니다. 새 값이 필요하면 `design.md`에 먼저 추가하고 이유를 한 줄 남깁니다.
- `design.md`에 없는 것은 Windows 기본 컨트롤(Fluent) 동작을 따릅니다.
- 끝내기 전에 `design.md`의 「완료 전 점검」을 확인하고, 지키지 못한 항목은 결과 보고에 적습니다.
- Excel · PPT · 아이콘 같은 문서 산출물은 GUI가 아닙니다. 그 기준은 아래 「문서 작업」과 `README.md`에 있습니다.

### ui-ux-pro-max 스킬

`.claude/skills/ui-ux-pro-max/`에 들어 있고, 로컬 · 클라우드 세션 모두 자동으로 불러옵니다. `design.md`를 보완하는 점검 도구로만 씁니다.

- 쓰는 것: `--domain ux`(UX 지침 · 안티패턴), `--domain chart`(차트 고르기), 정해진 기술 스택의 지침(`--stack winui` · `wpf` · `avalonia` · `uno` · `uwp` 등), `references/quick-reference.md`
- 쓰지 않는 것: `--design-system` · `--persist`(디자인 시스템은 `design.md`입니다). `style` · `color` · `typography` · `google-fonts` · `icons` · `gsap` · `landing` · `product` 검색의 추천(색 · 글꼴 · 아이콘 · 움직임은 `design.md`가 정하고, 외부 글꼴 · 아이콘 패키지는 폐쇄망에서 쓸 수 없습니다)
- 검색 결과의 수치가 `design.md`와 다르면 `design.md`를 따릅니다. 예: 스킬의 누르는 영역 44px · 본문 16px · 차트 색 → `design.md`의 32 · 14 · 용도 색
- 실행: 저장소 루트에서 `python3 .claude/skills/ui-ux-pro-max/scripts/search.py "<검색어>" --domain ux` (Python 3 표준 라이브러리만 쓰고 네트워크를 쓰지 않음)
- 출처: [nextlevelbuilder/ui-ux-pro-max-skill](https://github.com/nextlevelbuilder/ui-ux-pro-max-skill) v2.15.0(`623a9a9`), MIT. 바꾼 점: 스크립트 경로, `SKILL.md` 맨 위 안내, `scripts/tests` 제외. 업데이트할 때는 원본의 `.claude/skills/ui-ux-pro-max/`를 다시 복사하고 이 세 가지를 다시 적용합니다.

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

## 문서 글쓰기 (README · design.md · AGENTS.md · 설계 메모 · 제안서)

Markdown 문서를 새로 쓰거나 크게 고칠 때는 [anthropics/skills](https://github.com/anthropics/skills)의 `doc-coauthoring` 흐름을 따릅니다: 맥락 모으기 → 구조 잡고 다듬기 → 새로 읽는 사람 테스트(서브에이전트). 오타나 한두 줄 같은 작은 수정에는 쓰지 않습니다.

- **Claude Code (로컬)**: 저장소의 `.claude-plugin/marketplace.json`(마켓플레이스 `repo-skills`)이 이 스킬 하나만 원본 저장소에서 가져옵니다. 처음 한 번 `claude plugin install doc-coauthoring@repo-skills --scope project`를 실행합니다.
- **스킬이 없는 환경 (클라우드 세션 등)**: 원본 [`skills/doc-coauthoring/SKILL.md`](https://github.com/anthropics/skills/blob/main/skills/doc-coauthoring/SKILL.md)를 읽고 그 흐름을 따릅니다.
- 이 스킬은 라이선스 표시가 없어 저장소에 복사하지 않습니다.
- 저장소 문서는 저장소 안의 Markdown 파일로 씁니다. 아티팩트나 외부 문서로 만들지 않습니다.
- 문체는 `README.md`와 같게 씁니다: 짧은 문장, 표, `-습니다` · `-하세요`.
