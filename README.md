# 업무 작업 공간

회사 업무 PC에서 쓰는 PowerShell 스크립트, GUI 앱, 운영 문서 모음입니다.
환경·보안·코드 규칙은 [AGENTS.md](AGENTS.md)에 있습니다.

## 구성

- **[scripts/](scripts/)** — 복사 한 번으로 끝나는 단독 PowerShell 스크립트
  - 아직 없음. 새로 만들 때는 [scripts/README.md](scripts/README.md)
- **[apps/WorkTool/](apps/WorkTool/)** — 업무 통합 GUI 앱 (메모 · 근태 · 일정·알람, 트레이 상주)
  - [design/](apps/WorkTool/design/) — 새 화면 시안 (WPF 전환 기준)
- **[docs/guides/](docs/guides/)** — 읽는 가이드
  - [productivity/](docs/guides/productivity/) — 업무 효율·작업 환경
    - [갤럭시 S26 울트라 앱별 키보드 가이드](docs/guides/productivity/galaxy-s26-ultra-app-keyboard-guide.md)
    - [PowerShell 7 · Solarized Light 다듬기 가이드](docs/guides/productivity/powershell-solarized-light-guide.md)
  - ops/ — 운영 노하우·장애 대응 (아직 없음)
- **[docs/templates/](docs/templates/)** — 받아서 쓰는 Office 양식
  - [infra-kit/](docs/templates/infra-kit/) — IT 인프라 문서 세트: 장비 도면(Excel), PPT 구성 요소, 아이콘 60종
  - [rhel-config/](docs/templates/rhel-config/) — RHEL 계열 7·8·9 서버 구성표

## 업무 PC로 가져가기

**git이 있으면** (`git --version`으로 확인)

```powershell
git clone https://github.com/shbang0701/claude-test.git
```

그다음부터는 그 폴더에서 `git pull`만 하면 최신으로 바뀝니다.

**git이 없으면** 저장소 첫 화면의 초록색 **Code** 버튼 → **Download ZIP**을 받습니다.
ZIP에서 푼 파일에는 "인터넷에서 받은 파일" 표시가 붙어, 회사 실행 정책에 따라 `.ps1` 실행이 막힐 수 있습니다. 그때는 아래처럼 복사해 붙여 넣으세요.

**파일 하나만 필요하면**
- 스크립트: 파일을 열고 오른쪽 위 복사 아이콘(Copy raw file) → 업무 PC에서 같은 이름으로 저장
- Excel·PPT: 파일을 열고 오른쪽 위 다운로드 아이콘(Download raw file)

## 이름 규칙

- 폴더·파일 이름은 영문 소문자와 하이픈으로 짓고, 설명은 한글로 씁니다.
- 스크립트 이름은 `동사-명사.ps1` 형식입니다 (예: `Export-SheetList.ps1`).
- 양식 세트 안의 한글 파일 이름은 그대로 둡니다.
- 파일을 추가하거나 옮기면 위 **구성** 목록도 같이 고칩니다.
