# 업무 작업 공간

회사 업무 PC에서 쓰는 PowerShell 스크립트, GUI 앱, 운영 문서 모음입니다.
환경·보안·코드 규칙은 [AGENTS.md](AGENTS.md)에 있습니다.

## 구성

- **[scripts/](scripts/)** — 복사 한 번으로 끝나는 단독 PowerShell 스크립트
  - 아직 없음. 새로 만들 때는 [scripts/README.md](scripts/README.md)
- **[apps/design.md](apps/design.md)** — 모든 GUI 앱의 공통 디자인 원칙 (Windows 디자인 지침 기반)
- **[worktool.zip](worktool.zip)** — WorkTool 새 버전(WPF)을 바로 실행하는 묶음. 압축을 풀고 `Start-WorkTool.cmd`를 더블클릭 (개인 PC용, 아래 「업무 PC로 가져가기」 참고)
- **[apps/WorkTool/](apps/WorkTool/)** — 업무 통합 GUI 앱 (메모 · 근태 · 일정·알람, 트레이 상주)
  - [design/](apps/WorkTool/design/) — 새 화면 시안 (WPF 전환 기준)
  - [wpf/](apps/WorkTool/wpf/) — 새 버전(WPF) 1단계 뼈대 + 마크다운 보기. 파일 하나 버전은 `apps/WorkTool/dist/WorkTool-WPF.ps1`
- **[docs/guides/](docs/guides/)** — 읽는 가이드
  - [productivity/](docs/guides/productivity/) — 업무 효율·작업 환경
    - [갤럭시 S26 울트라 앱별 키보드 가이드](docs/guides/productivity/galaxy-s26-ultra-app-keyboard-guide.md)
    - [PowerShell 7 · Solarized Light 다듬기 가이드](docs/guides/productivity/powershell-solarized-light-guide.md)
  - ops/ — 운영 노하우·장애 대응 (아직 없음)
- **[docs/templates/](docs/templates/)** — 받아서 쓰는 Office 양식
  - [infra-kit/](docs/templates/infra-kit/) — IT 인프라 문서 세트: 장비 도면(Excel), PPT 구성 요소, 아이콘 60종
  - [rhel-config/](docs/templates/rhel-config/) — RHEL 계열 7·8·9 서버 구성표

## 업무 PC로 가져가기

흔적이 남지 않는 방법은 없습니다. 복붙해도 실행한 코드는 PowerShell 기록에 남을 수 있습니다.
그래서 **가져오는 방법보다 코드 내용이 설명되는지가 중요**하고, 그중에서 보안 장비가 덜 민감하게 보는 방법을 씁니다.

**스크립트·문서(텍스트): 복사해 붙여 넣기 (기본)**
- 파일을 열고 오른쪽 위 복사 아이콘(Copy raw file) → 업무 PC에서 같은 이름으로 저장
- WorkTool은 파일 하나로 묶은 `apps/WorkTool/dist/WorkTool.ps1`만 복사하면 됩니다.

**개인 PC: worktool.zip**
- 받은 `worktool.zip` 우클릭 → 속성 → 「차단 해제」 체크 → 압축 풀기 → `Start-WorkTool.cmd` 더블클릭
- 차단 해제를 안 하면 인터넷에서 받은 스크립트라 실행 정책(RemoteSigned)에 막힙니다. 실행 파일이 이 경우를 알려 줍니다.
- 업무 PC에는 쓰지 않습니다. 스크립트가 든 zip은 보안 장비가 가장 민감하게 봅니다 (아래 「쓰지 않는 것」). 업무 PC는 `WorkTool-WPF.ps1`을 복사해 붙여 넣습니다.

**Excel·PPT 양식: 파일 하나씩 다운로드**
- 파일을 열고 오른쪽 위 다운로드 아이콘(Download raw file)

**쓰지 않는 것**
- ZIP 통째 다운로드: 스크립트가 든 압축 파일은 보안 장비가 가장 민감하게 보고, 풀어도 실행이 막힐 수 있습니다.
- git 새로 설치: 승인되지 않은 프로그램 설치가 됩니다. 회사에서 이미 git을 업무에 쓰고 있을 때만 `git clone` / `git pull`을 씁니다.

## 이름 규칙

- 폴더·파일 이름은 영문 소문자와 하이픈으로 짓고, 설명은 한글로 씁니다.
- 스크립트 이름은 `동사-명사.ps1` 형식입니다 (예: `Export-SheetList.ps1`).
- 양식 세트 안의 한글 파일 이름은 그대로 둡니다.
- 파일을 추가하거나 옮기면 위 **구성** 목록도 같이 고칩니다.
