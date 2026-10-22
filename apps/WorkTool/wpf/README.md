# WorkTool 새 버전 (WPF) — 1단계 뼈대

지금 쓰는 WorkTool(WinForms, `apps/WorkTool/`)을 옮겨 갈 새 버전입니다.
계획과 결정 사항은 이슈 #8 「WorkTool 개편 계획 · 결정할 것」, 화면 기준은 `apps/WorkTool/design/README.md`에 있습니다.

## 1단계에 들어 있는 것

- 창 틀: 제목 줄을 직접 그린 창(최소화 · 최대화 · 닫기), 창 투명도, F11 전체 화면
- 왼쪽 메뉴: 애드온(탭) 목록으로 자동으로 만든다. 접기 · 펼치기(Ctrl+B), 아래에 설정 · 도움말
- 테마: 파스텔 7색(하늘 · 분홍 · 연두 · 노랑 · 보라 · 회색 · 검정), 창 투명도 60~100%, 글자 크기 3단계. 바꾸면 바로 반영
- 설정 화면: 테마, 메뉴에서 탭 끄기, 데이터 위치, 예전 데이터 다시 가져오기
- 예전 데이터 가져오기: 처음 실행하면 예전 WorkTool의 `data` 폴더에서 메모 · 일정 · 근태를 읽어 와서 보여 준다
- 탭: 홈(다른 탭의 카드 모음) · 일정 · 근태 · 메모 — **보기 전용**
- 트레이: 창의 X는 트레이로 숨김, 완전히 끄려면 트레이 아이콘 우클릭 → 종료

편집 · 알람 · 지금 출근 · 메모별 잠금은 2단계에서 옮깁니다. 그때까지는 예전 WorkTool을 계속 쓰세요.

## 업무 PC에서 실행하기

1. `apps/WorkTool/dist/WorkTool-WPF.ps1`을 복사해서, 예전 WorkTool 폴더(예전 `data` 폴더가 있는 곳)에 같은 이름으로 저장합니다.
2. `pwsh`에서 `.\WorkTool-WPF.ps1`로 실행합니다.
3. 확인할 것
   - 창이 뜨고, 왼쪽 메뉴에 홈 · 일정 · 근태 · 메모가 보이는지
   - 홈 · 일정 · 근태 · 메모에 예전 데이터가 보이는지
   - 설정에서 색 · 투명도 · 글자 크기를 바꾸면 바로 바뀌는지
   - 최대화 · F11 · Ctrl+B · 트레이로 숨기기 · 트레이에서 종료

오류가 나면 창에 뜬 문구(또는 콘솔의 빨간 글씨)를 그대로 알려 주세요.

- "STA 모드에서 실행해야 합니다"가 나오면 `pwsh -STA -File .\WorkTool-WPF.ps1`로 실행합니다.
- 예전 데이터를 못 찾으면 설정 → 데이터 → 「위치 고르기」에서 예전 `data` 폴더를 고릅니다.

## 데이터

- 새 버전 데이터는 앱 폴더 밖, **문서\WorkTool**에 저장합니다. 앱 파일을 새로 덮어써도 데이터는 그대로입니다.
  - 위치를 바꾸려면 앱 파일 옆에 `datapath.txt`를 만들고 첫 줄에 경로를 적습니다.
- 파일: `memo.json` · `schedule.json` · `attendance.json` · `settings.json`. 저장할 때마다 직전 파일을 `.bak`으로 하나 남깁니다.
- 각 파일에 `schema` 번호가 있습니다. 형식이 바뀌면 읽을 때 자동으로 변환하고, 변환 전 원본은 `.schema<번호>.bak`으로 남깁니다.
- `settings.json`에는 기본값과 다른 설정만 저장합니다.
- 예전 WorkTool의 CSV는 읽기만 하고 바꾸지 않습니다.

## 구조

```
wpf/
  Start-WorkTool.ps1      진입점. core → addons 순서로 불러오고 창을 띄운다
  core/
    Storage.psm1          저장 (데이터 폴더, JSON 저장소, schema 변환, 예전 CSV 읽기)
    Settings.psm1         설정 (기본값 + 바꾼 값만 저장, 바뀌면 처리기 호출)
    Theme.psm1            색 · 글자 크기 · 글꼴 (팔레트 7개, 대비 계산)
    Ui.psm1               공통 부품 (스타일, 버튼 · 카드 · 태그 · 숫자 칸 · 목록 줄 …)
    Addons.psm1           탭 목록 (Register-WtAddon)
    Pages.psm1            설정 · 도움말 화면
    Shell.psm1            창 틀 · 메뉴 · 트레이 · 단축키 · 화면 전환
  addons/
    home/Home.psm1        홈: 다른 탭의 홈 카드를 모아 보여 줌
    schedule/Schedule.psm1
    attendance/Attendance.psm1
    memo/Memo.psm1
  build/Build-Bundle.ps1  파일 하나로 묶기 → ../dist/WorkTool-WPF.ps1
  tests/Test-Core.ps1     창 없이 확인할 수 있는 부분 검사
```

## 탭(애드온) 더하기 · 빼기

- **더하기**: `addons/<이름>/<이름>.psm1`을 만들고 맨 아래에서 `Register-WtAddon`을 부릅니다. 메뉴 · 홈 카드 · 예전 데이터 가져오기가 저절로 연결됩니다.

  ```powershell
  Register-WtAddon -Id 'files' -Title '문서' -Icon 0xE8B7 -Order 50 `
      -NewView { New-WtFilesView } `
      -Refresh { param($view) Update-WtFilesView } `
      -HomeCard { New-WtFilesHomeCard }
  ```
- **빼기**: 그 폴더를 지웁니다. 다른 탭과 데이터는 그대로입니다. 잠깐 숨기기만 하려면 설정 → 메뉴에서 끕니다.
- 화면은 `core/Ui.psm1`의 `New-Wt*` 함수와 스타일 이름만 써서 만듭니다. 색 값이나 글꼴 이름을 직접 쓰지 않습니다.
- 저장은 `Get-WtStoreItems` · `Write-WtStore`만 씁니다. 형식을 바꿀 때는 `-Schema`를 올리고 `-Migrate`로 변환을 줍니다.
- 이벤트 처리기 안에서는 `$script:` 변수와 보낸 요소(`$s`)의 `Tag`만 씁니다. 처리기는 나중에 불리므로, 만든 함수의 지역 변수는 보이지 않습니다.

## 고친 뒤 할 일

```powershell
pwsh ./apps/WorkTool/wpf/tests/Test-Core.ps1        # 모두 통과해야 함
pwsh ./apps/WorkTool/wpf/build/Build-Bundle.ps1     # dist/WorkTool-WPF.ps1 다시 만들기 (같은 커밋에 넣기)
```

## 이 창에서 쓰지 않는 것

- 창 투명도를 위해 테두리 없는 투명 창(AllowsTransparency)으로 만들었습니다. 그래서 이 창 안에는 WinForms 컨트롤(WindowsFormsHost)이나 WebBrowser를 넣지 않습니다(그려지지 않음). 마크다운 미리보기 같은 것은 WPF 기본 컨트롤로 만듭니다.
- 윈도우 전체 단축키, 자동 실행 등록 등은 `AGENTS.md` 「코드 규칙」 4번에 따라 쓰지 않습니다.
