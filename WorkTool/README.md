# WorkTool

Windows에서 혼자 쓰는 업무 통합 GUI 툴. PowerShell 7 + WinForms, 외부 모듈 없음.

- 설치·실행·테스트 명령: [COMMANDS.md](COMMANDS.md)
- 요구사항과 진행 순서: [ROADMAP.md](ROADMAP.md)
- 이 문서: 구조와 각 부분이 하는 일

## 구조

```
WorkTool\
  Start-WorkTool.ps1   진입점. 모듈을 불러와 왼쪽 메뉴에 붙이고 메인 창을 띄운다
  modules\
    Storage.psm1       저장 담당
    UiHelper.psm1      화면 공통 (색, 글꼴, 배율, 버튼)
    Memo.psm1          메모 화면
  data\                실제 데이터 (git 제외)
```

툴은 자기 위치(`$PSScriptRoot`) 기준으로 동작하므로 WorkTool 폴더를 어디에 두든 된다. 네트워크 드라이브보다는 로컬 디스크가 낫다(실행 정책에 걸릴 수 있다).

## 원칙

- PowerShell 7 전용(`#Requires -Version 7`). 5.1 호환은 필요해지면 그때 고친다.
- 외부 반입 없음: `Install-Module`, 외부 DLL 안 쓴다. 기본 cmdlet과 .NET만 쓴다.
- 현장 값(경로, 서버 이름 등)은 코드에 넣지 않는다. 나중에 로컬 설정 파일에서 읽는다.
- 다른 모듈은 파일을 직접 읽고 쓰지 않고 Storage만 거친다.

## Start-WorkTool.ps1

- 메인 창: 왼쪽 메뉴(이름, 날짜, 기능 버튼) / 가운데 화면 / 아래 상태 표시줄.
- 창 크기는 화면 작업 영역 비율(65% × 80%)로 잡는다. 최대화 가능, F11은 테두리 없는 전체 화면.
- 기능 추가 방법: `modules\<Module>.psm1`을 만들고 `New-<Module>View` 함수가 화면(Control)을 돌려주게 한 뒤, `$features` 목록에 한 줄 넣는다.
  ```powershell
  @{ Module = 'Memo'; Title = '메모'; Icon = 0xE70B }
  ```
  `Icon`은 윈도우 기본 아이콘 글꼴(Segoe Fluent Icons / Segoe MDL2 Assets)의 글자 코드. 예: 홈 `0xE80F`, 메모 `0xE70B`, 달력 `0xE787`, 폴더 `0xE8B7`, 메일 `0xE715`, 설정 `0xE713`.

## UiHelper.psm1

| 함수 | 하는 일 |
|---|---|
| `Get-UiTheme` | 색 모음 (Back, Card, Sidebar, Accent, Text, SubText, Border 등) |
| `Get-UiFont [크기] [-Bold]` | 테마 글꼴 |
| `Get-UiSize <픽셀>` | 화면 배율에 맞춘 크기 (100% 기준 픽셀을 넣는다) |
| `New-UiPadding <왼> [위] [오] [아래]` | 배율 맞춘 Padding. 하나만 넣으면 사방 같은 값 |
| `New-UiButton <글자> [-Primary]` | 일반 버튼 / 강조 버튼 |
| `New-UiNavButton`, `Set-UiNavSelected` | 왼쪽 메뉴 버튼 (메인 전용) |

- **테마**: 하늘색 파스텔. 색은 파일 맨 위 `$script:theme`에서 바꾼다.
- **글꼴**: 후보 목록(Pretendard, SUIT, Noto Sans KR, 나눔 계열, 맑은 고딕) 중 PC에 설치된 첫 번째를 쓴다. 지금 쓰는 글꼴은 상태 표시줄에 나온다. 순서를 바꾸려면 `$fontCandidates`를 고친다. 글꼴을 새로 설치하는 건 외부 반입이라 하지 않는다.
- **배율**: 메인이 `SetHighDpiMode`를 먼저 부르고 UiHelper를 불러와야 배율이 맞게 읽힌다.

## Storage.psm1

| 함수 | 하는 일 |
|---|---|
| `Get-Records <이름>` | `data\<이름>.csv`를 읽는다. 없으면 아무것도 안 돌려주므로 `@(Get-Records x)`로 감싸서 쓴다 |
| `Save-Records <이름> <기록들>` | 기록 전체를 저장한다 |
| `New-RecordId` | 새 기록용 짧은 ID (12자) |

- **안 깨지는 저장**: 임시 파일(`.tmp`)에 먼저 쓰고 `File.Replace`로 바꿔치기한다. 저장 중에 꺼져도 원래 파일은 남는다. 직전 파일은 `.bak`으로 하나 남는다.
- **Excel 호환**: CSV를 UTF-8 BOM으로 저장해서 Excel로 열어도 한글이 안 깨진다. 단, Excel로 열어둔 상태에서는 파일이 잠겨 저장이 실패한다.
- **값은 전부 문자열**: CSV라서 읽으면 모두 문자열이다. 날짜·숫자는 쓰는 쪽에서 변환한다.
- 읽기·쓰기에 실패하면 항상 예외를 올린다. 읽기 실패를 빈 목록으로 넘기면 다음 저장에서 기존 데이터를 덮어쓰기 때문이다.
- 이름은 영문·숫자·`_`·`-`만 된다.
- 나중에 SQLite 등으로 바꿀 때는 이 파일만 고친다.

## Memo.psm1

- 화면: 왼쪽 32% 검색·목록(최근 수정 순), 오른쪽 68% 버튼 줄·제목·본문.
- 저장 파일: `data\memo.csv` (Id, Title, Body, Created, Updated)
- 자동 저장: 다른 메모를 고르거나, 새 메모를 누르거나, 창을 닫을 때 바뀐 내용이 있으면 저장한다.
- 단축키: Ctrl+S 저장, Ctrl+A 전체 선택(여러 줄 TextBox는 기본으로 안 돼서 직접 넣었다).
- 제목이 비어 있으면 저장 시각(`yyyy-MM-dd HH:mm:ss`)을 제목으로 쓴다.
- 파일을 못 읽으면 경고를 띄우고 그 실행 동안은 저장·삭제를 막는다(빈 목록으로 덮어쓰기 방지).
- 저장 실패(Excel이 파일을 잡고 있는 경우 등)는 경고 창으로 알리고, 편집 내용은 그대로 둔다.
- 구현 메모: WinForms 이벤트 핸들러는 함수가 끝난 뒤에 실행되므로, 핸들러에서 쓰는 컨트롤과 상태는 모두 `$script:` 변수에 둔다.
