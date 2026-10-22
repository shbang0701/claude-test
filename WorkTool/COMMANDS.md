# 명령어 모음

PowerShell 7(`pwsh`) 기준이다. 코드 블록 오른쪽 위 복사 버튼으로 복사해서 VS Code 터미널에 붙여넣는다.

**모든 명령은 WorkTool 폴더에서 실행한다.** VS Code로 WorkTool 폴더를 열면 터미널이 그 폴더에서 시작한다. 다른 곳이면 먼저 `cd`로 이동한다.

## 설치

### 폴더 만들기 (처음 한 번)

```powershell
New-Item -ItemType Directory -Force -Path modules, data | Out-Null
```

### 파일 붙여넣기

1. VS Code에서 파일 > 폴더 열기 → WorkTool 폴더
2. 탐색기에서 새 파일 → GitHub와 같은 경로·이름 (예: `Start-WorkTool.ps1`, `modules\Storage.psm1`)
3. GitHub 파일 화면의 "Copy raw file" 버튼으로 복사 → 붙여넣기 → 저장

## 실행

F5(디버그 실행)는 쓰지 않는다.

```powershell
.\Start-WorkTool.ps1
```

창을 닫으면 터미널로 돌아온다.

## 단계별 테스트

### Storage

`modules\Storage.psm1`을 붙여넣은 뒤 한 블록씩 실행한다.

불러오고 테스트 기록 1건 저장:

```powershell
Import-Module .\modules\Storage.psm1 -Force
Save-Records test @([pscustomobject]@{ Id = (New-RecordId); Title = '테스트'; Body = "첫 줄`n둘째 줄, 쉼표" })
Get-Records test | Format-List
```

한 건 더 붙여서 저장 → 파일 2개(`test.csv`, `test.csv.bak`)와 건수 2가 나오면 정상:

```powershell
$r = @(Get-Records test)
Save-Records test @($r + [pscustomobject]@{ Id = (New-RecordId); Title = '두번째'; Body = 'x' })
Get-ChildItem .\data | Select-Object Name
@(Get-Records test).Count
```

테스트 파일 지우기:

```powershell
Remove-Item .\data\test.csv*
```

### 메모

붙여넣을 파일: `modules\Memo.psm1`, `modules\Storage.psm1`, `Start-WorkTool.ps1`. 그다음 실행한다.

```powershell
.\Start-WorkTool.ps1
```

확인할 것:

1. 왼쪽 메뉴에 "메모"가 있다
2. 제목·본문 쓰고 Ctrl+S → 왼쪽 목록에 뜨고 "저장됨 시각"이 나온다
3. 새 메모 → 하나 더 쓰고 목록에서 첫 메모 클릭 → 두 번째 메모가 자동 저장된다
4. 제목을 비우고 저장 → 저장 시각이 제목이 된다
5. 검색칸에 글자 입력 → 목록이 걸러진다
6. 아무 메모나 고치고 저장 안 한 채 창 닫기 → 다시 열면 고친 내용이 남아 있다
7. 삭제 → 확인 창 → 목록에서 사라진다
8. `data\memo.csv`를 Excel로 열어둔 채 저장 → 경고 창이 뜬다 (Excel 닫고 다시 저장하면 된다)

### 왼쪽 메뉴·테마

붙여넣을 파일: `modules\UiHelper.psm1`(새 파일), `Start-WorkTool.ps1`(수정), `modules\Memo.psm1`(수정).

확인할 것:

1. 왼쪽에 하늘색 메뉴(WorkTool, 날짜, 홈, 메모)가 있고 아이콘이 보인다
2. 메뉴를 누르면 화면이 바뀌고, 고른 메뉴는 진하게 + 왼쪽에 파란 막대
3. 최대화 버튼이 되고, F11로 전체 화면 켜고 끄기가 된다
4. 아래 상태 표시줄에 지금 쓰는 글꼴 이름이 나온다
5. 메모 기능은 전과 똑같이 된다

PC에 설치된 한글 글꼴 중 쓸 만한 것 확인 (툴이 고를 수 있는 후보):

```powershell
Add-Type -AssemblyName System.Drawing
[System.Drawing.Text.InstalledFontCollection]::new().Families.Name | Where-Object { $_ -match 'Pretendard|SUIT|Noto|Nanum|나눔|Spoqa|맑은|Malgun|Gothic|고딕' }
```
