# 명령어 모음

PowerShell 7(`pwsh`) 기준이다. 코드 블록 오른쪽 위 복사 버튼으로 복사해서 VS Code 터미널에 붙여넣는다.

## 설치

### 폴더 만들기 (처음 한 번)

```powershell
New-Item -ItemType Directory -Force -Path "$HOME\WorkTool\modules", "$HOME\WorkTool\data" | Out-Null
```

### 파일 붙여넣기

1. VS Code에서 파일 > 폴더 열기 → `C:\Users\본인\WorkTool`
2. 탐색기에서 새 파일 → GitHub와 같은 경로·이름 (예: `Start-WorkTool.ps1`, `modules\Storage.psm1`)
3. GitHub 파일 화면의 "Copy raw file" 버튼으로 복사 → 붙여넣기 → 저장

## 실행

F5(디버그 실행)는 쓰지 않는다.

```powershell
& "$HOME\WorkTool\Start-WorkTool.ps1"
```

창을 닫으면 터미널로 돌아온다.

## 단계별 테스트

### Storage

`modules\Storage.psm1`을 붙여넣은 뒤 한 블록씩 실행한다.

불러오고 테스트 기록 1건 저장:

```powershell
Import-Module "$HOME\WorkTool\modules\Storage.psm1" -Force
Save-Records test @([pscustomobject]@{ Id = (New-RecordId); Title = '테스트'; Body = "첫 줄`n둘째 줄, 쉼표" })
Get-Records test | Format-List
```

한 건 더 붙여서 저장 → 파일 2개(`test.csv`, `test.csv.bak`)와 건수 2가 나오면 정상:

```powershell
$r = @(Get-Records test)
Save-Records test @($r + [pscustomobject]@{ Id = (New-RecordId); Title = '두번째'; Body = 'x' })
Get-ChildItem "$HOME\WorkTool\data" | Select-Object Name
@(Get-Records test).Count
```

테스트 파일 지우기:

```powershell
Remove-Item "$HOME\WorkTool\data\test.csv*"
```
