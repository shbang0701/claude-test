# 명령어 모음

코드 블록 오른쪽 위 복사 버튼으로 복사해서 PowerShell 창에 붙여넣는다. 붙여넣기는 창에서 우클릭.

## 1. 설치: 폴더 만들기 (처음 한 번)

```powershell
New-Item -ItemType Directory -Force -Path "$HOME\WorkTool\modules", "$HOME\WorkTool\data" | Out-Null
```

## 2. 설치: VS Code 설정 (처음 한 번)

ps1·psm1을 자동으로 UTF-8(BOM)으로 저장하게 한다.

1. Ctrl+Shift+P → `Preferences: Open User Settings (JSON)`
2. 맨 바깥 `{ }` 안에 아래를 추가 (앞 항목 끝에 쉼표가 있어야 한다)

```json
"[powershell]": {
    "files.encoding": "utf8bom"
}
```

## 3. 설치: 파일 붙여넣기

1. VS Code에서 파일 > 폴더 열기 → `C:\Users\본인\WorkTool`
2. 탐색기에서 새 파일 → GitHub와 같은 경로·이름 (예: `Start-WorkTool.ps1`, `modules\Storage.psm1`)
3. GitHub 파일 화면의 "Copy raw file" 버튼으로 복사 → 붙여넣기 → 저장
4. 오른쪽 아래 상태 표시줄에 **UTF-8 with BOM** 이 보이면 정상

인코딩 확인 (`239 187 191`이 나오면 BOM 포함):

```powershell
[IO.File]::ReadAllBytes("$HOME\WorkTool\Start-WorkTool.ps1")[0..2] -join ' '
```

## 4. 실행

VS Code 터미널이나 PowerShell 창에서 실행한다. F5(디버그 실행)는 쓰지 않는다.

```powershell
& "$HOME\WorkTool\Start-WorkTool.ps1"
```

창을 닫으면 터미널로 돌아온다.

## 참고: 환경 확인 (완료)

버전:

```powershell
$PSVersionTable.PSVersion.ToString()
```

실행 정책 (Undefined가 아닌 것만):

```powershell
Get-ExecutionPolicy -List | Where-Object ExecutionPolicy -ne 'Undefined'
```

WinForms 동작 ("test" 창이 뜨면 정상):

```powershell
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.MessageBox]::Show("test")
```

pwsh 7 스레드 (STA가 나와야 정상):

```powershell
[Threading.Thread]::CurrentThread.ApartmentState
```
