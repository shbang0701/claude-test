# 명령어 모음

코드 블록 오른쪽 위 복사 버튼으로 복사해서 PowerShell 창에 붙여넣는다. 붙여넣기는 창에서 우클릭.

## 1. 설치: 폴더 만들기 (처음 한 번)

```powershell
New-Item -ItemType Directory -Force -Path "$HOME\WorkTool\modules", "$HOME\WorkTool\data" | Out-Null
```

## 2. 설치: 파일 붙여넣기

GitHub에서 파일을 열고 "Copy raw file" 버튼으로 내용을 복사한 뒤, 메모장으로 같은 이름의 파일을 만든다.

```powershell
notepad "$HOME\WorkTool\Start-WorkTool.ps1"
```

1. "새 파일을 만드시겠습니까?" → 예
2. 붙여넣기
3. **다른 이름으로 저장**(Ctrl+Shift+S) → 파일 이름 그대로, 인코딩을 **UTF-8(BOM)** 으로 고르고 저장 → 덮어쓰기 예

인코딩 확인 (`239 187 191`이 나오면 BOM 포함):

```powershell
[IO.File]::ReadAllBytes("$HOME\WorkTool\Start-WorkTool.ps1")[0..2] -join ' '
```

## 3. 실행

```powershell
& "$HOME\WorkTool\Start-WorkTool.ps1"
```

창을 닫으면 콘솔로 돌아온다.

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
