# 확인용 명령어

코드 블록 오른쪽 위 복사 버튼으로 복사해서 PowerShell 창에 붙여넣는다.

## 1. Windows PowerShell 5.1 환경 확인

시작 메뉴에서 "Windows PowerShell"(파란 창)을 열고 실행한다.

버전 확인 (5.1이 나와야 맞는 창):

```powershell
$PSVersionTable.PSVersion.ToString()
```

실행 정책 (Undefined가 아닌 것만 출력):

```powershell
Get-ExecutionPolicy -List | Where-Object ExecutionPolicy -ne 'Undefined'
```

WinForms 동작 확인 ("test" 창이 뜨면 정상):

```powershell
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.MessageBox]::Show("test")
```

## 2. PowerShell 7 스레드 확인

pwsh 창에서 실행한다. STA가 나와야 클립보드·파일 대화상자가 문제없이 된다.

```powershell
[Threading.Thread]::CurrentThread.ApartmentState
```
