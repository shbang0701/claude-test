@echo off
setlocal DisableDelayedExpansion
chcp 65001 >nul
rem 개인 PC용 실행 파일. 실제 GUI는 콘솔 없는 별도 프로세스에서 실행한다.
where pwsh >nul 2>nul
if errorlevel 1 (
    echo PowerShell 7 - pwsh - 이 필요합니다.
    pause
    exit /b 1
)
pwsh -NoLogo -NoProfile -File "%~dp0Start-WorkTool-Launcher.ps1"
if errorlevel 1 (
    echo 받은 ZIP의 속성에서 차단 해제 후 다시 압축을 풀었는지 확인하세요.
    pause
    exit /b 1
)
exit /b 0
