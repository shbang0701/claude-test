@echo off
setlocal DisableDelayedExpansion
chcp 65001 >nul
rem 저장소를 받아 둔 개인 PC용 실행 파일. 묶음 파일 · zip 없이 이 폴더의 원본(Start-WorkTool.ps1)을 바로 띄운다.
rem 실제 GUI는 콘솔 없는 별도 프로세스에서 실행한다. 데이터는 zip으로 띄울 때와 같은 곳(문서\WorkTool)을 쓴다.
where pwsh >nul 2>nul
if errorlevel 1 (
    echo PowerShell 7 - pwsh - 이 필요합니다.
    pause
    exit /b 1
)
pwsh -NoLogo -NoProfile -File "%~dp0build\start-launcher.ps1" -ScriptPath "%~dp0Start-WorkTool.ps1"
if errorlevel 1 (
    pause
    exit /b 1
)
exit /b 0
