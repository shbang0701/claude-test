<#
.SYNOPSIS
    개인 PC 실행 파일이 콘솔 없이 GUI를 띄우고 보조 프로세스를 끝내는지 검사한다.
.DESCRIPTION
    임시 데이터 폴더와 전용 WorkTool 창만 사용한다. 사용자 데이터나 기존 창은 조회하지 않는다.
.EXAMPLE
    pwsh -NoProfile -STA -File apps/WorkTool/wpf/tests/Test-Launcher.ps1
#>
#Requires -Version 7.0
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
if (-not $IsWindows) { Write-Host 'SKIP 실행 파일 (Windows에서만)'; return }
$root = Split-Path $PSScriptRoot -Parent
. (Join-Path $root 'build/start-launcher.ps1')
$tmp = Join-Path ([IO.Path]::GetTempPath()) ('wt-launcher-' + [guid]::NewGuid().ToString('N'))
$fixture = Join-Path $tmp "한글 O'Brien & (검사)! %test%"
[void][IO.Directory]::CreateDirectory($fixture)
$bootstrap = $null; $gui = $null

function Assert-Launcher([string]$Name, [bool]$Ok) {
    if (-not $Ok) { throw "FAIL 실행 파일: $Name" }
    Write-Host "PASS 실행 파일: $Name"
}

try {
    Copy-Item -LiteralPath (Join-Path $root 'build/start-launcher.ps1') -Destination (Join-Path $fixture 'Start-WorkTool-Launcher.ps1')
    $cmd = [IO.File]::ReadAllText((Join-Path $root 'build/start-worktool.cmd')).Replace("`r`n", "`n").Replace("`n", "`r`n")
    [IO.File]::WriteAllText((Join-Path $fixture 'Start-WorkTool.cmd'), $cmd, [Text.UTF8Encoding]::new($false))
    Set-Content -LiteralPath (Join-Path $fixture 'datapath.txt') -Value (Join-Path $fixture 'data') -Encoding utf8
    $entry = Join-Path $fixture 'WorkTool-WPF.ps1'
    $sample = @'
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml, System.Windows.Forms, System.Drawing
$source = '__SOURCE__'
foreach ($name in 'Storage','Settings','Theme','Ui','Highlight','Markdown','Mermaid','Addons','Pages','Shell') {
    Import-Module (Join-Path $source "core/$name.psm1") -Force -DisableNameChecking
}
Import-Module (Join-Path $source 'addons/home/Home.psm1') -Force -DisableNameChecking
$script:sampleRoot = $PSScriptRoot
$script:sampleShell = Get-Module Shell
$script:sampleDeadline = [datetime]::UtcNow.AddSeconds(20)
$script:sampleReported = $false
$script:sampleTimer = [Windows.Threading.DispatcherTimer]::new()
$script:sampleTimer.Interval = [timespan]::FromMilliseconds(50)
$script:sampleTimer.Add_Tick({
    $window = Get-WtWindow
    if ($window -and $window.IsVisible -and -not $script:sampleReported) {
        [pscustomobject]@{
            Pid = $PID; Visible = $window.IsVisible; WindowState = [string]$window.WindowState
            IconWidth = $window.Icon.PixelWidth; HasOverlay = [bool]$window.TaskbarItemInfo.Overlay
            Description = $window.TaskbarItemInfo.Description
            Apartment = [string][Threading.Thread]::CurrentThread.GetApartmentState()
        } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $script:sampleRoot 'state.json') -Encoding utf8
        $script:sampleReported = $true
    }
    if ((Test-Path -LiteralPath (Join-Path $script:sampleRoot 'stop')) -or [datetime]::UtcNow -gt $script:sampleDeadline) {
        $script:sampleTimer.Stop()
        & $script:sampleShell { Exit-WtApp }
    }
})
$script:sampleTimer.Start()
try { Start-WtApp -AppRoot $PSScriptRoot } finally { $script:sampleTimer.Stop() }
'@
    $sample = $sample.Replace('__SOURCE__', $root.Replace("'", "''"))
    [IO.File]::WriteAllText($entry, $sample, [Text.UTF8Encoding]::new($false))
    $info = New-WtLaunchStartInfo $entry
    Assert-Launcher '콘솔 생성과 셸 실행을 끈다' ($info.CreateNoWindow -and -not $info.UseShellExecute)
    Assert-Launcher '특수문자 경로를 하나의 파일 인수로 전달한다' ($info.ArgumentList.Count -eq 5 -and $info.ArgumentList[4] -ceq $entry -and $info.WorkingDirectory -ceq $fixture)
    Assert-Launcher '실행 정책 우회 옵션을 넣지 않는다' (-not ($info.ArgumentList -match 'ExecutionPolicy|Bypass|EncodedCommand'))

    $cmdInfo = [Diagnostics.ProcessStartInfo]::new()
    $cmdInfo.FileName = Join-Path ([Environment]::SystemDirectory) 'cmd.exe'
    $cmdInfo.Arguments = '/d /c Start-WorkTool.cmd'
    $cmdInfo.WorkingDirectory = $fixture
    $cmdInfo.UseShellExecute = $false; $cmdInfo.CreateNoWindow = $true
    $cmdInfo.RedirectStandardOutput = $true; $cmdInfo.RedirectStandardError = $true
    $bootstrap = [Diagnostics.Process]::Start($cmdInfo)
    $outTask = $bootstrap.StandardOutput.ReadToEndAsync(); $errTask = $bootstrap.StandardError.ReadToEndAsync()
    if (-not $bootstrap.WaitForExit(15000)) { throw '실행 보조 프로세스가 종료되지 않았습니다.' }
    if ($bootstrap.ExitCode -ne 0) {
        $stdout = $outTask.GetAwaiter().GetResult(); $stderr = $errTask.GetAwaiter().GetResult()
        throw "실행 보조 실패: $stdout $stderr"
    }
    $statePath = Join-Path $fixture 'state.json'
    $watch = [Diagnostics.Stopwatch]::StartNew()
    while (-not (Test-Path -LiteralPath $statePath)) {
        if ($watch.Elapsed.TotalSeconds -gt 15) { throw '전용 GUI 창이 시작되지 않았습니다.' }
        Start-Sleep -Milliseconds 50
    }
    # 파일 작성이 끝난 뒤 읽는다.
    $state = $null
    for ($i=0; $i -lt 20 -and -not $state; $i++) {
        try { $state = Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json } catch { Start-Sleep -Milliseconds 25 }
    }
    if (-not $state) { throw '전용 GUI 검사 결과를 읽지 못했습니다.' }
    $gui = [Diagnostics.Process]::GetProcessById($state.Pid)
    Assert-Launcher 'CMD는 끝나고 별도 GUI 프로세스만 계속 실행한다' ($bootstrap.HasExited -and -not $gui.HasExited -and $state.Visible)
    # 상속된 표준 입력 핸들이 있어도 콘솔 창이 있다는 뜻은 아니다. 이 검사에서 만든 PID의 창만 확인한다.
    Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
    $ownedCondition = [Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::ProcessIdProperty, [int]$state.Pid)
    $ownedWindows = [Windows.Automation.AutomationElement]::RootElement.FindAll([Windows.Automation.TreeScope]::Children, $ownedCondition)
    Assert-Launcher '이 프로세스에는 WorkTool GUI 창 하나만 있다' ($ownedWindows.Count -eq 1 -and $ownedWindows[0].Current.Name -eq 'WorkTool')
    Assert-Launcher 'GUI는 STA·보통 창 상태로 실행한다' ($state.Apartment -eq 'STA' -and $state.WindowState -eq 'Normal')
    Assert-Launcher '전체 창 아이콘을 적용하고 작은 배지를 제거한다' ($state.IconWidth -eq 32 -and -not $state.HasOverlay -and $state.Description -eq 'WorkTool')
    Set-Content -LiteralPath (Join-Path $fixture 'stop') -Value 'stop' -Encoding utf8
    Assert-Launcher 'GUI 종료 뒤 프로세스도 끝난다' ($gui.WaitForExit(10000))
    [void]$outTask.GetAwaiter().GetResult(); [void]$errTask.GetAwaiter().GetResult()

    $bad = Join-Path $fixture 'bad.ps1'
    Set-Content -LiteralPath $bad -Value 'function {' -Encoding utf8
    $rejected = $false
    try { [void](New-WtLaunchStartInfo $bad) } catch { $rejected = $true }
    Assert-Launcher '문법이 깨진 파일은 숨겨진 프로세스로 보내지 않는다' $rejected
    $rejected = $false
    try { [void](New-WtLaunchStartInfo (Join-Path $fixture 'missing.ps1')) } catch { $rejected = $true }
    Assert-Launcher '없는 실행 파일은 시작 전에 알린다' $rejected
    Set-Content -LiteralPath $entry -Stream Zone.Identifier -Value "[ZoneTransfer]`r`nZoneId=3"
    $rejected = $false
    try { [void](New-WtLaunchStartInfo $entry) } catch { $rejected = $true }
    Assert-Launcher '다운로드 차단 표시는 우회하지 않고 알린다' $rejected
} finally {
    if (Test-Path -LiteralPath $fixture) { Set-Content -LiteralPath (Join-Path $fixture 'stop') -Value 'stop' -Encoding utf8 }
    foreach ($owned in @($gui, $bootstrap)) {
        if ($owned) {
            if (-not $owned.HasExited -and -not $owned.WaitForExit(5000)) { $owned.Kill(); $owned.WaitForExit() }
            $owned.Dispose()
        }
    }
    $safe = [IO.Path]::GetFullPath($tmp)
    if ((Split-Path $safe -Leaf) -notmatch '^wt-launcher-[a-f0-9]{32}$' -or -not $safe.StartsWith([IO.Path]::GetFullPath([IO.Path]::GetTempPath()), [StringComparison]::OrdinalIgnoreCase)) { throw '임시 경로 검증 실패' }
    Remove-Item -LiteralPath $safe -Recurse -Force
}
