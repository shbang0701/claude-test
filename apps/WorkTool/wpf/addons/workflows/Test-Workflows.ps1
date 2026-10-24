<#
.SYNOPSIS
    업무 탭을 임시 데이터와 가짜 스크립트로 검사한다: 스크립트 읽기 · 폼 칸 · 값 검사 · 인수 · 실행 · 취소 · 기록 · 화면.
    저장소 scripts/의 스크립트가 업무 탭 약속(docs/plans/workflows-design.md)을 지키는지도 본다.

.DESCRIPTION
    tests\Test-Core.ps1이 찾아서 돌린다. 혼자 돌려도 된다. 실제 데이터 폴더(문서\WorkTool)는 쓰지 않고,
    실행은 이 검사가 임시 폴더에 만든 가짜 스크립트만 한다. 콘솔 창은 띄우지 않는다(실행 정보만 검사).

.EXAMPLE
    pwsh -NoProfile -STA -File apps/WorkTool/wpf/addons/workflows/Test-Workflows.ps1
#>
#Requires -Version 7.0
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) 'tests/TestKit.ps1')
Import-WtTestCore
Import-Module (Join-Path $PSScriptRoot 'Workflows.psm1') -Force -DisableNameChecking
$mod = Get-Module Workflows
$test = New-WtTestData 'workflows'
$running = [System.Collections.Generic.List[object]]::new()

# 가짜 스크립트를 임시 스크립트 폴더에 쓴다
$scriptsRoot = Join-Path $test.Root '스크립트 폴더'
function Write-WtTestScript([string]$Relative, [string]$Text) {
    $path = Join-Path $scriptsRoot $Relative
    [void][System.IO.Directory]::CreateDirectory((Split-Path $path -Parent))
    [System.IO.File]::WriteAllText($path, $Text, [System.Text.UTF8Encoding]::new($false))
    $path
}

# 앱 안 실행을 끝날 때까지 돌린다 (UI 타이머 대신)
function Wait-WtTestRun($Run, [int]$Seconds = 30) {
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    while ($true) {
        [void](Read-WtFlowOutput $Run)
        if (Test-WtFlowRunDone $Run) { break }
        if ($watch.Elapsed.TotalSeconds -gt $Seconds) { throw '가짜 스크립트가 제때 끝나지 않았습니다.' }
        Start-Sleep -Milliseconds 50
    }
    [void](Read-WtFlowOutput $Run)
    Complete-WtFlowRun $Run
}

# 금지 명령 (AGENTS.md 「코드 규칙」 4번): 구문을 읽어서 찾는다
function Get-WtBannedUse([string]$Path) {
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$null)
    foreach ($cmd in $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $true)) {
        $name = $cmd.GetCommandName()
        if ($name -in 'Invoke-Expression', 'iex') { "$name ($($cmd.Extent.StartLineNumber)번째 줄)" }
        if ($name -eq 'Add-Type' -and ($cmd.CommandElements | Where-Object { $_ -is [System.Management.Automation.Language.CommandParameterAst] -and $_.ParameterName -in 'TypeDefinition', 'MemberDefinition' })) {
            "Add-Type 인라인 코드 ($($cmd.Extent.StartLineNumber)번째 줄)"
        }
    }
    foreach ($s in $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.StringConstantExpressionAst] -or $args[0] -is [System.Management.Automation.Language.CommandParameterAst] }, $true)) {
        $text = if ($s -is [System.Management.Automation.Language.CommandParameterAst]) { "-$($s.ParameterName)" } else { $s.Value }
        if ($text -match '(?i)-EncodedCommand|-ExecutionPolicy\b|DllImport') { "$($text.Trim()) ($($s.Extent.StartLineNumber)번째 줄)" }
    }
}

try {
    # ---- 가짜 스크립트 ----
    $sample = Write-WtTestScript 'net\Check-Sample.ps1' @'
<#
.SYNOPSIS
    가짜 점검: 받은 값을 그대로 보여 준다.
.DESCRIPTION
    업무 탭 검사용 스크립트.
.PARAMETER Target
    대상 이름 (도움말에 적은 설명)
.EXAMPLE
    .\Check-Sample.ps1 -Target SRV-01
.NOTES
    WorkTool: 콘솔
#>
#Requires -Version 7.0
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)][string]$Target,
    [ValidateSet('빠르게', '꼼꼼히')][string]$Mode = '빠르게',   # 확인 방식
    [ValidateRange(1, 10)][int]$Count = 3,                        # 몇 번 확인할지
    [string]$ListFile,                                            # 목록 파일
    [string]$OutFolder,                                           # 결과 폴더
    [string]$LogPath,                                             # 기록 경로
    [switch]$Loud,                                                # 크게 말하기
    [string[]]$Names,                                             # 이름 여러 개
    [double]$Ratio = 0.5,                                         # 비율
    $Untyped,                                                     # 형식 없는 값
    [ValidateRange(0, 255)][int]$ExitWith = 0                     # 끝낼 때 종료 코드
)
$ErrorActionPreference = 'Stop'
"Target=[$Target] Mode=[$Mode] Count=[$Count] Loud=[$Loud] Ratio=[$Ratio]"
"한글 출력 확인: 가나다"
if ($PSCmdlet.ShouldProcess($Target, '확인')) { '실제로 실행함' } else { '미리보기만 함' }
Write-Error '오류 줄 하나' -ErrorAction Continue
1..60 | ForEach-Object { "줄 $_" }
exit $ExitWith
'@
    [void](Write-WtTestScript 'Bad-Syntax.ps1' "<#`n.SYNOPSIS`n    망가진 스크립트`n#>`nfunction {")
    [void](Write-WtTestScript '_hidden.ps1' "'숨김'")
    [void](Write-WtTestScript 'deep\a\b\In-Depth.ps1' "<#`n.SYNOPSIS`n    세 단계 아래`n#>`n'ok'")
    [void](Write-WtTestScript 'deep\a\b\c\Too-Deep.ps1' "'너무 깊음'")
    [void](Write-WtTestScript 'NoHelp-Script.ps1' "'도움말 없음'")
    [void](Write-WtTestScript 'Needs-Array.ps1' "<#`n.SYNOPSIS`n    배열이 필수`n#>`nparam([Parameter(Mandatory)][string[]]`$Names)`n`$Names")
    $readInput = Write-WtTestScript 'Read-Input.ps1' "<#`n.SYNOPSIS`n    키 입력을 기다림`n#>`n`$ErrorActionPreference = 'Stop'`n`$x = Read-Host '입력'`n`"받음: `$x`""
    $sleeper = Write-WtTestScript 'Long-Sleep.ps1' "<#`n.SYNOPSIS`n    오래 걸림`n#>`n'시작'`nStart-Sleep -Seconds 30`n'끝'"

    # ---- 스크립트 찾기 ----
    $groups = @(Find-WtFlowScripts @($scriptsRoot, (Join-Path $test.Root '없는 폴더')))
    $names = @($groups[0].Scripts | ForEach-Object Name)
    Assert-WtTest '업무: 스크립트 폴더의 .ps1을 찾는다 (_로 시작 · 4단계 아래는 뺀다)' ((($names | Sort-Object) -join ',') -eq 'Bad-Syntax,Check-Sample,In-Depth,Long-Sleep,Needs-Array,NoHelp-Script,Read-Input') ($names -join ',')
    Assert-WtTest '업무: 하위 폴더 이름 · 없는 폴더 표시' (($groups[0].Scripts | Where-Object Name -eq 'Check-Sample').Folder -eq 'net' -and $groups[1].Missing)

    # ---- 스크립트 읽기 ----
    $info = Read-WtFlowScript $sample
    Assert-WtTest '업무: 도움말을 읽는다 (#Requires 줄은 빼고)' ($info.Synopsis -eq '가짜 점검: 받은 값을 그대로 보여 준다.' -and $info.Description -eq '업무 탭 검사용 스크립트.' -and $info.Examples.Count -eq 1 -and $info.Notes -notmatch 'Requires')
    Assert-WtTest '업무: .NOTES의 WorkTool: 콘솔 · SupportsShouldProcess를 알아본다' ($info.Console -and $info.WhatIf -and -not $info.Error)
    $kinds = ($info.Fields | ForEach-Object { "$($_.Name):$($_.Kind)" }) -join ' '
    Assert-WtTest '업무: param 종류별 폼 칸' ($kinds -eq 'Target:text Mode:choice Count:number ListFile:file OutFolder:folder LogPath:text Loud:switch Ratio:number Untyped:text ExitWith:number') $kinds
    $byName = @{}; foreach ($f in $info.Fields) { $byName[$f.Name] = $f }
    Assert-WtTest '업무: 필수 · 선택지 · 범위 · 기본값' ($byName.Target.Mandatory -and -not $byName.Mode.Mandatory -and ($byName.Mode.Choices -join ',') -eq '빠르게,꼼꼼히' -and $byName.Count.Integer -and $byName.Count.Min -eq 1 -and $byName.Count.Max -eq 10 -and $byName.Count.Default -eq 3 -and -not $byName.Ratio.Integer -and $byName.Ratio.Default -eq 0.5 -and $null -eq $byName.ListFile.Default)
    Assert-WtTest '업무: 칸 설명은 .PARAMETER, 없으면 param 줄 끝 주석' ($byName.Target.Help -eq '대상 이름 (도움말에 적은 설명)' -and $byName.Mode.Help -eq '확인 방식' -and $byName.Loud.Help -eq '크게 말하기')
    Assert-WtTest '업무: 폼에 넣을 수 없는 종류는 따로 알린다' (@($info.Skipped).Count -eq 1 -and $info.Skipped[0].Name -eq 'Names' -and $info.Skipped[0].TypeName -eq 'string[]')
    Assert-WtTest '업무: 구문 오류는 실행할 수 없다고 알린다' ((Read-WtFlowScript (Join-Path $scriptsRoot 'Bad-Syntax.ps1')).Error -like '구문 오류*')
    Assert-WtTest '업무: 폼에 넣을 수 없는 필수 값이 있으면 실행할 수 없다' ((Read-WtFlowScript (Join-Path $scriptsRoot 'Needs-Array.ps1')).Error -like '폼에서 넣을 수 없는 필수 값*')
    $noHelp = Read-WtFlowScript (Join-Path $scriptsRoot 'NoHelp-Script.ps1')
    Assert-WtTest '업무: 도움말 · param이 없어도 실행할 수 있다 (도움말 없음만 알림)' (-not $noHelp.Error -and $noHelp.Warning -and $noHelp.Fields.Count -eq 0)
    Assert-WtTest '업무: 없는 파일' ((Read-WtFlowScript (Join-Path $scriptsRoot '없음.ps1')).Error -eq '스크립트 파일이 없습니다.')

    # ---- 값 검사 · 인수 ----
    $problems = Test-WtFlowValues $info @{ Target = ' '; Count = 'abc'; Ratio = '1.5'; Mode = '아무거나' }
    Assert-WtTest '업무: 값 검사 (필수 · 정수 · 선택지)' ($problems.Target -eq '값을 넣으세요.' -and $problems.Count -like '정수를*' -and -not $problems.Contains('Ratio') -and $problems.Mode -like '고를 수 있는 값*')
    Assert-WtTest '업무: 값 검사 (범위 · 소수는 정수 칸에 안 됨)' ((Test-WtFlowValues $info @{ Target = 'a'; Count = '11' }).Count -like '1 ~ 10*' -and (Test-WtFlowValues $info @{ Target = 'a'; Count = '2.5' }).Count -like '정수를*' -and (Test-WtFlowValues $info @{ Target = 'a'; Count = '7' }).Count -eq 0)
    $argList = @(Get-WtFlowArgs $info @{ Target = ' SRV 01 '; Mode = ''; Count = '5'; Loud = $true; Ratio = ''; ListFile = 'C:\목록 폴더\a.txt' } -WhatIf)
    Assert-WtTest '업무: 인수는 -이름:값 하나씩, 비운 칸 · 끈 스위치는 넘기지 않는다' (($argList -join '|') -eq '-Target:SRV 01|-Count:5|-ListFile:C:\목록 폴더\a.txt|-Loud|-WhatIf') ($argList -join '|')
    Assert-WtTest '업무: -로 시작하는 값도 값으로 넘긴다' ((@(Get-WtFlowArgs $info @{ Target = '-x' }) -join '|') -eq '-Target:-x')
    $start = New-WtFlowStartInfo $info @{ Target = 'SRV 01' } 'app'
    $startArgs = @($start.ArgumentList)
    Assert-WtTest '업무: 앱 안 실행은 창 없이 · 입력을 닫고 · 색 코드 없이' ($start.CreateNoWindow -and -not $start.UseShellExecute -and $start.RedirectStandardInput -and $start.RedirectStandardOutput -and $start.Environment['NO_COLOR'] -eq '1' -and $startArgs -contains '-NonInteractive' -and $startArgs[-1] -eq '-Target:SRV 01' -and $start.WorkingDirectory -eq (Split-Path $sample -Parent))
    Assert-WtTest '업무: 실행 정책 우회 · 명령 문자열 조립을 쓰지 않는다' (-not ($startArgs -match '(?i)ExecutionPolicy|EncodedCommand|^-Command$|^-c$'))
    $console = New-WtFlowStartInfo $info @{ Target = 'SRV 01' } 'console'
    Assert-WtTest '업무: 콘솔 창 실행은 보이는 창 · -NoExit' ($console.UseShellExecute -and $console.Arguments -like '*-NoExit*' -and $console.Arguments -notlike '*-NonInteractive*' -and $console.Arguments -like '*"-Target:SRV 01"*')
    # 콘솔 창 인수 감싸기: 같은 규칙으로 읽는 pwsh에 넘겨 그대로 돌아오는지 본다
    $echo = Write-WtTestScript 'echo\Echo-Args.ps1' '$args | ForEach-Object { "[" + $_ + "]" }'
    $tricky = @('a b', 'c"d', 'e\', 'f\"g', '가 나', 'C:\폴더 이름\')
    $probe = [System.Diagnostics.ProcessStartInfo]::new((Join-Path $PSHOME 'pwsh.exe'))
    $probe.Arguments = (@('-NoLogo', '-NoProfile', '-NonInteractive', '-File', $echo) + $tricky | ForEach-Object { ConvertTo-WtFlowCommandArg $_ }) -join ' '
    $probe.UseShellExecute = $false; $probe.CreateNoWindow = $true; $probe.RedirectStandardOutput = $true
    $probe.StandardOutputEncoding = [System.Text.Encoding]::GetEncoding([System.Globalization.CultureInfo]::CurrentCulture.TextInfo.OEMCodePage)
    $p = [System.Diagnostics.Process]::Start($probe)
    $echoed = $p.StandardOutput.ReadToEnd().Trim() -split "`r?`n"
    $p.WaitForExit(); $p.Dispose()
    Assert-WtTest '업무: 콘솔 창 인수는 띄어쓰기 · 따옴표 · \ · 한글을 그대로 넘긴다' (($echoed -join '|') -eq (($tricky | ForEach-Object { "[$_]" }) -join '|')) ($echoed -join '|')

    # ---- 실행 (앱 안) ----
    $record = New-WtFlowRecord $info @{ Target = 'SRV 01'; Loud = $true; Count = '' } 'app' $false '가짜 업무'
    $run = Start-WtFlowProcess (New-WtFlowStartInfo $info @{ Target = 'SRV 01'; Loud = $true } 'app') $record
    $running.Add($run)
    $status = Wait-WtTestRun $run
    $lines = @($run.Lines)
    Assert-WtTest '업무: 앱 안 실행 · 한글 출력 · 오류 출력을 받는다' ($status -eq 'success' -and $lines -contains 'Target=[SRV 01] Mode=[빠르게] Count=[3] Loud=[True] Ratio=[0.5]' -and $lines -contains '한글 출력 확인: 가나다' -and $lines -contains '실제로 실행함' -and @($lines | Where-Object { $_ -like '*오류 줄 하나*' }).Count -eq 1) ($lines -join ' / ')
    Assert-WtTest '업무: 출력에 색 코드가 섞이지 않는다' (-not @($lines | Where-Object { $_ -match "`e" }).Count)
    $saved = @(Get-WtFlowRuns)[0]
    Assert-WtTest '업무: 실행 기록 (결과 · 넘긴 값 · 출력 마지막 40줄)' ($saved.status -eq 'success' -and $saved.exitCode -eq 0 -and $saved.flow -eq '가짜 업무' -and $saved.values.Target -eq 'SRV 01' -and $saved.values.Loud -eq $true -and -not ($saved.values.PSObject.Properties.Name -contains 'Count') -and @($saved.tail).Count -eq 40 -and @($saved.tail)[-1] -eq '줄 60')
    $run = Start-WtFlowProcess (New-WtFlowStartInfo $info @{ Target = 'A' } 'app' -WhatIf) (New-WtFlowRecord $info @{ Target = 'A' } 'app' $true '')
    $running.Add($run)
    $status = Wait-WtTestRun $run
    Assert-WtTest '업무: 미리보기는 -WhatIf로 할 일만 보인다' ($status -eq 'success' -and @($run.Lines) -contains '미리보기만 함' -and @(Get-WtFlowRuns)[0].whatIf)
    $run = Start-WtFlowProcess (New-WtFlowStartInfo $info @{ Target = 'A'; ExitWith = '3' } 'app') (New-WtFlowRecord $info @{ Target = 'A' } 'app' $false '')
    $running.Add($run)
    Assert-WtTest '업무: 종료 코드가 0이 아니면 실패' ((Wait-WtTestRun $run) -eq 'failed' -and @(Get-WtFlowRuns)[0].exitCode -eq 3)
    $inputInfo = Read-WtFlowScript $readInput
    $run = Start-WtFlowProcess (New-WtFlowStartInfo $inputInfo @{} 'app') (New-WtFlowRecord $inputInfo @{} 'app' $false '')
    $running.Add($run)
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    $status = Wait-WtTestRun $run 20
    Assert-WtTest '업무: 키 입력을 기다리는 스크립트는 멈추지 않고 실패로 끝난다' ($status -eq 'failed' -and $watch.Elapsed.TotalSeconds -lt 15 -and @($run.Lines | Where-Object { $_ -like '*NonInteractive*' }).Count -ge 1 -and @($run.Lines) -notcontains '받음: ') ("$status · $($run.Lines -join ' / ')")
    $sleepInfo = Read-WtFlowScript $sleeper
    $run = Start-WtFlowProcess (New-WtFlowStartInfo $sleepInfo @{} 'app') (New-WtFlowRecord $sleepInfo @{} 'app' $false '')
    $running.Add($run)
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    while (-not @($run.Lines).Count -and $watch.Elapsed.TotalSeconds -lt 15) { [void](Read-WtFlowOutput $run); Start-Sleep -Milliseconds 50 }
    $run.Cancelled = $true
    Stop-WtFlowProcess $run
    $status = Wait-WtTestRun $run 10
    Assert-WtTest '업무: [취소]는 프로세스를 끝내고 취소로 기록한다' ($status -eq 'cancelled' -and $run.Process -and @($run.Lines) -contains '시작' -and @($run.Lines) -notcontains '끝')

    # ---- 기록 · 업무 저장 ----
    $many = foreach ($i in 1..105) {
        [ordered]@{ id = "r$i"; script = 'x'; name = 'x'; flow = ''; mode = 'app'; whatIf = $false; values = @{}; started = (Get-Date '2026-01-01').AddMinutes($i).ToString('yyyy-MM-dd HH:mm:ss'); ended = ''; status = 'success'; exitCode = 0; tail = @() }
    }
    Write-WtStore -Name 'workflow-runs' -Schema 1 -Items @($many)
    Save-WtFlowRun ([ordered]@{ id = 'r200'; script = 'x'; name = 'x'; flow = ''; mode = 'app'; whatIf = $false; values = @{}; started = '2026-02-01 00:00:00'; ended = ''; status = 'success'; exitCode = 0; tail = @() })
    $kept = @(Get-WtFlowRuns)
    Assert-WtTest '업무: 실행 기록은 최근 100건만 둔다' ($kept.Count -eq 100 -and $kept[0].id -eq 'r200' -and $kept[-1].id -eq 'r7') "($($kept.Count) · $($kept[0].id) … $($kept[-1].id))"
    $flow = New-WtFlow 'PM 점검' $sample @{ Target = 'SRV-01'; Loud = $true; Gone = 'x' } 'console'
    Save-WtFlows @($flow)
    $back = @(Get-WtFlows)[0]
    $values = ConvertTo-WtFlowValues $back.values
    Assert-WtTest '업무: 저장한 업무를 다시 읽는다' ($back.title -eq 'PM 점검' -and $back.mode -eq 'console' -and $values.Target -eq 'SRV-01' -and $values.Loud -eq $true)
    Assert-WtTest '업무: 스크립트에 없는 값을 알아본다' ((@(Get-WtFlowStaleNames $info $values) -join ',') -eq 'Gone')
    Assert-WtTest '업무: 결과 표시는 글자 · 아이콘 · 색을 함께' ((Get-WtFlowStatusLook 'failed' $false).Text -eq '실패' -and (Get-WtFlowStatusLook 'failed' $false).Icon -and (Get-WtFlowStatusLook 'success' $true).Text -eq '미리보기 · 성공' -and (Format-WtFlowSpan ([timespan]::FromSeconds(65))) -eq '1분 5초')

    # ---- 저장소 scripts/가 약속을 지키는지 ----
    $repoScripts = Join-Path (Split-Path (Split-Path (Split-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) -Parent) -Parent) -Parent) 'scripts'
    if (Test-Path -LiteralPath $repoScripts -PathType Container) {
        $files = @(Get-ChildItem -LiteralPath $repoScripts -Filter '*.ps1' -File -Recurse | Where-Object Extension -eq '.ps1')
        $bad = @(foreach ($f in $files) {
                $x = Read-WtFlowScript $f.FullName
                if ($x.Error) { "$($f.Name): $($x.Error)" }
                if (-not $x.Synopsis) { "$($f.Name): 도움말(.SYNOPSIS)이 없습니다" }
                foreach ($use in @(Get-WtBannedUse $f.FullName)) { "$($f.Name): 금지 명령 $use" }
            })
        Assert-WtTest "업무: 저장소 scripts/의 스크립트 $($files.Count)개가 업무 탭 약속을 지킨다" (-not $bad.Count) ($bad -join '; ')
        $ping = Join-Path $repoScripts 'network/Watch-MultiPing.ps1'
        if (Test-Path -LiteralPath $ping) {
            $pingInfo = Read-WtFlowScript $ping
            Assert-WtTest '업무: Watch-MultiPing은 콘솔 창으로 띄우고, 목록 파일 칸에 [찾아보기]가 붙는다' ($pingInfo.Console -and ($pingInfo.Fields | Where-Object Name -eq 'HostFile').Kind -eq 'file')
        }
    } else { Write-Host 'SKIP 저장소 scripts/ 검사 (scripts 폴더 없음)' }

    if (-not (Test-WtTestWpf)) { Write-Host 'SKIP 업무 화면 (Windows STA에서만)'; return }

    # ---- 화면 ----
    Save-WtFlows @()
    Write-WtStore -Name 'workflow-runs' -Schema 1 -Items @()
    Set-WtSetting 'workflows.roots' @($scriptsRoot, (Join-Path $test.Root '없는 폴더'))
    $addon = Get-WtAddon -Id 'workflows'
    $view = New-WtTestView 'workflows'
    [void](& $addon.Refresh $view)
    $top = & $mod { @($script:ui.Tree.Items | ForEach-Object { $_.Header.Children[1].Text }) }
    Assert-WtTest 'WPF 업무: 목록은 저장한 업무 · 스크립트 폴더 · 없는 폴더 표시' ($top.Count -eq 3 -and $top[0] -eq '저장한 업무' -and $top[1] -eq '스크립트 폴더' -and $top[2] -eq '없는 폴더 · 폴더 없음') ($top -join ' / ')
    & $mod { param($p) Select-WtFlowNode @{ Kind = 'script'; Path = $p } } $sample
    $state = & $mod { @{ Fields = $script:fields.Count; Mode = (Get-WtFlowMode); Run = $script:ui.RunButton.IsEnabled; Preview = $script:ui.Preview.IsEnabled; Cancel = $script:ui.Cancel.IsEnabled; Title = $script:ui.Title.Text } }
    Assert-WtTest 'WPF 업무: 스크립트를 고르면 칸 · 실행 방식(콘솔) · 명령이 맞춰진다' ($state.Fields -eq 10 -and $state.Mode -eq 'console' -and $state.Run -and $state.Preview -and -not $state.Cancel -and $state.Title -eq 'Check-Sample')
    & $mod { Invoke-WtFlowRun }
    $err = & $mod { ($script:fields | Where-Object { $_.Field.Name -eq 'Target' }).Error }
    Assert-WtTest 'WPF 업무: 필수 칸을 비우고 실행하면 칸 아래에 알리고 실행하지 않는다' ($err.Visibility -eq 'Visible' -and $err.Text -eq '값을 넣으세요.' -and -not (& $mod { $script:run }))
    $countState = & $mod {
        $x = $script:fields | Where-Object { $_.Field.Name -eq 'Count' }
        $x.Input.Text = '99'; Set-WtFlowFieldError $x
        $bad = @{ Text = $x.Error.Text; Help = [System.Windows.Automation.AutomationProperties]::GetHelpText($x.Input) }
        $x.Input.Text = '4'; Set-WtFlowFieldError $x
        $bad.After = $x.Error.Visibility
        $bad
    }
    Assert-WtTest 'WPF 업무: 숫자 칸은 칸을 떠날 때 범위를 알리고, 고치면 지운다 (화면 읽기용 도움말도)' ($countState.Text -eq '1 ~ 10 사이로 넣으세요.' -and $countState.Help -eq $countState.Text -and $countState.After -eq 'Collapsed')
    & $mod {
        ($script:fields | Where-Object { $_.Field.Name -eq 'Target' }).Input.Text = 'SRV 02'
        foreach ($rb in $script:ui.ModeButtons) { $rb.IsChecked = ($rb.Tag.Value -eq 'app') }
        Invoke-WtFlowRun
    }
    Assert-WtTest 'WPF 업무: 실행 중에는 [실행]을 막고 [취소]를 연다' ((& $mod { [bool]$script:run }) -and -not (& $mod { $script:ui.RunButton.IsEnabled }) -and (& $mod { $script:ui.Cancel.IsEnabled }))
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    while ((& $mod { [bool]$script:run }) -and $watch.Elapsed.TotalSeconds -lt 30) { & $mod { Update-WtFlowRunState }; Start-Sleep -Milliseconds 50 }
    $after = & $mod { @{ Out = $script:ui.Output.Text; Tag = ($script:ui.RunStatus.Child.Child.Children | ForEach-Object { if ($_ -is [System.Windows.Controls.TextBlock]) { $_.Text } }) -join ''; Runs = @($script:ui.HelpPanel.Children | Where-Object { $_ -is [System.Windows.Controls.ListBox] }) } }
    Assert-WtTest 'WPF 업무: 출력을 보여 주고 결과 · 최근 실행을 갱신한다' ($after.Out -like '*Target=`[SRV 02`]*' -and $after.Out -like '*■*끝 · 성공*' -and $after.Tag -like '*성공*' -and $after.Runs.Count -eq 1 -and $after.Runs[0].Items.Count -eq 1)
    # 업무로 저장: 이름 묻는 창만 바꾼다
    & $mod { function script:Read-WtFlowTitle { param($Suggest) '가짜 업무' } }
    & $mod { Save-WtFlowFromForm }
    $flowState = & $mod { @{ Flow = $script:flow; Title = $script:ui.Title.Text; Saved = @($script:ui.Tree.Items[0].Items | ForEach-Object { $_.Header.Children[1].Text }) } }
    Assert-WtTest 'WPF 업무: 업무로 저장하면 저장한 업무에 생기고 그 업무를 고른다' ($flowState.Flow.title -eq '가짜 업무' -and $flowState.Title -eq '가짜 업무' -and $flowState.Saved -contains '가짜 업무' -and $flowState.Flow.mode -eq 'app' -and $flowState.Flow.values.Target -eq 'SRV 02')
    & $mod { ($script:fields | Where-Object { $_.Field.Name -eq 'Target' }).Input.Text = '고치는 중' }
    [void](& $addon.Refresh $view)
    Assert-WtTest 'WPF 업무: 다른 탭에 다녀와도 고치던 칸 값을 지킨다' ((& $mod { ($script:fields | Where-Object { $_.Field.Name -eq 'Target' }).Input.Text }) -eq '고치는 중' -and (& $mod { $script:flow.title }) -eq '가짜 업무')
    & $mod { Remove-WtFlowCurrent }
    Assert-WtTest 'WPF 업무: 업무 삭제 뒤 실행 취소가 보인다' (@(Get-WtFlows).Count -eq 0 -and (& $mod { $script:ui.Undo.Visibility }) -eq 'Visible')
    & $mod { Restore-WtFlowDeleted }
    Assert-WtTest 'WPF 업무: 실행 취소로 업무를 되살린다' (@(Get-WtFlows).Count -eq 1 -and (& $mod { $script:flow.title }) -eq '가짜 업무')
    $lost = New-WtFlow '옮겨진 업무' (Join-Path $scriptsRoot '옮겨짐.ps1') @{} 'app'
    Save-WtFlows (@(Get-WtFlows) + [pscustomobject]$lost)
    & $mod { param($id) $script:pendingFlow = $id; Update-WtFlowsView } $lost.id
    $lostState = & $mod { @{ Sub = $script:ui.Sub.Text; Repick = $script:ui.Repick.Visibility; Run = $script:ui.RunButton.IsEnabled; Save = $script:ui.SaveFlow.IsEnabled } }
    Assert-WtTest 'WPF 업무: 스크립트가 없어진 업무는 [스크립트 다시 고르기]를 보이고 실행을 막는다' ($lostState.Sub -eq '스크립트 파일이 없습니다.' -and $lostState.Repick -eq 'Visible' -and -not $lostState.Run -and -not $lostState.Save)
    Assert-WtTest 'WPF 업무: 홈 카드에 저장한 업무가 보인다' (@(& $addon.HomeCard | Where-Object { $_ -is [System.Windows.UIElement] }).Count -eq 1)
    # 앱을 끄면 앱 안 실행은 멈추고 기록한다
    & $mod { param($p) Select-WtFlowNode @{ Kind = 'script'; Path = $p }; foreach ($rb in $script:ui.ModeButtons) { $rb.IsChecked = ($rb.Tag.Value -eq 'app') }; Invoke-WtFlowRun } $sleeper
    $proc = & $mod { $script:run.Process }
    $running.Add(@{ Process = $proc })
    & $addon.OnExit
    Assert-WtTest 'WPF 업무: 앱을 끄면 앱 안 실행을 멈추고 멈춤으로 기록한다' (-not (& $mod { $script:run }) -and @(Get-WtFlowRuns)[0].status -eq 'stopped')
} finally {
    foreach ($r in $running) { try { if ($r.Process -and -not $r.Process.HasExited) { $r.Process.Kill($true) } } catch { } }
    & $mod { if ($script:timer) { $script:timer.Stop() } }
    Remove-WtTestData $test
}
