# 업무 탭: 스크립트 폴더의 PowerShell 스크립트를 폼으로 실행하고 기록한다. 설계는 docs/plans/workflows-design.md
# - 스크립트는 실행하지 않고 구문만 읽어서 목록 · 폼을 만든다 (Read-WtFlowScript). param 하나 = 폼 칸 하나.
# - 실행은 별도 pwsh 프로세스로 한다. 앱 안: 창 없이 출력을 받아 보여 줌 · 콘솔 창: 보이는 창에서 -NoExit.
#   출력은 UI 타이머에서 완료된 읽기만 꺼낸다 (다른 스레드에서 PowerShell 코드를 돌리지 않는다).
# - 저장소 workflows.json(저장한 업무) · workflow-runs.json(실행 기록, 최근 100건). 설정 workflows.roots(스크립트 폴더)
# 단축키(업무 화면 안): F5 스크립트 목록 새로 고침

$ErrorActionPreference = 'Stop'

$script:schema = 1
$script:maxRuns = 100        # 실행 기록은 모든 스크립트를 합쳐 이만큼만 둔다
$script:tailLines = 40       # 기록에 남기는 출력 줄 수
$script:maxLines = 2000      # 화면에 두는 출력 줄 수 (오래된 줄부터 버림)
$script:maxScripts = 1000    # 스크립트 폴더 하나에서 찾는 파일 수
$script:ui = @{}
$script:info = $null         # 고른 스크립트를 읽은 것
$script:flow = $null         # 고른 업무 (스크립트를 고른 것이면 $null)
$script:fields = @()         # 폼 칸: @{ Field; Input; Error }
$script:run = $null          # 앱 안에서 실행 중인 것
$script:consoles = [System.Collections.Generic.List[object]]::new()   # 띄운 콘솔 창: @{ Process; Record }
$script:timer = $null
$script:pendingFlow = $null  # 홈 카드에서 고른 업무 (화면을 열 때 고른다)
$script:lastDeleted = $null  # 지운 업무 (실행 취소용)
$script:filling = $false     # 코드로 트리를 고르는 중

Register-WtSettingDefault 'workflows.roots' @()   # 스크립트 폴더 목록
Register-WtSettingDefault 'workflows.help' $true  # 오른쪽 설명 패널을 펼침

# ---- 스크립트 읽기 (WPF 없이) -------------------------------------------------------------------

# 도움말 글 정리. 도움말 블록 바로 아래의 #Requires 줄은 마지막 절에 붙어 읽히므로 뺀다
function Format-WtFlowHelpText([string]$Text) {
    if (-not $Text) { return '' }
    (@($Text -split "`r?`n" | Where-Object { $_ -notmatch '^\s*Requires\s+-' }) -join "`n").Trim()
}

# 속성의 이름 있는 인수 값. [Parameter(Mandatory)]처럼 값을 생략하면 $true
function Get-WtFlowNamedArg($Attribute, [string]$Name) {
    foreach ($a in $Attribute.NamedArguments) {
        if ($a.ArgumentName -ne $Name) { continue }
        if ($a.ExpressionOmitted) { return $true }
        if ($a.Argument -is [System.Management.Automation.Language.ConstantExpressionAst]) { return [bool]$a.Argument.Value }
        if ($a.Argument -is [System.Management.Automation.Language.VariableExpressionAst]) { return $a.Argument.VariablePath.UserPath -ne 'false' }
        return $true
    }
    $false
}

# param 하나 → 폼 칸 정보. 폼에 넣을 수 없는 종류면 Kind가 비어 있다
function ConvertTo-WtFlowField($Param, $HelpInfo, $Tokens) {
    $name = $Param.Name.VariablePath.UserPath
    $field = [ordered]@{
        Name = $name; Kind = ''; TypeName = 'object'; Mandatory = $false; Choices = @()
        Min = $null; Max = $null; Integer = $false; Default = $null; Help = ''
    }
    foreach ($attr in $Param.Attributes) {
        if ($attr -is [System.Management.Automation.Language.TypeConstraintAst]) { $field.TypeName = $attr.TypeName.FullName; continue }
        switch ($attr.TypeName.Name) {
            'Parameter' { if (Get-WtFlowNamedArg $attr 'Mandatory') { $field.Mandatory = $true } }
            'ValidateSet' {
                $field.Choices = @($attr.PositionalArguments | Where-Object { $_ -is [System.Management.Automation.Language.ConstantExpressionAst] } |
                        ForEach-Object { [string]$_.Value })
            }
            'ValidateRange' {
                $range = @($attr.PositionalArguments | Where-Object { $_ -is [System.Management.Automation.Language.ConstantExpressionAst] })
                if ($range.Count -eq 2) { $field.Min = [double]$range[0].Value; $field.Max = [double]$range[1].Value }
            }
        }
    }
    $type = $Param.StaticType
    if ($type -eq [System.Management.Automation.SwitchParameter]) { $field.Kind = 'switch' }
    elseif ($type -eq [string] -or $type -eq [object]) {
        $field.Kind = if ($field.Choices.Count) { 'choice' } elseif ($name -match 'Folder$') { 'folder' } elseif ($name -match 'File$') { 'file' } else { 'text' }
    } elseif ($type -in [int], [long], [int16], [byte]) { $field.Kind = 'number'; $field.Integer = $true }
    elseif ($type -in [double], [single], [decimal]) { $field.Kind = 'number' }
    # 기본값은 글자 · 숫자 그대로일 때만 칸에 채운다
    if ($Param.DefaultValue -is [System.Management.Automation.Language.ConstantExpressionAst]) { $field.Default = $Param.DefaultValue.Value }
    # 칸 설명: .PARAMETER → 없으면 param 줄 끝 주석
    $text = $null
    if ($HelpInfo -and $HelpInfo.Parameters) {
        foreach ($k in $HelpInfo.Parameters.Keys) { if ($k -ieq $name) { $text = $HelpInfo.Parameters[$k] } }
    }
    if (-not $text) {
        $comment = $Tokens | Where-Object {
            $_.Kind -eq 'Comment' -and $_.Extent.StartLineNumber -eq $Param.Extent.EndLineNumber -and $_.Extent.StartOffset -ge $Param.Extent.EndOffset
        } | Select-Object -First 1
        if ($comment) { $text = $comment.Text.TrimStart('#') }
    }
    $field.Help = Format-WtFlowHelpText $text
    [pscustomobject]$field
}

# 스크립트 하나를 읽는다. 실행하지 않고 구문만 본다. 못 읽거나 실행할 수 없으면 Error에 이유
function Read-WtFlowScript([string]$Path) {
    $info = [pscustomobject]@{
        Path = $Path; Name = [System.IO.Path]::GetFileNameWithoutExtension($Path)
        Synopsis = ''; Description = ''; Examples = @(); Notes = ''
        Fields = @(); Skipped = @(); WhatIf = $false; Console = $false; Error = ''; Warning = ''
    }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { $info.Error = '스크립트 파일이 없습니다.'; return $info }
    $tokens = $null; $parseErrors = $null
    try { $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$parseErrors) }
    catch { $info.Error = "스크립트를 읽지 못했습니다: $($_.Exception.Message)"; return $info }
    if ($parseErrors.Count) {
        $info.Error = "구문 오류 ($($parseErrors[0].Extent.StartLineNumber)번째 줄): $($parseErrors[0].Message)"
        return $info
    }
    $helpInfo = $ast.GetHelpContent()
    if ($helpInfo) {
        $info.Synopsis = Format-WtFlowHelpText $helpInfo.Synopsis
        $info.Description = Format-WtFlowHelpText $helpInfo.Description
        $info.Examples = @($helpInfo.Examples | ForEach-Object { Format-WtFlowHelpText $_ } | Where-Object { $_ })
        $info.Notes = Format-WtFlowHelpText $helpInfo.Notes
    }
    if (-not $info.Synopsis) { $info.Warning = '도움말(.SYNOPSIS)이 없습니다.' }
    $info.Console = $info.Notes -match '(?m)^\s*WorkTool:\s*콘솔\s*$'
    if ($ast.ParamBlock) {
        foreach ($attr in $ast.ParamBlock.Attributes) {
            if ($attr.TypeName.Name -eq 'CmdletBinding' -and (Get-WtFlowNamedArg $attr 'SupportsShouldProcess')) { $info.WhatIf = $true }
        }
        $all = @(foreach ($p in $ast.ParamBlock.Parameters) { ConvertTo-WtFlowField $p $helpInfo $tokens })
        $info.Fields = @($all | Where-Object Kind)
        $info.Skipped = @($all | Where-Object { -not $_.Kind })
        $blocked = @($info.Skipped | Where-Object Mandatory)
        if ($blocked.Count) {
            $info.Error = "폼에서 넣을 수 없는 필수 값이 있어서 실행할 수 없습니다: $(($blocked | ForEach-Object { "$($_.Name)($($_.TypeName))" }) -join ', ')"
        }
    }
    $info
}

# 스크립트 폴더에서 .ps1을 찾는다 (하위 폴더 3단계까지, 이름이 _로 시작하면 뺀다). 없는 폴더도 표시하려고 돌려준다
function Find-WtFlowScripts([string[]]$Roots) {
    foreach ($root in $Roots) {
        if (-not (Test-Path -LiteralPath $root -PathType Container)) { [pscustomobject]@{ Root = $root; Missing = $true; Scripts = @() }; continue }
        $files = @(Get-ChildItem -LiteralPath $root -Filter '*.ps1' -File -Recurse -Depth 3 -ErrorAction SilentlyContinue |
                Where-Object { $_.Extension -eq '.ps1' -and $_.Name -notlike '_*' } | Sort-Object FullName | Select-Object -First $script:maxScripts)
        $found = foreach ($f in $files) {
            $rel = [System.IO.Path]::GetRelativePath($root, $f.DirectoryName)
            [pscustomobject]@{ Path = $f.FullName; Name = $f.BaseName; Folder = $(if ($rel -eq '.') { '' } else { $rel }) }
        }
        [pscustomobject]@{ Root = $root; Missing = $false; Scripts = @($found) }
    }
}

# ---- 값 검사 · 인수 (WPF 없이) ------------------------------------------------------------------

# 저장한 값(JSON에서 읽으면 PSCustomObject)을 해시 테이블로
function ConvertTo-WtFlowValues($Values) {
    $table = @{}
    if ($Values -is [System.Collections.IDictionary]) { foreach ($k in $Values.Keys) { $table[[string]$k] = $Values[$k] } }
    elseif ($Values) { foreach ($p in $Values.PSObject.Properties) { $table[$p.Name] = $p.Value } }
    $table
}

# 칸 값 검사. 틀린 칸만 @{ 이름 = '고칠 방법' }으로 돌려준다
function Test-WtFlowValues($Info, [hashtable]$Values) {
    $problems = [ordered]@{}
    foreach ($f in $Info.Fields) {
        if ($f.Kind -eq 'switch') { continue }
        $text = ([string]$Values[$f.Name]).Trim()
        if (-not $text) {
            if ($f.Mandatory) { $problems[$f.Name] = '값을 넣으세요.' }
            continue
        }
        if ($f.Kind -eq 'number') {
            $n = 0.0
            $ok = [double]::TryParse($text, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$n)
            if (-not $ok -or ($f.Integer -and $n -ne [Math]::Truncate($n))) {
                $problems[$f.Name] = $(if ($f.Integer) { '정수를 넣으세요. 예: 3' } else { '숫자를 넣으세요. 예: 1.5' })
                continue
            }
            if ($null -ne $f.Min -and ($n -lt $f.Min -or $n -gt $f.Max)) { $problems[$f.Name] = "$($f.Min) ~ $($f.Max) 사이로 넣으세요." }
        }
        if ($f.Kind -eq 'choice' -and $f.Choices -notcontains $text) { $problems[$f.Name] = "고를 수 있는 값: $($f.Choices -join ', ')" }
    }
    $problems
}

# 넘길 인수. 비운 칸은 넘기지 않고, 값은 -이름:값 하나로 넘긴다 (띄어쓰기 · 따옴표 · -로 시작하는 값도 그대로)
function Get-WtFlowArgs($Info, [hashtable]$Values, [switch]$WhatIf) {
    $list = [System.Collections.Generic.List[string]]::new()
    foreach ($f in $Info.Fields) {
        if ($f.Kind -eq 'switch') {
            if ($Values[$f.Name] -eq $true) { $list.Add("-$($f.Name)") }
            continue
        }
        $text = ([string]$Values[$f.Name]).Trim()
        if ($text) { $list.Add("-$($f.Name):$text") }
    }
    if ($WhatIf) { $list.Add('-WhatIf') }
    $list.ToArray()
}

# 저장한 업무의 값 중 지금 스크립트에 없는 칸 이름
function Get-WtFlowStaleNames($Info, [hashtable]$Values) {
    $names = @($Info.Fields | ForEach-Object Name)
    @($Values.Keys | Where-Object { $names -notcontains $_ } | Sort-Object)
}

# 명령줄 인수 하나를 Windows 규칙대로 감싼다. 콘솔 창 실행은 인수를 글자 한 줄로 넘겨야 해서 쓴다
function ConvertTo-WtFlowCommandArg([string]$Arg) {
    if ($Arg -and $Arg -notmatch '[\s"]') { return $Arg }
    $sb = [System.Text.StringBuilder]::new('"')
    $slashes = 0
    foreach ($c in $Arg.ToCharArray()) {
        if ($c -eq '\') { $slashes++; continue }
        if ($c -eq '"') { [void]$sb.Append('\', $slashes * 2 + 1).Append('"'); $slashes = 0; continue }
        if ($slashes) { [void]$sb.Append('\', $slashes); $slashes = 0 }
        [void]$sb.Append($c)
    }
    [void]$sb.Append('\', $slashes * 2).Append('"')
    $sb.ToString()
}

# 출력은 업무 PC 콘솔의 문자 코드(한국어 Windows는 CP949)로 나온다
function Get-WtFlowEncoding { [System.Text.Encoding]::GetEncoding([System.Globalization.CultureInfo]::CurrentCulture.TextInfo.OEMCodePage) }

# 실행할 프로세스 정보. 실행 정책은 바꾸지 않고, 입력값으로 명령 문자열을 조립하지 않는다
function New-WtFlowStartInfo($Info, [hashtable]$Values, [ValidateSet('app', 'console')][string]$Mode = 'app', [switch]$WhatIf) {
    $start = [System.Diagnostics.ProcessStartInfo]::new((Join-Path $PSHOME 'pwsh.exe'))
    $start.WorkingDirectory = Split-Path $Info.Path -Parent
    $scriptArgs = @(Get-WtFlowArgs $Info $Values -WhatIf:$WhatIf)
    if ($Mode -eq 'console') {
        # 보이는 새 창에서 연다. 끝나도 창을 남겨 출력을 읽게 하고, 창은 사용자가 닫는다
        $start.UseShellExecute = $true
        $start.Arguments = (@('-NoLogo', '-NoProfile', '-NoExit', '-File', $Info.Path) + $scriptArgs | ForEach-Object { ConvertTo-WtFlowCommandArg $_ }) -join ' '
        return $start
    }
    foreach ($a in @('-NoLogo', '-NoProfile', '-NonInteractive', '-File', $Info.Path) + $scriptArgs) { $start.ArgumentList.Add($a) }
    $start.UseShellExecute = $false
    $start.CreateNoWindow = $true
    $start.RedirectStandardInput = $true
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    $start.StandardOutputEncoding = Get-WtFlowEncoding
    $start.StandardErrorEncoding = Get-WtFlowEncoding
    $start.Environment['NO_COLOR'] = '1'   # 오류 글자에 색 코드가 섞이지 않게
    $start
}

# ---- 실행 상태 (WPF 없이) -----------------------------------------------------------------------

# 앱 안 실행을 시작한다. 입력은 바로 닫아서, 입력을 기다리는 스크립트도 멈추지 않고 오류로 끝나게 한다
function Start-WtFlowProcess($StartInfo, $Record) {
    $p = [System.Diagnostics.Process]::Start($StartInfo)
    $p.StandardInput.Close()
    [pscustomobject]@{
        Process = $p; Record = $Record; Started = Get-Date; ExitedAt = $null; Cancelled = $false; Stopped = $false
        Readers = @(
            @{ Reader = $p.StandardOutput; Task = $p.StandardOutput.ReadLineAsync(); Done = $false }
            @{ Reader = $p.StandardError; Task = $p.StandardError.ReadLineAsync(); Done = $false })
        Lines = [System.Collections.Generic.List[string]]::new()
    }
}

# 다 읽힌 줄을 꺼낸다. UI 타이머에서 부르고, 아직 안 끝난 읽기는 기다리지 않는다.
# 한 번에 $Max줄까지만 꺼낸다 (출력이 한꺼번에 쏟아져도 화면이 멈추지 않게. 남은 줄은 다음 차례에)
function Read-WtFlowOutput($Run, [int]$Max = 1000) {
    $fresh = [System.Collections.Generic.List[string]]::new()
    foreach ($r in $Run.Readers) {
        while (-not $r.Done -and $r.Task.IsCompleted -and $fresh.Count -lt $Max) {
            $line = if ($r.Task.IsFaulted -or $r.Task.IsCanceled) { $null } else { $r.Task.Result }
            if ($null -eq $line) { $r.Done = $true; break }
            $line = $line -replace "`e\[[0-9;?]*[ -/]*[@-~]", ''   # 남은 색 코드 정리
            $fresh.Add($line)
            $Run.Lines.Add($line)
            $r.Task = $r.Reader.ReadLineAsync()
        }
    }
    if ($Run.Lines.Count -gt $script:maxLines + 500) { $Run.Lines.RemoveRange(0, $Run.Lines.Count - $script:maxLines) }
    $fresh.ToArray()
}

# 프로세스가 끝나고 출력도 다 읽었으면 $true. 끝난 프로세스가 남긴 자식이 출력을 붙잡고 있어도 2초 뒤에는 끝으로 본다
function Test-WtFlowRunDone($Run) {
    if (-not $Run.Process.HasExited) { return $false }
    if (-not $Run.ExitedAt) { $Run.ExitedAt = Get-Date }
    -not @($Run.Readers | Where-Object { -not $_.Done }).Count -or ((Get-Date) - $Run.ExitedAt).TotalSeconds -gt 2
}

# 실행을 끝내고 기록을 저장한다. 결과(success · failed · cancelled · stopped)를 돌려준다.
# 기록을 저장하지 못해도 결과는 그대로 돌려주고, 그 이유는 실행 정보의 SaveError에 둔다
function Complete-WtFlowRun($Run) {
    $code = $null
    try { $code = $Run.Process.ExitCode } catch { }
    $status = if ($Run.Stopped) { 'stopped' } elseif ($Run.Cancelled) { 'cancelled' } elseif ($code -eq 0) { 'success' } else { 'failed' }
    $rec = $Run.Record
    $rec.ended = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $rec.status = $status
    $rec.exitCode = $code
    $rec.tail = @($Run.Lines | Select-Object -Last $script:tailLines)
    try { $Run.Process.Dispose() } catch { }
    try { Save-WtFlowRun $rec } catch { $Run | Add-Member -NotePropertyName SaveError -NotePropertyValue $_.Exception.Message -Force }
    $status
}

# 실행 중인 것을 끝낸다 (자식 프로세스까지). 이때 스크립트의 finally는 돌지 않을 수 있다
function Stop-WtFlowProcess($Run) {
    try { if (-not $Run.Process.HasExited) { $Run.Process.Kill($true) } } catch { }
}

# ---- 저장 (WPF 없이) ----------------------------------------------------------------------------

# 저장한 업무 · 실행 기록은 하나씩 내보낸다. Get-WtStoreItems는 목록을 한 덩어리로 돌려주므로 받아서 푼다
# (덩어리째 파이프에 넣으면 Where-Object · Sort-Object가 목록 전체를 항목 하나로 본다)
function Get-WtFlows {
    $items = Get-WtStoreItems -Name 'workflows' -Schema $script:schema
    $items | Sort-Object { [string]$_.title }
}

function Save-WtFlows([object[]]$Items) { Write-WtStore -Name 'workflows' -Schema $script:schema -Items @($Items) }

function New-WtFlow([string]$Title, [string]$ScriptPath, [hashtable]$Values, [string]$Mode) {
    $now = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    [ordered]@{ id = New-WtId; title = $Title; script = $ScriptPath; values = $Values; mode = $Mode; created = $now; updated = $now }
}

function Get-WtFlowRuns {
    $items = Get-WtStoreItems -Name 'workflow-runs' -Schema $script:schema
    $items
}

# 실행 기록 하나를 넣거나 바꾼다(같은 id면 바꿈). 시작 시각이 최근인 것부터 100건만 둔다
function Save-WtFlowRun($Record) {
    $rest = @(Get-WtFlowRuns | Where-Object { $_.id -ne $Record.id })
    $all = @(@($Record) + $rest | Sort-Object { [string]$_.started } -Descending | Select-Object -First $script:maxRuns)
    Write-WtStore -Name 'workflow-runs' -Schema $script:schema -Items $all
}

function New-WtFlowRecord($Info, [hashtable]$Values, [string]$Mode, [bool]$WhatIf, [string]$FlowTitle) {
    $kept = [ordered]@{}
    foreach ($f in $Info.Fields) {
        $v = $Values[$f.Name]
        if ($f.Kind -eq 'switch') { if ($v -eq $true) { $kept[$f.Name] = $true } }
        elseif (([string]$v).Trim()) { $kept[$f.Name] = ([string]$v).Trim() }
    }
    [ordered]@{
        id = New-WtId; script = $Info.Path; name = $Info.Name; flow = $FlowTitle; mode = $Mode; whatIf = $WhatIf; values = $kept
        started = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'); ended = ''; status = $(if ($Mode -eq 'console') { 'console' } else { 'running' })
        exitCode = $null; tail = @()
    }
}

# 결과 표시: 글자 · 아이콘 · 색(테마 리소스 이름). 색만으로 구분하지 않는다
function Get-WtFlowStatusLook([string]$Status, [bool]$WhatIf) {
    $look = switch ($Status) {
        'success' { @{ Text = '성공'; Icon = 0xE73E; Bg = 'SuccessBg'; Fg = 'Success' } }
        'failed' { @{ Text = '실패'; Icon = 0xE783; Bg = 'DangerBg'; Fg = 'Danger' } }
        'cancelled' { @{ Text = '취소'; Icon = 0xE711; Bg = 'Chip'; Fg = 'ChipText' } }
        'stopped' { @{ Text = '멈춤'; Icon = 0xE71A; Bg = 'WarnBg'; Fg = 'Warn' } }
        'console' { @{ Text = '콘솔 창'; Icon = 0xE756; Bg = 'Chip'; Fg = 'ChipText' } }
        default { @{ Text = '실행 중'; Icon = 0xE768; Bg = 'Chip'; Fg = 'ChipText' } }
    }
    if ($WhatIf) { $look.Text = "미리보기 · $($look.Text)" }
    $look
}

# 걸린 시간을 짧게: 12초 · 3분 5초 · 1시간 2분
function Format-WtFlowSpan([timespan]$Span) {
    if ($Span.TotalHours -ge 1) { return '{0}시간 {1}분' -f [int][Math]::Floor($Span.TotalHours), $Span.Minutes }
    if ($Span.TotalMinutes -ge 1) { return '{0}분 {1}초' -f [int][Math]::Floor($Span.TotalMinutes), $Span.Seconds }
    '{0}초' -f [int][Math]::Floor($Span.TotalSeconds)
}

function Get-WtFlowRunSpan($Record) {
    if (-not $Record.ended -or -not $Record.started) { return '' }
    Format-WtFlowSpan ([datetime]$Record.ended - [datetime]$Record.started)
}

# ---- 화면: 목록 ---------------------------------------------------------------------------------

function Get-WtFlowRoots { @(Get-WtSetting 'workflows.roots' | Where-Object { $_ }) }

function New-WtFlowTreeItem([string]$Label, [int]$Icon, $Tag, [string]$ToolTip, [switch]$Muted) {
    $item = [System.Windows.Controls.TreeViewItem]::new()
    $row = New-WtStack Horizontal 8
    $i = New-WtIcon $Icon
    Set-WtResource $i 'Foreground' 'B.SubText'
    $i.VerticalAlignment = 'Center'
    [void](Add-WtChild $row $i)
    $t = New-WtText $Label $(if ($Muted) { 'T.Sub' } else { 'T.Body' })
    $t.TextTrimming = 'CharacterEllipsis'
    $t.VerticalAlignment = 'Center'
    [void](Add-WtChild $row $t)
    $item.Header = $row
    $item.Tag = $Tag
    $item.IsExpanded = $true
    [System.Windows.Automation.AutomationProperties]::SetName($item, $Label)
    if ($ToolTip) { $item.ToolTip = $ToolTip }
    $item
}

# 왼쪽 목록: 저장한 업무 + 스크립트 폴더별 스크립트. 지금 고른 것을 다시 고른다
function Update-WtFlowTree {
    $tree = $script:ui.Tree
    $keep = if ($script:flow) { "flow:$($script:flow.id)" } elseif ($script:info) { "script:$($script:info.Path)" } else { '' }
    # 다른 탭에 다녀와도 고치던 칸 값과 실행 방식은 그대로 둔다
    $carry = if ($keep -and $script:fields.Count) { @{ Key = $keep; Values = (Get-WtFlowFormValues); Mode = (Get-WtFlowMode) } } else { $null }
    if ($script:pendingFlow) { $keep = "flow:$($script:pendingFlow)"; $script:pendingFlow = $null }
    $script:filling = $true
    try {
        $tree.Items.Clear()
        $found = $null
        $saved = New-WtFlowTreeItem '저장한 업무' 0xE734 @{ Kind = 'group' }
        $flows = @(Get-WtFlows)
        foreach ($f in $flows) {
            $missing = -not (Test-Path -LiteralPath ([string]$f.script) -PathType Leaf)
            $label = if ($missing) { "$($f.title) · 스크립트 없음" } else { [string]$f.title }
            $node = New-WtFlowTreeItem $label 0xE768 @{ Kind = 'flow'; Id = [string]$f.id } ([string]$f.script)
            [void]$saved.Items.Add($node)
            if ($keep -eq "flow:$($f.id)") { $found = $node }
        }
        if (-not $flows.Count) {
            $hint = New-WtFlowTreeItem '저장한 업무가 없습니다' 0xE946 @{ Kind = 'hint' } '스크립트를 고르고 [업무로 저장]하면 여기에 생깁니다' -Muted
            $hint.Focusable = $false
            [void]$saved.Items.Add($hint)
        }
        [void]$tree.Items.Add($saved)
        $groups = @(Find-WtFlowScripts (Get-WtFlowRoots))
        # 이름이 같은 스크립트 폴더가 둘 이상이면 위 폴더 이름을 붙여 구분한다 (예: mywork\scripts)
        $leafCount = @{}
        foreach ($g in $groups) { $leaf = Split-Path $g.Root -Leaf; $leafCount[$leaf] = 1 + [int]$leafCount[$leaf] }
        foreach ($group in $groups) {
            $rootName = Split-Path $group.Root -Leaf
            if (-not $rootName) { $rootName = $group.Root }
            elseif ($leafCount[$rootName] -gt 1) {
                $parentName = Split-Path (Split-Path $group.Root -Parent) -Leaf
                if ($parentName) { $rootName = "$parentName\$rootName" }
            }
            $label = if ($group.Missing) { "$rootName · 폴더 없음" } else { $rootName }
            $rootNode = New-WtFlowTreeItem $label 0xE8B7 @{ Kind = 'root'; Path = $group.Root } $group.Root
            $folders = @{}
            foreach ($s in $group.Scripts) {
                $parent = $rootNode
                if ($s.Folder) {
                    if (-not $folders.ContainsKey($s.Folder)) {
                        $folders[$s.Folder] = New-WtFlowTreeItem $s.Folder 0xE8B7 @{ Kind = 'folder' } (Join-Path $group.Root $s.Folder)
                        [void]$rootNode.Items.Add($folders[$s.Folder])
                    }
                    $parent = $folders[$s.Folder]
                }
                $node = New-WtFlowTreeItem $s.Name 0xE756 @{ Kind = 'script'; Path = $s.Path } $s.Path
                [void]$parent.Items.Add($node)
                if ($keep -eq "script:$($s.Path)") { $found = $node }
            }
            if (-not $group.Missing -and -not $group.Scripts.Count) {
                $hint = New-WtFlowTreeItem '.ps1 파일이 없습니다' 0xE946 @{ Kind = 'hint' } -Muted
                $hint.Focusable = $false
                [void]$rootNode.Items.Add($hint)
            }
            [void]$tree.Items.Add($rootNode)
        }
        if ($found) { $found.IsSelected = $true; $found.BringIntoView() }
    } finally { $script:filling = $false }
    if ($found) {
        Select-WtFlowNode $found.Tag
        if ($carry -and $carry.Key -eq $keep) {
            Set-WtFlowFormValues $carry.Values
            foreach ($rb in $script:ui.ModeButtons) { $rb.IsChecked = ($rb.Tag.Value -eq $carry.Mode) }
        }
    } elseif ($keep) { Clear-WtFlowSelection }
    else { Update-WtFlowCenter }
}

function Add-WtFlowRoot {
    $dlg = [System.Windows.Forms.FolderBrowserDialog]::new()
    $dlg.Description = '업무 탭에서 실행할 스크립트(.ps1)가 있는 폴더를 고르세요'
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $path = $dlg.SelectedPath
    $roots = @(Get-WtFlowRoots)
    if (-not ($roots | Where-Object { $_ -ieq $path })) { Set-WtSetting 'workflows.roots' @($roots + $path) }
    Update-WtFlowTree
    Set-WtFlowNote "스크립트 폴더를 더했습니다: $path"
}

function Remove-WtFlowRoot {
    $node = $script:ui.Tree.SelectedItem
    if (-not $node -or $node.Tag.Kind -ne 'root') { return }
    $path = $node.Tag.Path
    Set-WtSetting 'workflows.roots' @(Get-WtFlowRoots | Where-Object { $_ -ine $path })
    Clear-WtFlowSelection
    Update-WtFlowTree
    Set-WtFlowNote "목록에서 뺐습니다: $path (폴더와 파일은 그대로)"
}

# ---- 화면: 고른 것 · 폼 -------------------------------------------------------------------------

function Clear-WtFlowSelection {
    $script:info = $null
    $script:flow = $null
    $script:fields = @()
    Update-WtFlowCenter
}

function Select-WtFlowNode($Tag) {
    $script:ui.RemoveRoot.IsEnabled = ($Tag.Kind -eq 'root')
    switch ($Tag.Kind) {
        'script' {
            $script:flow = $null
            $script:info = Read-WtFlowScript $Tag.Path
            Update-WtFlowCenter
        }
        'flow' {
            $f = Get-WtFlows | Where-Object { $_.id -eq $Tag.Id } | Select-Object -First 1
            if (-not $f) { Clear-WtFlowSelection; return }
            $script:flow = $f
            $script:info = Read-WtFlowScript ([string]$f.script)
            Update-WtFlowCenter
        }
    }
}

# 가운데 · 오른쪽을 고른 것에 맞춘다
function Update-WtFlowCenter {
    $u = $script:ui
    if (-not $u.Count) { return }
    $u.Form.Children.Clear()
    $script:fields = @()
    $hasRoots = (Get-WtFlowRoots).Count -gt 0
    if (-not $script:info) {
        $u.Title.Text = '업무'
        $u.Sub.Text = if ($hasRoots -or (Get-WtFlows).Count) { '왼쪽에서 스크립트나 저장한 업무를 고르세요.' } else { '스크립트 폴더를 더하면 그 안의 스크립트를 폼으로 실행합니다.' }
        Set-WtResource $u.Sub 'Foreground' 'B.SubText'
        $u.Title.ToolTip = $null
        $u.Empty.Visibility = if ($hasRoots) { 'Collapsed' } else { 'Visible' }
        foreach ($x in @($u.Commands, $u.SaveFlow, $u.DeleteFlow, $u.Repick, $u.OutArea)) { $x.Visibility = 'Collapsed' }
        Update-WtFlowHelp
        return
    }
    $u.Empty.Visibility = 'Collapsed'
    $u.Commands.Visibility = 'Visible'
    $u.SaveFlow.Visibility = 'Visible'
    $u.OutArea.Visibility = 'Visible'
    $info = $script:info
    $u.Title.Text = if ($script:flow) { [string]$script:flow.title } else { $info.Name }
    $u.Title.ToolTip = $info.Path
    $u.Sub.Text = if ($info.Error) { $info.Error } elseif ($info.Synopsis) { $info.Synopsis } else { $info.Warning }
    if ($info.Error) { Set-WtResource $u.Sub 'Foreground' 'B.Danger' } else { Set-WtResource $u.Sub 'Foreground' 'B.SubText' }
    $values = if ($script:flow) { ConvertTo-WtFlowValues $script:flow.values } else { @{} }
    if ($script:flow) {
        $stale = @(Get-WtFlowStaleNames $info $values)
        if ($stale.Count) {
            $note = New-WtText "스크립트에 없는 값 $($stale.Count)개는 넘기지 않습니다: $($stale -join ', ')" 'T.Sub'
            Set-WtResource $note 'Foreground' 'B.Warn'
            $note.TextWrapping = 'Wrap'
            [void](Add-WtChild $u.Form $note)
        }
    }
    foreach ($f in $info.Fields) { New-WtFlowField $f $values }
    if (-not $info.Fields.Count -and -not $info.Error) { [void](Add-WtChild $u.Form (New-WtText '넣을 값이 없는 스크립트입니다. [실행]을 누르면 바로 실행합니다.' 'T.Sub')) }
    if ($info.Skipped.Count) {
        $names = ($info.Skipped | ForEach-Object { "$($_.Name)($($_.TypeName))" }) -join ', '
        $skip = New-WtText "폼에서 넣을 수 없는 값 $($info.Skipped.Count)개는 스크립트 기본값을 씁니다: $names" 'T.Sub'
        $skip.TextWrapping = 'Wrap'
        [void](Add-WtChild $u.Form $skip)
    }
    # 실행 방식: 업무에 저장한 것 → 스크립트가 정한 것(.NOTES의 WorkTool: 콘솔) → 앱 안
    $mode = if ($script:flow -and $script:flow.mode) { [string]$script:flow.mode } elseif ($info.Console) { 'console' } else { 'app' }
    foreach ($rb in $u.ModeButtons) { $rb.IsChecked = ($rb.Tag.Value -eq $mode) }
    $u.SaveFlow.Content = New-WtButtonContent $(if ($script:flow) { '업무 저장' } else { '업무로 저장' }) 0xE74E
    $u.SaveFlow.ToolTip = if ($script:flow) { '지금 값과 실행 방식을 이 업무에 저장합니다' } else { '지금 값과 실행 방식을 업무로 저장해 왼쪽 목록에 둡니다' }
    $u.DeleteFlow.Visibility = if ($script:flow) { 'Visible' } else { 'Collapsed' }
    $u.Repick.Visibility = if ($script:flow -and -not (Test-Path -LiteralPath $info.Path -PathType Leaf)) { 'Visible' } else { 'Collapsed' }
    Update-WtFlowCommands
    Update-WtFlowHelp
}

# 칸 하나: 라벨(필수) · 설명 · 입력 · 오류
function New-WtFlowField($Field, [hashtable]$Values) {
    $box = New-WtStack Vertical 8
    $value = if ($Values.ContainsKey($Field.Name)) { $Values[$Field.Name] } else { $Field.Default }
    $control = $null
    if ($Field.Kind -eq 'switch') {
        $control = [System.Windows.Controls.CheckBox]::new()
        Set-WtStyle $control 'Check'
        $control.Content = $Field.Name
        $control.IsChecked = ($value -eq $true)
        [void](Add-WtChild $box $control)
        if ($Field.Help) { $h = New-WtText $Field.Help 'T.Sub'; $h.TextWrapping = 'Wrap'; [void](Add-WtChild $box $h) }
    } else {
        $label = New-WtText $(if ($Field.Mandatory) { "$($Field.Name) (필수)" } else { $Field.Name }) 'T.Strong'
        [void](Add-WtChild $box $label)
        $helpText = $Field.Help
        if (-not $Field.Mandatory -and $null -eq $Field.Default) { $helpText = (@($helpText, '비우면 스크립트 기본값') | Where-Object { $_ }) -join ' · ' }
        if ($helpText) { $h = New-WtText $helpText 'T.Sub'; $h.TextWrapping = 'Wrap'; [void](Add-WtChild $box $h) }
        if ($Field.Kind -eq 'choice') {
            $control = [System.Windows.Controls.ComboBox]::new()
            Set-WtStyle $control 'Combo'
            if (-not $Field.Mandatory) { [void]$control.Items.Add('(기본값)') }
            foreach ($c in $Field.Choices) { [void]$control.Items.Add($c) }
            $pick = if ($null -ne $value -and $Field.Choices -contains [string]$value) { [string]$value } elseif (-not $Field.Mandatory) { '(기본값)' } else { $null }
            if ($pick) { $control.SelectedItem = ($control.Items | Where-Object { $_ -eq $pick } | Select-Object -First 1) }
            $control.HorizontalAlignment = 'Left'
            $control.MinWidth = 200
            [System.Windows.Automation.AutomationProperties]::SetName($control, $Field.Name)
            [void](Add-WtChild $box $control)
        } else {
            $control = [System.Windows.Controls.TextBox]::new()
            Set-WtStyle $control 'Input'
            if ($null -ne $value) { $control.Text = [string]$value }
            [System.Windows.Automation.AutomationProperties]::SetName($control, $Field.Name)
            if ($Field.Kind -eq 'number') {
                $control.Width = 160
                $control.HorizontalAlignment = 'Left'
                [void](Add-WtChild $box $control)
            } elseif ($Field.Kind -in 'file', 'folder') {
                $row = [System.Windows.Controls.DockPanel]::new()
                $browse = New-WtButton '찾아보기' -Icon 0xE838
                $browse.Margin = [System.Windows.Thickness]::new(8, 0, 0, 0)
                $browse.Tag = @{ Input = $control; Kind = $Field.Kind }
                $browse.Add_Click({ param($s, $e) Select-WtFlowPath $s.Tag.Input $s.Tag.Kind })
                [System.Windows.Controls.DockPanel]::SetDock($browse, 'Right')
                [void]$row.Children.Add($browse)
                [void]$row.Children.Add($control)
                [void](Add-WtChild $box $row)
            } else {
                [void](Add-WtChild $box $control)
            }
        }
    }
    $err = New-WtText '' 'T.Sub'
    Set-WtResource $err 'Foreground' 'B.Danger'
    $err.Visibility = 'Collapsed'
    [void](Add-WtChild $box $err)
    [void](Add-WtChild $script:ui.Form $box)
    $entry = @{ Field = $Field; Input = $control; Error = $err }
    if ($Field.Kind -eq 'number') {
        # 숫자 칸은 칸을 떠날 때 형식 · 범위를 바로 알린다. 비운 필수 칸은 [실행]할 때 알린다
        $control.Tag = $entry
        $control.Add_LostKeyboardFocus({ param($s, $e) if ($s.Text.Trim()) { Set-WtFlowFieldError $s.Tag } else { Set-WtFlowFieldError $s.Tag '' } })
    }
    $script:fields += $entry
}

# 칸 아래 오류 한 줄. 화면 읽기 프로그램도 읽게 칸의 도움말에 같은 글을 둔다. $Message를 안 주면 그 칸만 검사한다
function Set-WtFlowFieldError($Entry, $Message) {
    if ($null -eq $Message) {
        $Message = (Test-WtFlowValues ([pscustomobject]@{ Fields = @($Entry.Field) }) @{ $Entry.Field.Name = $Entry.Input.Text })[$Entry.Field.Name]
    }
    $Entry.Error.Text = [string]$Message
    $Entry.Error.Visibility = if ($Message) { 'Visible' } else { 'Collapsed' }
    [System.Windows.Automation.AutomationProperties]::SetHelpText($Entry.Input, [string]$Message)
}

function Select-WtFlowPath($TextBox, [string]$Kind) {
    if ($Kind -eq 'folder') {
        $dlg = [System.Windows.Forms.FolderBrowserDialog]::new()
        if ($TextBox.Text -and (Test-Path -LiteralPath $TextBox.Text -PathType Container)) { $dlg.SelectedPath = $TextBox.Text }
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $TextBox.Text = $dlg.SelectedPath }
        return
    }
    $dlg = [Microsoft.Win32.OpenFileDialog]::new()
    if ($TextBox.Text) {
        $dir = Split-Path $TextBox.Text -Parent
        if ($dir -and (Test-Path -LiteralPath $dir -PathType Container)) { $dlg.InitialDirectory = $dir }
    } elseif ($script:info) { $dlg.InitialDirectory = Split-Path $script:info.Path -Parent }
    if ($dlg.ShowDialog()) { $TextBox.Text = $dlg.FileName }
}

# 폼의 값: @{ 이름 = 글자 | $true/$false }. (기본값)을 고른 칸은 넘기지 않게 비운다
function Get-WtFlowFormValues {
    $values = @{}
    foreach ($x in $script:fields) {
        $values[$x.Field.Name] = switch ($x.Field.Kind) {
            'switch' { [bool]$x.Input.IsChecked }
            'choice' { if ($x.Input.SelectedItem -and $x.Input.SelectedItem -ne '(기본값)') { [string]$x.Input.SelectedItem } else { '' } }
            default { $x.Input.Text }
        }
    }
    $values
}

function Get-WtFlowMode { $rb = $script:ui.ModeButtons | Where-Object { $_.IsChecked } | Select-Object -First 1; if ($rb) { $rb.Tag.Value } else { 'app' } }

# 폼에 값을 다시 채운다 (화면을 다시 열 때 고치던 값을 지키려고)
function Set-WtFlowFormValues([hashtable]$Values) {
    foreach ($x in $script:fields) {
        if (-not $Values.ContainsKey($x.Field.Name)) { continue }
        $v = $Values[$x.Field.Name]
        switch ($x.Field.Kind) {
            'switch' { $x.Input.IsChecked = [bool]$v }
            'choice' {
                $pick = if ([string]$v) { [string]$v } else { '(기본값)' }
                $item = $x.Input.Items | Where-Object { $_ -eq $pick } | Select-Object -First 1
                if ($item) { $x.Input.SelectedItem = $item }
            }
            default { $x.Input.Text = [string]$v }
        }
    }
}

# 명령 버튼: 못 쓸 때는 숨기지 않고 비활성 + 이유
function Update-WtFlowCommands {
    $u = $script:ui
    if (-not $u.Count) { return }
    $running = [bool]$script:run
    $why = if (-not $script:info) { '스크립트를 먼저 고르세요' } elseif ($script:info.Error) { $script:info.Error } elseif ($running) { '다른 스크립트가 앱 안에서 실행 중입니다. 끝나거나 [취소]한 뒤에 실행하세요' } else { '' }
    $u.RunButton.IsEnabled = -not $why
    $u.RunButton.ToolTip = if ($why) { $why } else { '이 값으로 실행합니다' }
    $previewWhy = if ($why) { $why } elseif (-not $script:info.WhatIf) { '이 스크립트는 미리보기(-WhatIf)를 지원하지 않습니다' } else { '' }
    $u.Preview.IsEnabled = -not $previewWhy
    $u.Preview.ToolTip = if ($previewWhy) { $previewWhy } else { '-WhatIf로 할 일만 보여 줍니다. 바꾸는 것은 없습니다' }
    $u.Cancel.IsEnabled = $running
    $u.Cancel.ToolTip = if ($running) { '실행 중인 스크립트를 끝냅니다 (자식 프로세스까지)' } else { '앱 안에서 실행 중일 때만 쓸 수 있습니다' }
    $u.SaveFlow.IsEnabled = [bool]($script:info -and -not ($script:flow -and -not (Test-Path -LiteralPath $script:info.Path -PathType Leaf)))
}

function Set-WtFlowNote([string]$Text, [switch]$Undo) {
    if (-not $script:ui.Count) { return }
    $script:ui.Note.Text = $Text
    $script:ui.Undo.Visibility = if ($Undo) { 'Visible' } else { 'Collapsed' }
    $script:ui.NoteRow.Visibility = if ($Text) { 'Visible' } else { 'Collapsed' }
}

# ---- 실행 ---------------------------------------------------------------------------------------

function Invoke-WtFlowRun([switch]$WhatIf) {
    if ($script:run -or -not $script:info -or $script:info.Error) { return }
    $values = Get-WtFlowFormValues
    $problems = Test-WtFlowValues $script:info $values
    $first = $null
    foreach ($x in $script:fields) {
        $msg = [string]$problems[$x.Field.Name]
        Set-WtFlowFieldError $x $msg
        if ($msg -and -not $first) { $first = $x.Input }
    }
    if ($first) { [void]$first.Focus(); Set-WtFlowNote '칸을 고친 뒤 다시 실행하세요.'; return }
    $mode = Get-WtFlowMode
    $title = if ($script:flow) { [string]$script:flow.title } else { '' }
    $record = New-WtFlowRecord $script:info $values $mode ([bool]$WhatIf) $title
    try {
        $start = New-WtFlowStartInfo $script:info $values $mode -WhatIf:$WhatIf
        if ($mode -eq 'console') {
            $p = [System.Diagnostics.Process]::Start($start)
            if ($p) { $script:consoles.Add(@{ Process = $p; Record = $record }) }
            $note = "콘솔 창을 띄웠습니다: $($script:info.Name). 그 창에서 끝내고 닫으세요."
            try { Save-WtFlowRun $record } catch { $note += " (실행 기록은 저장하지 못했습니다: $($_.Exception.Message))" }
            Set-WtFlowNote $note
            Start-WtFlowTimer
            Update-WtFlowHelp
            return
        }
        $script:run = Start-WtFlowProcess $start $record
    } catch {
        Set-WtFlowNote "실행하지 못했습니다: $($_.Exception.Message)"
        return
    }
    $u = $script:ui
    $u.Output.Clear()
    $head = "▶ $($script:info.Name)$(if ($WhatIf) { ' (미리보기)' }) · $((Get-Date).ToString('HH:mm:ss')) 시작"
    $u.Output.AppendText($head)
    $u.OutputHint.Visibility = 'Collapsed'
    $u.Status.Visibility = 'Visible'
    $u.Busy.Visibility = 'Visible'
    $u.RunStatus.Child = $null
    $u.Elapsed.Text = '실행 중 · 0초'
    Set-WtFlowNote ''
    Update-WtFlowCommands
    Start-WtFlowTimer
}

function Stop-WtFlowRun {
    if (-not $script:run) { return }
    $script:run.Cancelled = $true
    Stop-WtFlowProcess $script:run
    Set-WtFlowNote '실행을 멈추는 중입니다…'
}

function Start-WtFlowTimer {
    if (-not $script:timer) {
        $script:timer = New-WtPollTimer { Update-WtFlowRunState }
        $script:timer.Interval = [timespan]::FromMilliseconds(100)
    }
    $script:timer.Start()
}

# 타이머: 새 출력을 붙이고, 끝났으면 결과를 보이고 기록한다. 콘솔 창이 닫혔는지도 본다
function Update-WtFlowRunState {
    try {
        $u = $script:ui
        if ($script:run) {
            $fresh = Read-WtFlowOutput $script:run
            if ($fresh.Count -and $u.Count) {
                if ($u.Output.LineCount -gt $script:maxLines) {
                    $u.Output.Text = ((@($script:run.Lines) | Select-Object -Last $script:maxLines) -join "`n")
                } else {
                    $u.Output.AppendText("`n" + ($fresh -join "`n"))
                }
                $u.Output.ScrollToEnd()
            }
            if ($u.Count) { $u.Elapsed.Text = "실행 중 · $(Format-WtFlowSpan ((Get-Date) - $script:run.Started))" }
            if ($fresh.Count -lt 1000 -and (Test-WtFlowRunDone $script:run)) {
                $done = $script:run
                $script:run = $null
                $status = Complete-WtFlowRun $done
                if ($u.Count) {
                    Show-WtFlowResult $done.Record $status
                    if ($done.PSObject.Properties['SaveError']) { Set-WtFlowNote "실행 기록을 저장하지 못했습니다: $($done.SaveError)" }
                }
            }
        }
        foreach ($c in @($script:consoles)) {
            $exited = $true
            try { $exited = $c.Process.HasExited } catch { }
            if (-not $exited) { continue }
            $c.Record.ended = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            try { Save-WtFlowRun $c.Record } catch { }
            try { $c.Process.Dispose() } catch { }
            [void]$script:consoles.Remove($c)
            if ($u.Count) { Update-WtFlowHelp }
        }
        if (-not $script:run -and -not $script:consoles.Count) { $script:timer.Stop() }
    } catch { $script:timer.Stop(); Show-WtError $_ '업무 실행' }
}

function Show-WtFlowResult($Record, [string]$Status) {
    $u = $script:ui
    $u.Busy.Visibility = 'Collapsed'
    $look = Get-WtFlowStatusLook $Status ([bool]$Record.whatIf)
    $u.RunStatus.Child = New-WtTag $look.Text $look.Bg $look.Fg $look.Icon
    $span = Get-WtFlowRunSpan $Record
    $u.Elapsed.Text = switch ($Status) {
        'success' { "$span · 종료 코드 0" }
        'failed' { "$span · 종료 코드 $($Record.exitCode)" }
        'cancelled' { "$span · 취소해서 멈췄습니다. 스크립트의 마무리 처리는 돌지 않았을 수 있습니다" }
        default { $span }
    }
    $u.Output.AppendText("`n■ $((Get-Date).ToString('HH:mm:ss')) 끝 · $($look.Text)")
    $u.Output.ScrollToEnd()
    Update-WtFlowCommands
    Update-WtFlowHelp
}

# ---- 화면: 업무 저장 · 삭제 ---------------------------------------------------------------------

function Save-WtFlowFromForm {
    if (-not $script:info) { return }
    $values = Get-WtFlowFormValues
    $kept = @{}
    foreach ($k in $values.Keys) { if ($values[$k] -eq $true -or ($values[$k] -isnot [bool] -and ([string]$values[$k]).Trim())) { $kept[$k] = $values[$k] } }
    $mode = Get-WtFlowMode
    $items = @(Get-WtFlows)
    if ($script:flow) {
        foreach ($f in $items) {
            if ($f.id -ne $script:flow.id) { continue }
            $f.values = $kept; $f.mode = $mode; $f.updated = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        }
        try { Save-WtFlows $items } catch { Show-WtError $_ '업무 저장'; return }
        $script:flow = $items | Where-Object { $_.id -eq $script:flow.id } | Select-Object -First 1
        Set-WtFlowNote "업무를 저장했습니다: $($script:flow.title)"
        return
    }
    $title = Read-WtFlowTitle $script:info.Name
    if (-not $title) { return }
    $new = New-WtFlow $title $script:info.Path $kept $mode
    try { Save-WtFlows (@($items) + [pscustomobject]$new) } catch { Show-WtError $_ '업무 저장'; return }
    $script:pendingFlow = $new.id
    Update-WtFlowTree
    Set-WtFlowNote "업무로 저장했습니다: $title"
}

# 업무 이름을 묻는 창. 취소하면 빈 글자
function Read-WtFlowTitle([string]$Suggest) {
    $dialog = New-WtModal '업무로 저장'
    [void](Add-WtChild $dialog.Body (New-WtText '왼쪽 목록에 보일 이름을 정하세요. 예: PM 점검 · 서버실' 'T.Sub'))
    $box = [System.Windows.Controls.TextBox]::new()
    Set-WtStyle $box 'Input'
    $box.Text = $Suggest
    [System.Windows.Automation.AutomationProperties]::SetName($box, '업무 이름')
    [void](Add-WtChild $dialog.Body $box)
    $save = New-WtButton '저장' -Kind Primary; $save.IsDefault = $true
    $cancel = New-WtButton '취소'; $cancel.IsCancel = $true
    $dialog.Window.Tag = @{ Box = $box; Error = $dialog.Error; Result = '' }
    $save.Tag = $dialog.Window
    $save.Add_Click({
            param($s, $e)
            $state = $s.Tag.Tag
            $name = $state.Box.Text.Trim()
            if (-not $name) { $state.Error.Text = '이름을 넣으세요.'; $state.Error.Visibility = 'Visible'; return }
            $state.Result = $name
            $s.Tag.DialogResult = $true
        })
    [void](Add-WtChild $dialog.Buttons $save)
    [void](Add-WtChild $dialog.Buttons $cancel)
    $dialog.Window.Add_ContentRendered({ param($s, $e) [void]$s.Tag.Box.Focus(); $s.Tag.Box.SelectAll() })
    [void]$dialog.Window.ShowDialog()
    [string]$dialog.Window.Tag.Result
}

function Remove-WtFlowCurrent {
    if (-not $script:flow) { return }
    $items = @(Get-WtFlows)
    $gone = $items | Where-Object { $_.id -eq $script:flow.id } | Select-Object -First 1
    try { Save-WtFlows @($items | Where-Object { $_.id -ne $script:flow.id }) } catch { Show-WtError $_ '업무 삭제'; return }
    $script:lastDeleted = $gone
    Clear-WtFlowSelection
    Update-WtFlowTree
    Set-WtFlowNote "업무를 지웠습니다: $($gone.title)" -Undo
}

function Restore-WtFlowDeleted {
    if (-not $script:lastDeleted) { return }
    $items = @(Get-WtFlows | Where-Object { $_.id -ne $script:lastDeleted.id })
    try { Save-WtFlows (@($items) + $script:lastDeleted) } catch { Show-WtError $_ '업무 되살리기'; return }
    $script:pendingFlow = $script:lastDeleted.id
    $script:lastDeleted = $null
    Update-WtFlowTree
    Set-WtFlowNote '지운 업무를 되살렸습니다.'
}

# 저장한 업무의 스크립트가 옮겨졌을 때 다시 고른다
function Select-WtFlowScriptAgain {
    if (-not $script:flow) { return }
    $dlg = [Microsoft.Win32.OpenFileDialog]::new()
    $dlg.Filter = 'PowerShell 스크립트 (*.ps1)|*.ps1'
    if (-not $dlg.ShowDialog()) { return }
    $items = @(Get-WtFlows)
    foreach ($f in $items) { if ($f.id -eq $script:flow.id) { $f.script = $dlg.FileName; $f.updated = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') } }
    try { Save-WtFlows $items } catch { Show-WtError $_ '스크립트 다시 고르기'; return }
    $script:pendingFlow = $script:flow.id
    Update-WtFlowTree
}

# ---- 화면: 설명 · 기록 --------------------------------------------------------------------------

function Update-WtFlowHelp {
    $u = $script:ui
    if (-not $u.Count) { return }
    $show = [bool](Get-WtSetting 'workflows.help')
    $u.HelpColumn.Width = if ($show) { [System.Windows.GridLength]::new(280) } else { [System.Windows.GridLength]::new(0) }
    $u.HelpCard.Visibility = if ($show) { 'Visible' } else { 'Collapsed' }
    $u.HelpOpen.Visibility = if ($show) { 'Collapsed' } else { 'Visible' }
    $panel = $u.HelpPanel
    $panel.Children.Clear()
    if (-not $show) { return }
    $info = $script:info
    if (-not $info) {
        $t = New-WtText '스크립트를 고르면 그 스크립트의 도움말과 최근 실행 기록이 여기에 보입니다.' 'T.Sub'
        $t.TextWrapping = 'Wrap'
        [void](Add-WtChild $panel $t)
    } else {
        $file = New-WtText $info.Path 'T.Sub'
        Set-WtResource $file 'FontFamily' 'Font.Mono'
        $file.TextWrapping = 'Wrap'
        [void](Add-WtChild $panel $file)
        foreach ($text in @($info.Description)) {
            if (-not $text) { continue }
            $t = New-WtText $text 'T.Body'; $t.TextWrapping = 'Wrap'
            [void](Add-WtChild $panel $t)
        }
        foreach ($ex in $info.Examples) {
            $b = [System.Windows.Controls.Border]::new()
            Set-WtResource $b 'Background' 'B.Code.Bg'
            $b.CornerRadius = [System.Windows.CornerRadius]::new(8)
            $b.Padding = [System.Windows.Thickness]::new(12, 8, 12, 8)
            $t = New-WtText $ex 'T.Sub'
            Set-WtResource $t 'FontFamily' 'Font.Mono'
            $t.TextWrapping = 'Wrap'
            $b.Child = $t
            [void](Add-WtChild $panel $b)
        }
        if ($info.Notes) { $n = New-WtText $info.Notes 'T.Sub'; $n.TextWrapping = 'Wrap'; [void](Add-WtChild $panel $n) }
        if (-not $info.Description -and -not $info.Examples.Count -and -not $info.Notes) { [void](Add-WtChild $panel (New-WtText '이 스크립트에는 설명(.DESCRIPTION · .EXAMPLE)이 없습니다.' 'T.Sub')) }
    }
    $head = New-WtText '최근 실행' 'T.H2'
    $head.Margin = [System.Windows.Thickness]::new(0, 12, 0, 0)
    [void](Add-WtChild $panel $head)
    $runs = @(Get-WtFlowRuns | Where-Object { -not $info -or $_.script -eq $info.Path } | Select-Object -First 10)
    if (-not $runs.Count) { [void](Add-WtChild $panel (New-WtText '아직 실행 기록이 없습니다.' 'T.Sub')); return }
    $list = [System.Windows.Controls.ListBox]::new()
    Set-WtStyle $list 'List'
    [System.Windows.Controls.ScrollViewer]::SetVerticalScrollBarVisibility($list, 'Disabled')
    foreach ($r in $runs) {
        $look = Get-WtFlowStatusLook ([string]$r.status) ([bool]$r.whatIf)
        $row = New-WtStack Vertical 4
        [void](Add-WtChild $row (New-WtTag $look.Text $look.Bg $look.Fg $look.Icon))
        $when = ([datetime]$r.started).ToString('M/d HH:mm')
        $span = Get-WtFlowRunSpan $r
        $line = (@($when, $(if (-not $info) { [string]$r.name }), $span, $(if ($r.flow) { [string]$r.flow })) | Where-Object { $_ }) -join ' · '
        [void](Add-WtChild $row (New-WtText $line 'T.Sub'))
        $item = [System.Windows.Controls.ListBoxItem]::new()
        $item.Content = $row
        $item.Tag = $r
        $item.ToolTip = '누르면 이 실행의 출력 마지막 줄을 아래 출력 칸에 보여 줍니다'
        [void]$list.Items.Add($item)
    }
    $list.Add_SelectionChanged({ param($s, $e) if ($s.SelectedItem) { Show-WtFlowRecordTail $s.SelectedItem.Tag } })
    [void](Add-WtChild $panel $list)
}

# 기록을 고르면 그때의 출력 마지막 줄을 출력 칸에 보여 준다 (실행 중이면 바꾸지 않는다)
function Show-WtFlowRecordTail($Record) {
    if ($script:run) { Set-WtFlowNote '실행 중에는 출력 칸을 바꾸지 않습니다.'; return }
    $u = $script:ui
    $u.OutputHint.Visibility = 'Collapsed'
    $look = Get-WtFlowStatusLook ([string]$Record.status) ([bool]$Record.whatIf)
    $lines = @("── $($Record.name) · $($Record.started) · $($look.Text) (기록: 출력 마지막 $($script:tailLines)줄) ──")
    if ($Record.status -eq 'console') { $lines += '콘솔 창 실행은 출력을 남기지 않습니다. 그 창에서 확인하세요.' }
    $lines += @($Record.tail)
    $u.Output.Text = $lines -join "`n"
    $u.Output.ScrollToEnd()
}

# ---- 화면 만들기 --------------------------------------------------------------------------------

function New-WtFlowsView {
    if ($script:timer) { $script:timer.Stop() }
    $script:ui = @{}
    $page = New-WtPage -Fill
    $add = New-WtButton '폴더 추가' -Icon 0xE710 -ToolTip '실행할 스크립트가 있는 폴더 추가'
    $add.Add_Click({ Add-WtFlowRoot })
    $header = New-WtPageHeader '업무' '' $add
    [System.Windows.Controls.DockPanel]::SetDock($header, 'Top')
    [void]$page.Panel.Children.Add($header)

    $grid = [System.Windows.Controls.Grid]::new()
    $c1 = [System.Windows.Controls.ColumnDefinition]::new(); $c1.Width = [System.Windows.GridLength]::new(240)
    $c2 = [System.Windows.Controls.ColumnDefinition]::new(); $c2.Width = [System.Windows.GridLength]::new(1, 'Star'); $c2.MinWidth = 360
    $c3 = [System.Windows.Controls.ColumnDefinition]::new(); $c3.Width = [System.Windows.GridLength]::new(280)
    foreach ($c in @($c1, $c2, $c3)) { $grid.ColumnDefinitions.Add($c) }

    # 왼쪽: 저장한 업무 + 스크립트 폴더
    $left = New-WtCard
    $left.Padding = [System.Windows.Thickness]::new(12)
    $left.Margin = [System.Windows.Thickness]::new(0, 0, 16, 0)
    $ldock = [System.Windows.Controls.DockPanel]::new()
    $remove = New-WtButton '목록에서 빼기' -Kind Link -Icon 0xE738
    $remove.ToolTip = '고른 스크립트 폴더를 목록에서 뺍니다 (폴더와 파일은 그대로)'
    [System.Windows.Controls.ToolTipService]::SetShowOnDisabled($remove, $true)
    $remove.HorizontalAlignment = 'Left'
    $remove.IsEnabled = $false
    $remove.Add_Click({ Remove-WtFlowRoot })
    [System.Windows.Controls.DockPanel]::SetDock($remove, 'Bottom')
    [void]$ldock.Children.Add($remove)
    $tree = [System.Windows.Controls.TreeView]::new()
    Set-WtStyle $tree 'Tree'
    $tree.Add_SelectedItemChanged({
            param($s, $e)
            $node = $s.SelectedItem
            if ($script:filling -or -not $node -or $node -isnot [System.Windows.Controls.TreeViewItem]) { return }
            try { Select-WtFlowNode $node.Tag } catch { Show-WtError $_ '업무 고르기' }
        })
    [void]$ldock.Children.Add($tree)
    $left.Child = $ldock
    [System.Windows.Controls.Grid]::SetColumn($left, 0)
    [void]$grid.Children.Add($left)

    # 가운데: 제목 · 업무 저장 / 폼 / 실행 명령 / 결과 / 알림 / 출력
    $center = New-WtCard
    $center.Padding = [System.Windows.Thickness]::new(16, 12, 16, 16)
    $cgrid = [System.Windows.Controls.Grid]::new()
    # 폼과 출력이 남는 높이를 3:2로 나눈다. 칸이 많으면 폼만 스크롤하고, 명령은 늘 보인다
    foreach ($h in @(@('Auto', 0, 0), @('Star', 80, 3), @('Auto', 0, 0), @('Auto', 0, 0), @('Auto', 0, 0), @('Star', 100, 2))) {
        $rd = [System.Windows.Controls.RowDefinition]::new()
        $rd.Height = if ($h[0] -eq 'Auto') { [System.Windows.GridLength]::Auto } else { [System.Windows.GridLength]::new($h[2], 'Star') }
        if ($h[1]) { $rd.MinHeight = $h[1] }
        $cgrid.RowDefinitions.Add($rd)
    }
    $head = [System.Windows.Controls.DockPanel]::new()
    $head.Margin = [System.Windows.Thickness]::new(0, 0, 0, 12)
    $headActions = New-WtStack Horizontal 8
    $headActions.VerticalAlignment = 'Top'
    $saveFlow = New-WtButton '업무로 저장' -Icon 0xE74E
    [System.Windows.Controls.ToolTipService]::SetShowOnDisabled($saveFlow, $true)
    $saveFlow.Add_Click({ Save-WtFlowFromForm })
    $deleteFlow = New-WtButton '업무 삭제' -Kind Link -Icon 0xE74D
    $deleteFlow.ToolTip = '저장한 업무를 지웁니다 (스크립트 파일은 그대로)'
    $deleteFlow.Add_Click({ Remove-WtFlowCurrent })
    $helpOpen = New-WtButton '설명 펼치기' -Kind Link -Icon 0xE946
    $helpOpen.ToolTip = '오른쪽에 스크립트 도움말과 최근 실행을 보여 줍니다'
    $helpOpen.Add_Click({ Set-WtSetting 'workflows.help' $true; Update-WtFlowHelp })
    foreach ($b in @($deleteFlow, $saveFlow, $helpOpen)) { [void](Add-WtChild $headActions $b) }
    [System.Windows.Controls.DockPanel]::SetDock($headActions, 'Right')
    [void]$head.Children.Add($headActions)
    $titleStack = New-WtStack Vertical 4
    $titleStack.Margin = [System.Windows.Thickness]::new(0, 0, 12, 0)
    $title = New-WtText '업무' 'T.H2'
    $title.TextTrimming = 'CharacterEllipsis'
    $sub = New-WtText '' 'T.Sub'
    $sub.TextWrapping = 'Wrap'
    $repick = New-WtButton '스크립트 다시 고르기' -Icon 0xE8E5
    $repick.HorizontalAlignment = 'Left'
    $repick.Visibility = 'Collapsed'
    $repick.Add_Click({ Select-WtFlowScriptAgain })
    foreach ($x in @($title, $sub, $repick)) { [void](Add-WtChild $titleStack $x) }
    [void]$head.Children.Add($titleStack)
    [System.Windows.Controls.Grid]::SetRow($head, 0)
    [void]$cgrid.Children.Add($head)

    $formScroll = [System.Windows.Controls.ScrollViewer]::new()
    $formScroll.VerticalScrollBarVisibility = 'Auto'
    $formScroll.HorizontalScrollBarVisibility = 'Disabled'
    $formScroll.Focusable = $false
    $form = New-WtStack Vertical 16
    $form.Margin = [System.Windows.Thickness]::new(0, 0, 8, 0)
    $formScroll.Content = $form
    $formArea = [System.Windows.Controls.Grid]::new()
    [void]$formArea.Children.Add($formScroll)
    $empty = New-WtStack Vertical 12
    $empty.HorizontalAlignment = 'Center'
    $empty.VerticalAlignment = 'Center'
    $emptyText = New-WtText '스크립트 폴더가 없습니다.' 'T.Sub'
    $emptyText.HorizontalAlignment = 'Center'
    $emptyAdd = New-WtButton '폴더 추가' -Icon 0xE710
    $emptyAdd.HorizontalAlignment = 'Center'
    $emptyAdd.Add_Click({ Add-WtFlowRoot })
    [void](Add-WtChild $empty $emptyText)
    [void](Add-WtChild $empty $emptyAdd)
    [void]$formArea.Children.Add($empty)
    [System.Windows.Controls.Grid]::SetRow($formArea, 1)
    [void]$cgrid.Children.Add($formArea)

    # 실행 명령: [실행] [미리보기] [취소] 실행 방식. 좁으면 실행 방식이 다음 줄로 내려간다
    $commands = [System.Windows.Controls.WrapPanel]::new()
    $commands.Margin = [System.Windows.Thickness]::new(0, 12, 0, 0)
    $runButton = New-WtButton '실행' -Kind Primary -Icon 0xE768
    $runButton.Add_Click({ Invoke-WtFlowRun })
    $preview = New-WtButton '미리보기' -Icon 0xE7B3
    $preview.Add_Click({ Invoke-WtFlowRun -WhatIf })
    $cancel = New-WtButton '취소' -Icon 0xE711
    $cancel.Add_Click({ Stop-WtFlowRun })
    $mode = New-WtSegmented @(@('app', '앱 안에서'), @('console', '콘솔 창에서')) 'app' 'WtFlowMode' { param($value) }
    $mode.ToolTip = '앱 안: 출력을 아래 칸에 보여 줍니다 · 콘솔 창: 자기 화면이 있거나 키 입력을 받는 스크립트'
    foreach ($b in @($runButton, $preview, $cancel, $mode)) {
        [System.Windows.Controls.ToolTipService]::SetShowOnDisabled($b, $true)
        $b.Margin = [System.Windows.Thickness]::new(0, 0, 8, 8)
        [void]$commands.Children.Add($b)
    }
    [System.Windows.Controls.Grid]::SetRow($commands, 2)
    [void]$cgrid.Children.Add($commands)

    # 결과: 상태(글자 · 아이콘 · 색) + 걸린 시간 · 종료 코드, 실행 중이면 진행 막대
    $statusStack = New-WtStack Vertical 4
    $statusStack.Margin = [System.Windows.Thickness]::new(0, 4, 0, 0)
    $statusStack.Visibility = 'Collapsed'
    $statusRow = [System.Windows.Controls.DockPanel]::new()
    $runStatus = [System.Windows.Controls.Border]::new()
    $runStatus.Margin = [System.Windows.Thickness]::new(0, 0, 8, 0)
    [System.Windows.Controls.DockPanel]::SetDock($runStatus, 'Left')
    $elapsed = New-WtText '' 'T.Sub'
    $elapsed.VerticalAlignment = 'Center'
    $elapsed.TextWrapping = 'Wrap'
    [void]$statusRow.Children.Add($runStatus)
    [void]$statusRow.Children.Add($elapsed)
    $busy = [System.Windows.Controls.ProgressBar]::new()
    Set-WtStyle $busy 'Busy'
    $busy.Visibility = 'Collapsed'
    [void](Add-WtChild $statusStack $statusRow)
    [void](Add-WtChild $statusStack $busy)
    [System.Windows.Controls.Grid]::SetRow($statusStack, 3)
    [void]$cgrid.Children.Add($statusStack)

    # 알림 한 줄 (+ 업무 삭제의 실행 취소)
    $noteRow = [System.Windows.Controls.DockPanel]::new()
    $noteRow.Margin = [System.Windows.Thickness]::new(0, 4, 0, 0)
    $noteRow.Visibility = 'Collapsed'
    $undo = New-WtButton '실행 취소' -Kind Link -Icon 0xE7A7
    $undo.Visibility = 'Collapsed'
    $undo.Add_Click({ Restore-WtFlowDeleted })
    [System.Windows.Controls.DockPanel]::SetDock($undo, 'Right')
    [void]$noteRow.Children.Add($undo)
    $note = New-WtText '' 'T.Sub'
    $note.TextWrapping = 'Wrap'
    $note.VerticalAlignment = 'Center'
    [void]$noteRow.Children.Add($note)
    [System.Windows.Controls.Grid]::SetRow($noteRow, 4)
    [void]$cgrid.Children.Add($noteRow)

    $outArea = [System.Windows.Controls.Grid]::new()
    $outArea.Margin = [System.Windows.Thickness]::new(0, 8, 0, 0)
    $output = [System.Windows.Controls.TextBox]::new()
    Set-WtStyle $output 'Output'
    [System.Windows.Automation.AutomationProperties]::SetName($output, '실행 출력')
    [void]$outArea.Children.Add($output)
    $outHint = New-WtText '실행하면 출력이 여기에 보입니다.' 'T.Sub'
    $outHint.HorizontalAlignment = 'Center'
    $outHint.VerticalAlignment = 'Center'
    $outHint.IsHitTestVisible = $false
    [void]$outArea.Children.Add($outHint)
    [System.Windows.Controls.Grid]::SetRow($outArea, 5)
    [void]$cgrid.Children.Add($outArea)
    $center.Child = $cgrid
    [System.Windows.Controls.Grid]::SetColumn($center, 1)
    [void]$grid.Children.Add($center)

    # 오른쪽: 설명 · 최근 실행. 머리글의 [접기]는 늘 같은 자리에 둔다
    $right = New-WtCard
    $right.Padding = [System.Windows.Thickness]::new(16, 12, 16, 12)
    $right.Margin = [System.Windows.Thickness]::new(16, 0, 0, 0)
    $rdock = [System.Windows.Controls.DockPanel]::new()
    $rhead = [System.Windows.Controls.DockPanel]::new()
    $rhead.Margin = [System.Windows.Thickness]::new(0, 0, 0, 8)
    $helpClose = New-WtButton '접기' -Kind Link -Icon 0xE76C
    $helpClose.ToolTip = '설명을 접어 폼을 넓게 씁니다'
    $helpClose.Add_Click({ Set-WtSetting 'workflows.help' $false; Update-WtFlowHelp })
    [System.Windows.Controls.DockPanel]::SetDock($helpClose, 'Right')
    [void]$rhead.Children.Add($helpClose)
    $rtitle = New-WtText '설명' 'T.H2'
    $rtitle.VerticalAlignment = 'Center'
    [void]$rhead.Children.Add($rtitle)
    [System.Windows.Controls.DockPanel]::SetDock($rhead, 'Top')
    [void]$rdock.Children.Add($rhead)
    $helpScroll = [System.Windows.Controls.ScrollViewer]::new()
    $helpScroll.VerticalScrollBarVisibility = 'Auto'
    $helpScroll.HorizontalScrollBarVisibility = 'Disabled'
    $helpScroll.Focusable = $false
    $helpPanel = New-WtStack Vertical 12
    $helpScroll.Content = $helpPanel
    [void]$rdock.Children.Add($helpScroll)
    $right.Child = $rdock
    [System.Windows.Controls.Grid]::SetColumn($right, 2)
    [void]$grid.Children.Add($right)
    [void]$page.Panel.Children.Add($grid)

    # 업무 화면 안 단축키
    $page.Root.Add_PreviewKeyDown({
            param($s, $e)
            if ($e.Key -eq 'F5') { Update-WtFlowTree; Set-WtFlowNote '스크립트 목록을 새로 읽었습니다.'; $e.Handled = $true }
        })

    $script:ui = @{
        Tree = $tree; RemoveRoot = $remove; Title = $title; Sub = $sub; Repick = $repick
        Form = $form; Empty = $empty; Commands = $commands; ModeButtons = @($mode.Child.Children)
        SaveFlow = $saveFlow; DeleteFlow = $deleteFlow; RunButton = $runButton; Preview = $preview; Cancel = $cancel
        Status = $statusStack; RunStatus = $runStatus; Elapsed = $elapsed; Busy = $busy; NoteRow = $noteRow; Note = $note; Undo = $undo
        Output = $output; OutputHint = $outHint; OutArea = $outArea
        HelpColumn = $c3; HelpCard = $right; HelpPanel = $helpPanel; HelpOpen = $helpOpen
    }
    $page.Root
}

function Update-WtFlowsView {
    Update-WtFlowTree
    Update-WtFlowCommands
}

# 홈: 저장한 업무와 그 업무의 마지막 결과
function New-WtFlowsHomeCard {
    $flows = @(Get-WtFlows | Select-Object -First 4)
    if (-not $flows.Count) { return }
    $runs = @(Get-WtFlowRuns)
    $card = New-WtCard
    $col = New-WtStack Vertical
    [void](Add-WtChild $col (New-WtCardHeader '업무' '' '업무 열기' 'workflows'))
    foreach ($f in $flows) {
        $last = $runs | Where-Object { $_.flow -eq $f.title -and $_.script -eq $f.script } | Select-Object -First 1
        $name = New-WtButton ([string]$f.title) -Kind Link
        $name.Tag = [string]$f.id
        $name.HorizontalAlignment = 'Left'
        $name.Add_Click({ param($s, $e) $script:pendingFlow = [string]$s.Tag; Select-WtPage 'workflows' })
        $state = if ($last) {
            $look = Get-WtFlowStatusLook ([string]$last.status) ([bool]$last.whatIf)
            New-WtTag $look.Text $look.Bg $look.Fg $look.Icon
        } else { New-WtText '아직 실행 안 함' 'T.Sub' }
        [void](Add-WtChild $col (New-WtListRow @($state, $name)))
    }
    $card.Child = $col
    $card
}

# 앱을 끄면 앱 안 실행은 멈춘다 (창 없는 프로세스를 남기지 않는다). 콘솔 창은 따로 도는 창이라 그대로 둔다
function Stop-WtFlows {
    if ($script:timer) { $script:timer.Stop() }
    if ($script:run) {
        $done = $script:run
        $script:run = $null
        $done.Stopped = $true
        Stop-WtFlowProcess $done
        try { [void]$done.Process.WaitForExit(2000) } catch { }
        try { [void](Read-WtFlowOutput $done) } catch { }
        try { [void](Complete-WtFlowRun $done) } catch { }
    }
    foreach ($c in @($script:consoles)) { try { $c.Process.Dispose() } catch { } }
    $script:consoles.Clear()
}

Register-WtAddon -Id 'workflows' -Title '업무' -Icon 0xE756 -Order 70 `
    -NewView { New-WtFlowsView } `
    -Refresh { param($view) Update-WtFlowsView } `
    -HomeCard { New-WtFlowsHomeCard } `
    -OnExit { Stop-WtFlows } `
    -Help '업무는 스크립트 폴더의 PowerShell 스크립트를 폼으로 실행합니다. 스크립트의 param이 입력 칸이 되고, 자주 쓰는 값은 [업무로 저장]해 두고 다시 씁니다.' `
    -Shortcuts (, @('F5', '스크립트 목록 새로 고침'))

Export-ModuleMember -Function Read-WtFlowScript, Find-WtFlowScripts, Test-WtFlowValues, Get-WtFlowArgs, Get-WtFlowStaleNames,
    ConvertTo-WtFlowValues, ConvertTo-WtFlowCommandArg, New-WtFlowStartInfo, Start-WtFlowProcess, Read-WtFlowOutput, Test-WtFlowRunDone,
    Complete-WtFlowRun, Stop-WtFlowProcess, Get-WtFlows, Save-WtFlows, New-WtFlow, Get-WtFlowRuns, Save-WtFlowRun, New-WtFlowRecord,
    Get-WtFlowStatusLook, Format-WtFlowSpan
