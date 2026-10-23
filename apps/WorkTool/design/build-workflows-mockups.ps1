<#
.SYNOPSIS
    가짜 업무 정의에서 모의 결과와 검토용 SVG·PNG를 만든다.
.DESCRIPTION
    PowerShell 7과 Windows 기본 WPF만 사용한다. 앱 모듈은 불러오지 않는다.
    파일 생성·알람·ping은 데이터 계산만 한다. 샘플 경로에 접근하지 않는다.
    자기 폴더의 workflows-results.json과 workflows-*.svg·png만 다시 만든다.
.EXAMPLE
    pwsh -NoProfile -STA -File apps/WorkTool/design/build-workflows-mockups.ps1
#>
#Requires -Version 7.6
[CmdletBinding()]
param()

# ===== 설정 (검토용 산출물 생성, 앱 설정 아님) =====
$Config = @{
    SamplePath = Join-Path $PSScriptRoot 'workflows-sample.json'
    OutputFolder = $PSScriptRoot
    PngScale = 1.5
}
# ===============================================
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase
$sample = Get-Content -LiteralPath $Config.SamplePath -Raw -Encoding utf8 | ConvertFrom-Json -AsHashtable -DateKind String
if ($sample.isMock -ne $true -or $sample.mockSchema -ne 1) { throw '가짜 목업 형식 1만 지원합니다.' }

function Copy-MockValue($Value) { ConvertFrom-Json -InputObject (ConvertTo-Json -InputObject $Value -Depth 70) -AsHashtable -DateKind String }
function Get-MockReference($Context, [string]$Reference) {
    $value = $Context
    foreach ($part in $Reference.Split('.')) {
        if ($value -isnot [Collections.IDictionary] -or -not $value.Contains($part)) { throw "없는 참조: $Reference" }
        $value = $value[$part]
    }
    return ,$value
}
function Resolve-MockValue($Value, $Context, [int]$Depth = 0) {
    if ($Depth -gt 16) { throw '목업 참조가 순환하거나 너무 깊습니다.' }
    if ($Value -is [string]) {
        if ($Value -match '^\$\{([a-zA-Z0-9_.]+)\}$') {
            return ,(Resolve-MockValue (Get-MockReference $Context $Matches[1]) $Context ($Depth + 1))
        }
        $resolved = [regex]::Replace($Value, '\$\{([a-zA-Z0-9_.]+)\}', {
            param($match)
            $part = Resolve-MockValue (Get-MockReference $Context $match.Groups[1].Value) $Context ($Depth + 1)
            if ($part -is [Collections.IDictionary] -or $part -is [array]) { throw '문자열 안에는 단일 값만 연결할 수 있습니다.' }
            [string]$part
        })
        return $resolved
    }
    if ($Value -is [Collections.IDictionary]) {
        $result = [ordered]@{}
        foreach ($key in $Value.Keys) { $result[$key] = Resolve-MockValue $Value[$key] $Context ($Depth + 1) }
        return $result
    }
    if ($Value -is [array]) {
        $result = @($Value | ForEach-Object { Resolve-MockValue $_ $Context ($Depth + 1) })
        return ,$result
    }
    return $Value
}
function New-MockError([string]$Code, [string]$Message, [string]$Field = '') {
    @{ code = $Code; message = $Message; field = $Field }
}
function Test-MockForm($Definition, $Values) {
    foreach ($field in $Definition.fields) {
        $value = $Values[$field.id]
        $empty = $null -eq $value -or ($value -is [string] -and [string]::IsNullOrWhiteSpace($value)) -or ($value -is [array] -and $value.Count -eq 0)
        if ($field.required -and ($empty -or ($field.type -eq 'checkbox' -and $value -ne $true))) {
            return New-MockError 'INPUT_REQUIRED' ($field.label + ' 항목을 입력하거나 확인하세요.') $field.id
        }
        if ($field.type -eq 'integer' -and ($value -isnot [long] -and $value -isnot [int] -or $value -lt $field.min)) {
            return New-MockError 'INPUT_INTEGER' ($field.label + ' 값은 허용 범위의 정수여야 합니다.') $field.id
        }
        if ($field.type -eq 'choice' -and $value -notin $field.choices) {
            return New-MockError 'INPUT_CHOICE' ($field.label + ' 목록에서 고르세요.') $field.id
        }
        if ($field.type -eq 'datetime-offset') {
            $parsed = [datetimeoffset]::MinValue
            if (-not [datetimeoffset]::TryParse([string]$value, [ref]$parsed) -or $value -notmatch '(Z|[+-]\d{2}:\d{2})$') {
                return New-MockError 'INPUT_DATETIME' '시각과 시간대를 입력하세요.' $field.id
            }
        }
    }
    return $null
}

# 모의 처리기만 존재한다. 샘플에서 명령을 불러오거나 실행하지 않는다.
$handlers = @{
    'path.compose' = { param($p, $context, $fixture, $step)
        @{ path = $p.folder.TrimEnd('/') + '/' + $p.name }
    }
    'content.render' = { param($p, $context, $fixture, $step)
        @{ text = $p.template }
    }
    'file.create' = { param($p, $context, $fixture, $step)
        @{ path = $p.path; text = $p.text; state = 'created' }
    }
    'record.append' = { param($p, $context, $fixture, $step)
        @{ recordId = "mock-$($fixture.id)-$($step.id)"; collection = $p.collection; values = $p.values }
    }
    'network.ping' = { param($p, $context, $fixture, $step)
        $rows = @(foreach ($target in $p.targets) {
            if (-not $fixture.mockResponses.Contains($target)) { throw "가짜 응답표에 없는 대상: $target" }
            $latency = $fixture.mockResponses[$target]
            $ok = $latency -le $p.timeoutMs
            @{ target = $target; status = $(if ($ok) { 'success' } else { 'timeout' }); latencyMs = $(if ($ok) { $latency } else { $null }) }
        })
        @{ rows = $rows; failedCount = @($rows | Where-Object status -eq 'timeout').Count }
    }
    'alarm.schedule' = { param($p, $context, $fixture, $step)
        @{ alarmId = "mock-$($fixture.id)-$($step.id)"; state = 'scheduled'; at = $p.at; title = $p.title; message = $p.message; taskId = $p.taskId }
    }
}

$apiMap = @{}
foreach ($api in $sample.apis) {
    if ($apiMap.ContainsKey($api.id)) { throw "중복 API: $($api.id)" }
    $apiMap[$api.id] = $api
}
$definitionMap = @{}
foreach ($definition in $sample.workflows) {
    if ($definitionMap.ContainsKey($definition.id)) { throw "중복 업무: $($definition.id)" }
    $definitionMap[$definition.id] = $definition
    $stepIds = @()
    foreach ($step in $definition.steps) {
        if ($step.id -in $stepIds -or -not $handlers.ContainsKey($step.action) -or -not $apiMap.ContainsKey($step.action) -or $step.version -ne $apiMap[$step.action].version) { throw "모의 단계 정의 오류: $($step.id)" }
        $stepIds += $step.id
    }
}
$runs = [Collections.Generic.List[object]]::new()
$runMap = @{}
foreach ($fixture in $sample.cases) {
    if ($runMap.ContainsKey($fixture.id)) { throw "중복 사례: $($fixture.id)" }
    $definition = Copy-MockValue $definitionMap[$fixture.workflowId]
    if ($fixture.Contains('settingsOverride')) {
        foreach ($key in $fixture.settingsOverride.Keys) { $definition.settings[$key] = $fixture.settingsOverride[$key] }
        $definition.version = $fixture.definitionVersion
    }
    $context = @{
        input = $fixture.input; settings = $definition.settings; steps = @{}
        context = @{ instanceId = "mock-task-$($fixture.id)"; date = ([datetimeoffset]::Parse($sample.clock)).ToString('yyyyMMdd'); sequence = $fixture.sequence }
    }
    $trace = [Collections.Generic.List[object]]::new()
    $failure = Test-MockForm $definition $fixture.input
    $status = if ($failure) { 'failed' } else { 'succeeded' }
    $cancelledBefore = $null
    if (-not $failure) {
        foreach ($step in $definition.steps) {
            if ($fixture.Contains('cancelBefore') -and $fixture.cancelBefore -eq $step.id) {
                $status = 'cancelled'; $cancelledBefore = $step.id; break
            }
            if ($step.Contains('when')) {
                if ($step.when.operator -ne 'gt') { throw '목업 조건은 gt만 지원합니다.' }
                if ((Get-MockReference $context $step.when.ref) -le $step.when.value) {
                    $trace.Add(@{ id = $step.id; action = $step.action; version = $step.version; args = $null; response = @{ mode = 'simulation'; status = 'skipped'; data = $null; error = $null } }); continue
                }
            }
            $parameters = Resolve-MockValue $step.args $context
            $stepFailure = $null
            $data = $null
            if ($fixture.Contains('failAt') -and $fixture.failAt -eq $step.id) {
                $stepFailure = New-MockError 'MOCK_WRITE_DENIED' '가짜 저장 실패입니다. 저장 위치를 확인하세요.' 'path'
                $status = 'failed'; $failure = $stepFailure
            } else {
                $data = & $handlers[$step.action] $parameters $context $fixture $step
                $context.steps[$step.id] = $data
            }
            $trace.Add(@{ id = $step.id; action = $step.action; version = $step.version; args = $parameters; response = @{ mode = 'simulation'; status = $(if ($stepFailure) { 'failed' } else { 'succeeded' }); data = $data; error = $stepFailure } })
            if ($stepFailure) { break }
        }
    }
    $run = [ordered]@{ id = $fixture.id; kind = $fixture.kind; mode = 'simulation'; workflowId = $definition.id; definitionVersion = $definition.version; definitionSnapshot = $definition; instanceId = $context.context.instanceId; at = $sample.clock; input = $fixture.input; status = $status; error = $failure; cancelledBefore = $cancelledBefore; trace = @($trace.ToArray()) }
    if ($status -ne $fixture.expect.status -or $trace.Count -ne $fixture.expect.stepCount) { throw "기대 상태 불일치: $($fixture.id) / $status / 단계 $($trace.Count) / $($failure | ConvertTo-Json -Compress)" }
    if ($fixture.expect.Contains('errorCode') -and $failure.code -ne $fixture.expect.errorCode) { throw "오류 코드 불일치: $($fixture.id)" }
    if ($fixture.expect.Contains('lastStepStatus') -and $trace[-1].response.status -ne $fixture.expect.lastStepStatus) { throw "마지막 단계 상태 불일치: $($fixture.id)" }
    if ($fixture.expect.Contains('values')) {
        foreach ($key in $fixture.expect.values.Keys) {
            if ((Get-MockReference $context $key) -cne $fixture.expect.values[$key]) { throw "기대 반환값 불일치: $($fixture.id) / $key" }
        }
    }
    # 그림 설명문은 업무 정의가 아니다. 2판에 1판 설명이 남지 않게 사본에서 제외한다.
    [void]$run.definitionSnapshot.Remove('presentation')
    $runs.Add($run); $runMap[$fixture.id] = $run
}

# 디자인 README의 색·글꼴·크기. 실제 앱에서는 공통 Theme/Ui를 사용한다.
$C = @{ Back='#F5F8FC'; Nav='#EDF3FA'; Card='#FFFFFF'; Border='#DCE3EC'; Input='#7F8C9C'; Button='#C4CEDA'; Text='#1D2B3A'; Sub='#56667A'; Accent='#2370C2'; AccentText='#1C5A9B'; Selected='#E0EBF7'; Warn='#8A5A08'; WarnBg='#FDF1DC'; Danger='#B83227'; DangerBg='#FBE7E4' }
$inv = [Globalization.CultureInfo]::InvariantCulture
$culture = [Globalization.CultureInfo]::GetCultureInfo('ko-KR')
$installed = @([Windows.Media.Fonts]::SystemFontFamilies | ForEach-Object Source)
$font = @('Pretendard', 'Noto Sans KR', 'Malgun Gothic' | Where-Object { $installed -contains $_ }) | Select-Object -First 1
if (-not $font) { $font = 'Malgun Gothic' }
$script:canvas = [Collections.Generic.List[object]]::new()
$script:brushes = @{}
function Get-Brush([string]$Color) {
    if (-not $script:brushes.ContainsKey($Color)) { $b = [Windows.Media.BrushConverter]::new().ConvertFromInvariantString($Color); $b.Freeze(); $script:brushes[$Color] = $b }
    $script:brushes[$Color]
}
function New-TextLayout([string]$Value, [double]$Size, [bool]$Bold, [string]$Color) {
    $weight = if ($Bold) { [Windows.FontWeights]::SemiBold } else { [Windows.FontWeights]::Normal }
    $face = [Windows.Media.Typeface]::new([Windows.Media.FontFamily]::new($font), [Windows.FontStyles]::Normal, $weight, [Windows.FontStretches]::Normal)
    [Windows.Media.FormattedText]::new($Value, $culture, [Windows.FlowDirection]::LeftToRight, $face, $Size, (Get-Brush $Color), 1.0)
}
function Add-Box([double]$X, [double]$Y, [double]$W, [double]$H, [string]$Fill, [string]$Stroke = '', [double]$Radius = 0) {
    $script:canvas.Add(@{ Kind='rect'; X=$X; Y=$Y; W=$W; H=$H; Fill=$Fill; Stroke=$Stroke; Radius=$Radius })
}
function Add-Label([double]$X, [double]$Y, [string]$Value, [double]$Size = 14, [string]$Color = $C.Text, [bool]$Bold = $false) {
    $layout = New-TextLayout $Value $Size $Bold $Color
    $script:canvas.Add(@{ Kind='text'; X=$X; Y=$Y; W=$layout.WidthIncludingTrailingWhitespace; H=$layout.Height; Value=$Value; Size=$Size; Color=$Color; Bold=$Bold; Baseline=$layout.Baseline })
}
function Get-WrappedLines([string]$Text, [double]$Width, [double]$Size = 14) {
    $lines = [Collections.Generic.List[string]]::new()
    foreach ($paragraph in ($Text -split "`n")) {
        $line = ''
        foreach ($token in [regex]::Matches($paragraph, '\S+\s*')) {
            $word = $token.Value
            if ($line -and (New-TextLayout ($line + $word) $Size $false $C.Text).Width -gt $Width) { $lines.Add($line.TrimEnd()); $line = '' }
            # 경로처럼 한 단어가 폭보다 길 때만 글자 단위로 줄을 바꾼다.
            foreach ($letter in $word.ToCharArray()) {
                $next = $line + $letter
                if ($line -and (New-TextLayout $next $Size $false $C.Text).Width -gt $Width) { $lines.Add($line.TrimEnd()); $line = [string]$letter } else { $line = $next }
            }
        }
        $lines.Add($line)
    }
    return ,$lines.ToArray()
}
function Add-Paragraph([double]$X, [double]$Y, [string]$Text, [double]$Width, [double]$Size = 14, [string]$Color = $C.Text) {
    $lines = Get-WrappedLines $Text $Width $Size
    foreach ($line in $lines) { Add-Label $X $Y $line $Size $Color; $Y += 24 }
    return $Y
}
function Add-Button([double]$X, [double]$Y, [double]$W, [string]$Text, [bool]$Primary = $false) {
    Add-Box $X $Y $W 40 $(if ($Primary) { $C.Accent } else { $C.Card }) $C.Button 8
    Add-Label ($X + 12) ($Y + 10) $Text 14 $(if ($Primary) { $C.Card } else { $C.Text })
}
function Add-Section([double]$Y, [string]$Title, [string[]]$Lines, [bool]$Warning = $false) {
    $wrapped = @($Lines | ForEach-Object { Get-WrappedLines $_ 384 } | ForEach-Object { $_ })
    $height = 80 + $wrapped.Count * 24
    Add-Box 24 $Y 432 $height $(if ($Warning) { $C.WarnBg } else { $C.Card }) $C.Border 12
    Add-Label 48 ($Y + 20) $Title 20 $C.Text $true
    $rowY = $Y + 60
    foreach ($line in $wrapped) { Add-Label 48 $rowY $line; $rowY += 24 }
    return ($Y + $height + 16)
}
function Get-InputLines($Run) {
    @(foreach ($field in $Run.definitionSnapshot.fields) {
        $value = $Run.input[$field.id]
        if ($value -is [array]) { $value = $value -join ', ' }
        if ($value -is [bool]) { $value = if ($value) { '확인함' } else { '미확인' } }
        if ($value -eq '') { $value = '(비어 있음)' }
        "$($field.label): $value"
    })
}
function Get-ResultLines($Run) {
    $lines = [Collections.Generic.List[string]]::new()
    $labels = @{ succeeded='모의 완료'; failed='실패'; skipped='조건 미충족 · 건너뜀'; cancelled='취소' }
    $lines.Add("전체 상태: $($labels[$Run.status]) · 실제 실행 없음")
    foreach ($item in $Run.trace) {
        $response = $item.response
        $lines.Add("$($apiMap[$item.action].title) → $($labels[$response.status])")
        if ($response.status -ne 'succeeded') { continue }
        $data = $response.data
        switch ($item.action) {
            'file.create' { $lines.Add($data.path); $lines.Add($data.text) }
            'record.append' {
                $lines.Add("기록 ID: $($data.recordId)")
                if ($data.collection -eq 'requests') { $lines.Add("$($data.values.title) / $($data.values.item) $($data.values.quantity)개") }
            }
            'network.ping' {
                foreach ($row in $data.rows) { $lines.Add("$($row.target): $(if ($row.status -eq 'success') { "$($row.latencyMs)ms · 성공" } else { '시간 초과' })") }
                $lines.Add("실패 개수: $($data.failedCount)")
            }
            'alarm.schedule' { $lines.Add("시각: $($data.at)"); $lines.Add("제목: $($data.title)"); $lines.Add("연결 업무: $($data.taskId)") }
        }
    }
    if ($Run.error) { $lines.Add("$($Run.error.code): $($Run.error.message)") }
    if ($Run.cancelledBefore) { $lines.Add("중단 위치: $($Run.cancelledBefore) 단계 시작 전") }
    return $lines.ToArray()
}

function Export-Scene([string]$Name, [int]$Width, [int]$Height, [double]$Scale = 1) {
    foreach ($shape in $script:canvas) {
        if ($shape.X -lt 0 -or $shape.Y -lt 0 -or $shape.X + $shape.W -gt $Width + 0.1 -or $shape.Y + $shape.H -gt $Height + 0.1) { throw "그림 경계 초과: $Name / $($shape | ConvertTo-Json -Compress)" }
    }
    $settings = [Xml.XmlWriterSettings]::new()
    $settings.Encoding = [Text.UTF8Encoding]::new($false); $settings.Indent = $true; $settings.NewLineChars = "`n"
    $writer = [Xml.XmlWriter]::Create((Join-Path $Config.OutputFolder "$Name.svg"), $settings)
    try {
        $writer.WriteStartElement('svg', 'http://www.w3.org/2000/svg')
        $writer.WriteAttributeString('width', [string]($Width * $Scale))
        $writer.WriteAttributeString('height', [string]($Height * $Scale))
        $writer.WriteAttributeString('viewBox', "0 0 $Width $Height")
        $writer.WriteAttributeString('role', 'img'); $writer.WriteAttributeString('aria-labelledby', 'title desc')
        $writer.WriteAttributeString('font-family', "$font, 'Noto Sans KR', 'Malgun Gothic', sans-serif")
        $writer.WriteStartElement('title'); $writer.WriteAttributeString('id', 'title'); $writer.WriteString('WorkTool 사용자 구성 업무 목업'); $writer.WriteEndElement()
        $writer.WriteStartElement('desc'); $writer.WriteAttributeString('id', 'desc'); $writer.WriteString('가짜 데이터로 계산한 구성, 사용, 모의 결과와 변경 비교. 실제 파일 생성, 알람, 서버 검사 없음.'); $writer.WriteEndElement()
        foreach ($shape in $script:canvas) {
            if ($shape.Kind -eq 'rect') {
                $writer.WriteStartElement('rect')
                $attrs = @{ x=$shape.X; y=$shape.Y; width=$shape.W; height=$shape.H; rx=$shape.Radius; fill=$shape.Fill; stroke=$(if ($shape.Stroke) { $shape.Stroke } else { 'none' }) }
            } else {
                $writer.WriteStartElement('text')
                $attrs = @{ x=$shape.X; y=$shape.Y + $shape.Baseline; fill=$shape.Color; 'font-size'=$shape.Size; 'font-weight'=$(if ($shape.Bold) { '600' } else { '400' }) }
            }
            foreach ($key in ($attrs.Keys | Sort-Object)) {
                $value = $attrs[$key]
                if ($value -is [ValueType]) { $value = [Convert]::ToString($value, $inv) }
                $writer.WriteAttributeString($key, [string]$value)
            }
            if ($shape.Kind -eq 'text') { $writer.WriteString($shape.Value) }
            $writer.WriteEndElement()
        }
        $writer.WriteEndElement()
    } finally { $writer.Dispose() }
    $visual = [Windows.Media.DrawingVisual]::new()
    $drawing = $visual.RenderOpen()
    try {
        foreach ($shape in $script:canvas) {
            if ($shape.Kind -eq 'rect') {
                $pen = if ($shape.Stroke) { [Windows.Media.Pen]::new((Get-Brush $shape.Stroke), 1) } else { $null }
                $drawing.DrawRoundedRectangle((Get-Brush $shape.Fill), $pen, [Windows.Rect]::new($shape.X, $shape.Y, $shape.W, $shape.H), $shape.Radius, $shape.Radius)
            } else { $drawing.DrawText((New-TextLayout $shape.Value $shape.Size $shape.Bold $shape.Color), [Windows.Point]::new($shape.X, $shape.Y)) }
        }
    } finally { $drawing.Close() }
    $renderScale = $Config.PngScale * $Scale
    $bitmap = [Windows.Media.Imaging.RenderTargetBitmap]::new([int]($Width * $renderScale), [int]($Height * $renderScale), 96 * $renderScale, 96 * $renderScale, [Windows.Media.PixelFormats]::Pbgra32)
    $bitmap.Render($visual)
    $encoder = [Windows.Media.Imaging.PngBitmapEncoder]::new(); $encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
    $stream = [IO.File]::Create((Join-Path $Config.OutputFolder "$Name.png"))
    try { $encoder.Save($stream) } finally { $stream.Dispose() }
    Write-Host "목업 생성: $Name.svg / .png"
}

# 전체 화면. 요청 정의와 동일한 입력에서 그린다.
$baseRequest = $runMap['request-base']
Add-Box 0 0 1280 672 $C.Back
Add-Box 0 40 68 608 $C.Nav
Add-Box 12 8 24 24 $C.Accent '' 8
Add-Label 17 8 'W' 16 $C.Card $true
Add-Label 48 10 'WorkTool' 16 $C.Text $true
Add-Label 976 12 '검토용 목업 · 실제 실행 없음' 12.5 $C.Sub
Add-Label 20 72 '홈' 14 $C.Sub
Add-Label 14 132 '일정' 14 $C.Sub
Add-Label 14 192 '메모' 14 $C.Sub
Add-Box 8 244 52 48 $C.Selected '' 8
Add-Box 8 256 3 24 $C.Accent
Add-Label 14 258 '업무' 14 $C.AccentText $true
Add-Label 14 600 '설정' 14 $C.Sub
Add-Box 68 40 252 608 $C.Card $C.Border
Add-Label 92 60 '내 업무' 20 $C.Text $true
Add-Box 92 100 204 40 $C.Back $C.Input 8
Add-Label 104 110 '업무 찾기' 14 $C.Sub
Add-Label 92 168 '⌄ 일반 업무' 14 $C.Sub $true
Add-Box 80 200 228 40 $C.Selected '' 8
Add-Box 80 212 3 16 $C.Accent
Add-Label 112 210 '요청 처리' 14 $C.AccentText $true
Add-Label 112 258 '자료 작성'
Add-Label 92 314 '⌄ 점검' 14 $C.Sub $true
Add-Label 112 362 '서버 점검'
Add-Label 92 440 '분류용 가상 폴더' 12.5 $C.Sub
Add-Label 92 464 '실제 폴더는 만들지 않습니다.' 12.5 $C.Sub
Add-Button 92 532 204 '폴더 추가'
Add-Button 92 584 204 '업무 만들기'
Add-Label 344 58 '일반 업무  /  요청 처리' 20 $C.Text $true
Add-Button 344 100 144 '새 업무 시작'
Add-Button 500 100 132 '구성 편집'
Add-Button 644 100 132 '실행 이력'
Add-Box 320 156 640 492 $C.Card $C.Border
Add-Label 344 172 '요청 내용' 20 $C.Text $true
Add-Label 344 212 '요청 번호 (필수)' 12.5 $C.Sub
Add-Label 644 212 '물품 (필수)' 12.5 $C.Sub
Add-Box 344 236 280 40 $C.Card $C.Input 8
Add-Box 644 236 292 40 $C.Card $C.Input 8
Add-Label 356 246 $baseRequest.input.requestNo
Add-Label 656 246 ($baseRequest.input.item + '  ⌄')
Add-Label 344 288 '요청 제목 (필수)' 12.5 $C.Sub
Add-Box 344 312 592 40 $C.Card $C.Input 8
Add-Label 356 322 $baseRequest.input.title
Add-Label 344 364 '수량 (필수)' 12.5 $C.Sub
Add-Label 644 364 '후속 확인 시각 (필수)' 12.5 $C.Sub
Add-Box 344 388 280 40 $C.Card $C.Input 8
Add-Box 644 388 292 40 $C.Card $C.Input 8
Add-Label 356 398 ([string]$baseRequest.input.quantity)
Add-Label 656 398 (([datetimeoffset]::Parse($baseRequest.input.followupAt)).ToString('yyyy-MM-dd HH:mm'))
Add-Label 344 452 '☑ 대여 가능 확인'
Add-Label 644 452 '☑ 반납 시각 확인'
Add-Box 344 496 592 56 $C.Back $C.Border 8
Add-Label 360 504 '처리 순서: 요청 기록 → 후속 알람' 14 $C.Text
Add-Label 360 528 '현재 그림의 결과는 시뮬레이션입니다.' 12.5 $C.Sub
Add-Button 344 584 156 '새 창에서 열기'
Add-Label 516 596 '메인과 함께 사용' 12.5 $C.Sub
Add-Button 816 584 120 '실행' $true
Add-Box 960 40 320 608 $C.Card $C.Border
Add-Label 984 60 '설명' 20 $C.Text $true
Add-Button 1136 52 120 '설명 접기'
Add-Label 984 124 '대여 확인 절차' 16 $C.Text $true
$null = Add-Paragraph 984 164 $baseRequest.definitionSnapshot.description 272
Add-Label 984 276 '확인할 것' 16 $C.Text $true
$null = Add-Paragraph 984 316 "1. 물품과 수량 확인`n2. 반납 시각 확인`n3. 후속 알람 시각 선택" 272
Add-Box 984 448 272 120 $C.Back $C.Border 8
Add-Label 1000 464 '후속 확인 시각' 14 $C.Text $true
$null = Add-Paragraph 1000 492 '호버·포커스 도움말: 이 시각에 요청을 다시 확인합니다.' 240 12.5 $C.Sub
Add-Box 68 648 1212 24 $C.Nav
Add-Label 84 651 '가짜 데이터 · 업무 구성 1판 · 아직 실행하지 않은 입력 화면' 12.5 $C.Sub
Export-Scene 'workflows-overview' 1280 672

# 같은 앱 틀에서 자료 작성의 구성 화면과 결과 화면을 따로 보여 준다.
$reportFrame = @($script:canvas | Where-Object { $_.X -lt 320 -or $_.Y -lt 40 -or $_.Y -ge 648 } | ForEach-Object { $_.Clone() })
function New-ReportFrame([string]$Title) {
    $script:canvas = [Collections.Generic.List[object]]::new()
    foreach ($original in $reportFrame) {
        $shape = $original.Clone()
        if ($shape.X -eq 80 -and $shape.Y -in @(200, 212)) { $shape.Y += 48 }
        if ($shape.Kind -eq 'text' -and $shape.Value -eq '요청 처리') { $shape.Color=$C.Text; $shape.Bold=$false }
        if ($shape.Kind -eq 'text' -and $shape.Value -eq '자료 작성') { $shape.Color=$C.AccentText; $shape.Bold=$true }
        if ($shape.Kind -eq 'text' -and $shape.Y -eq 651) { continue }
        $script:canvas.Add($shape)
    }
    Add-Label 344 58 $Title 20 $C.Text $true
    Add-Box 320 156 640 492 $C.Card $C.Border
    Add-Box 960 40 320 608 $C.Card $C.Border
    Add-Label 84 651 '가짜 데이터 · 자료 작성 1판 · 실제 파일 생성 없음' 12.5 $C.Sub
}
New-ReportFrame '자료 작성 / 구성 편집'
Add-Button 344 100 144 '입력 항목'
Add-Button 500 100 144 '출력 규칙'
Add-Box 512 142 120 3 $C.Accent
Add-Button 656 100 144 '창·설명'
Add-Label 344 172 '출력 규칙' 20 $C.Text $true
Add-Label 344 212 '작업 폴더' 12.5 $C.Sub
Add-Box 344 236 436 40 $C.Card $C.Input 8
Add-Label 356 246 '작업폴더 (선택한 경로의 예시)'
Add-Button 792 236 144 '폴더 선택'
Add-Label 344 288 '하위 폴더' 12.5 $C.Sub
Add-Box 344 312 592 40 $C.Card $C.Input 8
Add-Label 356 322 '[작업 폴더] / [기준 월]'
Add-Label 344 364 '파일명' 12.5 $C.Sub
Add-Box 344 388 592 40 $C.Card $C.Input 8
Add-Label 356 398 '[자료 제목]_[날짜]_[순번].txt'
Add-Label 344 440 '본문 양식' 12.5 $C.Sub
Add-Box 344 464 592 100 $C.Card $C.Input 8
$null = Add-Paragraph 356 472 "제목: [자료 제목]`n기준 월: [기준 월]`n요약: [요약]" 568
Add-Button 344 584 144 '값 연결'
Add-Button 500 584 144 '미리보기'
Add-Button 792 584 144 '구성 저장' $true
Add-Label 984 60 '단계 연결' 20 $C.Text $true
Add-Label 984 116 '1. 경로 조합'
Add-Label 984 156 '2. 본문 구성'
Add-Box 976 196 288 40 $C.Selected '' 8
Add-Label 984 206 '3. 새 파일 생성' 14 $C.AccentText $true
Add-Label 984 256 '4. 기록 추가'
Add-Button 984 308 272 '단계 추가'
Add-Label 984 376 '선택한 기능의 입력' 16 $C.Text $true
$null = Add-Paragraph 984 420 "저장 경로 ← 경로 조합 / 경로`n본문 ← 본문 구성 / 텍스트" 272
$null = Add-Paragraph 984 520 '앞 단계 결과를 골라 연결합니다. 이름과 규칙은 사용자가 바꿉니다.' 272 12.5 $C.Sub
Export-Scene 'workflows-builder' 1280 672

New-ReportFrame '자료 작성 / 사용 결과'
Add-Button 344 100 144 '입력값'
Add-Button 500 100 144 '결과'
Add-Box 512 142 120 3 $C.Accent
Add-Button 656 100 144 '실행 이력'
$report = $runMap['report-base']
$created = @($report.trace | Where-Object id -eq 'save')[0].response.data
Add-Label 344 172 '결과 미리보기' 20 $C.Text $true
Add-Box 344 212 592 40 $C.Selected '' 8
Add-Label 356 222 '모의 완료 · 실제 파일은 생성하지 않았습니다.' 14 $C.AccentText $true
Add-Label 344 276 '생성 예정 경로' 12.5 $C.Sub
$null = Add-Paragraph 344 304 $created.path 592
Add-Label 344 352 '파일 내용' 12.5 $C.Sub
Add-Box 344 380 592 164 $C.Back $C.Border 8
$null = Add-Paragraph 360 400 $created.text 560
Add-Button 344 584 144 '구성 편집'
Add-Button 500 584 144 '입력값 보기'
Add-Button 792 584 144 '새 업무 시작' $true
Add-Label 984 60 '실행 이력' 20 $C.Text $true
Add-Label 984 116 '자료 작성 · 정의 1판' 16 $C.Text $true
Add-Label 984 156 '2026-10-23 09:00 +09:00' 12.5 $C.Sub
$resultY = 220
foreach ($item in $report.trace) {
    Add-Label 984 $resultY ($apiMap[$item.action].title + ' · 모의 완료')
    $resultY += 48
}
$null = Add-Paragraph 984 452 '당시 입력값과 구성 1판을 함께 보관합니다. 구성 2판을 만들어도 이 결과는 바뀌지 않습니다.' 272
Export-Scene 'workflows-result' 1280 672

foreach ($definition in $sample.workflows) {
    $script:canvas = [Collections.Generic.List[object]]::new()
    # 배경 높이는 섹션 배치가 끝난 뒤 확정한다.
    Add-Box 0 0 480 1 $C.Back
    Add-Label 24 24 'WorkTool  /  사용자 구성 업무' 12.5 $C.Sub
    Add-Label 24 56 $definition.title 26 $C.Text $true
    Add-Label 24 100 '검토용 목업 · 가짜 데이터 · 실제 실행 없음' 12.5 $C.Sub
    $y = 140
    $y = Add-Section $y '1. 사용자가 만드는 구성' $definition.presentation.setup
    $base = $runMap["$($definition.id)-base"]
    $changed = $runMap["$($definition.id)-changed"]
    $edge = @($runs | Where-Object { $_.workflowId -eq $definition.id -and $_.kind -eq 'edge' })[0]
    $y = Add-Section $y '2. 사용 순서와 입력값' (@($definition.presentation.usage) + @('') + @(Get-InputLines $base))
    $y = Add-Section $y '3. 사용 결과 미리보기' (Get-ResultLines $base)
    $changeLines = @($definition.presentation.change, "정의 버전: $($base.definitionVersion) → $($changed.definitionVersion)")
    if ($definition.id -eq 'report') { $changeLines += @('예전 실행 이력은 1판 그대로 보존합니다.') }
    $changeLines += @(Get-ResultLines $changed)
    $y = Add-Section $y '4. 바꿔서 다시 사용하면' $changeLines
    $y = Add-Section $y '5. 실패·취소도 확인' (@($definition.presentation.edge) + @(Get-ResultLines $edge)) $true
    $apiLines = @(foreach ($step in $definition.steps) { "$($step.id): $($step.action) v$($step.version)" })
    $apiLines += '내부 API 이름·형식은 제안입니다.'
    $y = Add-Section $y '연결한 공통 기능' $apiLines
    Add-Label 24 $y '실행 데이터: workflows-results.json' 12.5 $C.Sub
    $height = [int]($y + 48)
    $script:canvas[0].H = $height
    Export-Scene "workflows-$($definition.id)" 480 $height 1.5
}
$output = [ordered]@{ isMock = $true; mockSchema = 1; mode = 'simulation'; source = 'workflows-sample.json'; clock = $sample.clock; checks = @{ cases = $runs.Count; expectations = 'passed'; shapeBounds = 'passed' }; runs = $runs.ToArray() }
function ConvertTo-StableMock($Value) {
    if ($Value -is [Collections.IDictionary]) {
        $stable = [ordered]@{}
        foreach ($key in ($Value.Keys | Sort-Object)) { $stable[$key] = ConvertTo-StableMock $Value[$key] }
        return $stable
    }
    if ($Value -is [array]) { return ,@($Value | ForEach-Object { ConvertTo-StableMock $_ }) }
    return $Value
}
(ConvertTo-StableMock $output) | ConvertTo-Json -Depth 70 | Set-Content -LiteralPath (Join-Path $Config.OutputFolder 'workflows-results.json') -Encoding utf8
Write-Host "검산 통과: $($runs.Count)개 사례 / 결과 JSON / SVG·PNG 6쌍"
