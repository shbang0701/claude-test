<#
.SYNOPSIS
    활성 작업 가짜 데이터에서 검토용 SVG와 PNG를 만든다.
.DESCRIPTION
    PowerShell 7과 Windows 기본 WPF 어셈블리만 쓴다.
    실제 창·탭을 조회하거나 화면을 캡처하지 않는다. 같은 가짜 데이터로 목업 파일만 다시 만든다.
.EXAMPLE
    pwsh -NoProfile -STA -File apps/WorkTool/design/build-active-tasks-mockups.ps1
#>
#Requires -Version 7.0
[CmdletBinding()]
param()

# ===== 설정 (목업 생성용, 앱 설정이 아님) =====
$Config = @{
    SamplePath = Join-Path $PSScriptRoot 'active-tasks-sample.json'
    OutputFolder = $PSScriptRoot
    PngScale = 2
}
# ===========================================

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase

$sample = Get-Content -LiteralPath $Config.SamplePath -Raw -Encoding utf8 | ConvertFrom-Json
if ($sample.isMock -ne $true) { throw 'isMock: true인 가짜 데이터만 목업에 사용할 수 있습니다.' }

# 디자인 README 「기준 값」. 실제 앱 모듈에서는 Theme/Ui 리소스를 사용한다.
$C = @{
    Back = '#F5F8FC'; Nav = '#EDF3FA'; Card = '#FFFFFF'; Border = '#DCE3EC'
    Divider = '#EEF2F7'; Input = '#7F8C9C'; Button = '#C4CEDA'
    Text = '#1D2B3A'; Sub = '#56667A'; NavText = '#2E3E50'
    Accent = '#2370C2'; AccentText = '#1C5A9B'; Selected = '#E0EBF7'
    Warn = '#8A5A08'; WarnBg = '#FDF1DC'; Danger = '#B83227'; DangerBg = '#FBE7E4'
}
$inv = [Globalization.CultureInfo]::InvariantCulture
$culture = [Globalization.CultureInfo]::GetCultureInfo('ko-KR')
$installed = @([Windows.Media.Fonts]::SystemFontFamilies | ForEach-Object Source)
$font = @('Pretendard', 'Noto Sans KR', 'Malgun Gothic' | Where-Object { $installed -contains $_ }) | Select-Object -First 1
if (-not $font) { $font = 'Malgun Gothic' }
$iconFont = if ($installed -contains 'Segoe Fluent Icons') { 'Segoe Fluent Icons' } else { 'Segoe MDL2 Assets' }
$script:canvas = [Collections.Generic.List[object]]::new()
$script:brushes = @{}

function Get-Brush([string]$Color) {
    if (-not $script:brushes.ContainsKey($Color)) {
        $b = [Windows.Media.BrushConverter]::new().ConvertFromInvariantString($Color)
        $b.Freeze()
        $script:brushes[$Color] = $b
    }
    $script:brushes[$Color]
}

function New-TextLayout([string]$Value, [double]$Size, [bool]$Bold, [string]$Color, [string]$Family = $font) {
    $weight = if ($Bold) { [Windows.FontWeights]::SemiBold } else { [Windows.FontWeights]::Normal }
    $face = [Windows.Media.Typeface]::new([Windows.Media.FontFamily]::new($Family), [Windows.FontStyles]::Normal, $weight, [Windows.FontStretches]::Normal)
    [Windows.Media.FormattedText]::new($Value, $culture, [Windows.FlowDirection]::LeftToRight, $face, $Size, (Get-Brush $Color), 1.0)
}

function Add-Box([double]$X, [double]$Y, [double]$W, [double]$H, [string]$Fill, [string]$Stroke = '', [double]$Radius = 0, [double]$Thickness = 1) {
    $script:canvas.Add(@{ Kind = 'rect'; X = $X; Y = $Y; W = $W; H = $H; Fill = $Fill; Stroke = $Stroke; Radius = $Radius; Thickness = $Thickness })
}
function Add-Label([double]$X, [double]$Y, [string]$Value, [double]$Size = 14, [string]$Color = $C.Text, [bool]$Bold = $false) {
    $layout = New-TextLayout $Value $Size $Bold $Color
    $script:canvas.Add(@{ Kind = 'text'; X = $X; Y = $Y; Value = $Value; Size = $Size; Color = $Color; Bold = $Bold; Baseline = $layout.Baseline })
}
function Add-Path([string]$Data, [string]$Fill, [string]$Stroke = '', [double]$Thickness = 1) {
    $script:canvas.Add(@{ Kind = 'path'; Data = $Data; Fill = $Fill; Stroke = $Stroke; Thickness = $Thickness })
}
function Add-Icon([double]$X, [double]$Y, [int]$Code, [double]$Size = 20, [string]$Color = $C.Sub) {
    # Windows 기본 아이콘을 경로로 바꾼다. 폰에 Segoe 글꼴이 없어도 같은 모양으로 보인다.
    $layout = New-TextLayout ([string][char]$Code) $Size $false $Color $iconFont
    $geometry = $layout.BuildGeometry([Windows.Point]::new($X, $Y))
    $data = $geometry.GetOutlinedPathGeometry().ToString($inv) -replace '^F[01]\s*', ''
    Add-Path $data $Color
}
function Add-Check([double]$X, [double]$Y, [string]$State) {
    if ($State -eq 'disabled') {
        Add-Box $X $Y 16 16 $C.Divider $C.Border 4
        return
    }
    $fill = if ($State -eq 'off') { $C.Card } else { $C.Accent }
    $stroke = if ($State -eq 'off') { $C.Input } else { $C.Accent }
    Add-Box $X $Y 16 16 $fill $stroke 4
    if ($State -eq 'mixed') { Add-Box ($X + 4) ($Y + 7) 8 2 $C.Card }
    if ($State -eq 'on') {
        Add-Path ('M {0} {1} L {2} {3} L {4} {5}' -f ($X + 3), ($Y + 8), ($X + 7), ($Y + 12), ($X + 13), ($Y + 4)) '' $C.Card 2
    }
}
function Add-Button([double]$X, [double]$Y, [double]$W, [string]$Text, [string]$Style = 'default', [int]$Icon = 0) {
    $fill = if ($Style -eq 'primary') { $C.Accent } elseif ($Style -eq 'danger-fill') { $C.DangerBg } else { $C.Card }
    $color = if ($Style -eq 'primary') { $C.Card } elseif ($Style -like 'danger*') { $C.Danger } else { $C.Text }
    $stroke = if ($Style -eq 'primary') { $C.Accent } else { $C.Button }
    Add-Box $X $Y $W 40 $fill $stroke 8
    if ($Icon) {
        Add-Icon ($X + 12) ($Y + 10) $Icon 20 $color
        Add-Label ($X + 40) ($Y + 10) $Text 14 $color
    } else {
        $tw = (New-TextLayout $Text 14 $false $color).Width
        Add-Label ($X + ($W - $tw) / 2) ($Y + 10) $Text 14 $color
    }
}
function Add-Crop([double]$X, [double]$Y, [double]$W, [double]$H, [double]$SourceX, [double]$SourceY, [double]$SourceW, [double]$SourceH, $Shapes) {
    $script:canvas.Add(@{ Kind = 'crop'; X = $X; Y = $Y; W = $W; H = $H; SX = $SourceX; SY = $SourceY; SW = $SourceW; SH = $SourceH; Shapes = $Shapes })
}

$selected = @($sample.items | Where-Object checked)
$selectedWindows = @($selected | Where-Object kind -eq 'window').Count
$selectedTabs = @($selected | Where-Object kind -eq 'tab').Count
$windowCount = @($sample.items | ForEach-Object { '{0}-{1}' -f $_.app, $_.window } | Sort-Object -Unique).Count
$tabCount = @($sample.items | Where-Object kind -eq 'tab').Count
$unavailable = @($sample.items | Where-Object kind -eq 'unavailable').Count

# 전체 화면: WPF 앱의 기준 창 1280 × 672.
$main = $script:canvas
Add-Box 0 0 1280 672 $C.Back $C.Border 8
Add-Box 0 40 232 632 $C.Nav
Add-Box 16 10 24 24 $C.Accent '' 8
Add-Label 22 11 'W' 16 $C.Card $true
Add-Label 52 10 'WorkTool' 16 $C.Text $true
Add-Label 1056 12 '목업 · 가짜 데이터' 12.5 $C.Sub
Add-Icon 1156 10 0xE921 16
Add-Icon 1200 10 0xE922 16
Add-Icon 1244 10 0xE8BB 16
Add-Label 24 64 '10월 23일 금요일' 12.5 $C.Sub
$navigation = @(
    @('홈', 0xE80F), @('일정', 0xE787), @('근태', 0xE823),
    @('메모', 0xE70B), @('문서', 0xE8B7), @('활성 작업', 0xE7C4)
)
for ($i = 0; $i -lt $navigation.Count; $i++) {
    $ny = 96 + $i * 48
    $isActive = $i -eq 5
    if ($isActive) {
        Add-Box 12 $ny 208 40 $C.Selected '' 8
        Add-Box 12 ($ny + 12) 3 16 $C.Accent '' 1.5
    }
    $nc = if ($isActive) { $C.AccentText } else { $C.NavText }
    Add-Icon 28 ($ny + 10) $navigation[$i][1] 20 $nc
    Add-Label 60 ($ny + 10) $navigation[$i][0] 14 $nc $isActive
}
Add-Icon 28 620 0xE713 20
Add-Label 56 620 '설정' 14 $C.Sub
Add-Icon 128 620 0xE897 20
Add-Label 156 620 '도움말' 14 $C.Sub

Add-Label 256 56 '활성 작업' 26 $C.Text $true
Add-Label 256 92 ('프로그램 {0}개 · 창 {1}개 · 확인한 탭 {2}개' -f $sample.groups.Count, $windowCount, $tabCount) 12.5 $C.Sub
Add-Box 256 120 360 40 $C.Card $C.Input 8
Add-Icon 268 130 0xE721 20
Add-Label 300 130 '이름으로 찾기' 14 $C.Sub
Add-Button 628 120 160 '모든 프로그램'
Add-Icon 764 134 0xE70D 12
Add-Button 956 120 136 '새로 고침' 'primary' 0xE72C
Add-Button 1104 120 152 ('선택 닫기 {0}' -f $selected.Count) 'danger' 0xE8BB

Add-Box 256 176 1000 448 $C.Card $C.Border 12
Add-Label 268 182 '선택' 12.5 $C.Sub $true
Add-Label 316 182 '이름' 14 $C.Sub $true
Add-Label 852 182 '종류' 14 $C.Sub $true
Add-Label 932 182 '창' 14 $C.Sub $true
Add-Label 1028 182 '상태' 14 $C.Sub $true
Add-Box 256 207 1000 1 $C.Border

$rowY = 208
foreach ($group in $sample.groups) {
    $rows = @($sample.items | Where-Object app -eq $group.id)
    $available = @($rows | Where-Object canClose)
    $checked = @($available | Where-Object checked).Count
    $checkState = if ($checked -eq 0) { 'off' } elseif ($checked -eq $available.Count) { 'on' } else { 'mixed' }
    Add-Box 257 $rowY 998 32 $C.Nav
    Add-Check 272 ($rowY + 8) $checkState
    Add-Icon 312 ($rowY + 10) $(if ($group.expanded) { 0xE70D } else { 0xE76C }) 12
    Add-Label 336 ($rowY + 6) $group.name 14 $C.Text $true
    $windows = @($rows | Select-Object -ExpandProperty window -Unique).Count
    $tabs = @($rows | Where-Object kind -eq 'tab').Count
    $missing = @($rows | Where-Object kind -eq 'unavailable').Count
    $summary = '창 {0}개' -f $windows
    if ($tabs) { $summary += ' · 확인한 탭 {0}개' -f $tabs }
    if ($missing) { $summary += ' · 읽기 제한 {0}개' -f $missing }
    Add-Label 556 ($rowY + 7) $summary 12.5 $C.Sub
    $rowY += 32
    if (-not $group.expanded) { continue }
    foreach ($item in $rows) {
        if ($item.checked) {
            Add-Box 257 $rowY 998 32 $C.Selected
            Add-Box 257 $rowY 3 32 $C.Accent
        }
        $checkState = if (-not $item.canClose) { 'disabled' } elseif ($item.checked) { 'on' } else { 'off' }
        Add-Check 272 ($rowY + 8) $checkState
        Add-Label 316 ($rowY + 6) $item.title
        Add-Label 852 ($rowY + 6) $(if ($item.kind -eq 'tab') { '탭' } else { '창' }) 14 $C.Sub
        Add-Label 932 ($rowY + 6) ('창 {0}' -f $item.window) 14 $C.Sub
        if ($item.kind -eq 'unavailable') {
            Add-Icon 1028 ($rowY + 8) 0xE7BA 16 $C.Warn
            Add-Label 1052 ($rowY + 7) $item.status 12.5 $C.Warn
        } elseif ($item.status) { Add-Label 1028 ($rowY + 7) $item.status 12.5 $C.Sub }
        Add-Box 260 ($rowY + 31) 995 1 $C.Divider
        $rowY += 32
    }
}
if ($rowY -gt 624) { throw '기준 창에서 표 높이가 넘칩니다. 데이터나 배치를 확인하세요.' }
Add-Label 256 632 ('{0}개 선택 · 창 {1}개, 탭 {2}개' -f $selected.Count, $selectedWindows, $selectedTabs) 12.5 $C.AccentText
Add-Label 828 632 ('탭 확인 불가 {0}개 · 마지막 수집 {1}' -f $unavailable, ([datetimeoffset]$sample.capturedAt).ToString('HH:mm', $inv)) 12.5 $C.Sub

# 확인창: 선택한 실제 목업 행에서 목록과 개수를 만든다.
$script:canvas = [Collections.Generic.List[object]]::new()
$dialog = $script:canvas
Add-Box 0 0 480 464 $C.Card $C.Border 12
Add-Label 24 24 ('선택한 {0}개 항목을 닫을까요?' -f $selected.Count) 20 $C.Text $true
Add-Label 24 60 ('창 {0}개 · 탭 {1}개' -f $selectedWindows, $selectedTabs) 14 $C.Sub
Add-Box 24 88 432 240 $C.Back $C.Border 8
$dy = 96
foreach ($item in $selected) {
    Add-Label 40 $dy $item.title
    $appName = ($sample.groups | Where-Object id -eq $item.app).name -replace '^Google ', '' -replace '^Microsoft ', ''
    $kindName = if ($item.kind -eq 'tab') { '탭' } else { '창' }
    Add-Label 40 ($dy + 20) ('{0} · 창 {1} · {2}' -f $appName, $item.window, $kindName) 12.5 $C.Sub
    $dy += 44
}
Add-Label 24 348 '저장 확인은 원래 프로그램에서 처리하세요.' 12.5 $C.Sub
Add-Label 24 370 '닫기 요청 뒤 결과를 다시 확인합니다.' 12.5 $C.Sub
Add-Button 248 400 104 '닫기' 'danger-fill'
Add-Box 358 398 100 44 $C.Card $C.Text 10 2
Add-Button 360 400 96 '취소'

# 폰 열람용: 같은 벡터 장면의 필요한 부분을 확대한다.
$script:canvas = [Collections.Generic.List[object]]::new()
$detail = $script:canvas
Add-Box 0 0 720 1776 $C.Back
Add-Label 24 24 '활성 작업 · 확대해서 보기' 26 $C.Text $true
Add-Label 24 68 '검토용 가짜 데이터 · 앱 화면의 일부를 나눠 확대한 그림' 14 $C.Sub
Add-Label 24 108 '01  프로그램별 묶음과 체크' 16 $C.Text $true
Add-Crop 24 144 672 (192 * 672 / 464) 256 176 464 192 $main
Add-Label 24 480 '02  다른 창의 같은 제목 구별' 16 $C.Text $true
Add-Crop 24 516 672 64 916 176 336 32 $main
Add-Crop 24 580 672 256 916 240 336 128 $main
Add-Label 24 868 '03  읽지 못한 Edge 창도 목록에 표시' 16 $C.Text $true
Add-Label 24 904 '해당 브라우저 창을 확인한 뒤 새로 고침합니다.' 14 $C.Sub
Add-Crop 24 940 672 64 916 464 336 32 $main
Add-Label 24 1044 '04  선택한 대상만 확인하고 닫기' 16 $C.Text $true
Add-Crop 24 1088 672 649.6 0 0 480 464 $dialog

function Write-SvgShapes($Writer, $Shapes) {
    foreach ($s in $Shapes) {
        $tag = switch ($s.Kind) { rect { 'rect' } text { 'text' } path { 'path' } crop { 'svg' } }
        $Writer.WriteStartElement($tag, 'http://www.w3.org/2000/svg')
        $attrs = @{}
        switch ($s.Kind) {
            rect { $attrs = @{ x = $s.X; y = $s.Y; width = $s.W; height = $s.H; rx = $s.Radius; fill = $s.Fill; stroke = $s.Stroke; 'stroke-width' = $s.Thickness } }
            text { $attrs = @{ x = $s.X; y = $s.Y + $s.Baseline; fill = $s.Color; 'font-size' = $s.Size; 'font-weight' = $(if ($s.Bold) { '600' } else { '400' }) } }
            path { $attrs = @{ d = $s.Data; fill = $s.Fill; stroke = $s.Stroke; 'stroke-width' = $s.Thickness; 'stroke-linecap' = 'round'; 'stroke-linejoin' = 'round' } }
            crop { $attrs = @{ x = $s.X; y = $s.Y; width = $s.W; height = $s.H; viewBox = ($s.SX, $s.SY, $s.SW, $s.SH | ForEach-Object { $_.ToString($inv) }) -join ' '; overflow = 'hidden' } }
        }
        foreach ($name in @($attrs.Keys | Sort-Object)) {
            $value = $attrs[$name]
            if ($value -is [double] -or $value -is [int]) { $value = $value.ToString($inv) }
            if (($name -eq 'fill' -or $name -eq 'stroke') -and -not $value) { $value = 'none' }
            $Writer.WriteAttributeString($name, [string]$value)
        }
        if ($s.Kind -eq 'text') { $Writer.WriteString($s.Value) }
        if ($s.Kind -eq 'crop') { Write-SvgShapes $Writer $s.Shapes }
        $Writer.WriteEndElement()
    }
}

function Draw-Shapes($Context, $Shapes) {
    foreach ($s in $Shapes) {
        switch ($s.Kind) {
            rect {
                $fill = if ($s.Fill) { Get-Brush $s.Fill } else { $null }
                $pen = if ($s.Stroke) { [Windows.Media.Pen]::new((Get-Brush $s.Stroke), $s.Thickness) } else { $null }
                $Context.DrawRoundedRectangle($fill, $pen, [Windows.Rect]::new($s.X, $s.Y, $s.W, $s.H), $s.Radius, $s.Radius)
            }
            text {
                $layout = New-TextLayout $s.Value $s.Size $s.Bold $s.Color
                $Context.DrawText($layout, [Windows.Point]::new($s.X, $s.Y))
            }
            path {
                $fill = if ($s.Fill) { Get-Brush $s.Fill } else { $null }
                $pen = if ($s.Stroke) { [Windows.Media.Pen]::new((Get-Brush $s.Stroke), $s.Thickness) } else { $null }
                $Context.DrawGeometry($fill, $pen, [Windows.Media.Geometry]::Parse($s.Data))
            }
            crop {
                $scale = $s.W / $s.SW
                $Context.PushClip([Windows.Media.RectangleGeometry]::new([Windows.Rect]::new($s.X, $s.Y, $s.W, $s.H)))
                $matrix = [Windows.Media.Matrix]::new($scale, 0, 0, $scale, $s.X - $s.SX * $scale, $s.Y - $s.SY * $scale)
                $Context.PushTransform([Windows.Media.MatrixTransform]::new($matrix))
                Draw-Shapes $Context $s.Shapes
                $Context.Pop()
                $Context.Pop()
            }
        }
    }
}

function Export-Scene([string]$Name, [int]$Width, [int]$Height, $Shapes, [string]$Title) {
    $svg = Join-Path $Config.OutputFolder "$Name.svg"
    $png = Join-Path $Config.OutputFolder "$Name.png"
    $settings = [Xml.XmlWriterSettings]::new()
    $settings.Encoding = [Text.UTF8Encoding]::new($false)
    $settings.Indent = $true
    $settings.NewLineChars = "`n"
    $writer = [Xml.XmlWriter]::Create($svg, $settings)
    try {
        $writer.WriteStartElement('svg', 'http://www.w3.org/2000/svg')
        $writer.WriteAttributeString('width', [string]$Width)
        $writer.WriteAttributeString('height', [string]$Height)
        $writer.WriteAttributeString('viewBox', "0 0 $Width $Height")
        $writer.WriteAttributeString('role', 'img')
        $writer.WriteAttributeString('aria-labelledby', 'title desc')
        $writer.WriteAttributeString('font-family', "$font, 'Noto Sans KR', 'Malgun Gothic', sans-serif")
        $writer.WriteStartElement('title'); $writer.WriteAttributeString('id', 'title'); $writer.WriteString($Title); $writer.WriteEndElement()
        $writer.WriteStartElement('desc'); $writer.WriteAttributeString('id', 'desc')
        $writer.WriteString('실제 창을 수집하지 않은 가짜 데이터 목업. 프로그램별 표에서 창 2개와 탭 3개를 선택했습니다. Edge 창 하나는 탭 확인 불가 상태입니다.')
        $writer.WriteEndElement()
        Write-SvgShapes $writer $Shapes
        $writer.WriteEndElement()
    } finally { $writer.Dispose() }

    $visual = [Windows.Media.DrawingVisual]::new()
    $context = $visual.RenderOpen()
    try { Draw-Shapes $context $Shapes } finally { $context.Close() }
    $scale = $Config.PngScale
    $bitmap = [Windows.Media.Imaging.RenderTargetBitmap]::new($Width * $scale, $Height * $scale, 96 * $scale, 96 * $scale, [Windows.Media.PixelFormats]::Pbgra32)
    $bitmap.Render($visual)
    $encoder = [Windows.Media.Imaging.PngBitmapEncoder]::new()
    $encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
    $stream = [IO.File]::Create($png)
    try { $encoder.Save($stream) } finally { $stream.Dispose() }
    Write-Host "만듦: $svg"
    Write-Host "만듦: $png"
}

Export-Scene 'active-tasks' 1280 672 $main 'WorkTool 활성 작업 — 전체 화면 목업'
Export-Scene 'active-tasks-detail' 720 1776 $detail 'WorkTool 활성 작업 — 폰 열람용 확대본'
