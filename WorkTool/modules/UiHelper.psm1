# 화면 공통: 색, 글꼴, 크기 배율, 버튼, 왼쪽 메뉴 버튼.
# 모든 화면 모듈이 이걸 써서 모양을 맞춘다. 색이나 글꼴을 바꿀 때는 이 파일만 고친다.

$ErrorActionPreference = 'Stop'

function ConvertTo-UiColor([string]$Hex) { [System.Drawing.ColorTranslator]::FromHtml($Hex) }

# 하늘색 파스텔 테마
$script:theme = @{
    Back            = ConvertTo-UiColor '#F4F9FE'   # 창 바탕
    Card            = [System.Drawing.Color]::White # 내용 영역
    Sidebar         = ConvertTo-UiColor '#E3F0FB'   # 왼쪽 메뉴
    SidebarHover    = ConvertTo-UiColor '#D3E7F8'
    SidebarSelected = ConvertTo-UiColor '#C3DEF5'
    Accent          = ConvertTo-UiColor '#2F80C9'   # 주 버튼, 선택 표시
    AccentHover     = ConvertTo-UiColor '#2A72B5'
    Text            = ConvertTo-UiColor '#23364A'
    SubText         = ConvertTo-UiColor '#6B7C8F'
    Border          = ConvertTo-UiColor '#CFE0F0'
}

# 설치된 글꼴 중 앞에서부터 처음 찾은 것을 쓴다. 맑은 고딕은 윈도우 기본이라 항상 있다.
$fontCandidates = @(
    'Pretendard', 'SUIT', 'Noto Sans KR', 'NanumSquare Neo', 'NanumBarunGothic', 'NanumGothic', 'Malgun Gothic'
)
function Find-UiFontName([string[]]$Candidates) {
    foreach ($name in $Candidates) {
        try {
            [System.Drawing.FontFamily]::new($name).Dispose()
            return $name
        } catch { }
    }
    $null
}
$script:fontName = Find-UiFontName $fontCandidates
if (-not $script:fontName) { $script:fontName = 'Malgun Gothic' }

# 메뉴 아이콘용 윈도우 기본 아이콘 글꼴 (Windows 11 / 10)
$iconFontName = Find-UiFontName @('Segoe Fluent Icons', 'Segoe MDL2 Assets')
$script:iconFont = if ($iconFontName) { [System.Drawing.Font]::new($iconFontName, 13) }

# 화면 배율(100% = 1.0). 픽셀로 정하는 크기는 Get-UiSize로 맞춘다.
$g = [System.Drawing.Graphics]::FromHwnd([IntPtr]::Zero)
$script:scale = $g.DpiX / 96
$g.Dispose()

function Get-UiTheme { $script:theme }

function Get-UiFontName { $script:fontName }

function Get-UiFont([float]$Size = 10, [switch]$Bold) {
    $style = if ($Bold) { [System.Drawing.FontStyle]::Bold } else { [System.Drawing.FontStyle]::Regular }
    [System.Drawing.Font]::new($script:fontName, $Size, $style)
}

function Get-UiSize([int]$Pixels) {
    [int][Math]::Round($Pixels * $script:scale)
}

function New-UiPadding([int]$Left, [int]$Top = $Left, [int]$Right = $Left, [int]$Bottom = $Top) {
    [System.Windows.Forms.Padding]::new((Get-UiSize $Left), (Get-UiSize $Top), (Get-UiSize $Right), (Get-UiSize $Bottom))
}

# 일반 버튼. -Primary면 강조색.
function New-UiButton([string]$Text, [switch]$Primary) {
    $b = [System.Windows.Forms.Button]::new()
    $b.Text = $Text
    $b.AutoSize = $true
    $b.FlatStyle = 'Flat'
    $b.Cursor = 'Hand'
    $b.Padding = New-UiPadding 8 2
    $b.Margin = New-UiPadding 0 0 6 0
    if ($Primary) {
        $b.BackColor = $script:theme.Accent
        $b.ForeColor = [System.Drawing.Color]::White
        $b.FlatAppearance.BorderColor = $script:theme.Accent
        $b.FlatAppearance.MouseOverBackColor = $script:theme.AccentHover
    } else {
        $b.BackColor = $script:theme.Card
        $b.ForeColor = $script:theme.Text
        $b.FlatAppearance.BorderColor = $script:theme.Border
        $b.FlatAppearance.MouseOverBackColor = $script:theme.SidebarHover
    }
    $b
}

$script:navFont = Get-UiFont 10.5
$script:navFontBold = Get-UiFont 10.5 -Bold

# 왼쪽 메뉴 버튼. 아이콘과 글자는 Paint에서 직접 그린다.
# $Icon은 아이콘 글꼴의 글자 코드 (예: 0xE80F 홈). 상태는 Tag(Text, Icon, Selected, Index)에 둔다.
function New-UiNavButton([string]$Text, [int]$Icon) {
    $b = [System.Windows.Forms.Button]::new()
    $b.Dock = 'Top'
    $b.Height = Get-UiSize 42
    $b.FlatStyle = 'Flat'
    $b.FlatAppearance.BorderSize = 0
    $b.Font = $script:navFont
    $b.ForeColor = $script:theme.Text
    $b.Cursor = 'Hand'
    $b.TabStop = $false
    $b.Tag = @{ Text = $Text; Icon = $(if ($Icon) { [string][char]$Icon } else { '' }); Selected = $false; Index = -1 }
    Set-UiNavSelected $b $false

    $b.add_Paint({
        param($s, $e)
        $t = $s.Tag
        $h = $s.Height
        $iconWidth = Get-UiSize 46
        if ($t.Selected) {
            $brush = [System.Drawing.SolidBrush]::new($script:theme.Accent)
            $e.Graphics.FillRectangle($brush, 0, (Get-UiSize 9), (Get-UiSize 4), $h - (Get-UiSize 18))
            $brush.Dispose()
        }
        if ($t.Icon -and $script:iconFont) {
            $r = [System.Drawing.Rectangle]::new(0, 0, $iconWidth, $h)
            [System.Windows.Forms.TextRenderer]::DrawText($e.Graphics, $t.Icon, $script:iconFont, $r, $s.ForeColor,
                [System.Windows.Forms.TextFormatFlags]'HorizontalCenter, VerticalCenter, NoPrefix')
        }
        # 메뉴를 접어서 폭이 좁으면 글자는 그리지 않는다
        $r = [System.Drawing.Rectangle]::new($iconWidth, 0, $s.Width - $iconWidth, $h)
        if ($r.Width -gt (Get-UiSize 24)) {
            [System.Windows.Forms.TextRenderer]::DrawText($e.Graphics, $t.Text, $s.Font, $r, $s.ForeColor,
                [System.Windows.Forms.TextFormatFlags]'Left, VerticalCenter, NoPrefix, EndEllipsis')
        }
    })
    $b
}

function Set-UiNavSelected($Button, [bool]$Selected) {
    $Button.Tag.Selected = $Selected
    if ($Selected) {
        $Button.BackColor = $script:theme.SidebarSelected
        $Button.FlatAppearance.MouseOverBackColor = $script:theme.SidebarSelected
        $Button.Font = $script:navFontBold
    } else {
        $Button.BackColor = $script:theme.Sidebar
        $Button.FlatAppearance.MouseOverBackColor = $script:theme.SidebarHover
        $Button.Font = $script:navFont
    }
    $Button.FlatAppearance.MouseDownBackColor = $script:theme.SidebarSelected
    $Button.Invalidate()
}

# 창·트레이에 쓰는 앱 아이콘. 파일 없이 강조색 둥근 사각형에 W를 그려서 만든다.
function New-UiAppIcon {
    $size = 32
    $bmp = [System.Drawing.Bitmap]::new($size, $size)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $g.TextRenderingHint = 'AntiAliasGridFit'
    $path = [System.Drawing.Drawing2D.GraphicsPath]::new()
    $d = 12
    $path.AddArc(0, 0, $d, $d, 180, 90)
    $path.AddArc($size - $d - 1, 0, $d, $d, 270, 90)
    $path.AddArc($size - $d - 1, $size - $d - 1, $d, $d, 0, 90)
    $path.AddArc(0, $size - $d - 1, $d, $d, 90, 90)
    $path.CloseFigure()
    $brush = [System.Drawing.SolidBrush]::new($script:theme.Accent)
    $g.FillPath($brush, $path)
    $font = [System.Drawing.Font]::new('Segoe UI', 20, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
    $format = [System.Drawing.StringFormat]::new()
    $format.Alignment = 'Center'
    $format.LineAlignment = 'Center'
    $g.DrawString('W', $font, [System.Drawing.Brushes]::White, [System.Drawing.RectangleF]::new(0, 1, $size, $size), $format)
    $g.Dispose()
    [System.Drawing.Icon]::FromHandle($bmp.GetHicon())
}

# 입력 칸 왼쪽에 붙는 이름표
function New-UiFieldLabel([string]$Text) {
    $l = [System.Windows.Forms.Label]::new()
    $l.Text = $Text
    $l.AutoSize = $true
    $l.Anchor = 'Left'
    $l.ForeColor = $script:theme.SubText
    $l.Margin = New-UiPadding 0 4 10 4
    $l
}

# 날짜·시각 입력. $Format 예: 'yyyy-MM-dd (ddd)', 'HH:mm' (시각이면 -Time)
function New-UiDatePicker([string]$Format, [switch]$Time) {
    $p = [System.Windows.Forms.DateTimePicker]::new()
    $p.Format = 'Custom'
    $p.CustomFormat = $Format
    $p.ShowUpDown = [bool]$Time
    $p.Dock = 'Fill'
    $p.Margin = New-UiPadding 0 4 0 4
    $p
}

# 표 형태 목록. $Columns는 @(@('이름', 폭), ...). 마지막 칸이 남는 폭을 채운다.
function New-UiListView([object[]]$Columns) {
    $list = [System.Windows.Forms.ListView]::new()
    $list.Dock = 'Fill'
    $list.Margin = New-UiPadding 0
    $list.View = 'Details'
    $list.FullRowSelect = $true
    $list.HideSelection = $false
    $list.MultiSelect = $false
    $list.HeaderStyle = 'Nonclickable'
    $list.ForeColor = $script:theme.Text
    foreach ($c in $Columns) { [void]$list.Columns.Add($c[0], (Get-UiSize $c[1])) }
    $list.add_Resize({
        param($s, $e)
        $c = $s.Columns
        $used = 0
        for ($i = 0; $i -lt $c.Count - 1; $i++) { $used += $c[$i].Width }
        $c[$c.Count - 1].Width = [Math]::Max((Get-UiSize 60), $s.ClientSize.Width - $used - 2)
    })
    $list
}

function Show-UiMessage([string]$Text, [string]$Title = 'WorkTool') {
    [void][System.Windows.Forms.MessageBox]::Show($Text, $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
}

# 예/아니요를 묻는다. 예면 $true.
function Confirm-UiYesNo([string]$Text, [string]$Title = 'WorkTool') {
    $answer = [System.Windows.Forms.MessageBox]::Show($Text, $Title,
        [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    $answer -eq [System.Windows.Forms.DialogResult]::Yes
}

Export-ModuleMember -Function Get-UiTheme, Get-UiFontName, Get-UiFont, Get-UiSize, New-UiPadding, New-UiButton,
    New-UiNavButton, Set-UiNavSelected, New-UiAppIcon, New-UiFieldLabel, New-UiDatePicker, New-UiListView,
    Show-UiMessage, Confirm-UiYesNo
