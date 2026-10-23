# WorkTool — 파일 하나로 묶은 버전 (자동 생성)
# 이 파일은 직접 고치지 않는다. apps/WorkTool의 원본을 고친 뒤 build/Build-Bundle.ps1로 다시 만든다.
# 데이터는 이 파일 옆 data 폴더에 저장된다. 기존 WorkTool 폴더에 이 파일을 두면 기존 데이터를 그대로 쓴다.
#Requires -Version 7
# WorkTool 진입점. 공통 모듈과 기능 모듈을 불러와 왼쪽 메뉴에 붙이고 메인 창을 띄운다.

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# 아래 설정은 창을 만들기 전에 해야 한다. 같은 터미널에서 이미 창을 띄웠으면 실패하므로 무시한다.
[System.Windows.Forms.Application]::EnableVisualStyles()
try { [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false) } catch { }
# 고해상도 화면에서 글자가 흐리지 않게 한다. UiHelper가 배율을 읽기 전에 해야 한다.
[void][System.Windows.Forms.Application]::SetHighDpiMode('SystemAware')

# ---- modules\Storage.psm1 ----
New-Module -Name Storage -ScriptBlock {
# 저장 담당. 다른 모듈은 파일을 직접 읽고 쓰지 않고 이 모듈만 거친다.
# 지금은 data\<이름>.csv에 저장한다. 나중에 저장 방식을 바꿀 때는 이 파일만 고친다.

# 읽기 실패를 빈 결과로 넘기면 다음 저장에서 기존 데이터를 덮어쓰므로, 실패는 항상 예외로 올린다.
$ErrorActionPreference = 'Stop'
$dataDir = Join-Path $PSScriptRoot 'data'
# Excel에서 열어도 한글이 깨지지 않게 BOM을 붙인다.
$encoding = [System.Text.UTF8Encoding]::new($true)

# 핸들을 앱이 끝날 때까지 유지한다. 비정상 종료되어도 OS가 잠금을 풀어 준다.
function Open-WorkToolDataLock {
    [void][System.IO.Directory]::CreateDirectory($dataDir)
    try {
        [System.IO.File]::Open((Join-Path $dataDir '.worktool.lock'), 'OpenOrCreate', 'ReadWrite', 'None')
    } catch {
        throw '같은 데이터 폴더를 사용하는 WorkTool이 실행 중이거나 폴더를 잠글 수 없습니다. 기존 창·트레이를 확인하세요.'
    }
}

function Get-DataPath {
    param([Parameter(Mandatory)][ValidatePattern('^[A-Za-z0-9_-]+$')][string]$Name)
    if (-not (Test-Path $dataDir)) { New-Item -ItemType Directory -Path $dataDir | Out-Null }
    Join-Path $dataDir "$Name.csv"
}

# 기록을 읽는다. 파일이 없으면 아무것도 돌려주지 않으므로 호출하는 쪽에서 @()로 감싼다.
# Import-Csv는 모든 값을 문자열로 읽는다. 날짜·숫자는 호출하는 쪽에서 변환한다.
function Get-Records {
    param([Parameter(Mandatory)][string]$Name)
    $path = Get-DataPath $Name
    if (Test-Path $path) { Import-Csv -Path $path -Encoding utf8 }
}

# 기록 전체를 저장한다. 임시 파일에 먼저 쓰고 바꿔치기해서, 저장 중에 꺼져도 원래 파일은 남는다.
# 직전 파일은 <이름>.csv.bak으로 하나 남긴다.
function Save-Records {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Records
    )
    $path = Get-DataPath $Name
    $tmp = "$path.tmp"
    $text = if ($Records.Count) { ($Records | ConvertTo-Csv -NoTypeInformation) -join "`r`n" } else { '' }
    [System.IO.File]::WriteAllText($tmp, $text, $encoding)
    if (Test-Path $path) {
        [System.IO.File]::Replace($tmp, $path, "$path.bak")
    } else {
        Move-Item -Path $tmp -Destination $path
    }
}

# 새 기록에 붙일 짧은 ID
function New-RecordId {
    [guid]::NewGuid().ToString('N').Substring(0, 12)
}

Export-ModuleMember -Function Get-Records, Save-Records, New-RecordId, Open-WorkToolDataLock
} | Import-Module -Force
# ---- modules\Storage.psm1 끝 ----
# ---- modules\UiHelper.psm1 ----
New-Module -Name UiHelper -ScriptBlock {
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
} | Import-Module -Force
# ---- modules\UiHelper.psm1 끝 ----

$dataLock = Open-WorkToolDataLock
$notify = $null
$form = $null
$pages = @()
try {

# 기능 목록. 기능을 추가할 때 modules\<Module>.psm1을 만들고 여기 한 줄 넣는다.
# 각 모듈은 New-<Module>View 함수로 화면(Control)을 돌려준다. Icon은 Segoe 아이콘 글꼴의 글자 코드.
$features = @(
    @{ Module = 'Schedule'; Title = '일정'; Icon = 0xE787 }
    @{ Module = 'Attendance'; Title = '근태'; Icon = 0xE823 }
    @{ Module = 'Memo'; Title = '메모'; Icon = 0xE70B }
)

$theme = Get-UiTheme
$appIcon = New-UiAppIcon

# 메인 창. 크기는 화면 작업 영역 비율로 잡는다. F11은 전체 화면, Ctrl+B는 메뉴 접기.
$area = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$form = [System.Windows.Forms.Form]::new()
$form.Text = 'WorkTool'
$form.Icon = $appIcon
$form.Font = Get-UiFont 10
$form.BackColor = $theme.Back
$form.ForeColor = $theme.Text
$form.StartPosition = 'CenterScreen'
$form.Size = [System.Drawing.Size]::new([int]($area.Width * 0.65), [int]($area.Height * 0.8))
$form.MinimumSize = [System.Drawing.Size]::new([int]($area.Width * 0.4), [int]($area.Height * 0.5))
$form.KeyPreview = $true

# 가운데: 화면이 들어가는 자리
$content = [System.Windows.Forms.Panel]::new()
$content.Dock = 'Fill'
$content.Padding = New-UiPadding 16

# 홈 화면
$homeView = [System.Windows.Forms.Panel]::new()
$homeView.Dock = 'Fill'
$homeView.BackColor = $theme.Card
$homeView.Padding = New-UiPadding 24
$homeDate = [System.Windows.Forms.Label]::new()
$homeDate.Text = (Get-Date).ToString('M월 d일 dddd')
$homeDate.Font = Get-UiFont 20 -Bold
$homeDate.AutoSize = $true
$homeDate.Dock = 'Top'
$homeNote = [System.Windows.Forms.Label]::new()
$homeNote.Text = '오늘 일정과 알람이 여기 표시될 예정'
$homeNote.ForeColor = $theme.SubText
$homeNote.AutoSize = $true
$homeNote.Dock = 'Top'
$homeNote.Padding = New-UiPadding 2 8 0 0
# Dock=Top은 나중에 넣은 것이 위로 간다
$homeView.Controls.Add($homeNote)
$homeView.Controls.Add($homeDate)

$pages = [System.Collections.Generic.List[hashtable]]::new()
$pages.Add(@{ Title = '홈'; Icon = 0xE80F; View = $homeView })
# ---- modules\Attendance.psm1 ----
New-Module -Name Attendance -ScriptBlock {
# 근태 화면. 출근 / 업무외시간 근무 / 출장 / 휴가를 기록하고 월별로 본다.
# 저장 파일: data\attendance.csv

$ErrorActionPreference = 'Stop'

# 색·글꼴은 UiHelper에서 가져온다 (메인이 먼저 불러온다)
$script:theme = Get-UiTheme
$script:inv = [System.Globalization.CultureInfo]::InvariantCulture

# 종류별로 보여줄 칸. Detail 칸은 종류마다 이름과 선택지가 다르다.
$script:types = [ordered]@{
    '출근'           = @{ EndDate = $false; StartTime = $true;  EndTime = $false; Place = $false; DetailLabel = '메모'; Choices = @() }
    '업무외시간 근무' = @{ EndDate = $false; StartTime = $true;  EndTime = $true;  Place = $false; DetailLabel = '내용'; Choices = @() }
    '출장'           = @{ EndDate = $true;  StartTime = $false; EndTime = $false; Place = $true;  DetailLabel = '목적'; Choices = @() }
    '휴가'           = @{ EndDate = $true;  StartTime = $false; EndTime = $false; Place = $false; DetailLabel = '종류'
                         Choices = @('연차', '오전 반차', '오후 반차', '병가', '경조사', '대체 휴무') }
}

$script:records    = @()
$script:month      = [datetime]::new((Get-Date).Year, (Get-Date).Month, 1)  # 보고 있는 달의 1일
$script:editId     = $null   # 편집 중인 기록 ID. 새 기록이면 $null
$script:busy       = $false  # 코드로 화면을 바꾸는 동안 이벤트를 무시한다
$script:readFailed = $false  # 파일을 못 읽었으면 빈 목록으로 덮어쓰지 않게 저장을 막는다

function ConvertTo-AttDate([string]$Text) {
    [datetime]::ParseExact($Text, 'yyyy-MM-dd', $script:inv)
}

function ConvertTo-AttTime([string]$Text) {
    [timespan]::ParseExact($Text, 'hh\:mm', $script:inv)
}

function Format-AttDay([datetime]$Date) {
    $Date.ToString('MM-dd (ddd)')
}

function Format-AttDuration([int]$Minutes) {
    '{0}:{1:00}' -f [Math]::Floor($Minutes / 60), ($Minutes % 60)
}

# 시작~끝 시각 사이 분. 끝이 시작보다 이르거나 같으면 다음 날로 본다 (야간 온콜 등).
function Get-AttMinutes($Record) {
    if (-not $Record.StartTime -or -not $Record.EndTime) { return 0 }
    $m = ((ConvertTo-AttTime $Record.EndTime) - (ConvertTo-AttTime $Record.StartTime)).TotalMinutes
    if ($m -le 0) { $m += 1440 }
    [int]$m
}

# 기록 기간 중 보고 있는 달에 들어가는 날 수. -WeekdaysOnly면 주말은 뺀다.
function Get-AttDaysInMonth($Record, [switch]$WeekdaysOnly) {
    $first = $script:month
    $last = $first.AddMonths(1).AddDays(-1)
    $s = ConvertTo-AttDate $Record.StartDate
    $e = ConvertTo-AttDate $Record.EndDate
    if ($s -lt $first) { $s = $first }
    if ($e -gt $last) { $e = $last }
    $n = 0
    for ($d = $s; $d -le $e; $d = $d.AddDays(1)) {
        if ($WeekdaysOnly -and $d.DayOfWeek -in 'Saturday', 'Sunday') { continue }
        $n++
    }
    $n
}

function Test-AttInMonth($Record) {
    $first = $script:month
    $last = $first.AddMonths(1).AddDays(-1)
    (ConvertTo-AttDate $Record.StartDate) -le $last -and (ConvertTo-AttDate $Record.EndDate) -ge $first
}

# 이번 달 요약 문장
function Get-AttSummary($Rows) {
    $work = @($Rows | Where-Object Type -eq '출근').Count
    $overtime = @($Rows | Where-Object Type -eq '업무외시간 근무')
    $overtimeMin = 0
    foreach ($r in $overtime) { $overtimeMin += Get-AttMinutes $r }
    $trip = 0
    foreach ($r in @($Rows | Where-Object Type -eq '출장')) { $trip += Get-AttDaysInMonth $r }
    # 휴가는 주말을 빼고 세고, 반차는 0.5일
    $leave = 0.0
    foreach ($r in @($Rows | Where-Object Type -eq '휴가')) {
        $days = Get-AttDaysInMonth $r -WeekdaysOnly
        $leave += $(if ($r.Detail -like '*반차*') { $days * 0.5 } else { $days })
    }
    "출근 ${work}일 · 업무외 근무 $($overtime.Count)건 $(Format-AttDuration $overtimeMin) · 출장 ${trip}일 · 휴가 ${leave}일"
}

function Set-AttStatus([string]$Text) {
    $script:statusLabel.Text = $Text
}

# 목록과 요약을 다시 그린다
function Update-AttList {
    $rows = @($script:records | Where-Object { Test-AttInMonth $_ } | Sort-Object StartDate, StartTime)
    $script:monthLabel.Text = $script:month.ToString('yyyy년 M월')
    $script:summaryLabel.Text = Get-AttSummary $rows

    $script:busy = $true
    try {
        $script:list.BeginUpdate()
        $script:list.Items.Clear()
        foreach ($r in $rows) {
            $s = ConvertTo-AttDate $r.StartDate
            $e = ConvertTo-AttDate $r.EndDate
            $day = if ($e -gt $s) { "$(Format-AttDay $s) ~ $(Format-AttDay $e)" } else { Format-AttDay $s }
            $time = ''
            if ($r.StartTime -and $r.EndTime) {
                $next = if ($r.EndTime -le $r.StartTime) { ' (다음날)' } else { '' }
                $time = "$($r.StartTime)~$($r.EndTime)$next · $(Format-AttDuration (Get-AttMinutes $r))"
            } elseif ($r.StartTime) {
                $time = $r.StartTime
            }
            $item = [System.Windows.Forms.ListViewItem]::new($day)
            [void]$item.SubItems.Add([string]$r.Type)
            [void]$item.SubItems.Add($time)
            [void]$item.SubItems.Add([string]$r.Place)
            [void]$item.SubItems.Add([string]$r.Detail)
            $item.Tag = $r.Id
            [void]$script:list.Items.Add($item)
            if ($r.Id -eq $script:editId) { $item.Selected = $true }
        }
    } finally {
        $script:list.EndUpdate()
        $script:busy = $false
    }
}

# 종류에 맞게 편집 칸을 보이고 숨긴다
function Set-AttType([string]$Type) {
    $f = $script:types[$Type]
    foreach ($key in 'EndDate', 'StartTime', 'EndTime', 'Place') {
        foreach ($c in $script:fieldRows[$key]) { $c.Visible = $f[$key] }
    }
    $script:startDateLabel.Text = if ($f.EndDate) { '시작일' } else { '날짜' }
    $script:detailLabel.Text = $f.DetailLabel
    $text = $script:detailBox.Text
    $script:detailBox.Items.Clear()
    if ($f.Choices.Count) { $script:detailBox.Items.AddRange([object[]]$f.Choices) }
    $script:detailBox.Text = $text
}

# 편집 칸에 기록을 띄운다. $null이면 새 기록.
function Show-AttRecord($Record) {
    $script:busy = $true
    try {
        if ($Record) {
            $script:editId = $Record.Id
            $script:editTitle.Text = '기록 수정'
            $script:typeBox.SelectedItem = $Record.Type
            Set-AttType $Record.Type
            $script:startDate.Value = ConvertTo-AttDate $Record.StartDate
            $script:endDate.Value = ConvertTo-AttDate $Record.EndDate
            if ($Record.StartTime) { $script:startTime.Value = [datetime]::Today.Add((ConvertTo-AttTime $Record.StartTime)) }
            if ($Record.EndTime) { $script:endTime.Value = [datetime]::Today.Add((ConvertTo-AttTime $Record.EndTime)) }
            $script:placeBox.Text = $Record.Place
            $script:detailBox.Text = $Record.Detail
        } else {
            $script:editId = $null
            $script:editTitle.Text = '새 기록'
            # 이번 달을 보고 있으면 오늘, 다른 달이면 그 달 1일
            $today = [datetime]::Today
            $day = if ($today.Year -eq $script:month.Year -and $today.Month -eq $script:month.Month) { $today } else { $script:month }
            $script:startDate.Value = $day
            $script:endDate.Value = $day
            $script:startTime.Value = [datetime]::Now
            $script:endTime.Value = [datetime]::Now
            $script:placeBox.Text = ''
            $script:detailBox.Text = ''
            Set-AttType ([string]$script:typeBox.SelectedItem)
        }
    } finally {
        $script:busy = $false
    }
}

# 기록 하나를 넣거나 바꿔서 저장한다. 성공하면 $true.
function Write-AttRecord($Record) {
    if ($script:readFailed) {
        Show-UiMessage '근태 파일을 읽지 못해 저장하지 않았습니다. 툴을 다시 열어 주세요.' '근태'
        return $false
    }
    $all = @($script:records | Where-Object Id -ne $Record.Id) + $Record
    try {
        Save-Records attendance $all
    } catch {
        Show-UiMessage "저장하지 못했습니다. attendance.csv가 Excel 등에서 열려 있는지 확인해 주세요.`n`n$($_.Exception.Message)" '근태'
        return $false
    }
    $script:records = $all
    $script:editId = $Record.Id
    # 저장한 기록이 있는 달로 이동
    $start = ConvertTo-AttDate $Record.StartDate
    $script:month = [datetime]::new($start.Year, $start.Month, 1)
    Update-AttList
    Set-AttStatus "저장됨 $((Get-Date).ToString('HH:mm:ss'))"
    $true
}

function Save-AttEditor {
    $type = [string]$script:typeBox.SelectedItem
    $f = $script:types[$type]
    $start = $script:startDate.Value.Date
    $end = if ($f.EndDate) { $script:endDate.Value.Date } else { $start }
    if ($end -lt $start) {
        Show-UiMessage '끝나는 날이 시작일보다 빠릅니다.' '근태'
        return
    }
    $record = [pscustomobject]@{
        Id        = if ($script:editId) { $script:editId } else { New-RecordId }
        Type      = $type
        StartDate = $start.ToString('yyyy-MM-dd', $script:inv)
        EndDate   = $end.ToString('yyyy-MM-dd', $script:inv)
        StartTime = if ($f.StartTime) { $script:startTime.Value.ToString('HH:mm', $script:inv) } else { '' }
        EndTime   = if ($f.EndTime) { $script:endTime.Value.ToString('HH:mm', $script:inv) } else { '' }
        Place     = if ($f.Place) { $script:placeBox.Text.Trim() } else { '' }
        Detail    = $script:detailBox.Text.Trim()
        Updated   = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss', $script:inv)
    }
    if ($f.EndTime -and $record.StartTime -eq $record.EndTime) {
        Show-UiMessage '시작 시각과 끝 시각이 같습니다.' '근태'
        return
    }
    if (Write-AttRecord $record) { $script:editTitle.Text = '기록 수정' }
}

# 지금 시각으로 오늘 출근 기록을 바로 남긴다
function Add-AttCheckIn {
    $today = [datetime]::Today.ToString('yyyy-MM-dd', $script:inv)
    $exist = $script:records | Where-Object { $_.Type -eq '출근' -and $_.StartDate -eq $today } | Select-Object -First 1
    if ($exist) {
        Show-UiMessage "오늘 출근 기록이 이미 있습니다 ($($exist.StartTime))." '근태'
        return
    }
    $record = [pscustomobject]@{
        Id        = New-RecordId
        Type      = '출근'
        StartDate = $today
        EndDate   = $today
        StartTime = (Get-Date).ToString('HH:mm', $script:inv)
        EndTime   = ''
        Place     = ''
        Detail    = ''
        Updated   = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss', $script:inv)
    }
    if (Write-AttRecord $record) { Show-AttRecord $record }
}

function Remove-AttRecord {
    if (-not $script:editId) {
        Show-AttRecord $null
        return
    }
    if (-not (Confirm-UiYesNo '이 기록을 삭제할까요?' '근태 삭제')) { return }
    if ($script:readFailed) {
        Show-UiMessage '근태 파일을 읽지 못해 삭제하지 않았습니다. 툴을 다시 열어 주세요.' '근태'
        return
    }
    $rest = @($script:records | Where-Object Id -ne $script:editId)
    try {
        Save-Records attendance $rest
    } catch {
        Show-UiMessage "삭제하지 못했습니다. attendance.csv가 Excel 등에서 열려 있는지 확인해 주세요.`n`n$($_.Exception.Message)" '근태'
        return
    }
    $script:records = $rest
    Show-AttRecord $null
    Update-AttList
    Set-AttStatus '삭제됨'
}

function Move-AttMonth([int]$Offset) {
    $script:month = if ($Offset -eq 0) {
        [datetime]::new((Get-Date).Year, (Get-Date).Month, 1)
    } else {
        $script:month.AddMonths($Offset)
    }
    Update-AttList
}

function New-AttLabel([string]$Text) {
    $l = [System.Windows.Forms.Label]::new()
    $l.Text = $Text
    $l.AutoSize = $true
    $l.Anchor = 'Left'
    $l.ForeColor = $script:theme.SubText
    $l.Margin = New-UiPadding 0 4 10 4
    $l
}

function New-AttDatePicker([string]$Format) {
    $p = [System.Windows.Forms.DateTimePicker]::new()
    $p.Format = 'Custom'
    $p.CustomFormat = $Format
    $p.Dock = 'Fill'
    $p.Margin = New-UiPadding 0 4 0 4
    $p
}

function New-AttendanceView {
    $view = [System.Windows.Forms.Panel]::new()
    $view.BackColor = $script:theme.Card
    $view.Padding = New-UiPadding 16

    # 전체: 위 달 이동 줄 / 요약 / 아래 목록+편집
    $main = [System.Windows.Forms.TableLayoutPanel]::new()
    $main.Dock = 'Fill'
    $main.ColumnCount = 1
    $main.RowCount = 3
    [void]$main.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize))
    [void]$main.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize))
    [void]$main.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))

    $bar = [System.Windows.Forms.FlowLayoutPanel]::new()
    $bar.Dock = 'Fill'
    $bar.AutoSize = $true
    $bar.WrapContents = $false
    $bar.Margin = New-UiPadding 0
    $prevButton = New-UiButton '◀'
    $script:monthLabel = [System.Windows.Forms.Label]::new()
    $script:monthLabel.AutoSize = $true
    $script:monthLabel.Font = Get-UiFont 14 -Bold
    $script:monthLabel.Margin = New-UiPadding 4 2 10 0
    $nextButton = New-UiButton '▶'
    $thisMonthButton = New-UiButton '이번 달'
    $checkInButton = New-UiButton '지금 출근' -Primary
    $checkInButton.Margin = New-UiPadding 24 0 0 0
    $bar.Controls.AddRange(@($prevButton, $script:monthLabel, $nextButton, $thisMonthButton, $checkInButton))

    $script:summaryLabel = [System.Windows.Forms.Label]::new()
    $script:summaryLabel.AutoSize = $true
    $script:summaryLabel.ForeColor = $script:theme.SubText
    $script:summaryLabel.Margin = New-UiPadding 2 8 0 10

    # 아래: 왼쪽 62% 목록, 오른쪽 38% 편집
    $body = [System.Windows.Forms.TableLayoutPanel]::new()
    $body.Dock = 'Fill'
    $body.Margin = New-UiPadding 0
    $body.ColumnCount = 2
    $body.RowCount = 1
    [void]$body.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 62))
    [void]$body.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 38))
    [void]$body.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))

    $script:list = [System.Windows.Forms.ListView]::new()
    $script:list.Dock = 'Fill'
    $script:list.Margin = New-UiPadding 0 0 12 0
    $script:list.View = 'Details'
    $script:list.FullRowSelect = $true
    $script:list.HideSelection = $false
    $script:list.MultiSelect = $false
    $script:list.HeaderStyle = 'Nonclickable'
    $script:list.ForeColor = $script:theme.Text
    [void]$script:list.Columns.Add('날짜', (Get-UiSize 160))
    [void]$script:list.Columns.Add('종류', (Get-UiSize 110))
    [void]$script:list.Columns.Add('시간', (Get-UiSize 170))
    [void]$script:list.Columns.Add('장소', (Get-UiSize 90))
    [void]$script:list.Columns.Add('내용', (Get-UiSize 160))

    # 편집 칸: 왼쪽 이름, 오른쪽 입력
    $editor = [System.Windows.Forms.TableLayoutPanel]::new()
    $editor.Dock = 'Fill'
    $editor.Margin = New-UiPadding 0
    $editor.ColumnCount = 2
    [void]$editor.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::AutoSize))
    [void]$editor.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))

    $script:editTitle = [System.Windows.Forms.Label]::new()
    $script:editTitle.AutoSize = $true
    $script:editTitle.Font = Get-UiFont 12 -Bold
    $script:editTitle.Margin = New-UiPadding 0 0 0 8

    $script:typeBox = [System.Windows.Forms.ComboBox]::new()
    $script:typeBox.DropDownStyle = 'DropDownList'
    $script:typeBox.Dock = 'Fill'
    $script:typeBox.Margin = New-UiPadding 0 4 0 4
    $script:typeBox.Items.AddRange([object[]]@($script:types.Keys))
    $script:typeBox.SelectedIndex = 0

    $script:startDateLabel = New-AttLabel '날짜'
    $script:startDate = New-AttDatePicker 'yyyy-MM-dd (ddd)'
    $endDateLabel = New-AttLabel '종료일'
    $script:endDate = New-AttDatePicker 'yyyy-MM-dd (ddd)'
    $startTimeLabel = New-AttLabel '시작 시각'
    $script:startTime = New-AttDatePicker 'HH:mm'
    $script:startTime.ShowUpDown = $true
    $endTimeLabel = New-AttLabel '끝 시각'
    $script:endTime = New-AttDatePicker 'HH:mm'
    $script:endTime.ShowUpDown = $true
    $placeLabel = New-AttLabel '장소'
    $script:placeBox = [System.Windows.Forms.TextBox]::new()
    $script:placeBox.Dock = 'Fill'
    $script:placeBox.Margin = New-UiPadding 0 4 0 4
    $script:detailLabel = New-AttLabel '내용'
    $script:detailBox = [System.Windows.Forms.ComboBox]::new()
    $script:detailBox.DropDownStyle = 'DropDown'
    $script:detailBox.Dock = 'Fill'
    $script:detailBox.Margin = New-UiPadding 0 4 0 4

    # 종류에 따라 숨기는 줄
    $script:fieldRows = @{
        EndDate   = @($endDateLabel, $script:endDate)
        StartTime = @($startTimeLabel, $script:startTime)
        EndTime   = @($endTimeLabel, $script:endTime)
        Place     = @($placeLabel, $script:placeBox)
    }

    $buttons = [System.Windows.Forms.FlowLayoutPanel]::new()
    $buttons.Dock = 'Fill'
    $buttons.AutoSize = $true
    $buttons.WrapContents = $false
    $buttons.Margin = New-UiPadding 0 12 0 0
    $saveButton = New-UiButton '저장' -Primary
    $newButton = New-UiButton '새로'
    $deleteButton = New-UiButton '삭제'
    $buttons.Controls.AddRange(@($saveButton, $newButton, $deleteButton))

    $script:statusLabel = [System.Windows.Forms.Label]::new()
    $script:statusLabel.AutoSize = $true
    $script:statusLabel.ForeColor = $script:theme.SubText
    $script:statusLabel.Margin = New-UiPadding 0 8 0 0

    $editorRows = @(
        @($script:editTitle),
        @((New-AttLabel '종류'), $script:typeBox),
        @($script:startDateLabel, $script:startDate),
        @($endDateLabel, $script:endDate),
        @($startTimeLabel, $script:startTime),
        @($endTimeLabel, $script:endTime),
        @($placeLabel, $script:placeBox),
        @($script:detailLabel, $script:detailBox),
        @($buttons),
        @($script:statusLabel)
    )
    $editor.RowCount = $editorRows.Count + 1
    for ($i = 0; $i -lt $editorRows.Count; $i++) {
        [void]$editor.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize))
        $cells = $editorRows[$i]
        if ($cells.Count -eq 1) {
            $editor.Controls.Add($cells[0], 0, $i)
            $editor.SetColumnSpan($cells[0], 2)
        } else {
            $editor.Controls.Add($cells[0], 0, $i)
            $editor.Controls.Add($cells[1], 1, $i)
        }
    }
    # 남는 공간은 맨 아래 빈 줄이 차지한다
    [void]$editor.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))

    $body.Controls.Add($script:list, 0, 0)
    $body.Controls.Add($editor, 1, 0)
    $main.Controls.Add($bar, 0, 0)
    $main.Controls.Add($script:summaryLabel, 0, 1)
    $main.Controls.Add($body, 0, 2)
    $view.Controls.Add($main)

    # 이벤트
    $script:list.add_Resize({
        # 마지막 칸(내용)이 남는 폭을 채우게 한다
        $c = $script:list.Columns
        $used = 0
        for ($i = 0; $i -lt $c.Count - 1; $i++) { $used += $c[$i].Width }
        $c[$c.Count - 1].Width = [Math]::Max((Get-UiSize 80), $script:list.ClientSize.Width - $used - 2)
    })
    $script:list.add_SelectedIndexChanged({
        if ($script:busy -or $script:list.SelectedItems.Count -eq 0) { return }
        $id = $script:list.SelectedItems[0].Tag
        Show-AttRecord ($script:records | Where-Object Id -eq $id | Select-Object -First 1)
        Set-AttStatus ''
    })
    $script:typeBox.add_SelectedIndexChanged({
        if (-not $script:busy) { Set-AttType ([string]$script:typeBox.SelectedItem) }
    })
    $prevButton.add_Click({ Move-AttMonth -1 })
    $nextButton.add_Click({ Move-AttMonth 1 })
    $thisMonthButton.add_Click({ Move-AttMonth 0 })
    $checkInButton.add_Click({ Add-AttCheckIn })
    $saveButton.add_Click({ Save-AttEditor })
    $newButton.add_Click({
        Show-AttRecord $null
        Update-AttList
        Set-AttStatus ''
    })
    $deleteButton.add_Click({ Remove-AttRecord })

    # 데이터 읽기
    try {
        $script:records = @(Get-Records attendance)
    } catch {
        $script:readFailed = $true
        Show-UiMessage "근태 파일을 읽지 못했습니다. 이번 실행에서는 저장하지 않습니다.`n`n$($_.Exception.Message)" '근태'
    }
    Show-AttRecord $null
    Update-AttList

    $view
}

Export-ModuleMember -Function New-AttendanceView
} | Import-Module -Force
# ---- modules\Attendance.psm1 끝 ----

# ---- modules\Memo.psm1 ----
New-Module -Name Memo -ScriptBlock {
# 메모 화면. 왼쪽은 검색과 목록, 오른쪽은 제목과 본문 편집.
# 다른 메모로 넘어가거나 새 메모를 시작하거나 창을 닫을 때, 바뀐 내용은 자동 저장한다.

$ErrorActionPreference = 'Stop'

# 색·글꼴은 UiHelper에서 가져온다 (메인이 먼저 불러온다)
$script:theme = Get-UiTheme

$script:memos      = @()
$script:currentId  = $null   # 편집 중인 메모 ID. 아직 저장 안 한 새 메모면 $null
$script:dirty      = $false  # 편집 내용이 저장되지 않았는지
$script:busy       = $false  # 코드로 화면을 바꾸는 동안 이벤트를 무시한다
$script:readFailed = $false  # 파일을 못 읽었으면 빈 목록으로 덮어쓰지 않게 저장을 막는다
$script:closeHooked = $false

function Show-MemoMessage([string]$Text) {
    [void][System.Windows.Forms.MessageBox]::Show($Text, '메모',
        [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
}

function Set-MemoStatus([string]$Text) {
    $script:statusLabel.Text = $Text
}

# 목록을 다시 그린다. 검색어로 거르고 최근 수정 순으로 정렬한다.
function Update-MemoList {
    $q = $script:searchBox.Text.Trim()
    $rows = $script:memos | Sort-Object Updated -Descending
    if ($q) {
        $cmp = [System.StringComparison]::OrdinalIgnoreCase
        $rows = $rows | Where-Object { ([string]$_.Title).Contains($q, $cmp) -or ([string]$_.Body).Contains($q, $cmp) }
    }
    $script:busy = $true
    try {
        $script:list.BeginUpdate()
        $script:list.Items.Clear()
        foreach ($m in $rows) {
            $updated = [string]$m.Updated
            if ($updated.Length -ge 16) { $updated = $updated.Substring(0, 16) }
            $item = [System.Windows.Forms.ListViewItem]::new([string]$m.Title)
            [void]$item.SubItems.Add($updated)
            $item.Tag = $m.Id
            [void]$script:list.Items.Add($item)
            if ($m.Id -eq $script:currentId) { $item.Selected = $true }
        }
    } finally {
        $script:list.EndUpdate()
        $script:busy = $false
    }
}

# 편집 칸에 메모를 띄운다. $null이면 빈 새 메모.
function Show-Memo($Memo) {
    $script:busy = $true
    try {
        if ($Memo) {
            $script:currentId = $Memo.Id
            $script:titleBox.Text = $Memo.Title
            # TextBox는 \r\n이어야 줄이 바뀐다
            $script:bodyBox.Text = ([string]$Memo.Body) -replace "(?<!`r)`n", "`r`n"
        } else {
            $script:currentId = $null
            $script:titleBox.Text = ''
            $script:bodyBox.Text = ''
        }
        $script:dirty = $false
    } finally {
        $script:busy = $false
    }
}

# 편집 중인 메모를 저장한다. 성공하면 $true.
function Save-CurrentMemo {
    if ($script:readFailed) {
        Show-MemoMessage '메모 파일을 읽지 못해 저장하지 않았습니다. 툴을 다시 열어 주세요.'
        return $false
    }
    $title = $script:titleBox.Text.Trim()
    $body = $script:bodyBox.Text
    if (-not $script:currentId -and -not $title -and -not $body.Trim()) {
        $script:dirty = $false
        return $true
    }
    $now = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    # 제목이 비어 있으면 저장 시각을 제목으로 쓴다
    if (-not $title) { $title = $now }

    $id = if ($script:currentId) { $script:currentId } else { New-RecordId }
    $old = $script:memos | Where-Object Id -eq $id | Select-Object -First 1
    $memo = [pscustomobject]@{
        Id      = $id
        Title   = $title
        Body    = $body
        Created = if ($old) { $old.Created } else { $now }
        Updated = $now
    }
    $all = @($script:memos | Where-Object Id -ne $id) + $memo

    try {
        Save-Records memo $all
    } catch {
        Show-MemoMessage "저장하지 못했습니다. memo.csv가 Excel 등에서 열려 있는지 확인해 주세요.`n`n$($_.Exception.Message)"
        return $false
    }
    $script:memos = $all
    $script:currentId = $id
    if ($script:titleBox.Text.Trim() -ne $title) {
        $script:busy = $true
        $script:titleBox.Text = $title
        $script:busy = $false
    }
    $script:dirty = $false
    Set-MemoStatus "저장됨 $((Get-Date).ToString('HH:mm:ss'))"
    Update-MemoList
    return $true
}

function Remove-CurrentMemo {
    if (-not $script:currentId) {
        Show-Memo $null
        return
    }
    $answer = [System.Windows.Forms.MessageBox]::Show('이 메모를 삭제할까요?', '메모 삭제',
        [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    if ($script:readFailed) {
        Show-MemoMessage '메모 파일을 읽지 못해 삭제하지 않았습니다. 툴을 다시 열어 주세요.'
        return
    }

    $rest = @($script:memos | Where-Object Id -ne $script:currentId)
    try {
        Save-Records memo $rest
    } catch {
        Show-MemoMessage "삭제하지 못했습니다. memo.csv가 Excel 등에서 열려 있는지 확인해 주세요.`n`n$($_.Exception.Message)"
        return
    }
    $script:memos = $rest
    Show-Memo $null
    Update-MemoList
    Set-MemoStatus '삭제됨'
}

function New-MemoView {
    $view = [System.Windows.Forms.Panel]::new()
    $view.BackColor = $script:theme.Card
    $view.Padding = New-UiPadding 16

    # 왼쪽 32% 목록, 오른쪽 68% 편집
    $grid = [System.Windows.Forms.TableLayoutPanel]::new()
    $grid.Dock = 'Fill'
    $grid.ColumnCount = 2
    $grid.RowCount = 1
    [void]$grid.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 32))
    [void]$grid.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 68))
    [void]$grid.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))

    # 왼쪽: 검색 + 목록
    $left = [System.Windows.Forms.TableLayoutPanel]::new()
    $left.Dock = 'Fill'
    $left.Margin = New-UiPadding 0 0 12 0
    $left.ColumnCount = 1
    $left.RowCount = 2
    [void]$left.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize))
    [void]$left.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))

    $script:searchBox = [System.Windows.Forms.TextBox]::new()
    $script:searchBox.Dock = 'Fill'
    $script:searchBox.PlaceholderText = '검색 (제목·본문)'
    $script:searchBox.Margin = New-UiPadding 0 0 0 8

    $script:list = [System.Windows.Forms.ListView]::new()
    $script:list.Dock = 'Fill'
    $script:list.Margin = [System.Windows.Forms.Padding]::new(0)
    $script:list.View = 'Details'
    $script:list.FullRowSelect = $true
    $script:list.HideSelection = $false
    $script:list.MultiSelect = $false
    $script:list.HeaderStyle = 'Nonclickable'
    $script:list.ForeColor = $script:theme.Text
    [void]$script:list.Columns.Add('제목', (Get-UiSize 200))
    [void]$script:list.Columns.Add('수정', (Get-UiSize 130))

    $left.Controls.Add($script:searchBox, 0, 0)
    $left.Controls.Add($script:list, 0, 1)

    # 오른쪽: 버튼 줄 + 제목 + 본문
    $right = [System.Windows.Forms.TableLayoutPanel]::new()
    $right.Dock = 'Fill'
    $right.Margin = [System.Windows.Forms.Padding]::new(0)
    $right.ColumnCount = 1
    $right.RowCount = 3
    [void]$right.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize))
    [void]$right.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize))
    [void]$right.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))

    $bar = [System.Windows.Forms.FlowLayoutPanel]::new()
    $bar.Dock = 'Fill'
    $bar.AutoSize = $true
    $bar.WrapContents = $false
    $bar.Margin = New-UiPadding 0 0 0 8
    $newButton = New-UiButton '새 메모'
    $saveButton = New-UiButton '저장 (Ctrl+S)' -Primary
    $deleteButton = New-UiButton '삭제'
    $script:statusLabel = [System.Windows.Forms.Label]::new()
    $script:statusLabel.AutoSize = $true
    $script:statusLabel.ForeColor = $script:theme.SubText
    $script:statusLabel.Margin = New-UiPadding 6 7 0 0
    $bar.Controls.AddRange(@($newButton, $saveButton, $deleteButton, $script:statusLabel))

    $script:titleBox = [System.Windows.Forms.TextBox]::new()
    $script:titleBox.Dock = 'Fill'
    $script:titleBox.Font = Get-UiFont 12 -Bold
    $script:titleBox.PlaceholderText = '제목 (비우면 저장 시각)'
    $script:titleBox.Margin = New-UiPadding 0 0 0 8

    $script:bodyBox = [System.Windows.Forms.TextBox]::new()
    $script:bodyBox.Dock = 'Fill'
    $script:bodyBox.Margin = [System.Windows.Forms.Padding]::new(0)
    $script:bodyBox.Multiline = $true
    $script:bodyBox.AcceptsReturn = $true
    $script:bodyBox.AcceptsTab = $true
    $script:bodyBox.ScrollBars = 'Vertical'

    $right.Controls.Add($bar, 0, 0)
    $right.Controls.Add($script:titleBox, 0, 1)
    $right.Controls.Add($script:bodyBox, 0, 2)

    $grid.Controls.Add($left, 0, 0)
    $grid.Controls.Add($right, 1, 0)
    $view.Controls.Add($grid)

    # 이벤트
    $script:list.add_Resize({
        # 제목 칸이 남는 폭을 채우게 한다
        $c = $script:list.Columns
        $c[0].Width = [Math]::Max(80, $script:list.ClientSize.Width - $c[1].Width - 2)
    })
    $script:list.add_SelectedIndexChanged({
        if ($script:busy -or $script:list.SelectedItems.Count -eq 0) { return }
        $id = $script:list.SelectedItems[0].Tag
        if ($id -eq $script:currentId) { return }
        if ($script:dirty -and -not (Save-CurrentMemo)) {
            Update-MemoList
            return
        }
        Show-Memo ($script:memos | Where-Object Id -eq $id | Select-Object -First 1)
        Update-MemoList
        Set-MemoStatus ''
    })
    $script:searchBox.add_TextChanged({ Update-MemoList })

    $markDirty = {
        if (-not $script:busy) {
            $script:dirty = $true
            Set-MemoStatus '수정 중'
        }
    }
    $script:titleBox.add_TextChanged($markDirty)
    $script:bodyBox.add_TextChanged($markDirty)

    # Ctrl+S 저장, Ctrl+A 전체 선택 (여러 줄 TextBox는 Ctrl+A가 기본으로 안 된다)
    $shortcutKeys = {
        param($s, $e)
        if (-not $e.Control) { return }
        if ($e.KeyCode -eq [System.Windows.Forms.Keys]::S) {
            $e.SuppressKeyPress = $true
            [void](Save-CurrentMemo)
        } elseif ($e.KeyCode -eq [System.Windows.Forms.Keys]::A) {
            $e.SuppressKeyPress = $true
            $s.SelectAll()
        }
    }
    $script:titleBox.add_KeyDown($shortcutKeys)
    $script:bodyBox.add_KeyDown($shortcutKeys)

    $newButton.add_Click({
        if ($script:dirty -and -not (Save-CurrentMemo)) { return }
        Show-Memo $null
        Update-MemoList
        Set-MemoStatus ''
        [void]$script:titleBox.Focus()
    })
    $saveButton.add_Click({ [void](Save-CurrentMemo) })
    $deleteButton.add_Click({ Remove-CurrentMemo })

    # 창을 닫을 때 저장 안 한 내용을 저장한다. 화면이 창에 붙은 뒤에야 창을 찾을 수 있다.
    $view.add_HandleCreated({
        $form = $this.FindForm()
        if ($form -and -not $script:closeHooked) {
            $form.add_FormClosing({
                param($sender, $event)
                if ($script:dirty -and -not (Save-CurrentMemo)) { $event.Cancel = $true }
            })
            $script:closeHooked = $true
        }
    })

    # 데이터 읽기
    try {
        $script:memos = @(Get-Records memo)
    } catch {
        $script:readFailed = $true
        Show-MemoMessage "메모 파일을 읽지 못했습니다. 이번 실행에서는 저장하지 않습니다.`n`n$($_.Exception.Message)"
    }
    Update-MemoList

    $view
}

Export-ModuleMember -Function New-MemoView
} | Import-Module -Force
# ---- modules\Memo.psm1 끝 ----

# ---- modules\Schedule.psm1 ----
New-Module -Name Schedule -ScriptBlock {
# 일정·알람 화면. 달력에서 날짜를 고르고 그날 일정을 편집한다.
# 알람은 툴이 켜져 있는 동안(트레이 포함) 타이머로 확인하고, 시간이 되면 소리와 함께 알림 창을 띄운다.
# 저장 파일: data\schedule.csv

$ErrorActionPreference = 'Stop'

# 색·글꼴은 UiHelper에서 가져온다 (메인이 먼저 불러온다)
$script:theme = Get-UiTheme
$script:inv = [System.Globalization.CultureInfo]::InvariantCulture

# 화면에 보이는 이름 → 저장 값
$script:alarmChoices = [ordered]@{
    '없음' = ''; '정시' = '0'; '5분 전' = '5'; '10분 전' = '10'; '30분 전' = '30'; '1시간 전' = '60'; '하루 전' = '1440'
}
$script:repeatChoices = [ordered]@{
    '없음' = ''; '매일' = 'daily'; '평일' = 'weekdays'; '매주' = 'weekly'
}

$script:items      = @()
$script:day        = [datetime]::Today   # 달력에서 고른 날
$script:editId     = $null   # 편집 중인 일정 ID. 새 일정이면 $null
$script:busy       = $false  # 코드로 화면을 바꾸는 동안 이벤트를 무시한다
$script:readFailed = $false  # 파일을 못 읽었으면 빈 목록으로 덮어쓰지 않게 저장을 막는다
$script:fired      = [System.Collections.Generic.HashSet[string]]::new()   # 이번 실행에서 울린 알람
$script:snoozes    = [System.Collections.Generic.List[hashtable]]::new()   # 다시 알림 예약

function ConvertTo-SchDate([string]$Text) {
    [datetime]::ParseExact($Text, 'yyyy-MM-dd', $script:inv)
}

function ConvertTo-SchTime([string]$Text) {
    [timespan]::ParseExact($Text, 'hh\:mm', $script:inv)
}

# 저장 값 → 화면에 보이는 이름
function Get-SchChoiceName($Choices, [string]$Value) {
    foreach ($k in $Choices.Keys) {
        if ($Choices[$k] -eq $Value) { return $k }
    }
    $null
}

# 일정이 그날 있는지. 반복 일정은 시작일부터 규칙대로 반복한다.
function Test-SchOccurs($Item, [datetime]$Day) {
    $start = ConvertTo-SchDate $Item.Date
    if ($Day -lt $start) { return $false }
    switch ($Item.Repeat) {
        'daily'    { return $true }
        'weekdays' { return $Day.DayOfWeek -notin 'Saturday', 'Sunday' }
        'weekly'   { return $Day.DayOfWeek -eq $start.DayOfWeek }
    }
    $Day -eq $start
}

# 그날 일정. 하루 종일 일정이 먼저, 나머지는 시각 순.
function Get-SchDayItems([datetime]$Day) {
    $script:items | Where-Object { Test-SchOccurs $_ $Day } | Sort-Object Time, Title
}

function Set-SchStatus([string]$Text) {
    $script:statusLabel.Text = $Text
}

function Update-SchDayList {
    $script:dayLabel.Text = $script:day.ToString('M월 d일 dddd')
    $rows = @(Get-SchDayItems $script:day)
    $script:busy = $true
    try {
        $script:dayList.BeginUpdate()
        $script:dayList.Items.Clear()
        foreach ($r in $rows) {
            $item = [System.Windows.Forms.ListViewItem]::new($(if ($r.Time) { $r.Time } else { '하루 종일' }))
            [void]$item.SubItems.Add([string]$r.Title)
            [void]$item.SubItems.Add($(if ($r.Alarm -ne '') { Get-SchChoiceName $script:alarmChoices $r.Alarm } else { '' }))
            [void]$item.SubItems.Add($(if ($r.Repeat) { Get-SchChoiceName $script:repeatChoices $r.Repeat } else { '' }))
            $item.Tag = $r.Id
            [void]$script:dayList.Items.Add($item)
            if ($r.Id -eq $script:editId) { $item.Selected = $true }
        }
    } finally {
        $script:dayList.EndUpdate()
        $script:busy = $false
    }
}

# 오늘부터 7일 동안의 일정. 오늘 이미 지난 시각은 뺀다.
function Update-SchUpcoming {
    $now = Get-Date
    $script:busy = $true
    try {
        $script:upcoming.BeginUpdate()
        $script:upcoming.Items.Clear()
        for ($d = 0; $d -lt 7; $d++) {
            $day = [datetime]::Today.AddDays($d)
            foreach ($r in @(Get-SchDayItems $day)) {
                if ($d -eq 0 -and $r.Time -and $day.Add((ConvertTo-SchTime $r.Time)) -lt $now) { continue }
                $when = $day.ToString('MM-dd (ddd)') + $(if ($r.Time) { " $($r.Time)" } else { '' })
                $item = [System.Windows.Forms.ListViewItem]::new($when)
                [void]$item.SubItems.Add([string]$r.Title)
                $item.Tag = @{ Id = $r.Id; Day = $day }
                [void]$script:upcoming.Items.Add($item)
                # 타이머가 15초마다 다시 그려도 고른 줄이 풀리지 않게 한다
                if ($r.Id -eq $script:editId -and $day -eq $script:day) { $item.Selected = $true }
            }
        }
    } finally {
        $script:upcoming.EndUpdate()
        $script:busy = $false
    }
}

# 달력에서 반복 없는 일정이 있는 날을 굵게 표시한다 (매일 반복 등은 표시하면 전부 굵어져서 뺀다)
function Update-SchCalendar {
    $dates = @($script:items | Where-Object { -not $_.Repeat } | ForEach-Object { ConvertTo-SchDate $_.Date })
    $script:calendar.BoldedDates = [datetime[]]$dates
    $script:calendar.UpdateBoldedDates()
}

function Update-SchAll {
    Update-SchDayList
    Update-SchUpcoming
    Update-SchCalendar
}

# 하루 종일이면 시각과 알람은 쓰지 않는다
function Update-SchAllDay {
    $on = -not $script:allDayBox.Checked
    $script:timePicker.Enabled = $on
    $script:alarmBox.Enabled = $on
}

# 편집 칸에 일정을 띄운다. $null이면 고른 날짜의 새 일정.
function Show-SchItem($Item) {
    $script:busy = $true
    try {
        if ($Item) {
            $script:editId = $Item.Id
            $script:editTitle.Text = '일정 수정'
            $script:titleBox.Text = $Item.Title
            $script:datePicker.Value = ConvertTo-SchDate $Item.Date
            $script:allDayBox.Checked = -not $Item.Time
            if ($Item.Time) { $script:timePicker.Value = [datetime]::Today.Add((ConvertTo-SchTime $Item.Time)) }
            $alarm = Get-SchChoiceName $script:alarmChoices $Item.Alarm
            $script:alarmBox.SelectedItem = $(if ($alarm) { $alarm } else { '없음' })
            $repeat = Get-SchChoiceName $script:repeatChoices $Item.Repeat
            $script:repeatBox.SelectedItem = $(if ($repeat) { $repeat } else { '없음' })
            $script:noteBox.Text = ([string]$Item.Note) -replace "(?<!`r)`n", "`r`n"
        } else {
            $script:editId = $null
            $script:editTitle.Text = '새 일정'
            $script:titleBox.Text = ''
            $script:datePicker.Value = $script:day
            $script:allDayBox.Checked = $false
            # 기본 시각은 다음 정각
            $script:timePicker.Value = [datetime]::Today.AddHours([datetime]::Now.Hour + 1)
            $script:alarmBox.SelectedItem = '10분 전'
            $script:repeatBox.SelectedItem = '없음'
            $script:noteBox.Text = ''
        }
        Update-SchAllDay
    } finally {
        $script:busy = $false
    }
}

# 일정 목록 전체를 저장한다. 성공하면 $true.
function Write-SchItems([object[]]$All, [string]$Action) {
    if ($script:readFailed) {
        Show-UiMessage "일정 파일을 읽지 못해 ${Action}하지 않았습니다. 툴을 다시 열어 주세요." '일정'
        return $false
    }
    try {
        Save-Records schedule $All
    } catch {
        Show-UiMessage "${Action}하지 못했습니다. schedule.csv가 Excel 등에서 열려 있는지 확인해 주세요.`n`n$($_.Exception.Message)" '일정'
        return $false
    }
    $script:items = $All
    $true
}

function Save-SchEditor {
    $title = $script:titleBox.Text.Trim()
    if (-not $title) {
        Show-UiMessage '제목을 입력해 주세요.' '일정'
        return
    }
    $allDay = $script:allDayBox.Checked
    $old = $script:items | Where-Object Id -eq $script:editId | Select-Object -First 1
    $date = $script:datePicker.Value.Date
    $record = [pscustomobject]@{
        Id        = if ($script:editId) { $script:editId } else { New-RecordId }
        Date      = $date.ToString('yyyy-MM-dd', $script:inv)
        Time      = if ($allDay) { '' } else { $script:timePicker.Value.ToString('HH:mm', $script:inv) }
        Title     = $title
        Note      = $script:noteBox.Text
        Alarm     = if ($allDay) { '' } else { $script:alarmChoices[[string]$script:alarmBox.SelectedItem] }
        Repeat    = $script:repeatChoices[[string]$script:repeatBox.SelectedItem]
        LastAlarm = if ($old) { [string]$old.LastAlarm } else { '' }
        Updated   = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss', $script:inv)
    }
    $all = @($script:items | Where-Object Id -ne $record.Id) + $record
    if (-not (Write-SchItems $all '저장')) { return }

    $script:editId = $record.Id
    $script:editTitle.Text = '일정 수정'
    $script:day = $date
    $script:busy = $true
    $script:calendar.SelectionStart = $date
    $script:busy = $false
    Update-SchAll
    Set-SchStatus "저장됨 $((Get-Date).ToString('HH:mm:ss'))"
}

function Remove-SchItem {
    if (-not $script:editId) {
        Show-SchItem $null
        return
    }
    if (-not (Confirm-UiYesNo '이 일정을 삭제할까요? 반복 일정이면 모든 반복이 지워집니다.' '일정 삭제')) { return }
    $rest = @($script:items | Where-Object Id -ne $script:editId)
    if (-not (Write-SchItems $rest '삭제')) { return }
    Show-SchItem $null
    Update-SchAll
    Set-SchStatus '삭제됨'
}

# 알림 창. 화면 오른쪽 아래에 맨 위로 띄운다.
function Show-SchAlarm($Item, [datetime]$Occurrence) {
    [System.Media.SystemSounds]::Exclamation.Play()

    $f = [System.Windows.Forms.Form]::new()
    $f.Text = 'WorkTool 알람'
    $f.Tag = 'alarm'
    $f.FormBorderStyle = 'FixedToolWindow'
    $f.TopMost = $true
    $f.ShowInTaskbar = $true
    $f.StartPosition = 'Manual'
    $f.AutoSize = $true
    $f.AutoSizeMode = 'GrowAndShrink'
    $f.BackColor = $script:theme.Card
    $f.ForeColor = $script:theme.Text
    $f.Font = Get-UiFont 10
    $f.Padding = New-UiPadding 16

    $layout = [System.Windows.Forms.TableLayoutPanel]::new()
    $layout.AutoSize = $true
    $layout.ColumnCount = 1
    $layout.MinimumSize = [System.Drawing.Size]::new((Get-UiSize 300), 0)

    $when = [System.Windows.Forms.Label]::new()
    $when.AutoSize = $true
    $when.ForeColor = $script:theme.Accent
    $when.Text = $Occurrence.ToString('M월 d일 (ddd) HH:mm')
    $title = [System.Windows.Forms.Label]::new()
    $title.AutoSize = $true
    $title.MaximumSize = [System.Drawing.Size]::new((Get-UiSize 360), 0)
    $title.Font = Get-UiFont 13 -Bold
    $title.Margin = New-UiPadding 0 4 0 4
    $title.Text = $Item.Title
    $note = [System.Windows.Forms.Label]::new()
    $note.AutoSize = $true
    $note.MaximumSize = [System.Drawing.Size]::new((Get-UiSize 360), 0)
    $note.ForeColor = $script:theme.SubText
    $note.Text = [string]$Item.Note
    $note.Visible = [bool]$Item.Note

    $buttons = [System.Windows.Forms.FlowLayoutPanel]::new()
    $buttons.AutoSize = $true
    $buttons.WrapContents = $false
    $buttons.Margin = New-UiPadding 0 12 0 0
    $ok = New-UiButton '확인' -Primary
    $snooze = New-UiButton '5분 뒤 다시'
    $buttons.Controls.AddRange(@($ok, $snooze))

    $layout.Controls.AddRange(@($when, $title, $note, $buttons))
    $f.Controls.Add($layout)

    $ok.Tag = @{ Form = $f }
    $ok.add_Click({ $this.Tag.Form.Close() })
    $snooze.Tag = @{ Form = $f; Id = $Item.Id; Occurrence = $Occurrence }
    $snooze.add_Click({
        $t = $this.Tag
        $script:snoozes.Add(@{ Id = $t.Id; Occurrence = $t.Occurrence; At = (Get-Date).AddMinutes(5) })
        $t.Form.Close()
    })

    # 이미 떠 있는 알림 창이 있으면 그만큼 위로 올려서 겹치지 않게 한다
    $f.add_Load({
        $area = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
        $open = @([System.Windows.Forms.Application]::OpenForms | Where-Object { $_.Tag -eq 'alarm' -and $_ -ne $this }).Count
        $gap = Get-UiSize 16
        $this.Location = [System.Drawing.Point]::new($area.Right - $this.Width - $gap, $area.Bottom - ($this.Height + $gap) * ($open + 1))
    })
    $f.Show()
    $f.Activate()
}

# 타이머가 부른다. 울릴 때가 된 알람을 띄운다.
# 어제·오늘·내일 일정을 보고, 알람 시각부터 일정 시각 30분 뒤까지 사이에 툴이 켜져 있으면 한 번 울린다.
function Invoke-SchAlarmCheck {
    $now = Get-Date

    foreach ($z in @($script:snoozes)) {
        if ($now -lt $z.At) { continue }
        [void]$script:snoozes.Remove($z)
        $r = $script:items | Where-Object Id -eq $z.Id | Select-Object -First 1
        if ($r) { Show-SchAlarm $r $z.Occurrence }
    }

    $changed = $false
    foreach ($r in @($script:items)) {
        if ($r.Alarm -eq '' -or -not $r.Time) { continue }
        foreach ($offset in -1, 0, 1) {
            $day = [datetime]::Today.AddDays($offset)
            if (-not (Test-SchOccurs $r $day)) { continue }
            $occurrence = $day.Add((ConvertTo-SchTime $r.Time))
            $alarmAt = $occurrence.AddMinutes(-[int]$r.Alarm)
            $key = $occurrence.ToString('yyyy-MM-dd HH:mm', $script:inv)
            if ($now -lt $alarmAt -or $now -gt $occurrence.AddMinutes(30)) { continue }
            if ($r.LastAlarm -eq $key -or $script:fired.Contains("$($r.Id)|$key")) { continue }
            [void]$script:fired.Add("$($r.Id)|$key")
            $r.LastAlarm = $key
            $changed = $true
            Show-SchAlarm $r $occurrence
        }
    }
    # 울린 기록을 남겨서 툴을 다시 켜도 같은 알람이 또 울리지 않게 한다. 실패해도 이번 실행에서는 안 울린다.
    if ($changed -and -not $script:readFailed) {
        try { Save-Records schedule $script:items } catch { }
    }
    if ($script:upcoming) { Update-SchUpcoming }
}

function New-ScheduleView {
    $view = [System.Windows.Forms.Panel]::new()
    $view.BackColor = $script:theme.Card
    $view.Padding = New-UiPadding 16

    # 왼쪽: 달력, 오늘, 다가오는 7일 / 오른쪽: 그날 일정, 편집
    $grid = [System.Windows.Forms.TableLayoutPanel]::new()
    $grid.Dock = 'Fill'
    $grid.ColumnCount = 2
    $grid.RowCount = 1
    [void]$grid.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::AutoSize))
    [void]$grid.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
    [void]$grid.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))

    $left = [System.Windows.Forms.TableLayoutPanel]::new()
    $left.Dock = 'Fill'
    $left.AutoSize = $true
    $left.Margin = New-UiPadding 0 0 16 0
    $left.ColumnCount = 1
    $left.RowCount = 4
    foreach ($i in 1..3) {
        [void]$left.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize))
    }
    [void]$left.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))

    $script:calendar = [System.Windows.Forms.MonthCalendar]::new()
    $script:calendar.MaxSelectionCount = 1
    $script:calendar.Margin = New-UiPadding 0
    $todayButton = New-UiButton '오늘'
    $todayButton.Margin = New-UiPadding 0 8 0 0
    $upcomingLabel = [System.Windows.Forms.Label]::new()
    $upcomingLabel.Text = '다가오는 7일'
    $upcomingLabel.AutoSize = $true
    $upcomingLabel.Font = Get-UiFont 10 -Bold
    $upcomingLabel.Margin = New-UiPadding 0 16 0 6
    $script:upcoming = New-UiListView @(@('언제', 130), @('제목', 100))

    $left.Controls.Add($script:calendar, 0, 0)
    $left.Controls.Add($todayButton, 0, 1)
    $left.Controls.Add($upcomingLabel, 0, 2)
    $left.Controls.Add($script:upcoming, 0, 3)

    $right = [System.Windows.Forms.TableLayoutPanel]::new()
    $right.Dock = 'Fill'
    $right.Margin = New-UiPadding 0
    $right.ColumnCount = 1
    $right.RowCount = 3
    [void]$right.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize))
    [void]$right.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))
    [void]$right.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize))

    $script:dayLabel = [System.Windows.Forms.Label]::new()
    $script:dayLabel.AutoSize = $true
    $script:dayLabel.Font = Get-UiFont 14 -Bold
    $script:dayLabel.Margin = New-UiPadding 0 0 0 8
    $script:dayList = New-UiListView @(@('시간', 90), @('제목', 260), @('알람', 80), @('반복', 80))
    $script:dayList.Margin = New-UiPadding 0 0 0 12

    # 편집 칸: 왼쪽 이름, 오른쪽 입력
    $editor = [System.Windows.Forms.TableLayoutPanel]::new()
    $editor.Dock = 'Fill'
    $editor.AutoSize = $true
    $editor.Margin = New-UiPadding 0
    $editor.ColumnCount = 2
    [void]$editor.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::AutoSize))
    [void]$editor.ColumnStyles.Add([System.Windows.Forms.ColumnStyle]::new([System.Windows.Forms.SizeType]::Percent, 100))

    $script:editTitle = [System.Windows.Forms.Label]::new()
    $script:editTitle.AutoSize = $true
    $script:editTitle.Font = Get-UiFont 12 -Bold
    $script:editTitle.Margin = New-UiPadding 0 0 0 8

    $script:titleBox = [System.Windows.Forms.TextBox]::new()
    $script:titleBox.Dock = 'Fill'
    $script:titleBox.Margin = New-UiPadding 0 4 0 4
    $script:datePicker = New-UiDatePicker 'yyyy-MM-dd (ddd)'

    # 시각 줄: 시각 입력 + 하루 종일
    $timeRow = [System.Windows.Forms.FlowLayoutPanel]::new()
    $timeRow.AutoSize = $true
    $timeRow.WrapContents = $false
    $timeRow.Margin = New-UiPadding 0
    $script:timePicker = New-UiDatePicker 'HH:mm' -Time
    $script:timePicker.Dock = 'None'
    $script:timePicker.Width = Get-UiSize 90
    $script:allDayBox = [System.Windows.Forms.CheckBox]::new()
    $script:allDayBox.Text = '하루 종일'
    $script:allDayBox.AutoSize = $true
    $script:allDayBox.Margin = New-UiPadding 12 7 0 0
    $timeRow.Controls.AddRange(@($script:timePicker, $script:allDayBox))

    $script:alarmBox = [System.Windows.Forms.ComboBox]::new()
    $script:alarmBox.DropDownStyle = 'DropDownList'
    $script:alarmBox.Width = Get-UiSize 140
    $script:alarmBox.Anchor = 'Left'
    $script:alarmBox.Margin = New-UiPadding 0 4 0 4
    $script:alarmBox.Items.AddRange([object[]]@($script:alarmChoices.Keys))
    $script:repeatBox = [System.Windows.Forms.ComboBox]::new()
    $script:repeatBox.DropDownStyle = 'DropDownList'
    $script:repeatBox.Width = Get-UiSize 140
    $script:repeatBox.Anchor = 'Left'
    $script:repeatBox.Margin = New-UiPadding 0 4 0 4
    $script:repeatBox.Items.AddRange([object[]]@($script:repeatChoices.Keys))

    $script:noteBox = [System.Windows.Forms.TextBox]::new()
    $script:noteBox.Dock = 'Fill'
    $script:noteBox.Multiline = $true
    $script:noteBox.AcceptsReturn = $true
    $script:noteBox.ScrollBars = 'Vertical'
    $script:noteBox.Height = Get-UiSize 70
    $script:noteBox.Margin = New-UiPadding 0 4 0 4

    $buttons = [System.Windows.Forms.FlowLayoutPanel]::new()
    $buttons.AutoSize = $true
    $buttons.WrapContents = $false
    $buttons.Margin = New-UiPadding 0 10 0 0
    $saveButton = New-UiButton '저장' -Primary
    $newButton = New-UiButton '새로'
    $deleteButton = New-UiButton '삭제'
    $script:statusLabel = [System.Windows.Forms.Label]::new()
    $script:statusLabel.AutoSize = $true
    $script:statusLabel.ForeColor = $script:theme.SubText
    $script:statusLabel.Margin = New-UiPadding 6 7 0 0
    $buttons.Controls.AddRange(@($saveButton, $newButton, $deleteButton, $script:statusLabel))

    $editorRows = @(
        @($script:editTitle),
        @((New-UiFieldLabel '제목'), $script:titleBox),
        @((New-UiFieldLabel '날짜'), $script:datePicker),
        @((New-UiFieldLabel '시각'), $timeRow),
        @((New-UiFieldLabel '알람'), $script:alarmBox),
        @((New-UiFieldLabel '반복'), $script:repeatBox),
        @((New-UiFieldLabel '메모'), $script:noteBox),
        @($buttons)
    )
    $editor.RowCount = $editorRows.Count
    for ($i = 0; $i -lt $editorRows.Count; $i++) {
        [void]$editor.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize))
        $cells = $editorRows[$i]
        $editor.Controls.Add($cells[0], 0, $i)
        if ($cells.Count -eq 1) {
            $editor.SetColumnSpan($cells[0], 2)
        } else {
            $editor.Controls.Add($cells[1], 1, $i)
        }
    }

    $right.Controls.Add($script:dayLabel, 0, 0)
    $right.Controls.Add($script:dayList, 0, 1)
    $right.Controls.Add($editor, 0, 2)

    $grid.Controls.Add($left, 0, 0)
    $grid.Controls.Add($right, 1, 0)
    $view.Controls.Add($grid)

    # 이벤트
    $script:calendar.add_DateChanged({
        param($s, $e)
        if ($script:busy) { return }
        $script:day = $e.Start.Date
        Show-SchItem $null
        Update-SchDayList
        Set-SchStatus ''
    })
    $todayButton.add_Click({ $script:calendar.SelectionStart = [datetime]::Today })
    $script:dayList.add_SelectedIndexChanged({
        if ($script:busy -or $script:dayList.SelectedItems.Count -eq 0) { return }
        $id = $script:dayList.SelectedItems[0].Tag
        Show-SchItem ($script:items | Where-Object Id -eq $id | Select-Object -First 1)
        Set-SchStatus ''
    })
    $script:upcoming.add_SelectedIndexChanged({
        if ($script:busy -or $script:upcoming.SelectedItems.Count -eq 0) { return }
        $t = $script:upcoming.SelectedItems[0].Tag
        $script:day = $t.Day
        $script:busy = $true
        $script:calendar.SelectionStart = $t.Day
        $script:busy = $false
        Show-SchItem ($script:items | Where-Object Id -eq $t.Id | Select-Object -First 1)
        Update-SchDayList
        Set-SchStatus ''
    })
    $script:allDayBox.add_CheckedChanged({ Update-SchAllDay })
    $saveButton.add_Click({ Save-SchEditor })
    $newButton.add_Click({
        Show-SchItem $null
        Update-SchDayList
        Set-SchStatus ''
        [void]$script:titleBox.Focus()
    })
    $deleteButton.add_Click({ Remove-SchItem })

    # 데이터 읽기
    try {
        $script:items = @(Get-Records schedule)
    } catch {
        $script:readFailed = $true
        Show-UiMessage "일정 파일을 읽지 못했습니다. 이번 실행에서는 저장하지 않습니다.`n`n$($_.Exception.Message)" '일정'
    }
    Show-SchItem $null
    Update-SchAll

    # 알람 확인 타이머 (15초마다). 오류가 나도 툴이 멈추지 않게 잡아서 상태 칸에만 보여준다.
    $script:timer = [System.Windows.Forms.Timer]::new()
    $script:timer.Interval = 15000
    $script:timer.add_Tick({
        try { Invoke-SchAlarmCheck } catch { Set-SchStatus "알람 확인 오류: $($_.Exception.Message)" }
    })
    $script:timer.Start()
    # 창이 닫히면 타이머도 끈다. 같은 터미널에서 다시 실행할 때 이전 타이머가 남아 알람이 두 번 울리지 않게 한다.
    $view.add_Disposed({
        $script:timer.Stop()
        $script:timer.Dispose()
    })

    $view
}

Export-ModuleMember -Function New-ScheduleView
} | Import-Module -Force
# ---- modules\Schedule.psm1 끝 ----

foreach ($f in $features) {
    $pages.Add(@{ Title = $f.Title; Icon = $f.Icon; View = (& "New-$($f.Module)View") })
}

# 왼쪽 메뉴. 펼치면 200, 접으면 아이콘 폭만 남는다.
$sidebarWidth = 200
$sidebarCollapsedWidth = 46
$sidebar = [System.Windows.Forms.Panel]::new()
$sidebar.Dock = 'Left'
$sidebar.Width = Get-UiSize $sidebarWidth
$sidebar.BackColor = $theme.Sidebar
$sidebar.Padding = New-UiPadding 0 10 0 14
$tip = [System.Windows.Forms.ToolTip]::new()

# 맨 위: 이름 겸 메뉴 접기/펼치기 버튼
$brand = New-UiNavButton 'WorkTool' 0xE700
$brand.Font = Get-UiFont 13 -Bold
$brand.ForeColor = $theme.Accent
$brand.add_Click({ Set-SidebarCollapsed (-not $script:sidebarCollapsed) })
$tip.SetToolTip($brand, '메뉴 접기/펼치기 (Ctrl+B)')

$today = [System.Windows.Forms.Label]::new()
$today.Text = (Get-Date).ToString('yyyy-MM-dd (ddd)')
$today.ForeColor = $theme.SubText
$today.AutoSize = $true
$today.Dock = 'Top'
$today.Padding = New-UiPadding 18 0 8 16

# 위에서 아래 순서로 모은 뒤, Dock=Top은 나중에 넣은 것이 위로 가므로 거꾸로 넣는다
$sidebarItems = [System.Collections.Generic.List[System.Windows.Forms.Control]]::new()
$sidebarItems.Add($brand)
$sidebarItems.Add($today)
for ($i = 0; $i -lt $pages.Count; $i++) {
    $p = $pages[$i]
    $p.Button = New-UiNavButton $p.Title $p.Icon
    $p.Button.Tag.Index = $i
    $p.Button.add_Click({ Select-Page $this.Tag.Index })
    $p.View.Dock = 'Fill'
    $p.View.Visible = $false
    $content.Controls.Add($p.View)
    $sidebarItems.Add($p.Button)
}
for ($i = $sidebarItems.Count - 1; $i -ge 0; $i--) { $sidebar.Controls.Add($sidebarItems[$i]) }

# 접으면 날짜를 숨기고, 메뉴 이름은 마우스를 올렸을 때 풍선 도움말로 보여준다
$script:sidebarCollapsed = $false
function Set-SidebarCollapsed([bool]$Collapsed) {
    $script:sidebarCollapsed = $Collapsed
    $sidebar.Width = Get-UiSize $(if ($Collapsed) { $sidebarCollapsedWidth } else { $sidebarWidth })
    $today.Visible = -not $Collapsed
    foreach ($p in $pages) { $tip.SetToolTip($p.Button, $(if ($Collapsed) { $p.Title } else { $null })) }
    $sidebar.Invalidate($true)
}

function Select-Page([int]$Index) {
    for ($i = 0; $i -lt $pages.Count; $i++) {
        $on = $i -eq $Index
        Set-UiNavSelected $pages[$i].Button $on
        $pages[$i].View.Visible = $on
    }
}

# 하단 상태 표시줄
$status = [System.Windows.Forms.StatusStrip]::new()
$status.Font = Get-UiFont 9
$status.BackColor = $theme.Back
$status.SizingGrip = $false
$statusText = [System.Windows.Forms.ToolStripStatusLabel]::new()
$statusText.Text = "F11 전체 화면 · Ctrl+B 메뉴 접기 · 글꼴 $(Get-UiFontName)"
$statusText.ForeColor = $theme.SubText
[void]$status.Items.Add($statusText)

$form.Controls.Add($sidebar)
$form.Controls.Add($status)
$form.Controls.Add($content)
# Fill은 맨 앞에 있어야 메뉴·상태 표시줄을 가리지 않는다.
$content.BringToFront()

# 단축키: F11 전체 화면, Ctrl+B 메뉴 접기
$script:fullScreen = $false
$script:prevState = 'Normal'
$form.add_KeyDown({
    param($s, $e)
    if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::B) {
        $e.SuppressKeyPress = $true
        Set-SidebarCollapsed (-not $script:sidebarCollapsed)
        return
    }
    if ($e.KeyCode -ne [System.Windows.Forms.Keys]::F11) { return }
    $e.Handled = $true
    if ($script:fullScreen) {
        $s.FormBorderStyle = 'Sizable'
        $s.WindowState = $script:prevState
        $script:fullScreen = $false
    } else {
        $script:prevState = $s.WindowState
        # 이미 최대화된 상태에서 테두리만 없애면 작업 표시줄을 덮지 않아서, 한 번 풀었다가 다시 최대화한다
        $s.WindowState = 'Normal'
        $s.FormBorderStyle = 'None'
        $s.WindowState = 'Maximized'
        $script:fullScreen = $true
    }
})

# 트레이 상주: 창의 X는 트레이로 숨기고, 완전히 끄는 건 트레이 아이콘 우클릭 → 종료.
# 숨겨져 있는 동안에도 일정 알람은 울린다.
$notify = [System.Windows.Forms.NotifyIcon]::new()
$notify.Icon = $appIcon
$notify.Text = 'WorkTool'
$trayMenu = [System.Windows.Forms.ContextMenuStrip]::new()
$trayMenu.Font = Get-UiFont 10
[void]$trayMenu.Items.Add('열기', $null, { Show-MainWindow })
[void]$trayMenu.Items.Add('종료', $null, { Exit-WorkTool })
$notify.ContextMenuStrip = $trayMenu
$notify.add_MouseDoubleClick({ Show-MainWindow })
$notify.Visible = $true

function Show-MainWindow {
    $form.Show()
    if ($form.WindowState -eq 'Minimized') { $form.WindowState = 'Normal' }
    $form.Activate()
}

$script:exiting = $false
function Exit-WorkTool {
    $script:exiting = $true
    try { $form.Close() } finally {
        if (-not $form.IsDisposed) { $script:exiting = $false; Show-MainWindow }
    }
}

$script:trayTipShown = $false
$form.add_FormClosing({
    param($s, $e)
    # 트레이 메뉴의 종료나 윈도우 종료·로그오프면 그대로 닫는다
    if ($e.Cancel) { return }
    if ($script:exiting -or $e.CloseReason -ne [System.Windows.Forms.CloseReason]::UserClosing) { return }
    $e.Cancel = $true
    $s.Hide()
    if (-not $script:trayTipShown) {
        $notify.ShowBalloonTip(3000, 'WorkTool', '트레이에서 계속 실행 중입니다. 끄려면 트레이 아이콘 우클릭 → 종료.',
            [System.Windows.Forms.ToolTipIcon]::Info)
        $script:trayTipShown = $true
    }
})

Select-Page 0
[System.Windows.Forms.Application]::Run($form)
} finally {
    try {
    # 아이콘을 치우지 않으면 툴이 끝나도 트레이에 남아 보인다
    if ($notify) { $notify.Visible = $false; $notify.Dispose() }
    # 떠 있던 알림 창도 닫는다
    foreach ($f in @([System.Windows.Forms.Application]::OpenForms)) { $f.Dispose() }
    foreach ($page in $pages) { if ($page.View -and -not $page.View.IsDisposed) { $page.View.Dispose() } }
    if ($form -and -not $form.IsDisposed) { $form.Dispose() }
    if ($appIcon) { $appIcon.Dispose() }
    } finally { $dataLock.Dispose() }
}
