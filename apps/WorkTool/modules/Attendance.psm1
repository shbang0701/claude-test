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
