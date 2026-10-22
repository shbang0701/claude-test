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
