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
