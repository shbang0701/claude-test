# 메모 탭 (1단계: 보기 전용). 쓰기 · 자동 저장 · 메모별 잠금은 2단계에서 옮긴다.
# 저장소 memo.json (schema 1). 항목:
#   id, title, body, tags[], lock($null=잠기지 않음. 2단계에서 'dpapi' | 'password'), created, updated
# 본문은 마크다운으로 보여 준다 (core/Markdown.psm1). [[메모 제목]]은 그 메모로, #태그는 그 태그 검색으로 간다.

$ErrorActionPreference = 'Stop'

$script:schema = 1
$script:items = @()
$script:list = $null
$script:search = $null
$script:titleText = $null
$script:metaText = $null
$script:mdView = $null
$script:rawBox = $null

Register-WtSettingDefault 'memo.raw' $false   # 본문을 원문(마크다운 기호 그대로)으로 보기

function Get-WtMemoItems { Get-WtStoreItems -Name 'memo' -Schema $script:schema }

# 최근 고친 순서
function Get-WtMemoSorted([object[]]$Items) {
    @($Items | Sort-Object -Property @{ Expression = { [string]$_.updated }; Descending = $true })
}

# 제목 · 본문에 검색어가 모두 들어 있는 메모 (띄어 쓴 낱말마다 AND)
function Find-WtMemo([object[]]$Items, [string]$Query) {
    $words = @($Query -split '\s+' | Where-Object { $_ })
    if (-not $words.Count) { return $Items }
    $Items | Where-Object {
        $text = "$($_.title)`n$($_.body)"
        -not ($words | Where-Object { $text -notlike "*$_*" })
    }
}

function Format-WtMemoTime([string]$Stamp) {
    $t = [datetime]::MinValue
    if (-not [datetime]::TryParse($Stamp, [ref]$t)) { return $Stamp }
    if ($t.Date -eq [datetime]::Today) { return Format-WtDate $t "'오늘' HH:mm" }
    if ($t.Date -eq [datetime]::Today.AddDays(-1)) { return '어제' }
    Format-WtDate $t "M'/'d"
}

# 예전 memo.csv → memo.json. 가져온 개수를 돌려준다.
function Import-WtMemoLegacy([string]$Folder) {
    $rows = Import-WtLegacyCsv $Folder 'memo'
    if ($null -eq $rows) { return $null }
    $items = foreach ($r in $rows) {
        [ordered]@{
            id      = [string]$r.Id
            title   = [string]$r.Title
            body    = [string]$r.Body
            tags    = @()
            lock    = $null
            created = [string]$r.Created
            updated = [string]$r.Updated
        }
    }
    Write-WtStore -Name 'memo' -Schema $script:schema -Items @($items) -Extra @{ importedFrom = $Folder }
    @($items).Count
}

function New-WtMemoListItem($Memo) {
    $li = [System.Windows.Controls.ListBoxItem]::new()
    $li.Tag = $Memo.id
    $col = New-WtStack Vertical 2
    $head = [System.Windows.Controls.DockPanel]::new()
    $when = New-WtText (Format-WtMemoTime $Memo.updated) 'T.Sub'
    [System.Windows.Controls.DockPanel]::SetDock($when, 'Right')
    $when.Margin = [System.Windows.Thickness]::new(8, 0, 0, 0)
    [void]$head.Children.Add($when)
    $title = New-WtText $Memo.title 'T.Strong'
    $title.TextWrapping = 'NoWrap'
    $title.TextTrimming = 'CharacterEllipsis'
    [void]$head.Children.Add($title)
    [void](Add-WtChild $col $head)
    $snippet = New-WtText (Get-WtMarkdownSnippet ([string]$Memo.body)) 'T.Sub'
    $snippet.TextWrapping = 'NoWrap'
    $snippet.TextTrimming = 'CharacterEllipsis'
    [void](Add-WtChild $col $snippet)
    $li.Content = $col
    $li
}

function Show-WtMemo([string]$Id) {
    $m = $script:items | Where-Object { $_.id -eq $Id } | Select-Object -First 1
    if (-not $m) {
        $script:titleText.Text = ''
        $script:metaText.Text = ''
        $msg = if ($script:items.Count) { '왼쪽에서 메모를 고르세요.' } else { '메모가 없습니다.' }
        Set-WtMarkdownText $script:mdView $msg
        $script:rawBox.Text = $msg
        return
    }
    $script:titleText.Text = $m.title
    $script:metaText.Text = '만든 날 {0} · 고친 날 {1}' -f $m.created, $m.updated
    Set-WtMarkdownText $script:mdView ([string]$m.body)
    $script:rawBox.Text = [string]$m.body
}

# 미리보기(마크다운) · 원문 바꾸기
function Update-WtMemoMode {
    $raw = [bool](Get-WtSetting 'memo.raw')
    $script:mdView.Visibility = if ($raw) { 'Collapsed' } else { 'Visible' }
    $script:rawBox.Visibility = if ($raw) { 'Visible' } else { 'Collapsed' }
}

# 본문의 링크를 눌렀을 때. [[제목]]은 제목이 같은 메모를 열고, #태그는 그 태그로 찾는다. 처리했으면 $true
function Invoke-WtMemoLink($Link) {
    switch ($Link.Kind) {
        'wiki' {
            $name = ([string]$Link.Target -split '#')[0].Trim()
            $m = $script:items | Where-Object { ([string]$_.title).Trim() -eq $name } | Select-Object -First 1
            if (-not $m) { return $false }
            Select-WtMemoItem $m.id
            return $true
        }
        'tag' {
            $script:search.Text = "#$($Link.Target)"
            return $true
        }
    }
    $false
}

# 목록에서 메모 하나를 고른다. 검색에 걸러져 안 보이면 검색을 비운다
function Select-WtMemoItem([string]$Id) {
    $li = $script:list.Items | Where-Object { [string]$_.Tag -eq $Id } | Select-Object -First 1
    if (-not $li) {
        $script:search.Text = ''
        $li = $script:list.Items | Where-Object { [string]$_.Tag -eq $Id } | Select-Object -First 1
    }
    if ($li) {
        $script:list.SelectedItem = $li
        $li.BringIntoView()
    }
}

function Update-WtMemoList {
    $selected = if ($script:list.SelectedItem) { [string]$script:list.SelectedItem.Tag } else { $null }
    $script:list.Items.Clear()
    $shown = @(Find-WtMemo $script:items $script:search.Text)
    foreach ($m in $shown) {
        $li = New-WtMemoListItem $m
        [void]$script:list.Items.Add($li)
        if ($m.id -eq $selected) { $li.IsSelected = $true }
    }
    if (-not $script:list.SelectedItem -and $script:list.Items.Count) { $script:list.SelectedIndex = 0 }
    if (-not $script:list.Items.Count) { Show-WtMemo '' }
}

function New-WtMemoView {
    $page = New-WtPage -Fill
    $header = New-WtPageHeader '메모' '보기 전용 · 쓰기와 메모별 잠금은 2단계에서 옮깁니다'
    [System.Windows.Controls.DockPanel]::SetDock($header, 'Top')
    [void]$page.Panel.Children.Add($header)

    $grid = [System.Windows.Controls.Grid]::new()
    $c1 = [System.Windows.Controls.ColumnDefinition]::new()
    $c1.Width = [System.Windows.GridLength]::new(320)
    $c2 = [System.Windows.Controls.ColumnDefinition]::new()
    $c2.Width = [System.Windows.GridLength]::new(1, 'Star')
    $grid.ColumnDefinitions.Add($c1)
    $grid.ColumnDefinitions.Add($c2)

    # 왼쪽: 검색 + 목록
    $left = New-WtCard
    $left.Padding = [System.Windows.Thickness]::new(12)
    $left.Margin = [System.Windows.Thickness]::new(0, 0, 16, 0)
    $ldock = [System.Windows.Controls.DockPanel]::new()
    $script:search = [System.Windows.Controls.TextBox]::new()
    Set-WtStyle $script:search 'Input'
    $script:search.ToolTip = '검색 (제목 · 본문, 띄어 쓰면 모두 들어간 것)'
    $script:search.Margin = [System.Windows.Thickness]::new(0, 0, 0, 8)
    $script:search.Add_TextChanged({ Update-WtMemoList })
    [System.Windows.Controls.DockPanel]::SetDock($script:search, 'Top')
    [void]$ldock.Children.Add($script:search)
    $script:list = [System.Windows.Controls.ListBox]::new()
    Set-WtStyle $script:list 'List'
    $script:list.Add_SelectionChanged({
            if ($script:list.SelectedItem) { Show-WtMemo ([string]$script:list.SelectedItem.Tag) }
        })
    [void]$ldock.Children.Add($script:list)
    $left.Child = $ldock
    [System.Windows.Controls.Grid]::SetColumn($left, 0)
    [void]$grid.Children.Add($left)

    # 오른쪽: 읽기 (미리보기 | 원문)
    $right = New-WtCard
    $rdock = [System.Windows.Controls.DockPanel]::new()
    $head = New-WtStack Vertical 4
    $head.Margin = [System.Windows.Thickness]::new(0, 0, 0, 12)
    $mode = New-WtSegmented @(@('preview', '미리보기'), @('raw', '원문')) $(if (Get-WtSetting 'memo.raw') { 'raw' } else { 'preview' }) 'WtMemoMode' {
        param($value)
        Set-WtSetting 'memo.raw' ($value -eq 'raw')
        Update-WtMemoMode
    }
    $mode.Margin = [System.Windows.Thickness]::new(0, 0, 0, 12)
    [void](Add-WtChild $head $mode)
    $script:titleText = New-WtText '' 'T.H2'
    Set-WtResource $script:titleText 'FontSize' 'F.Number'
    $script:metaText = New-WtText '' 'T.Sub'
    [void](Add-WtChild $head $script:titleText)
    [void](Add-WtChild $head $script:metaText)
    $line = [System.Windows.Controls.Border]::new()
    Set-WtStyle $line 'Divider'
    $line.Margin = [System.Windows.Thickness]::new(0, 12, 0, 0)
    [void]$head.Children.Add($line)
    [System.Windows.Controls.DockPanel]::SetDock($head, 'Top')
    [void]$rdock.Children.Add($head)
    $body = [System.Windows.Controls.Grid]::new()
    $script:mdView = New-WtMarkdownView -Scroll -OnLink { param($link) Invoke-WtMemoLink $link }
    $script:rawBox = [System.Windows.Controls.TextBox]::new()
    Set-WtStyle $script:rawBox 'Reader'
    Set-WtResource $script:rawBox 'FontFamily' 'Font.Mono'
    [void]$body.Children.Add($script:mdView)
    [void]$body.Children.Add($script:rawBox)
    [void]$rdock.Children.Add($body)
    $right.Child = $rdock
    Update-WtMemoMode
    [System.Windows.Controls.Grid]::SetColumn($right, 1)
    [void]$grid.Children.Add($right)

    [void]$page.Panel.Children.Add($grid)
    $page.Root
}

function Update-WtMemoView {
    $script:items = @(Get-WtMemoSorted (Get-WtMemoItems))
    Update-WtMemoList
}

function New-WtMemoHomeCard {
    $items = @(Get-WtMemoSorted (Get-WtMemoItems))
    $card = New-WtCard
    $col = New-WtStack Vertical
    [void](Add-WtChild $col (New-WtCardHeader '최근 메모' '' '메모 열기' 'memo'))
    foreach ($m in ($items | Select-Object -First 3)) {
        $t = New-WtStack Vertical 2
        $title = New-WtText $m.title 'T.Strong'
        $title.TextTrimming = 'CharacterEllipsis'
        $title.TextWrapping = 'NoWrap'
        [void](Add-WtChild $t $title)
        $snip = New-WtText (Get-WtMarkdownSnippet ([string]$m.body)) 'T.Sub'
        $snip.TextTrimming = 'CharacterEllipsis'
        $snip.TextWrapping = 'NoWrap'
        [void](Add-WtChild $t $snip)
        $when = New-WtText (Format-WtMemoTime $m.updated) 'T.Sub'
        [void](Add-WtChild $col (New-WtListRow @($when, $t)))
    }
    if (-not $items.Count) { [void](Add-WtChild $col (New-WtText '메모가 없습니다.' 'T.Sub')) }
    $card.Child = $col
    $card
}

Register-WtAddon -Id 'memo' -Title '메모' -Icon 0xE70B -Order 40 `
    -NewView { New-WtMemoView } `
    -Refresh { param($view) Update-WtMemoView } `
    -HomeCard { New-WtMemoHomeCard } `
    -ImportLegacy { param($folder) Import-WtMemoLegacy $folder }

Export-ModuleMember -Function Get-WtMemoItems, Get-WtMemoSorted, Find-WtMemo, Format-WtMemoTime, Import-WtMemoLegacy
