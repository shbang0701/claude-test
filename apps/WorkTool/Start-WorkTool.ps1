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

Import-Module (Join-Path $root 'modules\Storage.psm1') -Force
Import-Module (Join-Path $root 'modules\UiHelper.psm1') -Force

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
foreach ($f in $features) {
    Import-Module (Join-Path $root "modules\$($f.Module).psm1") -Force
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
    # 트레이 메뉴의 종료나 윈도우 종료·로그오프면 그대로 닫는다 (메모를 저장하지 못했으면 메모 화면이 끝낼지 묻는다)
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
