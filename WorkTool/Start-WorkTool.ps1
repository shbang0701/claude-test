# WorkTool 진입점. 기능 모듈을 불러와 탭으로 붙이고 메인 창을 띄운다.
# Windows PowerShell 5.1과 PowerShell 7 양쪽에서 돌게 작성한다.

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# 아래 설정은 창을 만들기 전에 해야 한다. 같은 창에서 이미 창을 띄웠으면 실패하므로 무시한다.
[System.Windows.Forms.Application]::EnableVisualStyles()
try { [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false) } catch { }
# 고해상도 화면에서 글자가 흐리지 않게 한다. PowerShell 7에만 있는 기능이다.
if ([System.Windows.Forms.Application].GetMethod('SetHighDpiMode')) {
    [void][System.Windows.Forms.Application]::SetHighDpiMode('SystemAware')
}

# 기능 모듈 목록. 기능을 추가할 때 modules\<이름>.psm1을 만들고 여기 한 줄 넣는다.
# 각 모듈은 New-<이름>Tab 함수로 TabPage를 돌려준다.
$featureModules = @(
    # 'Memo'
)

# 색과 글꼴
$navy      = [System.Drawing.Color]::FromArgb(31, 58, 95)
$navyLight = [System.Drawing.Color]::FromArgb(180, 191, 203)
$back      = [System.Drawing.Color]::FromArgb(243, 245, 248)
$gray      = [System.Drawing.Color]::FromArgb(110, 120, 132)
$font      = New-Object System.Drawing.Font('Malgun Gothic', 10)
$titleFont = New-Object System.Drawing.Font('Malgun Gothic', 14, [System.Drawing.FontStyle]::Bold)

# 메인 창. 크기는 화면 작업 영역 비율로 잡아서 배율과 상관없이 적당하게 뜨게 한다.
$area = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
$form = New-Object System.Windows.Forms.Form
$form.Text = 'WorkTool'
$form.Font = $font
$form.BackColor = $back
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size([int]($area.Width * 0.6), [int]($area.Height * 0.75))
$form.MinimumSize = New-Object System.Drawing.Size([int]($area.Width * 0.4), [int]($area.Height * 0.5))

# 상단 머리글: 왼쪽 이름, 오른쪽 오늘 날짜
$header = New-Object System.Windows.Forms.TableLayoutPanel
$header.Dock = 'Top'
$header.AutoSize = $true
$header.BackColor = $navy
$header.Padding = New-Object System.Windows.Forms.Padding(14, 8, 14, 8)
$header.ColumnCount = 2
[void]$header.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$header.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))

$title = New-Object System.Windows.Forms.Label
$title.Text = 'WorkTool'
$title.Font = $titleFont
$title.ForeColor = [System.Drawing.Color]::White
$title.AutoSize = $true
$title.Anchor = 'Left'

$today = New-Object System.Windows.Forms.Label
$today.Text = (Get-Date).ToString('yyyy-MM-dd (ddd)')
$today.ForeColor = $navyLight
$today.AutoSize = $true
$today.Anchor = 'Right'

$header.Controls.Add($title, 0, 0)
$header.Controls.Add($today, 1, 0)

# 가운데: 탭
$body = New-Object System.Windows.Forms.Panel
$body.Dock = 'Fill'
$body.Padding = New-Object System.Windows.Forms.Padding(10)

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = 'Fill'
$tabs.Padding = New-Object System.Drawing.Point(16, 6)
$body.Controls.Add($tabs)

$homeTab = New-Object System.Windows.Forms.TabPage
$homeTab.Text = '홈'
$homeTab.BackColor = [System.Drawing.Color]::White
$homeTab.Padding = New-Object System.Windows.Forms.Padding(16)

$homeText = New-Object System.Windows.Forms.Label
$homeText.Text = '기능 모듈을 붙이면 옆에 탭이 생긴다.'
$homeText.ForeColor = $gray
$homeText.AutoSize = $true
$homeText.Dock = 'Top'
$homeTab.Controls.Add($homeText)
$tabs.TabPages.Add($homeTab)

# 하단 상태 표시줄
$status = New-Object System.Windows.Forms.StatusStrip
$status.Font = $font
$statusText = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusText.Text = '준비'
[void]$status.Items.Add($statusText)

$form.Controls.Add($header)
$form.Controls.Add($status)
$form.Controls.Add($body)
# Fill은 맨 앞에 있어야 머리글·상태 표시줄을 가리지 않는다.
$body.BringToFront()

foreach ($name in $featureModules) {
    Import-Module (Join-Path $root "modules\$name.psm1") -Force
    $tabs.TabPages.Add((& "New-${name}Tab"))
}

[System.Windows.Forms.Application]::Run($form)
