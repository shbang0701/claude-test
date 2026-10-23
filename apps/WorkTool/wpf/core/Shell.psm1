# 셸: 메인 창 틀, 왼쪽 메뉴, 제목 줄 버튼, 트레이, 단축키, 화면 전환.
# 메뉴는 Addons.psm1에 등록된 애드온 목록으로 만든다. 탭을 더하거나 빼도 이 파일은 고치지 않는다.
#
# 창은 투명도를 바꿀 수 있게 WindowStyle=None + AllowsTransparency로 만들고, 제목 줄은 직접 그린다.
# 그래서 이 창 안에는 WindowsFormsHost · WebBrowser를 넣지 않는다 (투명 창에서는 그려지지 않는다).
# 투명 창은 ClearType을 못 쓰고 회색조로 글자를 그린다. 이때 Display 모드는 획이 픽셀에 맞춰 꺾여 각져 보이므로 Ideal로 그린다.
#
# 이벤트 처리기는 나중에 디스패처가 부르므로, 만든 함수의 지역 변수를 볼 수 없다.
# 처리기 안에서는 $script: 변수와 $s(보낸 요소)의 Tag만 쓴다.

$ErrorActionPreference = 'Stop'

$script:window = $null
$script:ui = $null
$script:frame = $null
$script:themeDict = $null
$script:notify = $null
$script:views = @{}
$script:navButtons = @{}
$script:current = $null
$script:selecting = $false
$script:exiting = $false
$script:trayTipShown = $false
$script:maxed = $false
$script:fullscreen = $false
$script:restoreBounds = $null
$script:fullFromMax = $false   # 최대화 상태에서 전체 화면으로 갔는지
$script:settingBounds = $false # 코드로 창 위치 · 크기를 바꾸는 중
$script:icon = $null
$script:lastImport = $null

$script:shellXaml = @'
<Window {NS}
    Title="WorkTool" Width="1200" Height="780" MinWidth="760" MinHeight="520"
    WindowStartupLocation="CenterScreen" WindowStyle="None" AllowsTransparency="True" ResizeMode="CanResize"
    Background="Transparent" UseLayoutRounding="True" SnapsToDevicePixels="True"
    TextOptions.TextFormattingMode="Ideal"
    FontFamily="{DynamicResource Font.UI}" FontSize="{DynamicResource F.Body}" Foreground="{DynamicResource B.Text}">
  <WindowChrome.WindowChrome>
    <WindowChrome CaptionHeight="40" ResizeBorderThickness="6" GlassFrameThickness="0" CornerRadius="0" UseAeroCaptionButtons="False"/>
  </WindowChrome.WindowChrome>
  <Border x:Name="Frame" Background="{DynamicResource B.WindowBg}" BorderBrush="{DynamicResource B.FrameBorder}" BorderThickness="1">
    <Grid>
      <Grid.ColumnDefinitions>
        <ColumnDefinition x:Name="NavColumn" Width="232"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>

      <Border Grid.Column="0" Background="{DynamicResource B.Sidebar}" BorderBrush="{DynamicResource B.Border}" BorderThickness="0,0,1,0">
        <DockPanel Margin="12,0,12,12">
          <StackPanel DockPanel.Dock="Top">
            <DockPanel Height="40" Margin="-2,0,0,0">
              <Button x:Name="CollapseButton" DockPanel.Dock="Left" Style="{DynamicResource Btn.Icon}"
                      WindowChrome.IsHitTestVisibleInChrome="True" ToolTip="메뉴 접기/펼치기 (Ctrl+B)">
                <TextBlock Text="&#xE700;" Style="{DynamicResource T.Icon}"/>
              </Button>
              <Border x:Name="BrandLogo" DockPanel.Dock="Left" Width="22" Height="22" CornerRadius="5" Margin="6,0,0,0"
                      Background="{DynamicResource B.Accent}" VerticalAlignment="Center">
                <TextBlock Text="W" FontFamily="Segoe UI" FontWeight="Bold" FontSize="14" Foreground="{DynamicResource B.OnAccent}"
                           HorizontalAlignment="Center" VerticalAlignment="Center"/>
              </Border>
              <TextBlock x:Name="BrandText" Text="WorkTool" Margin="8,0,0,0" FontWeight="Bold"
                         FontSize="{DynamicResource F.H2}" VerticalAlignment="Center"/>
            </DockPanel>
            <TextBlock x:Name="TodayText" Style="{DynamicResource T.Sub}" Margin="12,6,0,14"/>
          </StackPanel>
          <StackPanel x:Name="FooterPanel" DockPanel.Dock="Bottom" Orientation="Horizontal" Margin="-2,8,0,0">
            <Button x:Name="SettingsButton" Style="{DynamicResource Btn.Icon}" ToolTip="설정 (Ctrl+,)">
              <TextBlock Text="&#xE713;" Style="{DynamicResource T.Icon}"/>
            </Button>
            <Button x:Name="HelpButton" Style="{DynamicResource Btn.Icon}" ToolTip="도움말 (F1)">
              <TextBlock Text="&#xE9CE;" Style="{DynamicResource T.Icon}"/>
            </Button>
          </StackPanel>
          <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" Focusable="False">
            <StackPanel x:Name="NavPanel"/>
          </ScrollViewer>
        </DockPanel>
      </Border>

      <DockPanel Grid.Column="1">
        <StackPanel DockPanel.Dock="Top" Orientation="Horizontal" HorizontalAlignment="Right" Height="40">
          <Button x:Name="MinButton" Style="{DynamicResource Btn.Chrome}" WindowChrome.IsHitTestVisibleInChrome="True" ToolTip="최소화">
            <TextBlock Text="&#xE921;" Style="{DynamicResource T.Icon}" FontSize="10"/>
          </Button>
          <Button x:Name="MaxButton" Style="{DynamicResource Btn.Chrome}" WindowChrome.IsHitTestVisibleInChrome="True" ToolTip="최대화">
            <TextBlock x:Name="MaxGlyph" Text="&#xE922;" Style="{DynamicResource T.Icon}" FontSize="10"/>
          </Button>
          <Button x:Name="CloseButton" Style="{DynamicResource Btn.Close}" WindowChrome.IsHitTestVisibleInChrome="True" ToolTip="닫기">
            <TextBlock Text="&#xE8BB;" Style="{DynamicResource T.Icon}" FontSize="10"/>
          </Button>
        </StackPanel>
        <ContentControl x:Name="PageHost" Focusable="False"/>
      </DockPanel>
    </Grid>
  </Border>
</Window>
'@

function Get-WtWindow { $script:window }

# ---- 화면 전환 --------------------------------------------------------------------------------

# 스크립트 블록의 출력 중 화면 요소만 고른다 (실수로 섞여 나온 숫자 등은 버린다)
function Get-WtElements($Output) {
    @($Output | Where-Object { $_ -is [System.Windows.UIElement] })
}

function Get-WtPageView([string]$Id) {
    switch ($Id) {
        '_settings' { return @{ View = (New-WtSettingsPage); Addon = $null } }
        '_help' { return @{ View = (New-WtHelpPage); Addon = $null } }
    }
    if (-not $script:views.ContainsKey($Id)) {
        $a = Get-WtAddon -Id $Id
        if (-not $a) { return $null }
        $view = Get-WtElements (& $a.NewView) | Select-Object -Last 1
        if (-not $view) { throw "애드온 '$Id'가 화면을 돌려주지 않았습니다." }
        $script:views[$Id] = @{ View = $view; Addon = $a }
    }
    $script:views[$Id]
}

function Select-WtPage([string]$Id) {
    if ($script:selecting -or -not $Id) { return }
    $script:selecting = $true
    try {
        if (-not (Save-WtPendingChanges)) { return }
        $page = Get-WtPageView $Id
        if (-not $page) { return }
        if ($page.Addon -and $page.Addon.Refresh) { [void](& $page.Addon.Refresh $page.View) }
        $script:ui.PageHost.Content = $page.View
        foreach ($k in @($script:navButtons.Keys)) { $script:navButtons[$k].Button.IsChecked = ($k -eq $Id) }
        $script:current = $Id
    } catch {
        foreach ($k in @($script:navButtons.Keys)) { $script:navButtons[$k].Button.IsChecked = ($k -eq $script:current) }
        Show-WtError $_ "화면 열기: $Id"
    } finally {
        foreach ($k in @($script:navButtons.Keys)) { $script:navButtons[$k].Button.IsChecked = ($k -eq $script:current) }
        $script:selecting = $false
    }
}

# 데이터를 다시 가져온 뒤 현재 화면을 갱신한다. 숨긴 화면도 재진입할 때 갱신한다.
function Reset-WtViews {
    if (-not (Save-WtPendingChanges)) { return }
    if ($script:current) {
        $id = $script:current
        $script:current = $null
        Select-WtPage $id
    }
}

# ---- 왼쪽 메뉴 --------------------------------------------------------------------------------

function Update-WtNav {
    $panel = $script:ui.NavPanel
    $panel.Children.Clear()
    $script:navButtons = @{}
    $i = 0
    foreach ($a in @(Get-WtAddon)) {
        $i++
        $rb = [System.Windows.Controls.RadioButton]::new()
        Set-WtStyle $rb 'Nav'
        $rb.GroupName = 'WtNav'
        $rb.Tag = $a.Id
        $rb.ToolTip = if ($i -le 9) { "$($a.Title) (Ctrl+$i)" } else { $a.Title }
        $row = [System.Windows.Controls.StackPanel]::new()
        $row.Orientation = 'Horizontal'
        $icon = New-WtIcon $a.Icon
        $icon.Width = 20
        $label = [System.Windows.Controls.TextBlock]::new()
        $label.Text = $a.Title
        $label.Margin = [System.Windows.Thickness]::new(12, 0, 0, 0)
        $label.VerticalAlignment = 'Center'
        [void]$row.Children.Add($icon)
        [void]$row.Children.Add($label)
        $rb.Content = $row
        $rb.IsChecked = ($a.Id -eq $script:current)
        $rb.Add_Checked({ param($s, $e) Select-WtPage ([string]$s.Tag) })
        [void]$panel.Children.Add($rb)
        $script:navButtons[$a.Id] = @{ Button = $rb; Label = $label }
    }
    Update-WtCollapsed
}

function Update-WtCollapsed {
    $c = [bool](Get-WtSetting 'nav.collapsed')
    $script:ui.NavColumn.Width = [System.Windows.GridLength]::new($(if ($c) { 68 } else { 232 }))
    $vis = if ($c) { 'Collapsed' } else { 'Visible' }
    $script:ui.BrandText.Visibility = $vis
    $script:ui.BrandLogo.Visibility = $vis
    $script:ui.TodayText.Visibility = $vis
    foreach ($n in $script:navButtons.Values) { $n.Label.Visibility = $vis }
    $script:ui.FooterPanel.Orientation = if ($c) { 'Vertical' } else { 'Horizontal' }
}

function Switch-WtCollapsed { Set-WtSetting 'nav.collapsed' (-not [bool](Get-WtSetting 'nav.collapsed')) }

# ---- 창 크기 · 투명도 -------------------------------------------------------------------------

function Set-WtOpacity([double]$Percent = -1) {
    if ($Percent -lt 0) { $Percent = [double](Get-WtSetting 'ui.opacity') }
    $script:window.Opacity = [Math]::Max(0.6, [Math]::Min(1.0, $Percent / 100))
}

# 창이 있는 화면의 작업 영역(또는 전체) 크기를 WPF 단위로
function Get-WtScreenRect([switch]$Full) {
    $hwnd = [System.Windows.Interop.WindowInteropHelper]::new($script:window).Handle
    $screen = [System.Windows.Forms.Screen]::FromHandle($hwnd)
    $r = if ($Full) { $screen.Bounds } else { $screen.WorkingArea }
    $m = [System.Windows.PresentationSource]::FromVisual($script:window).CompositionTarget.TransformFromDevice
    $tl = $m.Transform([System.Windows.Point]::new($r.Left, $r.Top))
    $br = $m.Transform([System.Windows.Point]::new($r.Right, $r.Bottom))
    [System.Windows.Rect]::new($tl, $br)
}

function Get-WtBounds { [System.Windows.Rect]::new($script:window.Left, $script:window.Top, $script:window.Width, $script:window.Height) }

function Set-WtBounds([System.Windows.Rect]$R) {
    $script:settingBounds = $true
    try {
        $script:window.Left = $R.Left
        $script:window.Top = $R.Top
        $script:window.Width = $R.Width
        $script:window.Height = $R.Height
    } finally {
        $script:settingBounds = $false
    }
}

function Update-WtMaxGlyph {
    $script:ui.MaxGlyph.Text = [string][char]$(if ($script:maxed -or $script:fullscreen) { 0xE923 } else { 0xE922 })
    $script:ui.MaxButton.ToolTip = if ($script:maxed -or $script:fullscreen) { '이전 크기로' } else { '최대화' }
    $script:ui.Frame.BorderThickness = [System.Windows.Thickness]::new($(if ($script:maxed -or $script:fullscreen) { 0 } else { 1 }))
}

# 최대화는 WindowState 대신 작업 영역 크기로 맞춘다 (테두리 없는 창을 Maximized로 두면 작업 표시줄을 덮는다)
function Switch-WtMaximize {
    if ($script:fullscreen) { Switch-WtFullScreen; return }
    if ($script:maxed) {
        Set-WtBounds $script:restoreBounds
        $script:maxed = $false
    } else {
        $script:restoreBounds = Get-WtBounds
        Set-WtBounds (Get-WtScreenRect)
        $script:maxed = $true
    }
    Update-WtMaxGlyph
}

function Switch-WtFullScreen {
    if ($script:fullscreen) {
        $script:fullscreen = $false
        if ($script:fullFromMax) {
            Set-WtBounds (Get-WtScreenRect)
            $script:maxed = $true
        } else {
            Set-WtBounds $script:restoreBounds
            $script:maxed = $false
        }
    } else {
        $script:fullFromMax = $script:maxed
        if (-not $script:maxed) { $script:restoreBounds = Get-WtBounds }
        Set-WtBounds (Get-WtScreenRect -Full)
        $script:fullscreen = $true
    }
    Update-WtMaxGlyph
}

# ---- 트레이 · 종료 ----------------------------------------------------------------------------

# 앱 아이콘 그림 하나(PNG): 강조색 둥근 사각형에 흰 W. 파일 없이 그린다
function New-WtAppIconPng([int]$Size, [string]$Hex) {
    $dv = [System.Windows.Media.DrawingVisual]::new()
    $dc = $dv.RenderOpen()
    $bg = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($Hex))
    $r = [Math]::Max(2, $Size * 0.22)
    $dc.DrawRoundedRectangle($bg, $null, [System.Windows.Rect]::new(0, 0, $Size, $Size), $r, $r)
    $tf = [System.Windows.Media.Typeface]::new([System.Windows.Media.FontFamily]::new('Segoe UI'), [System.Windows.FontStyles]::Normal,
        [System.Windows.FontWeights]::Bold, [System.Windows.FontStretches]::Normal)
    $ft = [System.Windows.Media.FormattedText]::new('W', [System.Globalization.CultureInfo]::InvariantCulture,
        [System.Windows.FlowDirection]::LeftToRight, $tf, $Size * 0.62, [System.Windows.Media.Brushes]::White, 1.0)
    $geo = $ft.BuildGeometry([System.Windows.Point]::new(0, 0))
    $b = $geo.Bounds
    # 글자 모양(잉크)의 가운데를 그림 가운데에 맞춘다. 작은 크기에서 흐리지 않게 정수 위치로
    $dc.PushTransform([System.Windows.Media.TranslateTransform]::new([Math]::Round(($Size - $b.Width) / 2 - $b.X), [Math]::Round(($Size - $b.Height) / 2 - $b.Y)))
    $dc.DrawGeometry([System.Windows.Media.Brushes]::White, $null, $geo)
    $dc.Pop()
    $dc.Close()
    $rtb = [System.Windows.Media.Imaging.RenderTargetBitmap]::new($Size, $Size, 96, 96, [System.Windows.Media.PixelFormats]::Pbgra32)
    $rtb.Render($dv)
    $enc = [System.Windows.Media.Imaging.PngBitmapEncoder]::new()
    $enc.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($rtb))
    $ms = [System.IO.MemoryStream]::new()
    $enc.Save($ms)
    , $ms.ToArray()
}

# 앱 아이콘(.ico 내용): 크기별 그림을 하나로 묶는다. 트레이(16~24) · 제목 줄 · 작업 표시줄(24~48) · 큰 아이콘(256)에서
# Windows가 화면 배율에 맞는 크기를 고르므로 늘리거나 줄여서 흐려지지 않는다.
function New-WtAppIconBytes {
    $hex = (Get-WtThemeColors ([string](Get-WtSetting 'theme.color'))).Accent
    $sizes = @(16, 20, 24, 32, 40, 48, 64, 256)
    $pngs = @(foreach ($s in $sizes) { , (New-WtAppIconPng $s $hex) })
    $ms = [System.IO.MemoryStream]::new()
    $w = [System.IO.BinaryWriter]::new($ms)
    $w.Write([uint16]0); $w.Write([uint16]1); $w.Write([uint16]$sizes.Count)   # ICO 머리: 예약 · 종류(1=아이콘) · 개수
    $offset = 6 + 16 * $sizes.Count
    for ($i = 0; $i -lt $sizes.Count; $i++) {
        $s = $sizes[$i]
        $w.Write([byte]($s % 256)); $w.Write([byte]($s % 256))   # 너비 · 높이 (256은 0)
        $w.Write([byte]0); $w.Write([byte]0)                     # 색 수 · 예약
        $w.Write([uint16]1); $w.Write([uint16]32)                # 면 · 색 깊이
        $w.Write([uint32]$pngs[$i].Length); $w.Write([uint32]$offset)
        $offset += $pngs[$i].Length
    }
    foreach ($p in $pngs) { $w.Write([byte[]]$p) }
    $w.Flush()
    , $ms.ToArray()
}

# 창 · 트레이 아이콘을 지금 테마 색으로 (다시) 만든다
function Update-WtAppIcon {
    $bytes = New-WtAppIconBytes
    $old = $script:icon
    $script:icon = [System.Drawing.Icon]::new([System.IO.MemoryStream]::new($bytes), [System.Windows.Forms.SystemInformation]::SmallIconSize)
    # 아이콘 해석기로 읽은 그림을 주면 WPF가 그 안에서 제목 줄 · 작업 표시줄에 맞는 크기를 고른다
    $dec = [System.Windows.Media.Imaging.IconBitmapDecoder]::new([System.IO.MemoryStream]::new($bytes),
        [System.Windows.Media.Imaging.BitmapCreateOptions]::None, [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad)
    # 창이 생성된 뒤에도 다시 설정해 Windows가 창의 정식 아이콘을 갱신하게 한다.
    $script:window.Icon = $null
    $script:window.Icon = ($dec.Frames | Where-Object PixelWidth -eq 32 | Select-Object -First 1)
    if (-not $script:window.TaskbarItemInfo) { $script:window.TaskbarItemInfo = [System.Windows.Shell.TaskbarItemInfo]::new() }
    $script:window.TaskbarItemInfo.Description = 'WorkTool'
    $script:window.TaskbarItemInfo.Overlay = $null
    if ($script:notify) { $script:notify.Icon = $script:icon }
    if ($old) { $old.Dispose() }
}

function Show-WtMainWindow {
    $script:window.Show()
    if ($script:window.WindowState -eq 'Minimized') { $script:window.WindowState = 'Normal' }
    [void]$script:window.Activate()
}

# 정리 처리 전에 모든 저장을 확인한다. 실패하면 타이머와 창을 그대로 유지한다.
function Save-WtPendingChanges {
    foreach ($a in @(Get-WtAddon -All)) {
        if (-not $a.SavePending) { continue }
        try {
            if ((& $a.SavePending) -eq $false) {
                Show-WtMessage '저장하지 못한 내용이 있습니다. 메모의 저장 오류를 해결한 뒤 다시 시도하세요.' '저장 확인' 'Warning'
                return $false
            }
        } catch { Show-WtError $_ "저장 확인: $($a.Title)"; return $false }
    }
    $true
}

# 저장에 성공한 뒤에만 트레이·타이머를 정리하고 메시지 돌리기를 멈춘다.
function Stop-WtApp {
    if (-not $script:exiting) {
        if (-not (Save-WtPendingChanges)) { return $false }
        foreach ($a in @(Get-WtAddon -All)) {
            if ($a.OnExit) { try { [void](& $a.OnExit) } catch { Show-WtError $_ "끝내기 전 저장: $($a.Title)" } }
        }
    }
    $script:exiting = $true
    try { $script:notify.Visible = $false; $script:notify.Dispose() } catch { }
    try { if ($script:icon) { $script:icon.Dispose() } } catch { }
    if ($script:frame) { $script:frame.Continue = $false }
    $true
}

function Exit-WtApp {
    if (-not (Stop-WtApp)) { return }
    try { $script:window.Close() } catch { }
}

function Update-WtCloseTip {
    $script:ui.CloseButton.ToolTip = if ([bool](Get-WtSetting 'window.closeToTray')) { '닫기 (트레이에서 계속 실행)' } else { '닫기' }
}

function New-WtTrayIcon([System.Drawing.Icon]$Icon) {
    $n = [System.Windows.Forms.NotifyIcon]::new()
    $n.Icon = $Icon
    $n.Text = 'WorkTool'
    $menu = [System.Windows.Forms.ContextMenuStrip]::new()
    [void]$menu.Items.Add('열기', $null, { Show-WtMainWindow })
    [void]$menu.Items.Add('종료', $null, { Exit-WtApp })
    $n.ContextMenuStrip = $menu
    $n.Add_MouseDoubleClick({ Show-WtMainWindow })
    $n.Visible = $true
    $n
}

# ---- 예전 데이터 가져오기 ----------------------------------------------------------------------

# 예전 WorkTool(WinForms) data 폴더에서 탭마다 데이터를 가져온다.
# -Force가 없으면 아직 저장소가 없는 탭만 가져온다 (처음 실행할 때).
function Import-WtLegacyData([switch]$Force) {
    if ($Force -and -not (Save-WtPendingChanges)) { throw '저장하지 못한 메모가 있어 가져오기를 취소했습니다.' }
    $folder = Find-WtLegacyFolder
    $result = @{ Folder = $folder; Count = 0; Tabs = @(); At = $null }
    if (-not $folder) { return $result }
    foreach ($a in @(Get-WtAddon -All)) {
        if (-not $a.ImportLegacy) { continue }
        if (-not $Force -and (Test-WtStore $a.Id)) { continue }
        $n = & $a.ImportLegacy $folder | Select-Object -Last 1
        if ($null -ne $n) {
            $result.Count += [int]$n
            $result.Tabs += "$($a.Title) $n"
        }
    }
    if ($result.Tabs.Count) {
        $result.At = Get-Date
        $script:lastImport = $result
    }
    $result
}

function Get-WtLastImport { $script:lastImport }

# ---- 시작 -------------------------------------------------------------------------------------

function Start-WtApp {
    param([Parameter(Mandatory)][string]$AppRoot)

    $script:exiting = $false
    $script:notify = $null
    $script:views = @{}
    $script:current = $null
    $script:maxed = $false
    $script:fullscreen = $false
    Set-WtAppRoot $AppRoot
    $dataLock = Open-WtDataLock
    try {
    Initialize-WtSettings
    $script:themeDict = New-WtThemeDictionary

    $w = ConvertFrom-WtXaml $script:shellXaml
    $script:window = $w
    $script:ui = Get-WtNamedElements $w $script:shellXaml
    $w.Resources.MergedDictionaries.Add($script:themeDict)
    $w.Resources.MergedDictionaries.Add((Get-WtStyleDictionary))

    Update-WtAppIcon
    $w.Add_ContentRendered({ Update-WtAppIcon })
    $script:ui.TodayText.Text = Format-WtDate (Get-Date) 'yyyy년 M월 d일 dddd'

    # 제목 줄 · 메뉴 아래 버튼
    $script:ui.CollapseButton.Add_Click({ Switch-WtCollapsed })
    $script:ui.MinButton.Add_Click({ $script:window.WindowState = 'Minimized' })
    $script:ui.MaxButton.Add_Click({ Switch-WtMaximize })
    $script:ui.CloseButton.Add_Click({ $script:window.Close() })
    $script:ui.SettingsButton.Add_Click({ Select-WtPage '_settings' })
    $script:ui.HelpButton.Add_Click({ Select-WtPage '_help' })

    # 제목 줄을 두 번 누르면 Windows가 Maximized로 바꾼다 → 작업 영역 최대화로 바꿔 준다
    $w.Add_StateChanged({
            if ($script:window.WindowState -eq 'Maximized') {
                $script:window.WindowState = 'Normal'
                Switch-WtMaximize
            }
        })

    # 최대화 · 전체 화면 상태에서 제목 줄을 끌어 옮기면 그 표시를 푼다 (다음 최대화 버튼이 바로 최대화하도록)
    $w.Add_LocationChanged({
            if ($script:settingBounds -or -not ($script:maxed -or $script:fullscreen)) { return }
            $script:maxed = $false
            $script:fullscreen = $false
            Update-WtMaxGlyph
        })

    # X(닫기)는 앱을 끝낸다. 설정에서 '트레이로 숨기기'를 켜면 트레이로 숨기고, 끄는 건 트레이 아이콘 우클릭 → 종료
    $w.Add_Closing({
            param($s, $e)
            if ($script:exiting) { return }
            if (-not [bool](Get-WtSetting 'window.closeToTray')) { $e.Cancel = -not (Stop-WtApp); return }
            $e.Cancel = $true
            $s.Hide()
            if (-not $script:trayTipShown) {
                $script:notify.ShowBalloonTip(3000, 'WorkTool', '트레이에서 계속 실행 중입니다. 끄려면 트레이 아이콘 우클릭 → 종료.',
                    [System.Windows.Forms.ToolTipIcon]::Info)
                $script:trayTipShown = $true
            }
        })

    $w.Add_PreviewKeyDown({
            param($s, $e)
            try {
                $ctrl = ([System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Control) -ne 0
                $key = if ($e.Key -eq 'System') { $e.SystemKey } else { $e.Key }
                if ($ctrl -and $key -eq 'B') { Switch-WtCollapsed; $e.Handled = $true; return }
                if ($key -eq 'F11') { Switch-WtFullScreen; $e.Handled = $true; return }
                if ($key -eq 'Escape' -and $script:fullscreen) { Switch-WtFullScreen; $e.Handled = $true; return }
                if ($key -eq 'F1') { Select-WtPage '_help'; $e.Handled = $true; return }
                if ($ctrl -and $key -eq 'OemComma') { Select-WtPage '_settings'; $e.Handled = $true; return }
                if ($ctrl -and [int]$key -ge [int][System.Windows.Input.Key]::D1 -and [int]$key -le [int][System.Windows.Input.Key]::D9) {
                    $ids = @(Get-WtAddon | ForEach-Object { $_.Id })
                    $idx = [int]$key - [int][System.Windows.Input.Key]::D1
                    if ($idx -lt $ids.Count) { Select-WtPage $ids[$idx] }
                    $e.Handled = $true
                }
            } catch { Show-WtError $_ '단축키' }
        })

    if ($global:WorkToolUnhandledHandler) {
        try { $w.Dispatcher.remove_UnhandledException($global:WorkToolUnhandledHandler) } catch { }
    }
    $global:WorkToolUnhandledHandler = [System.Windows.Threading.DispatcherUnhandledExceptionEventHandler] {
        param($s, $e)
        Show-WtError $e.Exception '처리되지 않은 오류'
        $e.Handled = $true
    }
    $w.Dispatcher.add_UnhandledException($global:WorkToolUnhandledHandler)

    # 설정이 바뀌면 바로 반영한다
    Register-WtSettingHandler {
        param($key, $value)
        switch -Wildcard ($key) {
            'theme.color' { Update-WtThemeDictionary $script:themeDict; Update-WtAppIcon }
            'ui.fontScale' { Update-WtThemeDictionary $script:themeDict }
            'ui.opacity' { Set-WtOpacity }
            'nav.collapsed' { Update-WtCollapsed }
            'window.closeToTray' { Update-WtCloseTip }
            'addons.disabled' { Update-WtNav }
        }
    }

    # 처음 실행이면 예전 WorkTool 데이터를 가져온다 (이미 가져온 탭은 건너뜀)
    try { [void](Import-WtLegacyData) } catch { Show-WtError $_ '예전 데이터 가져오기' }

    Update-WtNav
    Update-WtCloseTip
    Set-WtOpacity
    $script:notify = New-WtTrayIcon $script:icon

    # 창이 닫힐 때까지 메시지를 돌린다. Dispatcher를 끄지 않으므로 같은 창에서 다시 실행해도 된다.
    $script:frame = [System.Windows.Threading.DispatcherFrame]::new()
    $w.Show()
    # 실행 파일이 콘솔을 최소화해서 띄우면(start /min) 첫 창도 최소화돼서 뜬다 → 원래 크기로 앞에 띄운다
    if ($w.WindowState -eq 'Minimized') { $w.WindowState = 'Normal' }
    [void]$w.Activate()
    $first = @(Get-WtAddon)[0]
    if ($first) { Select-WtPage $first.Id } else { Select-WtPage '_settings' }
    if (-not $script:exiting) { [System.Windows.Threading.Dispatcher]::PushFrame($script:frame) }
    } finally {
        try {
            if ($script:notify) { $script:notify.Visible = $false; $script:notify.Dispose() }
            if ($script:icon) { $script:icon.Dispose() }
        } finally { $dataLock.Dispose() }
    }
}

Export-ModuleMember -Function Get-WtWindow, Select-WtPage, Reset-WtViews, Update-WtNav, Switch-WtCollapsed, Set-WtOpacity,
    Switch-WtMaximize, Switch-WtFullScreen, Show-WtMainWindow, Exit-WtApp, Import-WtLegacyData, Get-WtLastImport, Start-WtApp
