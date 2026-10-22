<#
.SYNOPSIS
    WorkTool 새 버전(WPF) 1단계 미리보기를 띄운다.

.DESCRIPTION
    - 창 틀 · 왼쪽 메뉴(애드온) · 테마(파스텔 7색 · 투명도 · 글자 크기) · 설정 · 도움말이 들어 있다.
    - 처음 실행하면 예전 WorkTool(WinForms)의 data 폴더를 찾아 메모 · 일정 · 근태를 읽어 와서 보여 준다.
      예전 데이터는 바꾸지 않는다. 새 버전 데이터는 문서\WorkTool 에 저장된다.
    - 편집 · 알람은 2단계에서 옮긴다. 그때까지는 예전 WorkTool을 계속 쓴다.

.EXAMPLE
    .\Start-WorkTool.ps1
    새 버전을 띄운다. 창의 X는 트레이로 숨기고, 완전히 끄려면 트레이 아이콘 우클릭 → 종료.
#>
# WorkTool 새 버전(WPF) — 파일 하나로 묶은 버전 (자동 생성)
# 이 파일은 직접 고치지 않는다. apps/WorkTool/wpf의 원본을 고친 뒤 wpf/build/Build-Bundle.ps1로 다시 만든다.
#Requires -Version 7.0
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

# WPF 창은 STA 스레드에서만 띄울 수 있다
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    Write-Host 'WorkTool은 STA 모드에서 실행해야 합니다. 다음처럼 실행하세요:' -ForegroundColor Yellow
    Write-Host "  pwsh -STA -File `"$PSCommandPath`""
    return
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
Add-Type -AssemblyName System.Windows.Forms, System.Drawing

#region 모듈 (묶음 파일: build\Build-Bundle.ps1이 넣음)
# ---- core\Storage.psm1 ----
New-Module -Name 'WorkTool.Storage' -ScriptBlock {
# 저장 담당. 기능 모듈은 파일을 직접 읽고 쓰지 않고 이 모듈만 거친다.
# - 데이터는 앱 폴더 밖(기본: 문서\WorkTool)에 둔다. 앱을 새 버전으로 덮어써도 데이터는 그대로다.
# - 저장소(store) 하나 = JSON 파일 하나. 파일마다 schema 번호를 붙여, 형식이 바뀌면 읽을 때 한 단계씩 변환한다.
# - 쓸 때는 임시 파일에 먼저 쓰고 바꿔치기한다. 직전 파일은 .bak으로 하나 남긴다.
# - 읽기 실패는 항상 예외로 올린다. 빈 결과로 넘기면 다음 저장에서 기존 데이터를 덮어쓰기 때문이다.

$ErrorActionPreference = 'Stop'

$script:appRoot = $null
$script:dataRoot = $null
$script:utf8 = [System.Text.UTF8Encoding]::new($false)

# 진입 스크립트가 자기 위치를 알려 준다. 묶음 파일이든 원본 폴더든 같은 방식으로 쓴다.
function Set-WtAppRoot([string]$Path) {
    $script:appRoot = $Path
    $script:dataRoot = $null
}

function Get-WtAppRoot { $script:appRoot }

# 데이터 폴더. 앱 폴더에 datapath.txt가 있으면 그 첫 줄의 경로, 없으면 문서\WorkTool
function Get-WtDataRoot {
    if ($script:dataRoot) { return $script:dataRoot }
    $path = $null
    if ($script:appRoot) {
        $override = Join-Path $script:appRoot 'datapath.txt'
        if (Test-Path -LiteralPath $override) {
            $line = Get-Content -LiteralPath $override -TotalCount 1 -Encoding utf8
            if ($line -and $line.Trim()) { $path = [Environment]::ExpandEnvironmentVariables($line.Trim()) }
        }
    }
    if (-not $path) {
        $docs = [Environment]::GetFolderPath('MyDocuments')
        if (-not $docs) { $docs = $HOME }
        $path = Join-Path $docs 'WorkTool'
    }
    if (-not (Test-Path -LiteralPath $path)) { New-Item -ItemType Directory -Path $path -Force | Out-Null }
    $script:dataRoot = (Resolve-Path -LiteralPath $path).ProviderPath
    $script:dataRoot
}

# JSON 파일 하나를 읽는다. 없으면 $null
function Read-WtJsonFile([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    [System.IO.File]::ReadAllText($Path, $script:utf8) | ConvertFrom-Json -Depth 20
}

# JSON 파일 하나를 바꿔치기로 쓴다. 직전 파일은 <파일>.bak
function Write-WtJsonFile([string]$Path, $Data) {
    $json = ConvertTo-Json -InputObject $Data -Depth 20
    $tmp = "$Path.tmp"
    [System.IO.File]::WriteAllText($tmp, $json, $script:utf8)
    if (Test-Path -LiteralPath $Path) {
        [System.IO.File]::Replace($tmp, $Path, "$Path.bak")
    } else {
        Move-Item -LiteralPath $tmp -Destination $Path
    }
}

function Get-WtStorePath {
    param([Parameter(Mandatory)][ValidatePattern('^[a-z0-9-]+$')][string]$Name)
    Join-Path (Get-WtDataRoot) "$Name.json"
}

function Test-WtStore([string]$Name) { Test-Path -LiteralPath (Get-WtStorePath $Name) }

# 저장소 하나를 읽는다. 파일이 없으면 $null.
# $Migrate는 { param($doc, $from) ... $doc } 형태로, schema $from 문서를 $from+1 형식으로 바꿔 돌려준다.
# 변환하기 전 원본은 <이름>.json.schema<번호>.bak으로 남긴다.
function Read-WtStore {
    param(
        [Parameter(Mandatory)][string]$Name,
        [int]$Schema = 1,
        [scriptblock]$Migrate
    )
    $path = Get-WtStorePath $Name
    $doc = Read-WtJsonFile $path
    if ($null -eq $doc) { return $null }
    $from = [int]$doc.schema
    if ($from -gt $Schema) {
        throw "$Name.json은 더 새 버전(schema $from)에서 저장됐습니다. WorkTool을 최신으로 바꿔 주세요."
    }
    if ($from -lt $Schema) {
        if (-not $Migrate) { throw "$Name.json을 schema $from에서 $Schema(으)로 바꾸는 방법이 없습니다." }
        Copy-Item -LiteralPath $path -Destination "$path.schema$from.bak" -Force
        while ($from -lt $Schema) {
            $doc = & $Migrate $doc $from
            $from++
            $doc.schema = $from
        }
        Write-WtStore -Name $Name -Schema $Schema -Items @($doc.items)
    }
    $doc
}

# 저장소의 항목만 배열로 돌려준다. 파일이 없으면 빈 배열
function Get-WtStoreItems {
    param([Parameter(Mandatory)][string]$Name, [int]$Schema = 1, [scriptblock]$Migrate)
    $doc = Read-WtStore -Name $Name -Schema $Schema -Migrate $Migrate
    if ($null -eq $doc) { return , @() }
    , @($doc.items | Where-Object { $null -ne $_ })
}

# 저장소 하나를 통째로 쓴다. $Extra의 키는 문서 맨 위에 함께 저장된다 (예: 가져온 위치)
function Write-WtStore {
    param(
        [Parameter(Mandatory)][string]$Name,
        [int]$Schema = 1,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Items,
        [hashtable]$Extra
    )
    $doc = [ordered]@{ schema = $Schema; saved = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss') }
    if ($Extra) { foreach ($k in $Extra.Keys) { $doc[$k] = $Extra[$k] } }
    $doc['items'] = @($Items | Where-Object { $null -ne $_ })
    Write-WtJsonFile (Get-WtStorePath $Name) $doc
}

# 새 항목 ID
function New-WtId { [guid]::NewGuid().ToString('N').Substring(0, 12) }

# ---- 예전(WinForms) WorkTool 데이터 ----------------------------------------------------------

# 예전 데이터 폴더를 찾는다. 설정에 적은 위치 → 이 파일 옆 data → 한 단계 위 data 순서
function Find-WtLegacyFolder {
    $candidates = [System.Collections.Generic.List[string]]::new()
    $custom = Get-WtSetting 'legacy.path'
    if ($custom) { $candidates.Add([string]$custom) }
    if ($script:appRoot) {
        $candidates.Add((Join-Path $script:appRoot 'data'))
        $candidates.Add((Join-Path (Split-Path $script:appRoot -Parent) 'data'))
    }
    foreach ($c in $candidates) {
        foreach ($f in 'memo.csv', 'schedule.csv', 'attendance.csv') {
            if (Test-Path -LiteralPath (Join-Path $c $f)) { return (Resolve-Path -LiteralPath $c).ProviderPath }
        }
    }
    $null
}

# 예전 CSV 하나를 읽는다. 파일이 없으면 $null (빈 파일이면 빈 배열)
function Import-WtLegacyCsv([string]$Folder, [string]$Name) {
    $path = Join-Path $Folder "$Name.csv"
    if (-not (Test-Path -LiteralPath $path)) { return $null }
    , @(Import-Csv -LiteralPath $path -Encoding utf8)
}

Export-ModuleMember -Function Set-WtAppRoot, Get-WtAppRoot, Get-WtDataRoot, Read-WtJsonFile, Write-WtJsonFile,
    Get-WtStorePath, Test-WtStore, Read-WtStore, Get-WtStoreItems, Write-WtStore, New-WtId,
    Find-WtLegacyFolder, Import-WtLegacyCsv
} | Import-Module -Force -DisableNameChecking
# ---- core\Storage.psm1 끝 ----

# ---- core\Settings.psm1 ----
New-Module -Name 'WorkTool.Settings' -ScriptBlock {
# 설정 담당. 기본값은 코드에 두고, 사용자가 바꾼 값만 데이터 폴더의 settings.json에 저장한다.
# 그래서 새 설정 항목이 생겨도 예전 settings.json이 그대로 동작하고, 기본값과 같아지면 파일에서 빠진다.

$ErrorActionPreference = 'Stop'

$script:defaults = @{
    'theme.color'     = 'sky'    # Theme.psm1의 팔레트 이름
    'ui.opacity'      = 100      # 창 불투명도(%) 60~100
    'ui.fontScale'    = 'm'      # 글자 크기 s · m · l
    'nav.collapsed'   = $false   # 왼쪽 메뉴 접힘
    'addons.disabled' = @()      # 메뉴에서 숨긴 애드온 ID
    'legacy.path'     = ''       # 예전 WorkTool data 폴더 (비우면 자동으로 찾음)
}
$script:values = $null
$script:handlers = [System.Collections.Generic.List[scriptblock]]::new()

function Get-WtSettingsPath { Join-Path (Get-WtDataRoot) 'settings.json' }

function Initialize-WtSettings {
    $script:values = @{}
    $doc = Read-WtJsonFile (Get-WtSettingsPath)
    if ($doc -and $doc.values) {
        foreach ($p in $doc.values.PSObject.Properties) { $script:values[$p.Name] = $p.Value }
    }
}

# 애드온이 자기 설정의 기본값을 등록한다
function Register-WtSettingDefault([string]$Key, $Value) { $script:defaults[$Key] = $Value }

function Get-WtSetting([string]$Key) {
    if ($null -eq $script:values) { Initialize-WtSettings }
    if ($script:values.ContainsKey($Key)) { return $script:values[$Key] }
    $script:defaults[$Key]
}

function Test-WtSameValue($A, $B) {
    (ConvertTo-Json -InputObject $A -Compress -Depth 5) -eq (ConvertTo-Json -InputObject $B -Compress -Depth 5)
}

# 값을 바꾸고 저장한 뒤, 등록된 처리기에 알린다
function Set-WtSetting([string]$Key, $Value) {
    if ($null -eq $script:values) { Initialize-WtSettings }
    if ($script:defaults.ContainsKey($Key) -and (Test-WtSameValue $script:defaults[$Key] $Value)) {
        [void]$script:values.Remove($Key)
    } else {
        $script:values[$Key] = $Value
    }
    $sorted = [ordered]@{}
    foreach ($k in ($script:values.Keys | Sort-Object)) { $sorted[$k] = $script:values[$k] }
    Write-WtJsonFile (Get-WtSettingsPath) ([ordered]@{ schema = 1; values = $sorted })
    foreach ($h in @($script:handlers)) { & $h $Key $Value }
}

# 설정이 바뀔 때 부를 처리기. { param($key, $value) ... }
function Register-WtSettingHandler([scriptblock]$Handler) { $script:handlers.Add($Handler) }

Export-ModuleMember -Function Get-WtSettingsPath, Initialize-WtSettings, Register-WtSettingDefault, Get-WtSetting,
    Set-WtSetting, Register-WtSettingHandler
} | Import-Module -Force -DisableNameChecking
# ---- core\Settings.psm1 끝 ----

# ---- core\Theme.psm1 ----
New-Module -Name 'WorkTool.Theme' -ScriptBlock {
# 테마 담당. 색 · 글자 크기 · 글꼴은 여기서만 정하고, 화면은 리소스 이름(B.Accent, F.Body …)으로만 쓴다.
# 값의 기준은 apps/WorkTool/design/README.md다. 팔레트를 바꾸거나 더할 때는 그 문서도 같이 고친다.
#
# 팔레트 하나 = 연한 바탕색(Tint) + 강조색(Accent).
# - Tint: 선택된 메뉴 바탕. 메뉴 · 창 바탕은 Tint를 흰색과 섞어 더 연하게 쓴다.
# - Accent: 주 버튼 · 선택 표시. 흰 글자와 4.5:1 이상이어야 한다 (tests/Test-Core.ps1이 검사).

$ErrorActionPreference = 'Stop'

$script:palettes = [ordered]@{
    sky    = @{ Label = '하늘'; Tint = '#DCEBFA'; Accent = '#2370C2' }
    pink   = @{ Label = '분홍'; Tint = '#FADDE8'; Accent = '#B83D69' }
    green  = @{ Label = '연두'; Tint = '#E1F1CF'; Accent = '#3F7D20' }
    yellow = @{ Label = '노랑'; Tint = '#FBEDBF'; Accent = '#856000' }
    purple = @{ Label = '보라'; Tint = '#E9E0F8'; Accent = '#7048B8' }
    gray   = @{ Label = '회색'; Tint = '#E3E7EC'; Accent = '#56606D' }
    black  = @{ Label = '검정'; Tint = '#DEDFE2'; Accent = '#2B2F36' }
}

# 테마와 상관없이 같은 색 (글자, 선, 상태, 근태 종류)
$script:fixed = @{
    Card         = '#FFFFFF'
    OnAccent     = '#FFFFFF'
    Text         = '#1D2B3A'
    SubText      = '#56667A'
    NavText      = '#2E3E50'
    Border       = '#DCE3EC'
    Divider      = '#EEF2F7'
    InputBorder  = '#7F8C9C'
    ButtonBorder = '#C4CEDA'
    Chip         = '#EEF2F7'
    ChipText     = '#4A5A6C'
    Danger       = '#B83227'
    DangerBg     = '#FBE7E4'
    Success      = '#1E7F4F'
    SuccessBg    = '#E3F4EB'
    Warn         = '#8A5A08'
    WarnBg       = '#FDF1DC'
    CloseHover   = '#C42B1C'
    'Att.Work'   = '#1C5A9B'; 'Att.WorkBg'  = '#E0EBF7'
    'Att.Over'   = '#5A42AE'; 'Att.OverBg'  = '#EEEAFB'
    'Att.Trip'   = '#0B6E6D'; 'Att.TripBg'  = '#DDF2F1'
    'Att.Leave'  = '#9C4E0F'; 'Att.LeaveBg' = '#FCEBDC'
}

$script:fontScales = @{ s = 0.93; m = 1.0; l = 1.12 }
$script:fontSizes = @{ Title = 26; Number = 22; H2 = 16; Body = 14; Small = 12.5; Icon = 16 }

# ---- 색 계산 (WPF 없이도 동작: 테스트용) ------------------------------------------------------

function ConvertFrom-WtHex([string]$Hex) {
    $h = $Hex.TrimStart('#')
    @([Convert]::ToInt32($h.Substring(0, 2), 16), [Convert]::ToInt32($h.Substring(2, 2), 16), [Convert]::ToInt32($h.Substring(4, 2), 16))
}

# $A를 $B 쪽으로 $T(0~1)만큼 섞는다
function Get-WtMixedColor([string]$A, [string]$B, [double]$T) {
    $x = ConvertFrom-WtHex $A
    $y = ConvertFrom-WtHex $B
    $c = for ($i = 0; $i -lt 3; $i++) { [int][Math]::Round($x[$i] * (1 - $T) + $y[$i] * $T) }
    '#{0:X2}{1:X2}{2:X2}' -f $c[0], $c[1], $c[2]
}

function Get-WtLuminance([string]$Hex) {
    $lin = foreach ($v in (ConvertFrom-WtHex $Hex)) {
        $s = $v / 255
        if ($s -le 0.03928) { $s / 12.92 } else { [Math]::Pow(($s + 0.055) / 1.055, 2.4) }
    }
    0.2126 * $lin[0] + 0.7152 * $lin[1] + 0.0722 * $lin[2]
}

function Get-WtContrast([string]$A, [string]$B) {
    $la = Get-WtLuminance $A
    $lb = Get-WtLuminance $B
    ([Math]::Max($la, $lb) + 0.05) / ([Math]::Min($la, $lb) + 0.05)
}

function Get-WtPalette { $script:palettes }

# 팔레트 이름 → 색 이름별 #RRGGBB
function Get-WtThemeColors([string]$Name) {
    $p = $script:palettes[$Name]
    if (-not $p) { $p = $script:palettes['sky'] }
    $c = @{}
    foreach ($k in $script:fixed.Keys) { $c[$k] = $script:fixed[$k] }
    $c.Accent       = $p.Accent
    $c.AccentHover  = Get-WtMixedColor $p.Accent '#000000' 0.12
    $c.AccentStrong = Get-WtMixedColor $p.Accent '#000000' 0.18
    $c.AccentSoft   = $p.Tint
    $c.Hover        = Get-WtMixedColor $p.Tint '#FFFFFF' 0.25
    $c.Sidebar      = Get-WtMixedColor $p.Tint '#FFFFFF' 0.45
    $c.WindowBg     = Get-WtMixedColor $p.Tint '#FFFFFF' 0.75
    $c.FrameBorder  = Get-WtMixedColor $p.Tint '#8C99A8' 0.55
    $c
}

# ---- WPF 리소스 ------------------------------------------------------------------------------

# 테마 리소스를 담은 사전을 채운다. 같은 사전의 값을 바꾸면 DynamicResource로 쓴 화면이 바로 바뀐다.
function Update-WtThemeDictionary([System.Windows.ResourceDictionary]$Dictionary) {
    $colors = Get-WtThemeColors ([string](Get-WtSetting 'theme.color'))
    foreach ($k in $colors.Keys) {
        $color = [System.Windows.Media.ColorConverter]::ConvertFromString($colors[$k])
        $brush = [System.Windows.Media.SolidColorBrush]::new($color)
        $brush.Freeze()
        $Dictionary["B.$k"] = $brush
    }
    $scale = $script:fontScales[[string](Get-WtSetting 'ui.fontScale')]
    if (-not $scale) { $scale = 1.0 }
    foreach ($k in $script:fontSizes.Keys) {
        $Dictionary["F.$k"] = [double][Math]::Round($script:fontSizes[$k] * $scale, 1)
    }
    # 설치된 글꼴 중 앞에서부터 쓴다. 맑은 고딕은 Windows 기본이라 항상 있다.
    $Dictionary['Font.UI'] = [System.Windows.Media.FontFamily]::new('Pretendard, Noto Sans KR, Malgun Gothic')
    $Dictionary['Font.Icon'] = [System.Windows.Media.FontFamily]::new('Segoe Fluent Icons, Segoe MDL2 Assets')
    $Dictionary['Font.Mono'] = [System.Windows.Media.FontFamily]::new('Consolas, D2Coding, Malgun Gothic')
}

function New-WtThemeDictionary {
    $d = [System.Windows.ResourceDictionary]::new()
    Update-WtThemeDictionary $d
    $d
}

Export-ModuleMember -Function ConvertFrom-WtHex, Get-WtMixedColor, Get-WtLuminance, Get-WtContrast, Get-WtPalette,
    Get-WtThemeColors, Update-WtThemeDictionary, New-WtThemeDictionary
} | Import-Module -Force -DisableNameChecking
# ---- core\Theme.psm1 끝 ----

# ---- core\Ui.psm1 ----
New-Module -Name 'WorkTool.Ui' -ScriptBlock {
# 공통 부품. 버튼 · 입력 칸 · 카드 · 메뉴 항목 같은 모양은 여기서만 정한다.
# 기능(애드온) 모듈은 아래 New-Wt* 함수와 스타일 이름만 쓰고, 색 값이나 글꼴 이름을 직접 쓰지 않는다.
# 색 · 크기는 Theme.psm1의 리소스(B.* · F.* · Font.*)를 DynamicResource로 참조해서, 테마를 바꾸면 바로 반영된다.

$ErrorActionPreference = 'Stop'

$script:ns = 'xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"'

# 스타일 모음. 색은 Style.Triggers에서 바꾼다 (템플릿 트리거로 바꾸면 BasedOn 스타일에서 덮어쓸 수 없어서).
$script:stylesXaml = @'
<ResourceDictionary {NS}>

  <Style x:Key="T.Body" TargetType="TextBlock">
    <Setter Property="FontFamily" Value="{DynamicResource Font.UI}"/>
    <Setter Property="FontSize" Value="{DynamicResource F.Body}"/>
    <Setter Property="Foreground" Value="{DynamicResource B.Text}"/>
    <Setter Property="TextWrapping" Value="Wrap"/>
  </Style>
  <Style x:Key="T.Sub" TargetType="TextBlock" BasedOn="{StaticResource T.Body}">
    <Setter Property="FontSize" Value="{DynamicResource F.Small}"/>
    <Setter Property="Foreground" Value="{DynamicResource B.SubText}"/>
  </Style>
  <Style x:Key="T.Title" TargetType="TextBlock" BasedOn="{StaticResource T.Body}">
    <Setter Property="FontSize" Value="{DynamicResource F.Title}"/>
    <Setter Property="FontWeight" Value="Bold"/>
  </Style>
  <Style x:Key="T.H2" TargetType="TextBlock" BasedOn="{StaticResource T.Body}">
    <Setter Property="FontSize" Value="{DynamicResource F.H2}"/>
    <Setter Property="FontWeight" Value="SemiBold"/>
  </Style>
  <Style x:Key="T.Number" TargetType="TextBlock" BasedOn="{StaticResource T.Body}">
    <Setter Property="FontSize" Value="{DynamicResource F.Number}"/>
    <Setter Property="FontWeight" Value="Bold"/>
  </Style>
  <Style x:Key="T.Strong" TargetType="TextBlock" BasedOn="{StaticResource T.Body}">
    <Setter Property="FontWeight" Value="SemiBold"/>
  </Style>
  <Style x:Key="T.Icon" TargetType="TextBlock">
    <Setter Property="FontFamily" Value="{DynamicResource Font.Icon}"/>
    <Setter Property="FontSize" Value="{DynamicResource F.Icon}"/>
    <Setter Property="VerticalAlignment" Value="Center"/>
    <Setter Property="HorizontalAlignment" Value="Center"/>
  </Style>

  <Style x:Key="Card" TargetType="Border">
    <Setter Property="Background" Value="{DynamicResource B.Card}"/>
    <Setter Property="BorderBrush" Value="{DynamicResource B.Border}"/>
    <Setter Property="BorderThickness" Value="1"/>
    <Setter Property="CornerRadius" Value="12"/>
    <Setter Property="Padding" Value="20,16"/>
  </Style>
  <Style x:Key="Tile" TargetType="Border">
    <Setter Property="Background" Value="{DynamicResource B.WindowBg}"/>
    <Setter Property="CornerRadius" Value="10"/>
    <Setter Property="Padding" Value="14,10"/>
  </Style>
  <Style x:Key="Divider" TargetType="Border">
    <Setter Property="BorderBrush" Value="{DynamicResource B.Divider}"/>
    <Setter Property="BorderThickness" Value="0,1,0,0"/>
  </Style>

  <Style x:Key="Btn" TargetType="Button">
    <Setter Property="FontFamily" Value="{DynamicResource Font.UI}"/>
    <Setter Property="FontSize" Value="{DynamicResource F.Body}"/>
    <Setter Property="Foreground" Value="{DynamicResource B.Text}"/>
    <Setter Property="Background" Value="{DynamicResource B.Card}"/>
    <Setter Property="BorderBrush" Value="{DynamicResource B.ButtonBorder}"/>
    <Setter Property="BorderThickness" Value="1"/>
    <Setter Property="Padding" Value="16,0"/>
    <Setter Property="MinHeight" Value="40"/>
    <Setter Property="Cursor" Value="Hand"/>
    <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="Button">
          <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                  BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="8" Padding="{TemplateBinding Padding}">
            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" RecognizesAccessKey="False"/>
          </Border>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
    <Style.Triggers>
      <Trigger Property="IsMouseOver" Value="True">
        <Setter Property="Background" Value="{DynamicResource B.Hover}"/>
      </Trigger>
      <Trigger Property="IsKeyboardFocused" Value="True">
        <Setter Property="BorderBrush" Value="{DynamicResource B.Accent}"/>
      </Trigger>
      <Trigger Property="IsEnabled" Value="False">
        <Setter Property="Opacity" Value="0.5"/>
      </Trigger>
    </Style.Triggers>
  </Style>
  <Style x:Key="Btn.Primary" TargetType="Button" BasedOn="{StaticResource Btn}">
    <Setter Property="Background" Value="{DynamicResource B.Accent}"/>
    <Setter Property="BorderBrush" Value="{DynamicResource B.Accent}"/>
    <Setter Property="Foreground" Value="{DynamicResource B.OnAccent}"/>
    <Setter Property="FontWeight" Value="SemiBold"/>
    <Style.Triggers>
      <Trigger Property="IsMouseOver" Value="True">
        <Setter Property="Background" Value="{DynamicResource B.AccentHover}"/>
        <Setter Property="BorderBrush" Value="{DynamicResource B.AccentHover}"/>
      </Trigger>
      <Trigger Property="IsKeyboardFocused" Value="True">
        <Setter Property="BorderBrush" Value="{DynamicResource B.Text}"/>
      </Trigger>
    </Style.Triggers>
  </Style>
  <Style x:Key="Btn.Danger" TargetType="Button" BasedOn="{StaticResource Btn}">
    <Setter Property="Background" Value="Transparent"/>
    <Setter Property="BorderBrush" Value="Transparent"/>
    <Setter Property="Foreground" Value="{DynamicResource B.Danger}"/>
    <Style.Triggers>
      <Trigger Property="IsMouseOver" Value="True">
        <Setter Property="Background" Value="{DynamicResource B.DangerBg}"/>
      </Trigger>
    </Style.Triggers>
  </Style>
  <Style x:Key="Btn.Link" TargetType="Button" BasedOn="{StaticResource Btn}">
    <Setter Property="Background" Value="Transparent"/>
    <Setter Property="BorderBrush" Value="Transparent"/>
    <Setter Property="Foreground" Value="{DynamicResource B.AccentStrong}"/>
    <Setter Property="FontSize" Value="{DynamicResource F.Small}"/>
    <Setter Property="MinHeight" Value="32"/>
    <Setter Property="Padding" Value="10,0"/>
  </Style>
  <Style x:Key="Btn.Icon" TargetType="Button" BasedOn="{StaticResource Btn}">
    <Setter Property="Background" Value="Transparent"/>
    <Setter Property="BorderBrush" Value="Transparent"/>
    <Setter Property="Foreground" Value="{DynamicResource B.NavText}"/>
    <Setter Property="Width" Value="40"/>
    <Setter Property="Height" Value="40"/>
    <Setter Property="MinHeight" Value="0"/>
    <Setter Property="Padding" Value="0"/>
  </Style>
  <Style x:Key="Btn.Chrome" TargetType="Button" BasedOn="{StaticResource Btn.Icon}">
    <Setter Property="Width" Value="46"/>
    <Setter Property="Height" Value="40"/>
  </Style>
  <Style x:Key="Btn.Close" TargetType="Button" BasedOn="{StaticResource Btn.Chrome}">
    <Style.Triggers>
      <Trigger Property="IsMouseOver" Value="True">
        <Setter Property="Background" Value="{DynamicResource B.CloseHover}"/>
        <Setter Property="Foreground" Value="#FFFFFF"/>
      </Trigger>
    </Style.Triggers>
  </Style>

  <Style x:Key="Nav" TargetType="RadioButton">
    <Setter Property="FontFamily" Value="{DynamicResource Font.UI}"/>
    <Setter Property="FontSize" Value="{DynamicResource F.Body}"/>
    <Setter Property="Foreground" Value="{DynamicResource B.NavText}"/>
    <Setter Property="Background" Value="Transparent"/>
    <Setter Property="Height" Value="40"/>
    <Setter Property="Margin" Value="0,2"/>
    <Setter Property="Cursor" Value="Hand"/>
    <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="RadioButton">
          <Border Background="{TemplateBinding Background}" CornerRadius="8" Padding="12,0">
            <ContentPresenter VerticalAlignment="Center" RecognizesAccessKey="False"/>
          </Border>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
    <Style.Triggers>
      <Trigger Property="IsMouseOver" Value="True">
        <Setter Property="Background" Value="{DynamicResource B.Hover}"/>
      </Trigger>
      <Trigger Property="IsChecked" Value="True">
        <Setter Property="Background" Value="{DynamicResource B.AccentSoft}"/>
        <Setter Property="Foreground" Value="{DynamicResource B.AccentStrong}"/>
        <Setter Property="FontWeight" Value="SemiBold"/>
      </Trigger>
    </Style.Triggers>
  </Style>

  <Style x:Key="Seg" TargetType="RadioButton">
    <Setter Property="FontFamily" Value="{DynamicResource Font.UI}"/>
    <Setter Property="FontSize" Value="{DynamicResource F.Body}"/>
    <Setter Property="Foreground" Value="{DynamicResource B.ChipText}"/>
    <Setter Property="Background" Value="Transparent"/>
    <Setter Property="MinHeight" Value="34"/>
    <Setter Property="Padding" Value="14,0"/>
    <Setter Property="Cursor" Value="Hand"/>
    <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="RadioButton">
          <Border Background="{TemplateBinding Background}" CornerRadius="8" Padding="{TemplateBinding Padding}">
            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" RecognizesAccessKey="False"/>
          </Border>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
    <Style.Triggers>
      <Trigger Property="IsChecked" Value="True">
        <Setter Property="Background" Value="{DynamicResource B.Card}"/>
        <Setter Property="Foreground" Value="{DynamicResource B.Text}"/>
        <Setter Property="FontWeight" Value="SemiBold"/>
      </Trigger>
    </Style.Triggers>
  </Style>

  <Style x:Key="Swatch" TargetType="RadioButton">
    <Setter Property="Width" Value="40"/>
    <Setter Property="Height" Value="40"/>
    <Setter Property="Margin" Value="0,0,6,0"/>
    <Setter Property="Cursor" Value="Hand"/>
    <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="RadioButton">
          <Grid Background="Transparent">
            <Ellipse x:Name="Ring" Stroke="{DynamicResource B.Text}" StrokeThickness="2" Visibility="Hidden"/>
            <Ellipse Margin="6" Fill="{TemplateBinding Background}" Stroke="{TemplateBinding BorderBrush}" StrokeThickness="1"/>
          </Grid>
          <ControlTemplate.Triggers>
            <Trigger Property="IsChecked" Value="True">
              <Setter TargetName="Ring" Property="Visibility" Value="Visible"/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <Style x:Key="Input" TargetType="TextBox">
    <Setter Property="FontFamily" Value="{DynamicResource Font.UI}"/>
    <Setter Property="FontSize" Value="{DynamicResource F.Body}"/>
    <Setter Property="Foreground" Value="{DynamicResource B.Text}"/>
    <Setter Property="Background" Value="{DynamicResource B.Card}"/>
    <Setter Property="BorderBrush" Value="{DynamicResource B.InputBorder}"/>
    <Setter Property="BorderThickness" Value="1"/>
    <Setter Property="Padding" Value="10,8"/>
    <Setter Property="MinHeight" Value="40"/>
    <Setter Property="VerticalContentAlignment" Value="Center"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="TextBox">
          <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                  BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="8">
            <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}" Focusable="False"
                          VerticalAlignment="{TemplateBinding VerticalContentAlignment}"/>
          </Border>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
    <Style.Triggers>
      <Trigger Property="IsKeyboardFocused" Value="True">
        <Setter Property="BorderBrush" Value="{DynamicResource B.Accent}"/>
      </Trigger>
    </Style.Triggers>
  </Style>
  <Style x:Key="Reader" TargetType="TextBox" BasedOn="{StaticResource Input}">
    <Setter Property="IsReadOnly" Value="True"/>
    <Setter Property="BorderThickness" Value="0"/>
    <Setter Property="Background" Value="Transparent"/>
    <Setter Property="Padding" Value="0"/>
    <Setter Property="TextWrapping" Value="Wrap"/>
    <Setter Property="AcceptsReturn" Value="True"/>
    <Setter Property="VerticalContentAlignment" Value="Top"/>
    <Setter Property="VerticalScrollBarVisibility" Value="Auto"/>
  </Style>

  <Style x:Key="Row" TargetType="ListBoxItem">
    <Setter Property="Background" Value="Transparent"/>
    <Setter Property="Padding" Value="12,10"/>
    <Setter Property="Margin" Value="0,1"/>
    <Setter Property="Cursor" Value="Hand"/>
    <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
    <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="ListBoxItem">
          <Border Background="{TemplateBinding Background}" CornerRadius="8" Padding="{TemplateBinding Padding}">
            <ContentPresenter HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}"/>
          </Border>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
    <Style.Triggers>
      <Trigger Property="IsMouseOver" Value="True">
        <Setter Property="Background" Value="{DynamicResource B.Hover}"/>
      </Trigger>
      <Trigger Property="IsSelected" Value="True">
        <Setter Property="Background" Value="{DynamicResource B.AccentSoft}"/>
      </Trigger>
    </Style.Triggers>
  </Style>
  <Style x:Key="List" TargetType="ListBox">
    <Setter Property="Background" Value="Transparent"/>
    <Setter Property="BorderThickness" Value="0"/>
    <Setter Property="ItemContainerStyle" Value="{StaticResource Row}"/>
    <Setter Property="ScrollViewer.HorizontalScrollBarVisibility" Value="Disabled"/>
  </Style>

  <Style x:Key="Check" TargetType="CheckBox">
    <Setter Property="FontFamily" Value="{DynamicResource Font.UI}"/>
    <Setter Property="FontSize" Value="{DynamicResource F.Body}"/>
    <Setter Property="Foreground" Value="{DynamicResource B.Text}"/>
    <Setter Property="VerticalContentAlignment" Value="Center"/>
    <Setter Property="Margin" Value="0,4"/>
    <Setter Property="Cursor" Value="Hand"/>
  </Style>

</ResourceDictionary>
'@

$script:styles = $null

# 스타일 사전 (처음 부를 때 한 번 만든다)
function Get-WtStyleDictionary {
    if (-not $script:styles) {
        $script:styles = [System.Windows.Markup.XamlReader]::Parse($script:stylesXaml.Replace('{NS}', $script:ns))
    }
    $script:styles
}

# XAML 문자열을 요소로 만든다. 맨 앞 태그에 {NS}를 쓰면 기본 이름공간으로 바뀐다.
function ConvertFrom-WtXaml([string]$Xaml) {
    [System.Windows.Markup.XamlReader]::Parse($Xaml.Replace('{NS}', $script:ns))
}

# x:Name을 붙인 요소들을 이름 → 요소 사전으로 돌려준다
function Get-WtNamedElements($Root, [string]$Xaml) {
    $map = @{}
    foreach ($m in [regex]::Matches($Xaml, 'x:Name="([A-Za-z0-9_]+)"')) {
        $name = $m.Groups[1].Value
        $map[$name] = $Root.FindName($name)
    }
    $map
}

function Set-WtStyle($Element, [string]$Key) {
    $Element.SetResourceReference([System.Windows.FrameworkElement]::StyleProperty, $Key)
}

# 의존 속성에 테마 리소스를 건다. 예: Set-WtBrush $tb 'Foreground' 'B.SubText'
function Set-WtResource($Element, [string]$Property, [string]$Key) {
    $field = $Element.GetType().GetField("${Property}Property",
        [System.Reflection.BindingFlags]'Public, Static, FlattenHierarchy')
    if (-not $field) { throw "$($Element.GetType().Name)에 $Property 속성이 없습니다." }
    $Element.SetResourceReference($field.GetValue($null), $Key)
}

function New-WtText([string]$Text, [string]$Style = 'T.Body') {
    $t = [System.Windows.Controls.TextBlock]::new()
    $t.Text = $Text
    Set-WtStyle $t $Style
    $t
}

# 아이콘 글꼴(Segoe Fluent Icons / MDL2)의 글자 하나
function New-WtIcon([int]$Code) {
    $t = [System.Windows.Controls.TextBlock]::new()
    $t.Text = [string][char]$Code
    Set-WtStyle $t 'T.Icon'
    $t
}

function New-WtStack([ValidateSet('Vertical', 'Horizontal')][string]$Orientation = 'Vertical', [double]$Gap = 0) {
    $p = [System.Windows.Controls.StackPanel]::new()
    $p.Orientation = $Orientation
    $p.Tag = $Gap
    $p
}

# 쌓기 패널에 요소를 넣는다. 패널 Tag의 간격만큼 앞에 여백을 둔다.
function Add-WtChild($Panel, $Child) {
    $gap = if ($Panel.Tag -is [double]) { [double]$Panel.Tag } else { 0 }
    if ($Panel.Children.Count -gt 0 -and $gap -gt 0 -and $Child -is [System.Windows.FrameworkElement]) {
        $m = $Child.Margin
        if ($Panel.Orientation -eq 'Horizontal') {
            $Child.Margin = [System.Windows.Thickness]::new($m.Left + $gap, $m.Top, $m.Right, $m.Bottom)
        } else {
            $Child.Margin = [System.Windows.Thickness]::new($m.Left, $m.Top + $gap, $m.Right, $m.Bottom)
        }
    }
    [void]$Panel.Children.Add($Child)
    $Child
}

function New-WtCard {
    $b = [System.Windows.Controls.Border]::new()
    Set-WtStyle $b 'Card'
    $b
}

# 카드 머리: 제목(+ 개수) 왼쪽, 오른쪽에 링크 버튼 하나 (선택)
function New-WtCardHeader([string]$Title, [string]$Count, [string]$LinkText, [string]$LinkPage) {
    $g = [System.Windows.Controls.DockPanel]::new()
    $g.LastChildFill = $true
    $g.Margin = [System.Windows.Thickness]::new(0, 0, 0, 8)
    if ($LinkText) {
        $b = New-WtButton $LinkText -Kind Link
        $b.Tag = $LinkPage
        $b.Add_Click({ param($s, $e) Select-WtPage $s.Tag })
        [System.Windows.Controls.DockPanel]::SetDock($b, 'Right')
        [void]$g.Children.Add($b)
    }
    $title = New-WtText $Title 'T.H2'
    if ($Count) {
        $run = [System.Windows.Documents.Run]::new("  $Count")
        $run.FontWeight = 'Normal'
        $run.SetResourceReference([System.Windows.Documents.TextElement]::ForegroundProperty, 'B.SubText')
        $title.Inlines.Add($run)
    }
    $title.VerticalAlignment = 'Center'
    [void]$g.Children.Add($title)
    $g
}

# 버튼. $Kind: Default · Primary · Danger · Icon · Link
function New-WtButton([string]$Text, [string]$Kind = 'Default', [int]$Icon = 0, [string]$ToolTip) {
    $b = [System.Windows.Controls.Button]::new()
    switch ($Kind) {
        'Primary' { Set-WtStyle $b 'Btn.Primary' }
        'Danger' { Set-WtStyle $b 'Btn.Danger' }
        'Icon' { Set-WtStyle $b 'Btn.Icon' }
        'Link' { Set-WtStyle $b 'Btn.Link' }
        default { Set-WtStyle $b 'Btn' }
    }
    if ($Icon -and $Text) {
        $row = New-WtStack Horizontal 8
        [void](Add-WtChild $row (New-WtIcon $Icon))
        $label = [System.Windows.Controls.TextBlock]::new()
        $label.Text = $Text
        $label.VerticalAlignment = 'Center'
        [void](Add-WtChild $row $label)
        $b.Content = $row
    } elseif ($Icon) {
        $b.Content = New-WtIcon $Icon
    } else {
        $b.Content = $Text
    }
    if ($ToolTip) { $b.ToolTip = $ToolTip }
    $b
}

# 작은 알약 모양 표시. $Bg/$Fg는 테마 리소스 이름 (예: 'Att.WorkBg', 'Att.Work')
function New-WtTag([string]$Text, [string]$Bg = 'Chip', [string]$Fg = 'ChipText', [int]$Icon = 0) {
    $b = [System.Windows.Controls.Border]::new()
    $b.CornerRadius = [System.Windows.CornerRadius]::new(12)
    $b.Padding = [System.Windows.Thickness]::new(9, 2, 9, 3)
    $b.VerticalAlignment = 'Center'
    $b.HorizontalAlignment = 'Left'
    Set-WtResource $b 'Background' "B.$Bg"
    $row = New-WtStack Horizontal 4
    if ($Icon) {
        $i = New-WtIcon $Icon
        Set-WtResource $i 'FontSize' 'F.Small'
        Set-WtResource $i 'Foreground' "B.$Fg"
        [void](Add-WtChild $row $i)
    }
    $t = New-WtText $Text 'T.Sub'
    Set-WtResource $t 'Foreground' "B.$Fg"
    $t.FontWeight = 'SemiBold'
    $t.TextWrapping = 'NoWrap'
    [void](Add-WtChild $row $t)
    $b.Child = $row
    $b
}

# 숫자 칸: 이름(점 색) + 큰 숫자 + 단위
function New-WtStatTile([string]$Label, [string]$Value, [string]$Unit, [string]$DotKey = 'Accent') {
    $tile = [System.Windows.Controls.Border]::new()
    Set-WtStyle $tile 'Tile'
    $col = New-WtStack Vertical 2
    $head = New-WtStack Horizontal 6
    $dot = [System.Windows.Shapes.Ellipse]::new()
    $dot.Width = 8; $dot.Height = 8
    $dot.VerticalAlignment = 'Center'
    Set-WtResource $dot 'Fill' "B.$DotKey"
    [void](Add-WtChild $head $dot)
    [void](Add-WtChild $head (New-WtText $Label 'T.Sub'))
    [void](Add-WtChild $col $head)
    $num = New-WtText $Value 'T.Number'
    $run = [System.Windows.Documents.Run]::new(" $Unit")
    $run.SetResourceReference([System.Windows.Documents.TextElement]::FontSizeProperty, 'F.Small')
    $run.SetResourceReference([System.Windows.Documents.TextElement]::ForegroundProperty, 'B.SubText')
    $run.FontWeight = 'Normal'
    $num.Inlines.Add($run)
    [void](Add-WtChild $col $num)
    $tile.Child = $col
    $tile
}

# 화면 맨 위: 작은 설명(선택) + 큰 제목 + 오른쪽 요소(선택)
function New-WtPageHeader([string]$Title, [string]$Sub, $Right) {
    $d = [System.Windows.Controls.DockPanel]::new()
    $d.Margin = [System.Windows.Thickness]::new(0, 0, 0, 20)
    if ($Right) {
        [System.Windows.Controls.DockPanel]::SetDock($Right, 'Right')
        $Right.VerticalAlignment = 'Bottom'
        [void]$d.Children.Add($Right)
    }
    $col = New-WtStack Vertical 4
    if ($Sub) { [void](Add-WtChild $col (New-WtText $Sub 'T.Sub')) }
    [void](Add-WtChild $col (New-WtText $Title 'T.Title'))
    [void]$d.Children.Add($col)
    $d
}

# 화면 하나의 바깥 틀: 스크롤 + 안쪽 여백. 돌려준 패널에 내용을 넣는다.
# -Fill: 바깥 스크롤 없이 창 높이를 채운다. 목록 · 본문처럼 스스로 스크롤하는 요소가 있는 화면에 쓴다.
#        돌려준 DockPanel에 머리글을 Top으로 붙이고, 마지막에 넣은 요소가 남는 공간을 채운다.
function New-WtPage([switch]$Fill) {
    if ($Fill) {
        $dock = [System.Windows.Controls.DockPanel]::new()
        $dock.LastChildFill = $true
        $dock.Margin = [System.Windows.Thickness]::new(36, 8, 36, 32)
        return @{ Root = $dock; Panel = $dock }
    }
    $sv = [System.Windows.Controls.ScrollViewer]::new()
    $sv.VerticalScrollBarVisibility = 'Auto'
    $sv.HorizontalScrollBarVisibility = 'Disabled'
    $sv.Focusable = $false
    $panel = New-WtStack Vertical 16
    $panel.Margin = [System.Windows.Thickness]::new(36, 8, 36, 32)
    $sv.Content = $panel
    @{ Root = $sv; Panel = $panel }
}

# 카드 안 한 줄: 구분선 위에 왼쪽부터 요소들을 차례로 놓는다. 마지막 요소가 남는 폭을 채운다.
function New-WtListRow([object[]]$Parts, [double]$Gap = 14) {
    $b = [System.Windows.Controls.Border]::new()
    Set-WtStyle $b 'Divider'
    $b.Padding = [System.Windows.Thickness]::new(0, 10, 0, 10)
    $d = [System.Windows.Controls.DockPanel]::new()
    $d.LastChildFill = $true
    for ($i = 0; $i -lt $Parts.Count; $i++) {
        $p = $Parts[$i]
        if ($null -eq $p) { continue }
        if ($i -lt $Parts.Count - 1) {
            [System.Windows.Controls.DockPanel]::SetDock($p, 'Left')
            $p.Margin = [System.Windows.Thickness]::new(0, 0, $Gap, 0)
        }
        $p.VerticalAlignment = 'Center'
        [void]$d.Children.Add($p)
    }
    $b.Child = $d
    $b
}

$script:ko = [System.Globalization.CultureInfo]::GetCultureInfo('ko-KR')

# 날짜를 한국어로. 예: Format-WtDate (Get-Date) 'M월 d일 dddd' → 10월 22일 목요일
function Format-WtDate([datetime]$Date, [string]$Format) { $Date.ToString($Format, $script:ko) }

function Show-WtMessage([string]$Text, [string]$Title = 'WorkTool', [string]$Icon = 'Information') {
    [void][System.Windows.MessageBox]::Show($Text, $Title, 'OK', $Icon)
}

function Confirm-WtYesNo([string]$Text, [string]$Title = 'WorkTool') {
    [System.Windows.MessageBox]::Show($Text, $Title, 'YesNo', 'Question') -eq 'Yes'
}

# 오류를 사용자에게 보여 준다. 문구는 그대로 복사해서 알려 줄 수 있게 한다.
function Show-WtError($ErrorRecord, [string]$Where = '') {
    $msg = if ($ErrorRecord -is [System.Management.Automation.ErrorRecord]) {
        $where = if ($ErrorRecord.InvocationInfo) { $ErrorRecord.InvocationInfo.PositionMessage } else { '' }
        "$($ErrorRecord.Exception.Message)`n`n위치: $where"
    } elseif ($ErrorRecord -is [System.Exception]) {
        "$($ErrorRecord.Message)`n`n$($ErrorRecord.StackTrace)"
    } else { [string]$ErrorRecord }
    if ($Where) { $msg = "[$Where]`n$msg" }
    try { Show-WtMessage $msg 'WorkTool 오류' 'Error' } catch { Write-Host $msg }
}

Export-ModuleMember -Function Get-WtStyleDictionary, ConvertFrom-WtXaml, Get-WtNamedElements, Set-WtStyle, Set-WtResource,
    New-WtText, New-WtIcon, New-WtStack, Add-WtChild, New-WtCard, New-WtCardHeader, New-WtButton, New-WtTag,
    New-WtStatTile, New-WtPageHeader, New-WtPage, New-WtListRow, Format-WtDate, Show-WtMessage, Confirm-WtYesNo, Show-WtError
} | Import-Module -Force -DisableNameChecking
# ---- core\Ui.psm1 끝 ----

# ---- core\Addons.psm1 ----
New-Module -Name 'WorkTool.Addons' -ScriptBlock {
# 애드온(탭) 목록. 탭 하나 = addons\<이름>\ 폴더 하나이고, 그 모듈이 불러와질 때 Register-WtAddon으로 자기를 등록한다.
# 셸은 이 목록만 보고 메뉴를 만든다. 탭을 더하거나 빼도 셸 코드는 고치지 않는다.

$ErrorActionPreference = 'Stop'

$script:addons = [System.Collections.Generic.List[hashtable]]::new()

<#
등록 항목
  Id           영문 소문자 ID (저장소 이름에도 씀)
  Title        메뉴에 보이는 이름
  Icon         Segoe 아이콘 글꼴 글자 코드 (예: 0xE70B)
  Order        메뉴 순서 (작을수록 위)
  Required     $true면 설정에서 끌 수 없음 (홈)
  NewView      { } → 화면 요소를 돌려준다. 처음 열 때 한 번 부른다.
  Refresh      { param($view) } → 화면을 열 때마다 부른다 (선택)
  HomeCard     { } → 홈에 놓을 카드 요소(여러 개 가능)를 돌려준다 (선택)
  ImportLegacy { param($folder) } → 예전 WorkTool data 폴더에서 이 탭 데이터를 가져와 저장한다. 가져온 개수를 돌려준다 (선택)
#>
function Register-WtAddon {
    param(
        [Parameter(Mandatory)][ValidatePattern('^[a-z0-9-]+$')][string]$Id,
        [Parameter(Mandatory)][string]$Title,
        [int]$Icon = 0xE10C,
        [int]$Order = 100,
        [switch]$Required,
        [Parameter(Mandatory)][scriptblock]$NewView,
        [scriptblock]$Refresh,
        [scriptblock]$HomeCard,
        [scriptblock]$ImportLegacy,
        [string]$Description = ''
    )
    for ($i = $script:addons.Count - 1; $i -ge 0; $i--) {
        if ($script:addons[$i].Id -eq $Id) { $script:addons.RemoveAt($i) }
    }
    $script:addons.Add(@{
            Id = $Id; Title = $Title; Icon = $Icon; Order = $Order; Required = [bool]$Required
            NewView = $NewView; Refresh = $Refresh; HomeCard = $HomeCard; ImportLegacy = $ImportLegacy
            Description = $Description
        })
}

# 순서대로 돌려준다. -All이 없으면 설정에서 끈 탭은 뺀다.
function Get-WtAddon {
    param([switch]$All, [string]$Id)
    $disabled = @(Get-WtSetting 'addons.disabled')
    $list = $script:addons | Sort-Object { $_.Order }, { $_.Id }
    if ($Id) { return $list | Where-Object { $_.Id -eq $Id } | Select-Object -First 1 }
    if ($All) { return $list }
    $list | Where-Object { $_.Required -or $disabled -notcontains $_.Id }
}

Export-ModuleMember -Function Register-WtAddon, Get-WtAddon
} | Import-Module -Force -DisableNameChecking
# ---- core\Addons.psm1 끝 ----

# ---- core\Pages.psm1 ----
New-Module -Name 'WorkTool.Pages' -ScriptBlock {
# 셸에 붙은 기본 화면: 설정, 도움말. 메뉴 아래 톱니바퀴 · 물음표 버튼으로 연다.
# 이벤트 처리기는 $script: 변수와 보낸 요소의 Tag만 쓴다 (Shell.psm1 머리말 참고).

$ErrorActionPreference = 'Stop'

$script:addonChecks = @()
$script:opacityLabel = $null
$script:opacityTimer = $null
$script:pendingOpacity = 100

# ---- 설정 -------------------------------------------------------------------------------------

function New-WtSettingRow([string]$Label, $Control, [string]$Note) {
    $row = [System.Windows.Controls.DockPanel]::new()
    $row.Margin = [System.Windows.Thickness]::new(0, 6, 0, 6)
    $left = New-WtStack Vertical 2
    $left.Width = 180
    $left.VerticalAlignment = 'Center'
    [void](Add-WtChild $left (New-WtText $Label 'T.Strong'))
    if ($Note) { [void](Add-WtChild $left (New-WtText $Note 'T.Sub')) }
    [System.Windows.Controls.DockPanel]::SetDock($left, 'Left')
    [void]$row.Children.Add($left)
    $Control.VerticalAlignment = 'Center'
    $Control.HorizontalAlignment = 'Left'
    [void]$row.Children.Add($Control)
    $row
}

function New-WtThemeCard {
    $card = New-WtCard
    $col = New-WtStack Vertical 4
    [void](Add-WtChild $col (New-WtCardHeader '화면'))

    # 색: 팔레트마다 동그라미 하나
    $swatches = New-WtStack Horizontal
    $current = [string](Get-WtSetting 'theme.color')
    $palettes = Get-WtPalette
    foreach ($key in $palettes.Keys) {
        $p = $palettes[$key]
        $rb = [System.Windows.Controls.RadioButton]::new()
        Set-WtStyle $rb 'Swatch'
        $rb.GroupName = 'WtTheme'
        $rb.Tag = $key
        $rb.ToolTip = $p.Label
        $rb.Background = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($p.Accent))
        $rb.BorderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.ColorConverter]::ConvertFromString($p.Tint))
        $rb.IsChecked = ($key -eq $current)
        $rb.Add_Checked({ param($s, $e) Set-WtSetting 'theme.color' ([string]$s.Tag) })
        [void]$swatches.Children.Add($rb)
    }
    $labels = ($palettes.Values | ForEach-Object { $_.Label }) -join ' · '
    [void](Add-WtChild $col (New-WtSettingRow '색' $swatches $labels))

    # 창 투명도: 끌면 바로 보이고, 손을 떼고 잠시 뒤 저장한다
    $op = New-WtStack Horizontal 12
    $slider = [System.Windows.Controls.Slider]::new()
    $slider.Minimum = 60
    $slider.Maximum = 100
    $slider.TickFrequency = 5
    $slider.IsSnapToTickEnabled = $true
    $slider.Width = 240
    $slider.Value = [double](Get-WtSetting 'ui.opacity')
    $slider.VerticalAlignment = 'Center'
    $script:opacityLabel = New-WtText ('{0}%' -f [int]$slider.Value) 'T.Strong'
    $script:opacityLabel.VerticalAlignment = 'Center'
    if (-not $script:opacityTimer) {
        $script:opacityTimer = [System.Windows.Threading.DispatcherTimer]::new()
        $script:opacityTimer.Interval = [TimeSpan]::FromMilliseconds(500)
        $script:opacityTimer.Add_Tick({
                param($s, $e)
                $s.Stop()
                Set-WtSetting 'ui.opacity' ([int]$script:pendingOpacity)
            })
    }
    $slider.Add_ValueChanged({
            param($s, $e)
            $script:pendingOpacity = [int]$e.NewValue
            $script:opacityLabel.Text = '{0}%' -f $script:pendingOpacity
            Set-WtOpacity $script:pendingOpacity
            $script:opacityTimer.Stop()
            $script:opacityTimer.Start()
        })
    [void](Add-WtChild $op $slider)
    [void](Add-WtChild $op $script:opacityLabel)
    [void](Add-WtChild $col (New-WtSettingRow '창 투명도' $op '60% ~ 100%'))

    # 글자 크기
    $segBox = [System.Windows.Controls.Border]::new()
    Set-WtResource $segBox 'Background' 'B.Chip'
    $segBox.CornerRadius = [System.Windows.CornerRadius]::new(10)
    $segBox.Padding = [System.Windows.Thickness]::new(4)
    $seg = New-WtStack Horizontal 4
    $scale = [string](Get-WtSetting 'ui.fontScale')
    foreach ($o in @(@('s', '작게'), @('m', '보통'), @('l', '크게'))) {
        $rb = [System.Windows.Controls.RadioButton]::new()
        Set-WtStyle $rb 'Seg'
        $rb.GroupName = 'WtFont'
        $rb.Tag = $o[0]
        $rb.Content = $o[1]
        $rb.IsChecked = ($o[0] -eq $scale)
        $rb.Add_Checked({ param($s, $e) Set-WtSetting 'ui.fontScale' ([string]$s.Tag) })
        [void](Add-WtChild $seg $rb)
    }
    $segBox.Child = $seg
    [void](Add-WtChild $col (New-WtSettingRow '글자 크기' $segBox))

    $card.Child = $col
    $card
}

function New-WtMenuCard {
    $card = New-WtCard
    $col = New-WtStack Vertical 4
    [void](Add-WtChild $col (New-WtCardHeader '메뉴'))
    [void](Add-WtChild $col (New-WtText '끄면 메뉴에서만 사라지고 데이터는 그대로입니다.' 'T.Sub'))
    $disabled = @(Get-WtSetting 'addons.disabled')
    $script:addonChecks = @()
    foreach ($a in @(Get-WtAddon -All)) {
        $cb = [System.Windows.Controls.CheckBox]::new()
        Set-WtStyle $cb 'Check'
        $cb.Content = if ($a.Description) { "$($a.Title) — $($a.Description)" } else { $a.Title }
        $cb.Tag = $a.Id
        $cb.IsChecked = $a.Required -or ($disabled -notcontains $a.Id)
        $cb.IsEnabled = -not $a.Required
        $cb.Add_Click({
                $off = @($script:addonChecks | Where-Object { $_.IsEnabled -and -not $_.IsChecked } | ForEach-Object { [string]$_.Tag })
                Set-WtSetting 'addons.disabled' $off
            })
        $script:addonChecks += $cb
        [void](Add-WtChild $col $cb)
    }
    $card.Child = $col
    $card
}

function New-WtDataCard {
    $card = New-WtCard
    $col = New-WtStack Vertical 6
    [void](Add-WtChild $col (New-WtCardHeader '데이터'))

    $root = Get-WtDataRoot
    $path = New-WtText $root 'T.Body'
    Set-WtResource $path 'FontFamily' 'Font.Mono'
    $open = New-WtButton '폴더 열기' -Icon 0xE8B7
    $open.Tag = $root
    $open.Add_Click({ param($s, $e) Invoke-Item -LiteralPath ([string]$s.Tag) })
    $box = New-WtStack Vertical 6
    [void](Add-WtChild $box $path)
    [void](Add-WtChild $box $open)
    [void](Add-WtChild $col (New-WtSettingRow '저장 위치' $box '바꾸려면 앱 폴더에 datapath.txt를 만들고 첫 줄에 경로를 적습니다'))

    $legacy = Find-WtLegacyFolder
    $lbox = New-WtStack Vertical 6
    $ltext = New-WtText $(if ($legacy) { $legacy } else { '찾지 못함' }) 'T.Body'
    Set-WtResource $ltext 'FontFamily' 'Font.Mono'
    [void](Add-WtChild $lbox $ltext)
    $btns = New-WtStack Horizontal 8
    $pick = New-WtButton '위치 고르기'
    $pick.Add_Click({
            $dlg = [System.Windows.Forms.FolderBrowserDialog]::new()
            $dlg.Description = '예전 WorkTool의 data 폴더를 고르세요 (memo.csv · schedule.csv · attendance.csv가 있는 폴더)'
            if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                Set-WtSetting 'legacy.path' $dlg.SelectedPath
                Select-WtPage '_settings'
            }
        })
    $again = New-WtButton '다시 가져오기' -Kind Primary
    $again.IsEnabled = [bool]$legacy
    $again.Add_Click({
            if (-not (Confirm-WtYesNo "예전 WorkTool 데이터로 새 버전의 메모 · 일정 · 근태를 다시 채웁니다.`n지금 파일은 .bak으로 하나 남습니다. 계속할까요?")) { return }
            try {
                $r = Import-WtLegacyData -Force
                Reset-WtViews
                Show-WtMessage ("가져왔습니다: {0}" -f ($r.Tabs -join ', '))
            } catch { Show-WtError $_ '다시 가져오기' }
        })
    [void](Add-WtChild $btns $pick)
    [void](Add-WtChild $btns $again)
    [void](Add-WtChild $lbox $btns)
    [void](Add-WtChild $col (New-WtSettingRow '예전 WorkTool 데이터' $lbox '1단계에서는 여기서 읽어 와서 보여 주기만 합니다'))

    $card.Child = $col
    $card
}

function New-WtSettingsPage {
    $page = New-WtPage
    [void](Add-WtChild $page.Panel (New-WtPageHeader '설정'))
    [void](Add-WtChild $page.Panel (New-WtThemeCard))
    [void](Add-WtChild $page.Panel (New-WtMenuCard))
    [void](Add-WtChild $page.Panel (New-WtDataCard))
    $page.Root
}

# ---- 도움말 -----------------------------------------------------------------------------------

function New-WtKeyRow([string]$Keys, [string]$What) {
    $k = New-WtTag $Keys
    $k.MinWidth = 120
    New-WtListRow @($k, (New-WtText $What 'T.Body'))
}

function New-WtHelpPage {
    $page = New-WtPage
    [void](Add-WtChild $page.Panel (New-WtPageHeader '도움말'))

    $about = New-WtCard
    $col = New-WtStack Vertical 6
    [void](Add-WtChild $col (New-WtCardHeader '이 버전 (1단계 미리보기)'))
    [void](Add-WtChild $col (New-WtText '창 틀 · 메뉴 · 테마 · 설정, 그리고 예전 WorkTool 데이터를 읽어서 보여 주는 기능까지 들어 있습니다.' 'T.Body'))
    [void](Add-WtChild $col (New-WtText '메모 · 일정 · 근태 편집과 알람은 2단계에서 옮깁니다. 그때까지는 기존 WorkTool을 계속 쓰세요.' 'T.Body'))
    $about.Child = $col
    [void](Add-WtChild $page.Panel $about)

    $keys = New-WtCard
    $col = New-WtStack Vertical
    [void](Add-WtChild $col (New-WtCardHeader '단축키'))
    foreach ($r in @(
            @('Ctrl + 1 ~ 9', '메뉴 순서대로 화면 이동'),
            @('Ctrl + B', '메뉴 접기 / 펼치기'),
            @('F11', '전체 화면 (Esc로 나가기)'),
            @('Ctrl + ,', '설정'),
            @('F1', '도움말'))) {
        [void](Add-WtChild $col (New-WtKeyRow $r[0] $r[1]))
    }
    $keys.Child = $col
    [void](Add-WtChild $page.Panel $keys)

    $win = New-WtCard
    $col = New-WtStack Vertical 6
    [void](Add-WtChild $col (New-WtCardHeader '창 · 데이터'))
    [void](Add-WtChild $col (New-WtText '창의 X는 트레이로 숨깁니다. 완전히 끄려면 트레이 아이콘을 우클릭 → 종료.' 'T.Body'))
    [void](Add-WtChild $col (New-WtText "데이터 위치: $(Get-WtDataRoot)" 'T.Body'))
    $last = Get-WtLastImport
    if ($last) {
        $when = Format-WtDate $last.At 'M월 d일 HH:mm'
        [void](Add-WtChild $col (New-WtText "예전 데이터를 가져옴: $($last.Folder) ($when · $($last.Tabs -join ', '))" 'T.Sub'))
    }
    $win.Child = $col
    [void](Add-WtChild $page.Panel $win)

    $page.Root
}

Export-ModuleMember -Function New-WtSettingsPage, New-WtHelpPage
} | Import-Module -Force -DisableNameChecking
# ---- core\Pages.psm1 끝 ----

# ---- core\Shell.psm1 ----
New-Module -Name 'WorkTool.Shell' -ScriptBlock {
# 셸: 메인 창 틀, 왼쪽 메뉴, 제목 줄 버튼, 트레이, 단축키, 화면 전환.
# 메뉴는 Addons.psm1에 등록된 애드온 목록으로 만든다. 탭을 더하거나 빼도 이 파일은 고치지 않는다.
#
# 창은 투명도를 바꿀 수 있게 WindowStyle=None + AllowsTransparency로 만들고, 제목 줄은 직접 그린다.
# 그래서 이 창 안에는 WindowsFormsHost · WebBrowser를 넣지 않는다 (투명 창에서는 그려지지 않는다).
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
    TextOptions.TextFormattingMode="Display"
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
              <TextBlock x:Name="BrandText" Text="WorkTool" Margin="6,0,0,0" FontWeight="Bold"
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
          <Button x:Name="CloseButton" Style="{DynamicResource Btn.Close}" WindowChrome.IsHitTestVisibleInChrome="True" ToolTip="닫기 (트레이에서 계속 실행)">
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
        $script:selecting = $false
    }
}

# 데이터를 다시 가져온 뒤 등: 만들어 둔 화면을 버리고 지금 화면을 다시 연다
function Reset-WtViews {
    $script:views = @{}
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

# 앱 아이콘: 파일 없이 강조색 둥근 사각형에 W를 그린다
function New-WtAppIcon {
    $hex = (Get-WtThemeColors ([string](Get-WtSetting 'theme.color'))).Accent
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
    $brush = [System.Drawing.SolidBrush]::new([System.Drawing.ColorTranslator]::FromHtml($hex))
    $g.FillPath($brush, $path)
    $font = [System.Drawing.Font]::new('Segoe UI', 20, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
    $format = [System.Drawing.StringFormat]::new()
    $format.Alignment = 'Center'
    $format.LineAlignment = 'Center'
    $g.DrawString('W', $font, [System.Drawing.Brushes]::White, [System.Drawing.RectangleF]::new(0, 1, $size, $size), $format)
    $g.Dispose()
    [System.Drawing.Icon]::FromHandle($bmp.GetHicon())
}

# 창 · 트레이 아이콘을 지금 테마 색으로 (다시) 만든다
function Update-WtAppIcon {
    $old = $script:icon
    $script:icon = New-WtAppIcon
    $script:window.Icon = [System.Windows.Interop.Imaging]::CreateBitmapSourceFromHIcon($script:icon.Handle,
        [System.Windows.Int32Rect]::Empty, [System.Windows.Media.Imaging.BitmapSizeOptions]::FromEmptyOptions())
    if ($script:notify) { $script:notify.Icon = $script:icon }
    if ($old) { $old.Dispose() }
}

function Show-WtMainWindow {
    $script:window.Show()
    if ($script:window.WindowState -eq 'Minimized') { $script:window.WindowState = 'Normal' }
    [void]$script:window.Activate()
}

function Exit-WtApp {
    $script:exiting = $true
    try { $script:notify.Visible = $false; $script:notify.Dispose() } catch { }
    try { if ($script:icon) { $script:icon.Dispose() } } catch { }
    try { $script:window.Close() } catch { }
    if ($script:frame) { $script:frame.Continue = $false }
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
    Initialize-WtSettings
    $script:themeDict = New-WtThemeDictionary

    $w = ConvertFrom-WtXaml $script:shellXaml
    $script:window = $w
    $script:ui = Get-WtNamedElements $w $script:shellXaml
    $w.Resources.MergedDictionaries.Add($script:themeDict)
    $w.Resources.MergedDictionaries.Add((Get-WtStyleDictionary))

    Update-WtAppIcon
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

    # X(닫기)는 트레이로 숨긴다. 완전히 끄는 건 트레이 아이콘 우클릭 → 종료
    $w.Add_Closing({
            param($s, $e)
            if ($script:exiting) { return }
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
            'addons.disabled' { $script:views = @{}; Update-WtNav }
        }
    }

    # 처음 실행이면 예전 WorkTool 데이터를 가져온다 (이미 가져온 탭은 건너뜀)
    try { [void](Import-WtLegacyData) } catch { Show-WtError $_ '예전 데이터 가져오기' }

    Update-WtNav
    Set-WtOpacity
    $script:notify = New-WtTrayIcon $script:icon

    # 창이 닫힐 때까지 메시지를 돌린다. Dispatcher를 끄지 않으므로 같은 창에서 다시 실행해도 된다.
    $script:frame = [System.Windows.Threading.DispatcherFrame]::new()
    $w.Show()
    $first = @(Get-WtAddon)[0]
    if ($first) { Select-WtPage $first.Id } else { Select-WtPage '_settings' }
    if (-not $script:exiting) { [System.Windows.Threading.Dispatcher]::PushFrame($script:frame) }
}

Export-ModuleMember -Function Get-WtWindow, Select-WtPage, Reset-WtViews, Update-WtNav, Switch-WtCollapsed, Set-WtOpacity,
    Switch-WtMaximize, Switch-WtFullScreen, Show-WtMainWindow, Exit-WtApp, Import-WtLegacyData, Get-WtLastImport, Start-WtApp
} | Import-Module -Force -DisableNameChecking
# ---- core\Shell.psm1 끝 ----

# ---- addons\attendance\Attendance.psm1 ----
New-Module -Name 'WorkTool.addon.attendance' -ScriptBlock {
# 근태 탭 (1단계: 보기 전용). 기록 · 지금 출근은 2단계에서 옮긴다.
# 저장소 attendance.json (schema 1). 항목:
#   id, type(출근 | 업무외시간 근무 | 출장 | 휴가), startDate, endDate, startTime, endTime, place, detail, updated

$ErrorActionPreference = 'Stop'

$script:schema = 1
$script:inv = [System.Globalization.CultureInfo]::InvariantCulture
$script:month = [datetime]::new([datetime]::Today.Year, [datetime]::Today.Month, 1)
$script:body = $null
$script:monthLabel = $null

# 종류 → 표시 색 (Theme.psm1의 Att.*)
$script:typeColor = @{ '출근' = 'Att.Work'; '업무외시간 근무' = 'Att.Over'; '출장' = 'Att.Trip'; '휴가' = 'Att.Leave' }

function ConvertTo-AttDate([string]$Text) { [datetime]::ParseExact($Text, 'yyyy-MM-dd', $script:inv) }

function Get-WtAttendanceItems { Get-WtStoreItems -Name 'attendance' -Schema $script:schema }

# 기록이 그 달과 겹치는지
function Test-WtAttInMonth($Item, [datetime]$Month) {
    $first = $Month.Date
    $last = $first.AddMonths(1).AddDays(-1)
    $s = ConvertTo-AttDate $Item.startDate
    $e = if ($Item.endDate) { ConvertTo-AttDate $Item.endDate } else { $s }
    $s -le $last -and $e -ge $first
}

# 업무외시간 근무 분. 끝 시각이 시작보다 이르면 다음 날로 본다.
function Get-WtAttMinutes($Item) {
    if (-not $Item.startTime -or -not $Item.endTime) { return 0 }
    $a = [timespan]::ParseExact($Item.startTime, 'hh\:mm', $script:inv)
    $b = [timespan]::ParseExact($Item.endTime, 'hh\:mm', $script:inv)
    $m = ($b - $a).TotalMinutes
    if ($m -le 0) { $m += 1440 }
    [int]$m
}

# 그 달 요약: 출근 일수, 업무외 시간, 출장 건수, 휴가 일수(반차 0.5)
function Get-WtAttendanceSummary([datetime]$Month, [object[]]$Items) {
    if ($null -eq $Items) { $Items = Get-WtAttendanceItems }
    $first = $Month.Date
    $last = $first.AddMonths(1).AddDays(-1)
    $inMonth = @($Items | Where-Object { Test-WtAttInMonth $_ $first })
    $work = @($inMonth | Where-Object { $_.type -eq '출근' } | ForEach-Object { $_.startDate } | Sort-Object -Unique).Count
    $over = ($inMonth | Where-Object { $_.type -eq '업무외시간 근무' } | ForEach-Object { Get-WtAttMinutes $_ } | Measure-Object -Sum).Sum
    $trip = @($inMonth | Where-Object { $_.type -eq '출장' }).Count
    $leave = 0.0
    foreach ($r in ($inMonth | Where-Object { $_.type -eq '휴가' })) {
        if ($r.detail -like '*반차*') { $leave += 0.5; continue }
        $s = ConvertTo-AttDate $r.startDate
        $e = if ($r.endDate) { ConvertTo-AttDate $r.endDate } else { $s }
        if ($s -lt $first) { $s = $first }
        if ($e -gt $last) { $e = $last }
        $leave += ($e - $s).Days + 1
    }
    [pscustomobject]@{ Work = $work; OverMinutes = [int]$over; Trip = $trip; Leave = $leave }
}

function Format-WtHours([int]$Minutes) {
    $h = [Math]::Round($Minutes / 60, 1)
    if ($h -eq [Math]::Floor($h)) { '{0:0}' -f $h } else { '{0:0.0}' -f $h }
}

# 예전 attendance.csv → attendance.json. 가져온 개수를 돌려준다.
function Import-WtAttendanceLegacy([string]$Folder) {
    $rows = Import-WtLegacyCsv $Folder 'attendance'
    if ($null -eq $rows) { return $null }
    $items = foreach ($r in $rows) {
        [ordered]@{
            id        = [string]$r.Id
            type      = [string]$r.Type
            startDate = [string]$r.StartDate
            endDate   = if ($r.EndDate) { [string]$r.EndDate } else { [string]$r.StartDate }
            startTime = [string]$r.StartTime
            endTime   = [string]$r.EndTime
            place     = [string]$r.Place
            detail    = [string]$r.Detail
            updated   = [string]$r.Updated
        }
    }
    Write-WtStore -Name 'attendance' -Schema $script:schema -Items @($items) -Extra @{ importedFrom = $Folder }
    @($items).Count
}

function New-WtAttTiles($Summary, [int]$Columns) {
    $g = [System.Windows.Controls.Primitives.UniformGrid]::new()
    $g.Columns = $Columns
    foreach ($t in @(
            @('출근', "$($Summary.Work)", '일', 'Att.Work'),
            @('업무외시간', (Format-WtHours $Summary.OverMinutes), '시간', 'Att.Over'),
            @('출장', "$($Summary.Trip)", '건', 'Att.Trip'),
            @('휴가', ('{0:0.#}' -f $Summary.Leave), '일', 'Att.Leave'))) {
        $tile = New-WtStatTile $t[0] $t[1] $t[2] $t[3]
        $tile.Margin = [System.Windows.Thickness]::new(0, 0, 10, 10)
        [void]$g.Children.Add($tile)
    }
    $g
}

function New-WtAttRow($Item) {
    $s = ConvertTo-AttDate $Item.startDate
    $e = if ($Item.endDate) { ConvertTo-AttDate $Item.endDate } else { $s }
    $dateText = if ($e -gt $s) { '{0} ~ {1}' -f (Format-WtDate $s "M'/'d"), (Format-WtDate $e "M'/'d") } else { Format-WtDate $s "M'/'d (ddd)" }
    $date = New-WtText $dateText 'T.Body'
    $date.Width = 110
    $color = $script:typeColor[[string]$Item.type]
    if (-not $color) { $color = 'Chip' }
    $tagBox = [System.Windows.Controls.Border]::new()
    $tagBox.Width = 130
    $tagBox.Child = New-WtTag $Item.type "${color}Bg" $color
    if (-not $script:typeColor[[string]$Item.type]) { $tagBox.Child = New-WtTag $Item.type }
    $time = switch ([string]$Item.type) {
        '출근' { $Item.startTime }
        '업무외시간 근무' { '{0}–{1} ({2}시간)' -f $Item.startTime, $Item.endTime, (Format-WtHours (Get-WtAttMinutes $Item)) }
        '휴가' { if ($Item.detail) { $Item.detail } else { '휴가' } }
        default { if ($e -gt $s) { '{0}일' -f (($e - $s).Days + 1) } else { '하루' } }
    }
    $timeText = New-WtText ([string]$time) 'T.Body'
    $timeText.Width = 170
    $noteParts = @($Item.place, $(if ($Item.type -ne '휴가') { $Item.detail })) | Where-Object { $_ }
    $note = New-WtText ($noteParts -join ' · ') 'T.Sub'
    New-WtListRow @($date, $tagBox, $timeText, $note)
}

function New-WtAttendanceView {
    $page = New-WtPage
    $nav = New-WtStack Horizontal 6
    $prev = New-WtButton '' -Icon 0xE76B -ToolTip '이전 달'
    $prev.Tag = -1
    $next = New-WtButton '' -Icon 0xE76C -ToolTip '다음 달'
    $next.Tag = 1
    $thisMonth = New-WtButton '이번 달'
    $thisMonth.Tag = 0
    foreach ($b in $prev, $next, $thisMonth) {
        $b.Add_Click({
                param($s, $e)
                $step = [int]$s.Tag
                $script:month = if ($step -eq 0) { [datetime]::new([datetime]::Today.Year, [datetime]::Today.Month, 1) } else { $script:month.AddMonths($step) }
                Update-WtAttendanceView
            })
    }
    $script:monthLabel = New-WtText '' 'T.H2'
    $script:monthLabel.VerticalAlignment = 'Center'
    $script:monthLabel.Margin = [System.Windows.Thickness]::new(6, 0, 6, 0)
    [void](Add-WtChild $nav $prev)
    [void](Add-WtChild $nav $script:monthLabel)
    [void](Add-WtChild $nav $next)
    [void](Add-WtChild $nav $thisMonth)
    [void](Add-WtChild $page.Panel (New-WtPageHeader '근태' '보기 전용 · 기록과 「지금 출근」은 2단계에서 옮깁니다' $nav))
    $script:body = New-WtStack Vertical 16
    [void](Add-WtChild $page.Panel $script:body)
    $page.Root
}

function Update-WtAttendanceView {
    $script:monthLabel.Text = Format-WtDate $script:month 'yyyy년 M월'
    $script:body.Children.Clear()
    $items = Get-WtAttendanceItems
    [void](Add-WtChild $script:body (New-WtAttTiles (Get-WtAttendanceSummary $script:month $items) 4))
    $card = New-WtCard
    $col = New-WtStack Vertical
    [void](Add-WtChild $col (New-WtCardHeader ('{0} 기록' -f (Format-WtDate $script:month 'M월'))))
    $rows = @($items | Where-Object { Test-WtAttInMonth $_ $script:month } |
        Sort-Object -Property @{ Expression = { $_.startDate }; Descending = $true }, @{ Expression = { $_.startTime }; Descending = $true })
    foreach ($r in $rows) { [void](Add-WtChild $col (New-WtAttRow $r)) }
    if (-not $rows.Count) { [void](Add-WtChild $col (New-WtText '이 달에는 기록이 없습니다.' 'T.Sub')) }
    $card.Child = $col
    [void](Add-WtChild $script:body $card)
}

function New-WtAttendanceHomeCard {
    $items = Get-WtAttendanceItems
    $month = [datetime]::new([datetime]::Today.Year, [datetime]::Today.Month, 1)
    $card = New-WtCard
    $col = New-WtStack Vertical 8
    [void](Add-WtChild $col (New-WtCardHeader ('근태 · {0}' -f (Format-WtDate $month 'M월')) '' '근태 열기' 'attendance'))
    $today = [datetime]::Today.ToString('yyyy-MM-dd', $script:inv)
    $in = $items | Where-Object { $_.type -eq '출근' -and $_.startDate -eq $today } | Select-Object -First 1
    $status = if ($in) { New-WtTag "오늘 출근 $($in.startTime)" 'SuccessBg' 'Success' } else { New-WtTag '아직 출근 기록 없음' 'WarnBg' 'Warn' }
    [void](Add-WtChild $col $status)
    [void](Add-WtChild $col (New-WtAttTiles (Get-WtAttendanceSummary $month $items) 2))
    $card.Child = $col
    $card
}

Register-WtAddon -Id 'attendance' -Title '근태' -Icon 0xE823 -Order 30 `
    -NewView { New-WtAttendanceView } `
    -Refresh { param($view) Update-WtAttendanceView } `
    -HomeCard { New-WtAttendanceHomeCard } `
    -ImportLegacy { param($folder) Import-WtAttendanceLegacy $folder }

Export-ModuleMember -Function Get-WtAttendanceItems, Test-WtAttInMonth, Get-WtAttMinutes, Get-WtAttendanceSummary,
    Format-WtHours, Import-WtAttendanceLegacy
} | Import-Module -Force -DisableNameChecking
# ---- addons\attendance\Attendance.psm1 끝 ----

# ---- addons\home\Home.psm1 ----
New-Module -Name 'WorkTool.addon.home' -ScriptBlock {
# 홈 탭. 다른 탭이 등록한 홈 카드(HomeCard)를 모아서 두 줄로 보여 준다.
# 탭을 더하면 그 탭의 카드도 여기에 저절로 나온다. 홈은 다른 탭의 데이터를 직접 읽지 않는다.

$ErrorActionPreference = 'Stop'

$script:dateText = $null
$script:columns = $null

function New-WtHomeView {
    $page = New-WtPage
    # 머리글: 날짜(열 때마다 새로 씀) + 제목
    $header = New-WtStack Vertical 4
    $header.Margin = [System.Windows.Thickness]::new(0, 0, 0, 4)
    $script:dateText = New-WtText '' 'T.Sub'
    [void](Add-WtChild $header $script:dateText)
    [void](Add-WtChild $header (New-WtText '오늘' 'T.Title'))
    [void](Add-WtChild $page.Panel $header)

    $grid = [System.Windows.Controls.Grid]::new()
    $script:columns = @()
    for ($i = 0; $i -lt 2; $i++) {
        $cd = [System.Windows.Controls.ColumnDefinition]::new()
        $cd.Width = [System.Windows.GridLength]::new(1, 'Star')
        $grid.ColumnDefinitions.Add($cd)
        $col = New-WtStack Vertical 16
        $col.Margin = [System.Windows.Thickness]::new($(if ($i -eq 0) { 0 } else { 8 }), 0, $(if ($i -eq 0) { 8 } else { 0 }), 0)
        [System.Windows.Controls.Grid]::SetColumn($col, $i)
        [void]$grid.Children.Add($col)
        $script:columns += $col
    }
    [void](Add-WtChild $page.Panel $grid)
    $page.Root
}

function Update-WtHomeView {
    $script:dateText.Text = Format-WtDate (Get-Date) 'yyyy년 M월 d일 dddd'
    foreach ($c in $script:columns) { $c.Children.Clear() }
    $cards = [System.Collections.Generic.List[object]]::new()
    foreach ($a in @(Get-WtAddon)) {
        if (-not $a.HomeCard -or $a.Id -eq 'home') { continue }
        try {
            foreach ($el in @(& $a.HomeCard)) { if ($el -is [System.Windows.UIElement]) { $cards.Add($el) } }
        } catch {
            $cards.Add((New-WtText "$($a.Title) 카드를 만들지 못했습니다: $($_.Exception.Message)" 'T.Sub'))
        }
    }
    # 왼쪽 · 오른쪽에 번갈아 놓는다
    for ($i = 0; $i -lt $cards.Count; $i++) {
        [void](Add-WtChild $script:columns[$i % 2] $cards[$i])
    }
    if (-not $cards.Count) {
        [void](Add-WtChild $script:columns[0] (New-WtText '홈에 보여 줄 카드가 없습니다.' 'T.Sub'))
    }
}

Register-WtAddon -Id 'home' -Title '홈' -Icon 0xE80F -Order 0 -Required `
    -NewView { New-WtHomeView } `
    -Refresh { param($view) Update-WtHomeView }

Export-ModuleMember -Function @()
} | Import-Module -Force -DisableNameChecking
# ---- addons\home\Home.psm1 끝 ----

# ---- addons\memo\Memo.psm1 ----
New-Module -Name 'WorkTool.addon.memo' -ScriptBlock {
# 메모 탭 (1단계: 보기 전용). 쓰기 · 자동 저장 · 메모별 잠금은 2단계에서 옮긴다.
# 저장소 memo.json (schema 1). 항목:
#   id, title, body, tags[], lock($null=잠기지 않음. 2단계에서 'dpapi' | 'password'), created, updated

$ErrorActionPreference = 'Stop'

$script:schema = 1
$script:items = @()
$script:list = $null
$script:search = $null
$script:titleText = $null
$script:metaText = $null
$script:bodyBox = $null

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
    $snippet = New-WtText ((([string]$Memo.body) -split "`r?`n" | Where-Object { $_.Trim() }) -join ' · ') 'T.Sub'
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
        $script:bodyBox.Text = if ($script:items.Count) { '왼쪽에서 메모를 고르세요.' } else { '메모가 없습니다.' }
        return
    }
    $script:titleText.Text = $m.title
    $script:metaText.Text = '만든 날 {0} · 고친 날 {1}' -f $m.created, $m.updated
    $script:bodyBox.Text = [string]$m.body
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

    # 오른쪽: 읽기
    $right = New-WtCard
    $rdock = [System.Windows.Controls.DockPanel]::new()
    $head = New-WtStack Vertical 4
    $head.Margin = [System.Windows.Thickness]::new(0, 0, 0, 12)
    $script:titleText = New-WtText '' 'T.H2'
    Set-WtResource $script:titleText 'FontSize' 'F.Number'
    $script:metaText = New-WtText '' 'T.Sub'
    [void](Add-WtChild $head $script:titleText)
    [void](Add-WtChild $head $script:metaText)
    [System.Windows.Controls.DockPanel]::SetDock($head, 'Top')
    [void]$rdock.Children.Add($head)
    $script:bodyBox = [System.Windows.Controls.TextBox]::new()
    Set-WtStyle $script:bodyBox 'Reader'
    [void]$rdock.Children.Add($script:bodyBox)
    $right.Child = $rdock
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
        $snip = New-WtText ((([string]$m.body) -split "`r?`n" | Where-Object { $_.Trim() }) -join ' · ') 'T.Sub'
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
} | Import-Module -Force -DisableNameChecking
# ---- addons\memo\Memo.psm1 끝 ----

# ---- addons\schedule\Schedule.psm1 ----
New-Module -Name 'WorkTool.addon.schedule' -ScriptBlock {
# 일정 탭 (1단계: 보기 전용). 편집 · 알람은 2단계에서 옮긴다.
# 저장소 schedule.json (schema 1). 항목:
#   id, title, startDate, endDate(기간 일정. 하루짜리는 startDate와 같음), time(''=하루 종일), endTime,
#   note, alarm(분 전, $null=없음), repeat(''|daily|weekdays|weekly), tags[], lastAlarm, updated

$ErrorActionPreference = 'Stop'

$script:schema = 1
$script:inv = [System.Globalization.CultureInfo]::InvariantCulture
$script:repeatNames = @{ daily = '매일'; weekdays = '평일'; weekly = '매주' }
$script:list = $null

function ConvertTo-SchDate([string]$Text) { [datetime]::ParseExact($Text, 'yyyy-MM-dd', $script:inv) }

function Get-WtScheduleItems { Get-WtStoreItems -Name 'schedule' -Schema $script:schema }

# $Item이 $Day에 있는지 (반복 포함)
function Test-WtScheduleOn($Item, [datetime]$Day) {
    $d = $Day.Date
    $start = ConvertTo-SchDate $Item.startDate
    $end = if ($Item.endDate) { ConvertTo-SchDate $Item.endDate } else { $start }
    if ($d -lt $start) { return $false }
    switch ([string]$Item.repeat) {
        'daily' { return $true }
        'weekdays' { return [int]$d.DayOfWeek -ge 1 -and [int]$d.DayOfWeek -le 5 }
        'weekly' { return $d.DayOfWeek -eq $start.DayOfWeek }
        default { return $d -le $end }
    }
}

# 그날 일정을 시간 순서로 (하루 종일이 먼저)
function Get-WtScheduleOn([datetime]$Day, [object[]]$Items) {
    if ($null -eq $Items) { $Items = Get-WtScheduleItems }
    @($Items | Where-Object { Test-WtScheduleOn $_ $Day } | Sort-Object { if ($_.time) { $_.time } else { '' } }, { $_.title })
}

function Format-WtAlarm($Alarm) {
    if ($null -eq $Alarm -or "$Alarm" -eq '') { return $null }
    $m = [int]$Alarm
    if ($m -eq 0) { '정시' } elseif ($m -eq 1440) { '하루 전' } elseif ($m % 60 -eq 0) { "$($m / 60)시간 전" } else { "${m}분 전" }
}

# 예전 schedule.csv → schedule.json. 가져온 개수를 돌려준다.
function Import-WtScheduleLegacy([string]$Folder) {
    $rows = Import-WtLegacyCsv $Folder 'schedule'
    if ($null -eq $rows) { return $null }
    $items = foreach ($r in $rows) {
        [ordered]@{
            id        = [string]$r.Id
            title     = [string]$r.Title
            startDate = [string]$r.Date
            endDate   = [string]$r.Date
            time      = [string]$r.Time
            endTime   = ''
            note      = [string]$r.Note
            alarm     = if ("$($r.Alarm)" -eq '') { $null } else { [int]$r.Alarm }
            repeat    = [string]$r.Repeat
            tags      = @()
            lastAlarm = [string]$r.LastAlarm
            updated   = [string]$r.Updated
        }
    }
    Write-WtStore -Name 'schedule' -Schema $script:schema -Items @($items) -Extra @{ importedFrom = $Folder }
    @($items).Count
}

# 한 줄: 시각 · 제목 · 표시(반복 · 알람)
function New-WtScheduleRow($Item, [string]$DayLabel) {
    $parts = [System.Collections.Generic.List[object]]::new()
    if ($DayLabel) {
        $d = New-WtText $DayLabel 'T.Sub'
        $d.Width = 84
        $parts.Add($d)
    }
    $t = New-WtText $(if ($Item.time) { $Item.time } else { '하루 종일' }) 'T.Strong'
    $t.Width = 72
    $parts.Add($t)
    $tags = New-WtStack Horizontal 6
    $rep = $script:repeatNames[[string]$Item.repeat]
    if ($rep) { [void](Add-WtChild $tags (New-WtTag $rep -Icon 0xE8EE)) }
    $al = Format-WtAlarm $Item.alarm
    if ($al) { [void](Add-WtChild $tags (New-WtTag $al -Icon 0xEA8F)) }
    foreach ($tag in @($Item.tags)) { if ($tag) { [void](Add-WtChild $tags (New-WtTag "#$tag")) } }
    if ($tags.Children.Count) { $parts.Add($tags) }
    $title = New-WtText $Item.title 'T.Body'
    $title.TextTrimming = 'CharacterEllipsis'
    $title.TextWrapping = 'NoWrap'
    $parts.Add($title)
    # 제목이 남는 폭을 채우도록 맨 끝에 둔다 (표시는 제목 앞)
    New-WtListRow $parts.ToArray()
}

function New-WtScheduleView {
    $page = New-WtPage
    [void](Add-WtChild $page.Panel (New-WtPageHeader '일정' '보기 전용 · 편집과 알람은 2단계에서 옮깁니다'))
    $card = New-WtCard
    $script:list = New-WtStack Vertical
    $card.Child = $script:list
    [void](Add-WtChild $page.Panel $card)
    $page.Root
}

function Update-WtScheduleView {
    $script:list.Children.Clear()
    [void](Add-WtChild $script:list (New-WtCardHeader '다가오는 14일'))
    $items = Get-WtScheduleItems
    $any = $false
    for ($i = 0; $i -lt 14; $i++) {
        $day = [datetime]::Today.AddDays($i)
        $label = if ($i -eq 0) { '오늘' } elseif ($i -eq 1) { '내일' } else { Format-WtDate $day "M'/'d (ddd)" }
        foreach ($it in (Get-WtScheduleOn $day $items)) {
            [void](Add-WtChild $script:list (New-WtScheduleRow $it $label))
            $label = ''
            $any = $true
        }
    }
    if (-not $any) { [void](Add-WtChild $script:list (New-WtText '다가오는 일정이 없습니다.' 'T.Sub')) }
}

function New-WtScheduleHomeCards {
    $items = Get-WtScheduleItems
    $today = @(Get-WtScheduleOn ([datetime]::Today) $items)

    $c1 = New-WtCard
    $col = New-WtStack Vertical
    [void](Add-WtChild $col (New-WtCardHeader '오늘 일정' "$($today.Count)" '일정 열기' 'schedule'))
    foreach ($it in $today) { [void](Add-WtChild $col (New-WtScheduleRow $it)) }
    if (-not $today.Count) { [void](Add-WtChild $col (New-WtText '오늘은 일정이 없습니다.' 'T.Sub')) }
    $c1.Child = $col

    $c2 = New-WtCard
    $col = New-WtStack Vertical
    [void](Add-WtChild $col (New-WtCardHeader '다가오는 7일'))
    $n = 0
    for ($i = 1; $i -le 7; $i++) {
        $day = [datetime]::Today.AddDays($i)
        foreach ($it in (Get-WtScheduleOn $day $items)) {
            if ($it.repeat -eq 'daily' -or $it.repeat -eq 'weekdays') { continue }   # 매일 반복은 오늘 카드에서만
            [void](Add-WtChild $col (New-WtScheduleRow $it (Format-WtDate $day "M'/'d (ddd)")))
            $n++
        }
    }
    if (-not $n) { [void](Add-WtChild $col (New-WtText '다가오는 일정이 없습니다.' 'T.Sub')) }
    $c2.Child = $col

    $c1
    $c2
}

Register-WtAddon -Id 'schedule' -Title '일정' -Icon 0xE787 -Order 20 `
    -NewView { New-WtScheduleView } `
    -Refresh { param($view) Update-WtScheduleView } `
    -HomeCard { New-WtScheduleHomeCards } `
    -ImportLegacy { param($folder) Import-WtScheduleLegacy $folder }

Export-ModuleMember -Function Get-WtScheduleItems, Test-WtScheduleOn, Get-WtScheduleOn, Format-WtAlarm, Import-WtScheduleLegacy
} | Import-Module -Force -DisableNameChecking
# ---- addons\schedule\Schedule.psm1 끝 ----
#endregion

try {
    Start-WtApp -AppRoot $PSScriptRoot
} catch {
    $msg = "$($_.Exception.Message)`n`n위치: $($_.InvocationInfo.PositionMessage)"
    Write-Host $msg -ForegroundColor Red
    try { [void][System.Windows.MessageBox]::Show($msg, 'WorkTool 시작 오류', 'OK', 'Error') } catch { }
}
