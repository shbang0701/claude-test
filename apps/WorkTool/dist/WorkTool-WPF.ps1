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

    # 마크다운 보기 (Markdown.psm1)
    'Md.QuoteBar' = '#C4CEDA'   # 인용 왼쪽 막대 (장식)
    'Md.Mark'     = '#FFF0A8'   # ==형광펜== 바탕
    # 콜아웃 (> [!note] …): 제목 · 막대 글자색 / 바탕
    'Callout.Blue'   = '#1C5A9B'; 'Callout.BlueBg'   = '#E0EBF7'
    'Callout.Teal'   = '#0B6E6D'; 'Callout.TealBg'   = '#DDF2F1'
    'Callout.Green'  = '#1A6E44'; 'Callout.GreenBg'  = '#E3F4EB'
    'Callout.Orange' = '#8A5A08'; 'Callout.OrangeBg' = '#FDF1DC'
    'Callout.Red'    = '#B83227'; 'Callout.RedBg'    = '#FBE7E4'
    'Callout.Purple' = '#5A42AE'; 'Callout.PurpleBg' = '#EEEAFB'
    'Callout.Gray'   = '#4A5A6C'; 'Callout.GrayBg'   = '#EEF2F7'
    # 코드 블록 바탕과 색칠 (Highlight.psm1의 Kind 이름). 모두 Code.Bg 위에서 4.5:1 이상
    'Code.Bg'       = '#F3F6F9'
    'Code.Comment'  = '#56667A'
    'Code.String'   = '#0A3069'
    'Code.Keyword'  = '#C4243B'
    'Code.Number'   = '#0550AE'
    'Code.Command'  = '#6639BA'
    'Code.Variable' = '#953800'
    'Code.Type'     = '#116329'
    'Code.Attr'     = '#0B6E6D'
    'Code.Inserted' = '#1A6E44'
    'Code.Deleted'  = '#B83227'
    # 다이어그램 (Mermaid.psm1)
    'Dia.NodeBg'    = '#EEF2F7'
    'Dia.NodeLine'  = '#7F8C9C'   # 흰 바탕과 3:1 이상
    'Dia.Line'      = '#56667A'   # 연결선 · 화살표
    'Dia.GroupBg'   = '#F7F9FB'
    'Dia.GroupLine' = '#C4CEDA'
    'Dia.NoteBg'    = '#FDF1DC'
    'Dia.NoteLine'  = '#D9B66A'
}

$script:fontScales = @{ s = 0.93; m = 1.0; l = 1.12 }
# MdH1~3 · Code · BodyLine(본문 줄 높이 1.5배)은 마크다운 보기용
$script:fontSizes = @{ Title = 26; Number = 22; H2 = 16; Body = 14; Small = 12.5; Icon = 16; MdH1 = 24; MdH2 = 20; MdH3 = 17; Code = 13; BodyLine = 21 }
$script:fonts = @{
    # 설치된 글꼴 중 앞에서부터 쓴다. 맑은 고딕은 Windows 기본이라 항상 있다.
    UI   = 'Pretendard, Noto Sans KR, Malgun Gothic'
    Icon = 'Segoe Fluent Icons, Segoe MDL2 Assets'
    Mono = 'Consolas, D2Coding, Malgun Gothic'
}

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
    foreach ($k in $script:fontSizes.Keys) { $Dictionary["F.$k"] = Get-WtFontSize $k }
    foreach ($k in $script:fonts.Keys) { $Dictionary["Font.$k"] = [System.Windows.Media.FontFamily]::new($script:fonts[$k]) }
}

# 지금 설정의 글자 크기 (F.* 리소스와 같은 값). 화면에 붙기 전에 크기를 재야 할 때 쓴다 (다이어그램 배치 등).
function Get-WtFontSize([string]$Name) {
    $scale = $script:fontScales[[string](Get-WtSetting 'ui.fontScale')]
    if (-not $scale) { $scale = 1.0 }
    [double][Math]::Round($script:fontSizes[$Name] * $scale, 1)
}

# 글꼴 이름 목록 (Font.* 리소스와 같은 값). $Name: UI · Icon · Mono
function Get-WtFontName([string]$Name) { $script:fonts[$Name] }

function New-WtThemeDictionary {
    $d = [System.Windows.ResourceDictionary]::new()
    Update-WtThemeDictionary $d
    $d
}

Export-ModuleMember -Function ConvertFrom-WtHex, Get-WtMixedColor, Get-WtLuminance, Get-WtContrast, Get-WtPalette,
    Get-WtThemeColors, Update-WtThemeDictionary, New-WtThemeDictionary, Get-WtFontSize, Get-WtFontName
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

  <!-- 마크다운 보기 (Markdown.psm1). 문서 요소(Paragraph · Section · Run …)의 모양 -->
  <!-- RichTextBox는 배치할 때 문서 좌우에 5씩 여백을 넣는다(PagePadding 5,0). 안쪽 여백 -5로 없애서 옆 글과 줄을 맞춘다 -->
  <Style x:Key="Md.View" TargetType="RichTextBox">
    <Setter Property="Background" Value="Transparent"/>
    <Setter Property="BorderThickness" Value="0"/>
    <Setter Property="Padding" Value="-5,0,-5,0"/>
    <Setter Property="AcceptsTab" Value="False"/>
    <Setter Property="Foreground" Value="{DynamicResource B.Text}"/>
    <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
    <Setter Property="HorizontalScrollBarVisibility" Value="Disabled"/>
  </Style>
  <Style x:Key="Md.Doc" TargetType="FlowDocument">
    <Setter Property="FontFamily" Value="{DynamicResource Font.UI}"/>
    <Setter Property="FontSize" Value="{DynamicResource F.Body}"/>
    <Setter Property="Foreground" Value="{DynamicResource B.Text}"/>
    <Setter Property="LineHeight" Value="{DynamicResource F.BodyLine}"/>
    <Setter Property="PagePadding" Value="0"/>
    <Setter Property="TextAlignment" Value="Left"/>
  </Style>
  <Style x:Key="Md.P" TargetType="Paragraph">
    <Setter Property="Margin" Value="0,0,0,12"/>
  </Style>
  <Style x:Key="Md.PTight" TargetType="Paragraph">
    <Setter Property="Margin" Value="0"/>
  </Style>
  <Style x:Key="Md.H1" TargetType="Paragraph">
    <Setter Property="FontSize" Value="{DynamicResource F.MdH1}"/>
    <Setter Property="FontWeight" Value="Bold"/>
    <Setter Property="LineHeight" Value="Auto"/>
    <Setter Property="Margin" Value="0,24,0,12"/>
  </Style>
  <Style x:Key="Md.H2" TargetType="Paragraph" BasedOn="{StaticResource Md.H1}">
    <Setter Property="FontSize" Value="{DynamicResource F.MdH2}"/>
    <Setter Property="FontWeight" Value="SemiBold"/>
    <Setter Property="Margin" Value="0,24,0,8"/>
  </Style>
  <Style x:Key="Md.H3" TargetType="Paragraph" BasedOn="{StaticResource Md.H2}">
    <Setter Property="FontSize" Value="{DynamicResource F.MdH3}"/>
    <Setter Property="Margin" Value="0,16,0,8"/>
  </Style>
  <Style x:Key="Md.H4" TargetType="Paragraph" BasedOn="{StaticResource Md.H2}">
    <Setter Property="FontSize" Value="{DynamicResource F.H2}"/>
    <Setter Property="Margin" Value="0,16,0,4"/>
  </Style>
  <Style x:Key="Md.H5" TargetType="Paragraph" BasedOn="{StaticResource Md.H4}">
    <Setter Property="FontSize" Value="{DynamicResource F.Body}"/>
  </Style>
  <Style x:Key="Md.H6" TargetType="Paragraph" BasedOn="{StaticResource Md.H5}">
    <Setter Property="Foreground" Value="{DynamicResource B.SubText}"/>
  </Style>
  <Style x:Key="Md.List" TargetType="List">
    <Setter Property="Margin" Value="0,0,0,12"/>
    <Setter Property="Padding" Value="24,0,0,0"/>
  </Style>
  <Style x:Key="Md.ListNested" TargetType="List" BasedOn="{StaticResource Md.List}">
    <Setter Property="Margin" Value="0,4,0,0"/>
  </Style>
  <Style x:Key="Md.ListItem" TargetType="ListItem">
    <Setter Property="Margin" Value="0,0,0,4"/>
  </Style>
  <Style x:Key="Md.Quote" TargetType="Section">
    <Setter Property="BorderBrush" Value="{DynamicResource B.Md.QuoteBar}"/>
    <Setter Property="BorderThickness" Value="3,0,0,0"/>
    <Setter Property="Padding" Value="16,4,0,4"/>
    <Setter Property="Margin" Value="0,0,0,12"/>
    <Setter Property="Foreground" Value="{DynamicResource B.SubText}"/>
  </Style>
  <Style x:Key="Md.Callout" TargetType="Section">
    <Setter Property="BorderThickness" Value="3,0,0,0"/>
    <Setter Property="Padding" Value="16,12,16,12"/>
    <Setter Property="Margin" Value="0,0,0,12"/>
  </Style>
  <Style x:Key="Md.CalloutTitle" TargetType="Paragraph">
    <Setter Property="FontWeight" Value="SemiBold"/>
    <Setter Property="Margin" Value="0,0,0,8"/>
  </Style>
  <Style x:Key="Md.FoldLink" TargetType="Hyperlink">
    <Setter Property="TextDecorations" Value="{x:Null}"/>
    <Setter Property="Cursor" Value="Hand"/>
    <Style.Triggers>
      <Trigger Property="IsMouseOver" Value="True">
        <Setter Property="TextDecorations" Value="Underline"/>
      </Trigger>
    </Style.Triggers>
  </Style>
  <Style x:Key="Md.CodeBlock" TargetType="Section">
    <Setter Property="Background" Value="{DynamicResource B.Code.Bg}"/>
    <Setter Property="BorderBrush" Value="{DynamicResource B.Border}"/>
    <Setter Property="BorderThickness" Value="1"/>
    <Setter Property="Padding" Value="12,0,8,12"/>
    <Setter Property="Margin" Value="0,0,0,12"/>
  </Style>
  <Style x:Key="Md.CodeHead" TargetType="BlockUIContainer">
    <Setter Property="Margin" Value="0"/>
  </Style>
  <Style x:Key="Md.CodeText" TargetType="Paragraph">
    <Setter Property="FontFamily" Value="{DynamicResource Font.Mono}"/>
    <Setter Property="FontSize" Value="{DynamicResource F.Code}"/>
    <Setter Property="LineHeight" Value="Auto"/>
    <Setter Property="Margin" Value="0"/>
  </Style>
  <Style x:Key="Md.CodeNote" TargetType="Paragraph">
    <Setter Property="FontSize" Value="{DynamicResource F.Small}"/>
    <Setter Property="Foreground" Value="{DynamicResource B.Warn}"/>
    <Setter Property="Margin" Value="0,0,0,8"/>
  </Style>
  <Style x:Key="Md.InlineCode" TargetType="Run">
    <Setter Property="FontFamily" Value="{DynamicResource Font.Mono}"/>
    <Setter Property="FontSize" Value="{DynamicResource F.Code}"/>
    <Setter Property="Background" Value="{DynamicResource B.Code.Bg}"/>
  </Style>
  <Style x:Key="Md.Math" TargetType="Paragraph">
    <Setter Property="FontFamily" Value="{DynamicResource Font.Mono}"/>
    <Setter Property="Background" Value="{DynamicResource B.Code.Bg}"/>
    <Setter Property="Padding" Value="12,8"/>
    <Setter Property="Margin" Value="0,0,0,12"/>
    <Setter Property="TextAlignment" Value="Center"/>
  </Style>
  <Style x:Key="Md.Mark" TargetType="Span">
    <Setter Property="Background" Value="{DynamicResource B.Md.Mark}"/>
  </Style>
  <Style x:Key="Md.Sup" TargetType="Span">
    <Setter Property="BaselineAlignment" Value="Superscript"/>
    <Setter Property="FontSize" Value="{DynamicResource F.Small}"/>
  </Style>
  <Style x:Key="Md.Sub" TargetType="Span">
    <Setter Property="BaselineAlignment" Value="Subscript"/>
    <Setter Property="FontSize" Value="{DynamicResource F.Small}"/>
  </Style>
  <Style x:Key="Md.Tag" TargetType="Span">
    <Setter Property="Background" Value="{DynamicResource B.Chip}"/>
    <Setter Property="Foreground" Value="{DynamicResource B.ChipText}"/>
  </Style>
  <Style x:Key="Md.Link" TargetType="Hyperlink">
    <Setter Property="Foreground" Value="{DynamicResource B.AccentStrong}"/>
    <Setter Property="Cursor" Value="Hand"/>
    <Style.Triggers>
      <Trigger Property="IsMouseOver" Value="True">
        <Setter Property="Foreground" Value="{DynamicResource B.AccentHover}"/>
      </Trigger>
    </Style.Triggers>
  </Style>
  <Style x:Key="Md.TagLink" TargetType="Hyperlink" BasedOn="{StaticResource Md.Link}">
    <Setter Property="Background" Value="{DynamicResource B.AccentSoft}"/>
    <Setter Property="TextDecorations" Value="{x:Null}"/>
  </Style>
  <Style x:Key="Md.FnRef" TargetType="Hyperlink" BasedOn="{StaticResource Md.Link}">
    <Setter Property="BaselineAlignment" Value="Superscript"/>
    <Setter Property="FontSize" Value="{DynamicResource F.Small}"/>
    <Setter Property="TextDecorations" Value="{x:Null}"/>
  </Style>
  <Style x:Key="Md.Hr" TargetType="BlockUIContainer">
    <Setter Property="Margin" Value="0,8,0,20"/>
  </Style>
  <Style x:Key="Md.Figure" TargetType="BlockUIContainer">
    <Setter Property="Margin" Value="0,0,0,12"/>
  </Style>
  <Style x:Key="Md.FigureTight" TargetType="BlockUIContainer">
    <Setter Property="Margin" Value="0"/>
  </Style>
  <Style x:Key="Md.FigureBox" TargetType="Border">
    <Setter Property="Background" Value="{DynamicResource B.Card}"/>
    <Setter Property="BorderBrush" Value="{DynamicResource B.Border}"/>
    <Setter Property="BorderThickness" Value="1"/>
    <Setter Property="CornerRadius" Value="8"/>
    <Setter Property="Padding" Value="12,0,8,16"/>
  </Style>
  <Style x:Key="Md.Table" TargetType="Table">
    <Setter Property="CellSpacing" Value="0"/>
    <Setter Property="BorderBrush" Value="{DynamicResource B.Border}"/>
    <Setter Property="BorderThickness" Value="1"/>
    <Setter Property="Margin" Value="0,0,0,12"/>
  </Style>
  <Style x:Key="Md.Td" TargetType="TableCell">
    <Setter Property="Padding" Value="12,6"/>
    <Setter Property="BorderBrush" Value="{DynamicResource B.Border}"/>
    <Setter Property="BorderThickness" Value="0,0,0,1"/>
  </Style>
  <Style x:Key="Md.Th" TargetType="TableCell" BasedOn="{StaticResource Md.Td}">
    <Setter Property="FontWeight" Value="SemiBold"/>
    <Setter Property="Background" Value="{DynamicResource B.Code.Bg}"/>
  </Style>
  <Style x:Key="Md.Props" TargetType="Table" BasedOn="{StaticResource Md.Table}">
    <Setter Property="Background" Value="{DynamicResource B.Code.Bg}"/>
    <Setter Property="Margin" Value="0,0,0,16"/>
  </Style>
  <Style x:Key="Md.PropKey" TargetType="TableCell">
    <Setter Property="Padding" Value="12,4"/>
    <Setter Property="Foreground" Value="{DynamicResource B.SubText}"/>
  </Style>
  <Style x:Key="Md.PropValue" TargetType="TableCell">
    <Setter Property="Padding" Value="12,4"/>
  </Style>
  <Style x:Key="Md.Footnotes" TargetType="Section">
    <Setter Property="BorderBrush" Value="{DynamicResource B.Border}"/>
    <Setter Property="BorderThickness" Value="0,1,0,0"/>
    <Setter Property="Padding" Value="0,12,0,0"/>
    <Setter Property="Margin" Value="0,12,0,0"/>
    <Setter Property="FontSize" Value="{DynamicResource F.Small}"/>
    <Setter Property="LineHeight" Value="Auto"/>
    <Setter Property="Foreground" Value="{DynamicResource B.SubText}"/>
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
    # PowerShell 변수는 대소문자를 가리지 않는다: $title로 쓰면 [string]$Title에 들어가 글자로 바뀐다
    $head = New-WtText $Title 'T.H2'
    if ($Count) {
        $run = [System.Windows.Documents.Run]::new("  $Count")
        $run.FontWeight = 'Normal'
        $run.SetResourceReference([System.Windows.Documents.TextElement]::ForegroundProperty, 'B.SubText')
        $head.Inlines.Add($run)
    }
    $head.VerticalAlignment = 'Center'
    [void]$g.Children.Add($head)
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
    $b.Content = New-WtButtonContent $Text $Icon
    if ($ToolTip) { $b.ToolTip = $ToolTip }
    $b
}

# 버튼 안 내용: 아이콘 + 글자 (둘 중 하나만 있어도 됨). 누른 뒤 글자를 바꿀 때도 쓴다 (예: 복사 → 복사됨)
function New-WtButtonContent([string]$Text, [int]$Icon = 0) {
    if ($Icon -and $Text) {
        $row = New-WtStack Horizontal 8
        [void](Add-WtChild $row (New-WtIcon $Icon))
        $label = [System.Windows.Controls.TextBlock]::new()
        $label.Text = $Text
        $label.VerticalAlignment = 'Center'
        [void](Add-WtChild $row $label)
        return $row
    }
    if ($Icon) { return New-WtIcon $Icon }
    $Text
}

# 나눔 버튼: 여러 개 중 하나를 고르는 버튼 묶음 (예: 미리보기 | 원문, 작게 | 보통 | 크게)
# $Options: @(@('값', '글'), ...). 고르면 & $OnChange '값'. $Group은 화면 안에서 겹치지 않는 이름
function New-WtSegmented([object[]]$Options, [string]$Selected, [string]$Group, [scriptblock]$OnChange) {
    $box = [System.Windows.Controls.Border]::new()
    Set-WtResource $box 'Background' 'B.Chip'
    $box.CornerRadius = [System.Windows.CornerRadius]::new(10)
    $box.Padding = [System.Windows.Thickness]::new(4)
    $box.HorizontalAlignment = 'Left'
    $row = New-WtStack Horizontal 4
    foreach ($o in $Options) {
        $rb = [System.Windows.Controls.RadioButton]::new()
        Set-WtStyle $rb 'Seg'
        $rb.GroupName = $Group
        $rb.Content = $o[1]
        $rb.Tag = @{ Value = $o[0]; OnChange = $OnChange }
        $rb.IsChecked = ($o[0] -eq $Selected)
        $rb.Add_Checked({ param($s, $e) & $s.Tag.OnChange $s.Tag.Value })
        [void](Add-WtChild $row $rb)
    }
    $box.Child = $row
    $box
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
        $pos = if ($ErrorRecord.InvocationInfo) { $ErrorRecord.InvocationInfo.PositionMessage } else { '' }
        "$($ErrorRecord.Exception.Message)`n`n위치: $pos"
    } elseif ($ErrorRecord -is [System.Exception]) {
        "$($ErrorRecord.Message)`n`n$($ErrorRecord.StackTrace)"
    } else { [string]$ErrorRecord }
    if ($Where) { $msg = "[$Where]`n$msg" }
    try { Show-WtMessage $msg 'WorkTool 오류' 'Error' } catch { Write-Host $msg }
}

Export-ModuleMember -Function Get-WtStyleDictionary, ConvertFrom-WtXaml, Get-WtNamedElements, Set-WtStyle, Set-WtResource,
    New-WtText, New-WtIcon, New-WtStack, Add-WtChild, New-WtCard, New-WtCardHeader, New-WtButton, New-WtButtonContent, New-WtSegmented, New-WtTag,
    New-WtStatTile, New-WtPageHeader, New-WtPage, New-WtListRow, Format-WtDate, Show-WtMessage, Confirm-WtYesNo, Show-WtError
} | Import-Module -Force -DisableNameChecking
# ---- core\Ui.psm1 끝 ----

# ---- core\Highlight.psm1 ----
New-Module -Name 'WorkTool.Highlight' -ScriptBlock {
# 코드 색칠. 마크다운 코드 블록 등에서 쓴다. WPF 없이 동작한다 (글을 조각 목록으로 나누기만 한다).
# 언어 하나 = 규칙 목록 하나 (Register-WtCodeLanguage). 새 언어는 이 파일 아래쪽 목록에 하나 더 등록하면 된다.
#
# 규칙: @{ Kind = '색 이름'; Pattern = '정규식' } 또는 @{ Kind = …; Words = '낱말 낱말 …' }
# - 같은 자리에서는 위에 있는 규칙이 먼저다. 주석 · 문자열을 맨 위에 둔다.
# - Kind는 Theme.psm1의 Code.* 색 이름이다: Comment · String · Keyword · Number · Command · Variable · Type · Attr · Inserted · Deleted
# - 정규식은 여러 줄 모드(^ $가 줄 단위)다. 규칙 안에서는 이름 붙은 그룹을 쓰지 않는다.

$ErrorActionPreference = 'Stop'

$script:languages = @{}   # 이름 → @{ Name; Regex; Kinds }
$script:aliases = @{}     # 별칭 → 이름

function Register-WtCodeLanguage {
    param(
        [Parameter(Mandatory)][string]$Name,
        [string[]]$Aliases = @(),
        [Parameter(Mandatory)][object[]]$Rules,
        [switch]$IgnoreCase
    )
    $parts = for ($i = 0; $i -lt $Rules.Count; $i++) {
        $r = $Rules[$i]
        $pattern = if ($r.Words) {
            $words = @($r.Words -split '\s+' | Where-Object { $_ } | ForEach-Object { [regex]::Escape($_) })
            '(?<![\w$-])(?:' + ($words -join '|') + ')(?![\w-])'
        } else { $r.Pattern }
        "(?<k$i>$pattern)"
    }
    $options = [System.Text.RegularExpressions.RegexOptions]'Multiline, CultureInvariant'
    if ($IgnoreCase) { $options = $options -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase }
    $script:languages[$Name] = @{
        Name  = $Name
        Regex = [regex]::new(($parts -join '|'), $options)
        Kinds = @($Rules | ForEach-Object { $_.Kind })
    }
    foreach ($a in @($Name) + $Aliases) { $script:aliases[$a.ToLowerInvariant()] = $Name }
}

# 코드 블록 이름표(```powershell 의 powershell)로 언어를 찾는다. 없으면 $null
function Get-WtCodeLanguage([string]$Name) {
    if (-not $Name) { return $null }
    $key = $Name.Trim().ToLowerInvariant()
    if ($script:aliases.ContainsKey($key)) { return $script:aliases[$key] }
    $null
}

function Get-WtCodeLanguageNames { @($script:languages.Keys | Sort-Object) }

# 글을 @{ Kind; Text } 조각들로 나눈다. 색칠하지 않는 부분은 Kind가 ''다. 조각을 이으면 원래 글이 된다.
function Get-WtCodeTokens([string]$Code, [string]$Language) {
    $out = [System.Collections.Generic.List[object]]::new()
    if (-not $Code) { return }
    $name = Get-WtCodeLanguage $Language
    if (-not $name) { return @{ Kind = ''; Text = $Code } }
    $lang = $script:languages[$name]
    $pos = 0
    foreach ($m in $lang.Regex.Matches($Code)) {
        if ($m.Length -eq 0) { continue }
        if ($m.Index -gt $pos) { $out.Add(@{ Kind = ''; Text = $Code.Substring($pos, $m.Index - $pos) }) }
        $kind = ''
        for ($i = 0; $i -lt $lang.Kinds.Count; $i++) {
            if ($m.Groups["k$i"].Success) { $kind = $lang.Kinds[$i]; break }
        }
        $out.Add(@{ Kind = $kind; Text = $m.Value })
        $pos = $m.Index + $m.Length
    }
    if ($pos -lt $Code.Length) { $out.Add(@{ Kind = ''; Text = $Code.Substring($pos) }) }
    $out.ToArray()
}

# ---- 언어 목록 ---------------------------------------------------------------------------------

$num = '\b(?:0x[0-9a-fA-F]+|\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)\b'
$dq = '"(?:\\.|[^"\\\n])*"'
$sq = '''(?:\\.|[^''\\\n])*'''
$cBlock = '/\*[\s\S]*?(?:\*/|\z)'

Register-WtCodeLanguage -Name 'powershell' -Aliases 'ps1', 'ps', 'pwsh', 'posh', 'psm1' -IgnoreCase -Rules @(
    @{ Kind = 'Comment'; Pattern = '<#[\s\S]*?(?:#>|\z)|(?<![\w`])#.*' }
    @{ Kind = 'String'; Pattern = '@"[\s\S]*?^"@|@''[\s\S]*?^''@|"(?:`[\s\S]|""|[^"`])*"|''(?:''''|[^''])*''' }
    @{ Kind = 'Variable'; Pattern = '\$(?:\{[^}\n]*\}|(?:[A-Za-z]+:)?[\w?]+|\$|\^)' }
    @{ Kind = 'Type'; Pattern = '\[[A-Za-z][\w.]*(?:\[\])?(?:,\s*[\w.]+)*\](?!\s*\()' }
    @{ Kind = 'Command'; Pattern = '\b[A-Za-z]+-[A-Za-z][A-Za-z0-9]*\b' }
    @{ Kind = 'Keyword'; Pattern = '(?<![\w-])-(?:c|i)?(?:eq|ne|gt|ge|lt|le|like|notlike|match|notmatch|contains|notcontains|in|notin|replace|split)\b|(?<![\w-])-(?:and|or|not|xor|band|bor|bxor|bnot|shl|shr|is|isnot|as|join|f)\b' }
    @{ Kind = 'Keyword'; Words = 'begin break catch class clean continue data default do dynamicparam else elseif end enum exit filter finally for foreach function hidden if in param process return static switch throw trap try until using while' }
    @{ Kind = 'Attr'; Pattern = '(?<![\w-])-[A-Za-z][\w]*' }
    @{ Kind = 'Number'; Pattern = '\b\d+(?:\.\d+)?(?:[kmgtp]b)?\b|\b0x[0-9a-f]+\b' }
)

Register-WtCodeLanguage -Name 'bash' -Aliases 'sh', 'shell', 'zsh', 'console', 'shellscript' -Rules @(
    @{ Kind = 'Comment'; Pattern = '(?<=^|[ \t;])#.*' }
    @{ Kind = 'String'; Pattern = "$dq|'[^']*'" }
    @{ Kind = 'Variable'; Pattern = '\$(?:\{[^}\n]*\}|\w+|[@*#?$!0-9-])' }
    @{ Kind = 'Keyword'; Words = 'if then else elif fi for while until do done case esac function in return break continue local export readonly declare select time' }
    @{ Kind = 'Command'; Words = 'echo printf cd ls cat grep egrep sed awk find xargs sort uniq head tail wc cut tr tee sudo su chmod chown chgrp cp mv rm mkdir rmdir ln touch tar gzip gunzip zip unzip curl wget ssh scp rsync systemctl journalctl service kill killall ps top htop df du free mount umount lsblk blkid fdisk parted mkfs ip ifconfig ping traceroute ss netstat nmcli firewall-cmd iptables yum dnf apt apt-get rpm dpkg pip source set unset exit test read sleep date hostname uname whoami id useradd usermod passwd crontab nohup git docker kubectl podman vi vim nano less more which env alias' }
    @{ Kind = 'Attr'; Pattern = '(?<=\s)--?[A-Za-z][\w-]*' }
    @{ Kind = 'Number'; Pattern = '\b\d+(?:\.\d+)?\b' }
)

Register-WtCodeLanguage -Name 'python' -Aliases 'py', 'python3' -Rules @(
    @{ Kind = 'Comment'; Pattern = '#.*' }
    @{ Kind = 'String'; Pattern = '(?i:[rbfu]{0,2})(?:"""[\s\S]*?(?:"""|\z)|''''''[\s\S]*?(?:''''''|\z)|' + $dq + '|' + $sq + ')' }
    @{ Kind = 'Attr'; Pattern = '^[ \t]*@[\w.]+' }
    @{ Kind = 'Keyword'; Words = 'False None True and as assert async await break class continue def del elif else except finally for from global if import in is lambda nonlocal not or pass raise return try while with yield match case' }
    @{ Kind = 'Command'; Pattern = '(?<=\b(?:def|class)\s+)\w+|\b(?:print|len|range|open|str|int|float|bool|list|dict|set|tuple|type|isinstance|enumerate|zip|map|filter|sorted|sum|min|max|abs|any|all|super|input|format|round|repr|getattr|setattr|hasattr)\b(?=\s*\()' }
    @{ Kind = 'Variable'; Words = 'self cls' }
    @{ Kind = 'Number'; Pattern = $num }
)

Register-WtCodeLanguage -Name 'javascript' -Aliases 'js', 'jsx', 'mjs', 'cjs', 'typescript', 'ts', 'tsx' -Rules @(
    @{ Kind = 'Comment'; Pattern = "//.*|$cBlock" }
    @{ Kind = 'String'; Pattern = "$dq|$sq|``(?:\\[\s\S]|[^``\\])*``" }
    @{ Kind = 'Keyword'; Words = 'break case catch class const continue debugger default delete do else export extends finally for function if import in instanceof let new return super switch this throw try typeof var void while with yield async await of static get set interface type enum implements private public protected readonly as from null undefined true false' }
    @{ Kind = 'Type'; Pattern = '\b[A-Z][A-Za-z0-9_]*\b' }
    @{ Kind = 'Command'; Pattern = '\b[A-Za-z_$][\w$]*(?=\s*\()' }
    @{ Kind = 'Number'; Pattern = $num }
)

Register-WtCodeLanguage -Name 'json' -Aliases 'jsonc', 'json5' -Rules @(
    @{ Kind = 'Comment'; Pattern = "//.*|$cBlock" }
    @{ Kind = 'Attr'; Pattern = '"(?:\\.|[^"\\\n])*"(?=\s*:)' }
    @{ Kind = 'String'; Pattern = $dq }
    @{ Kind = 'Keyword'; Words = 'true false null' }
    @{ Kind = 'Number'; Pattern = '-?\b\d+(?:\.\d+)?(?:[eE][+-]?\d+)?\b' }
)

Register-WtCodeLanguage -Name 'yaml' -Aliases 'yml' -Rules @(
    @{ Kind = 'Comment'; Pattern = '(?<=^|\s)#.*' }
    @{ Kind = 'Attr'; Pattern = '(?<=^[ \t]*(?:-[ \t]+)?)[^\s#:''"\-][^:#\n]*?(?=[ \t]*:(?:[ \t]|$))' }
    @{ Kind = 'String'; Pattern = "$dq|'(?:''|[^'\n])*'" }
    @{ Kind = 'Type'; Pattern = '^(?:---|\.\.\.)[ \t]*$|!!?[\w/]+' }
    @{ Kind = 'Variable'; Pattern = '[&*][\w-]+' }
    @{ Kind = 'Keyword'; Pattern = '(?i)(?<=:[ \t]+|-[ \t]+|^)(?:true|false|null|yes|no|on|off|~)(?=[ \t]*(?:#|$))' }
    @{ Kind = 'Number'; Pattern = '(?<=:[ \t]+|-[ \t]+)-?\d+(?:\.\d+)?(?=[ \t]*(?:#|$))' }
)

Register-WtCodeLanguage -Name 'xml' -Aliases 'html', 'htm', 'xaml', 'svg', 'xhtml', 'plist' -IgnoreCase -Rules @(
    @{ Kind = 'Comment'; Pattern = '<!--[\s\S]*?(?:-->|\z)' }
    @{ Kind = 'Keyword'; Pattern = '<!\[CDATA\[|\]\]>|<!DOCTYPE[^>]*>|<\?[\s\S]*?\?>' }
    @{ Kind = 'Type'; Pattern = '(?<=</?)[\w:.-]+' }
    @{ Kind = 'String'; Pattern = '(?<==\s*)(?:"[^"]*"|''[^'']*'')' }
    @{ Kind = 'Attr'; Pattern = '(?<=<[^<>]*\s)[\w:.-]+(?=\s*=)' }
    @{ Kind = 'Variable'; Pattern = '&[#\w]+;' }
)

Register-WtCodeLanguage -Name 'sql' -Aliases 'mysql', 'tsql', 'plsql', 'postgresql', 'psql', 'sqlite' -IgnoreCase -Rules @(
    @{ Kind = 'Comment'; Pattern = "--.*|$cBlock" }
    @{ Kind = 'String'; Pattern = "'(?:''|[^'])*'" }
    @{ Kind = 'Type'; Words = 'int integer bigint smallint tinyint varchar nvarchar char nchar text ntext date datetime datetime2 time timestamp decimal numeric float real bit boolean money uniqueidentifier blob clob' }
    @{ Kind = 'Keyword'; Words = 'select from where and or not insert into values update set delete create alter drop truncate table index view database schema join left right inner outer full cross on group by order having limit offset top distinct as is null like in between exists case when then else end union all primary key foreign references default constraint unique check begin commit rollback transaction declare exec execute procedure function returns return if while grant revoke with desc asc use go' }
    @{ Kind = 'Command'; Pattern = '\b(?:count|sum|avg|min|max|coalesce|isnull|ifnull|nullif|cast|convert|concat|substring|len|length|upper|lower|trim|getdate|now|row_number|rank|dense_rank)\b(?=\s*\()' }
    @{ Kind = 'Variable'; Pattern = '@@?\w+' }
    @{ Kind = 'Number'; Pattern = '\b\d+(?:\.\d+)?\b' }
)

Register-WtCodeLanguage -Name 'csharp' -Aliases 'cs', 'c#' -Rules @(
    @{ Kind = 'Comment'; Pattern = "//.*|$cBlock" }
    @{ Kind = 'String'; Pattern = '@"(?:""|[^"])*"|\$?' + $dq + "|'(?:\\.|[^'\\\n])'" }
    @{ Kind = 'Keyword'; Words = 'abstract as base bool break byte case catch char checked class const continue decimal default delegate do double else enum event explicit extern false finally fixed float for foreach goto if implicit in int interface internal is lock long namespace new null object operator out override params private protected public readonly ref return sbyte sealed short sizeof stackalloc static string struct switch this throw true try typeof uint ulong unchecked unsafe ushort using virtual void volatile while var async await get set value record init' }
    @{ Kind = 'Type'; Pattern = '\b[A-Z]\w*\b' }
    @{ Kind = 'Command'; Pattern = '\b[a-z_]\w*(?=\s*\()' }
    @{ Kind = 'Number'; Pattern = $num }
)

Register-WtCodeLanguage -Name 'ini' -Aliases 'toml', 'conf', 'cfg', 'properties', 'env', 'dotenv', 'editorconfig' -Rules @(
    @{ Kind = 'Comment'; Pattern = '^[ \t]*[;#].*' }
    @{ Kind = 'Type'; Pattern = '^[ \t]*\[[^\]\n]*\]' }
    @{ Kind = 'Attr'; Pattern = '^[ \t]*[^=:\s\[;#][^=:\n]*?(?=[ \t]*[=:])' }
    @{ Kind = 'String'; Pattern = "$dq|'[^'\n]*'" }
    @{ Kind = 'Keyword'; Pattern = '(?i)(?<=[=:][ \t]*)(?:true|false|yes|no|on|off)\b' }
    @{ Kind = 'Number'; Pattern = '(?<=[=:][ \t]*)-?\d+(?:\.\d+)?\b' }
)

Register-WtCodeLanguage -Name 'diff' -Aliases 'patch', 'udiff' -Rules @(
    @{ Kind = 'Type'; Pattern = '^(?:@@.*|diff .*|index .*|---(?: .*)?|\+\+\+(?: .*)?)$' }
    @{ Kind = 'Inserted'; Pattern = '^[+>].*' }
    @{ Kind = 'Deleted'; Pattern = '^[-<].*' }
)

Register-WtCodeLanguage -Name 'css' -Aliases 'scss', 'less' -Rules @(
    @{ Kind = 'Comment'; Pattern = "$cBlock|(?<=^|\s)//.*" }
    @{ Kind = 'String'; Pattern = "$dq|$sq" }
    @{ Kind = 'Keyword'; Pattern = '@[\w-]+|!important' }
    @{ Kind = 'Attr'; Pattern = '(?<=[{;\s])[\w-]+(?=\s*:[^{}\n]*[;}\n])' }
    @{ Kind = 'Number'; Pattern = '#[0-9a-fA-F]{3,8}\b|-?\b\d+(?:\.\d+)?(?:px|em|rem|%|vh|vw|s|ms|pt|deg)?' }
    @{ Kind = 'Type'; Pattern = '[.#][A-Za-z_][\w-]*' }
)

Register-WtCodeLanguage -Name 'dockerfile' -Aliases 'docker', 'containerfile' -IgnoreCase -Rules @(
    @{ Kind = 'Comment'; Pattern = '^[ \t]*#.*' }
    @{ Kind = 'Keyword'; Pattern = '^[ \t]*(?:FROM|RUN|CMD|LABEL|MAINTAINER|EXPOSE|ENV|ADD|COPY|ENTRYPOINT|VOLUME|USER|WORKDIR|ARG|ONBUILD|STOPSIGNAL|HEALTHCHECK|SHELL)\b|\bAS\b' }
    @{ Kind = 'String'; Pattern = "$dq|'[^'\n]*'" }
    @{ Kind = 'Variable'; Pattern = '\$(?:\{[^}\n]*\}|\w+)' }
    @{ Kind = 'Attr'; Pattern = '(?<=\s)--?[A-Za-z][\w-]*' }
)

Register-WtCodeLanguage -Name 'batch' -Aliases 'bat', 'cmd' -IgnoreCase -Rules @(
    @{ Kind = 'Comment'; Pattern = '^[ \t@]*(?:rem(?:\s.*)?|::.*)$' }
    @{ Kind = 'String'; Pattern = '"[^"\n]*"' }
    @{ Kind = 'Variable'; Pattern = '%~?[\w]*%|%%~?\w|%~?\d|![\w]+!' }
    @{ Kind = 'Type'; Pattern = '^[ \t]*:\w+' }
    @{ Kind = 'Keyword'; Words = 'if else for in do goto call set setlocal endlocal echo exit not exist defined errorlevel equ neq lss leq gtr geq pause shift start cd pushd popd title' }
    @{ Kind = 'Attr'; Pattern = '(?<=\s)/[\w?]+' }
    @{ Kind = 'Number'; Pattern = '\b\d+\b' }
)

Export-ModuleMember -Function Register-WtCodeLanguage, Get-WtCodeLanguage, Get-WtCodeLanguageNames, Get-WtCodeTokens
} | Import-Module -Force -DisableNameChecking
# ---- core\Highlight.psm1 끝 ----

# ---- core\Markdown.psm1 ----
New-Module -Name 'WorkTool.Markdown' -ScriptBlock {
# 마크다운 보기. 메모 · 문서 등 여러 화면이 같이 쓴다. 이 파일에는 특정 화면(메모 등)을 위한 코드를 넣지 않는다.
#
# 1) ConvertTo-WtMarkdownModel: 글 → 블록 · 인라인 목록(해시테이블). WPF 없이 동작한다 (tests/Test-Core.ps1이 검사).
#    PowerShell 7에 들어 있는 Markdig로 읽고, Markdig가 모르는 Obsidian 문법(콜아웃, [[위키 링크]], #태그, %%주석%%)은 여기서 더 읽는다.
# 2) New-WtMarkdownView + Set-WtMarkdownText: 모델 → WPF 문서(읽기 전용 RichTextBox).
#    모양은 Ui.psm1의 Md.* 스타일과 Theme.psm1의 색 · 크기 리소스만 쓴다.
#
# 화면마다 다른 것은 매개변수 · 등록으로 받는다
# - 링크 · [[위키 링크]] · #태그를 눌렀을 때: New-WtMarkdownView -OnLink { param($link) ...; $true }
#   $link = @{ Kind = 'url' | 'file' | 'wiki' | 'tag'; Target; Text }. $true를 돌려주면 처리한 것, 아니면 기본 동작(Invoke-WtMarkdownLink).
# - 상대 경로(그림 · 파일 링크)의 기준 폴더: -BaseDir (메모처럼 파일이 아닌 글은 비워 둔다)
# - 코드 블록을 언어별로 따로 그리기: Register-WtMarkdownCodeRenderer (예: Mermaid.psm1이 mermaid를 등록)
# - 콜아웃 종류 더하기: Register-WtMarkdownCallout
#
# 이벤트 처리기는 나중에 불리므로 보낸 요소($s)의 Tag와 $script: 변수만 쓴다.

$ErrorActionPreference = 'Stop'

$script:pipeline = $null
$script:codeRenderers = @{}   # 언어 이름(소문자) → @{ Label; Render; Remeasure }
$script:callouts = @{}        # 종류(소문자) → @{ Kind; Color; Icon }
$script:views = [System.Collections.Generic.List[System.WeakReference]]::new()

# 링크로 바로 여는 파일 종류. 실행 파일 · 스크립트는 열지 않고 폴더에서 보여 준다.
$script:openable = @('.md', '.txt', '.log', '.csv', '.tsv', '.json', '.xml', '.yaml', '.yml', '.ini', '.conf', '.cfg',
    '.pdf', '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp', '.svg',
    '.xlsx', '.xls', '.xlsm', '.docx', '.doc', '.pptx', '.ppt', '.vsdx', '.vsd', '.htm', '.html', '.zip')
$script:imageExt = @('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tif', '.tiff', '.ico', '.webp')

# ---- 등록 -------------------------------------------------------------------------------------

<#
코드 블록 그리기 등록
  Language   ```뒤에 쓰는 이름 (여러 개 가능)
  Label      코드 블록 머리에 보일 이름
  Render     { param($code) ... } → 화면 요소 하나. 못 그리면 throw → 원문을 코드 블록으로 보여 준다
  Remeasure  글자 크기를 바꾸면 다시 그린다 (글자 크기를 재서 배치하는 그림)
#>
function Register-WtMarkdownCodeRenderer {
    param(
        [Parameter(Mandatory)][string[]]$Language,
        [Parameter(Mandatory)][scriptblock]$Render,
        [string]$Label = '',
        [switch]$Remeasure
    )
    foreach ($l in $Language) {
        $script:codeRenderers[$l.ToLowerInvariant()] = @{ Label = $(if ($Label) { $Label } else { $l }); Render = $Render; Remeasure = [bool]$Remeasure }
    }
}

<#
콜아웃 종류 등록. > [!종류] 제목
  Color  Theme.psm1의 Callout.<Color> · Callout.<Color>Bg (Blue · Teal · Green · Orange · Red · Purple · Gray)
  Icon   Segoe 아이콘 글꼴 글자 코드
#>
function Register-WtMarkdownCallout {
    param(
        [Parameter(Mandatory)][string]$Kind,
        [Parameter(Mandatory)][string]$Color,
        [Parameter(Mandatory)][int]$Icon,
        [string[]]$Aliases = @()
    )
    $c = @{ Kind = $Kind; Color = $Color; Icon = $Icon }
    foreach ($k in @($Kind) + $Aliases) { $script:callouts[$k.ToLowerInvariant()] = $c }
}

# Obsidian 기본 콜아웃 (GitHub의 > [!NOTE] · [!TIP] · [!IMPORTANT] · [!WARNING] · [!CAUTION]도 여기에 들어 있다)
Register-WtMarkdownCallout -Kind 'note' -Color 'Blue' -Icon 0xE70F
Register-WtMarkdownCallout -Kind 'abstract' -Color 'Teal' -Icon 0xE8A5 -Aliases 'summary', 'tldr'
Register-WtMarkdownCallout -Kind 'info' -Color 'Blue' -Icon 0xE946
Register-WtMarkdownCallout -Kind 'todo' -Color 'Blue' -Icon 0xE73A
Register-WtMarkdownCallout -Kind 'tip' -Color 'Teal' -Icon 0xEA80 -Aliases 'hint', 'important'
Register-WtMarkdownCallout -Kind 'success' -Color 'Green' -Icon 0xE73E -Aliases 'check', 'done'
Register-WtMarkdownCallout -Kind 'question' -Color 'Orange' -Icon 0xE9CE -Aliases 'help', 'faq'
Register-WtMarkdownCallout -Kind 'warning' -Color 'Orange' -Icon 0xE7BA -Aliases 'caution', 'attention'
Register-WtMarkdownCallout -Kind 'failure' -Color 'Red' -Icon 0xE711 -Aliases 'fail', 'missing'
Register-WtMarkdownCallout -Kind 'danger' -Color 'Red' -Icon 0xEA39 -Aliases 'error'
Register-WtMarkdownCallout -Kind 'bug' -Color 'Red' -Icon 0xEBE8
Register-WtMarkdownCallout -Kind 'example' -Color 'Purple' -Icon 0xE8FD
Register-WtMarkdownCallout -Kind 'quote' -Color 'Gray' -Icon 0xE8BD -Aliases 'cite'

# ---- 글 → 모델 (WPF 없이) ---------------------------------------------------------------------

<#
모델
  블록: @{ T = 'heading' (Level, Inl, Text) | 'para' (Inl) | 'list' (Ordered, Start, Loose, Items[@{ Task; Blocks; Line }])
             | 'quote' (Blocks) | 'callout' (Kind, Color, Icon, Title[인라인], Fold '' · '+' · '-', Blocks)
             | 'code' (Lang, Text) | 'math' (Text) | 'hr' | 'html' (Text) | 'table' (Align[], Rows[@{ Header; Cells[블록[]] }])
             | 'props' (Items[@{ Key; Value }]) | 'footnotes' (Items[@{ Order; Blocks }]); Line = 원문 줄 번호(1부터) }
  인라인: @{ T = 'text' (Text) | 'soft'(줄바꿈 한 번) | 'br' | 'em' · 'strong' · 'del' · 'mark' · 'sup' · 'sub' · 'ins' · 'span' (Inl)
             | 'code' (Text) | 'math' (Text) | 'link' (Kind, Target, Url, Title, Inl) | 'image' (Url, Alt, Width)
             | 'tag' (Text) | 'fn' · 'fnback' (Order) | 'task' (Checked) }
  링크 Kind: url · file · wiki · anchor(같은 문서 안 제목 · 각주)
#>
function ConvertTo-WtMarkdownModel([AllowEmptyString()][string]$Text) {
    $src = ([string]$Text).Replace("`r`n", "`n").Replace("`r", "`n")
    $doc = [Markdig.Markdown]::Parse($src, (Get-WtMarkdownPipeline))
    $ctx = @{ Src = $src; Headings = [System.Collections.Generic.List[object]]::new() }
    $blocks = ConvertFrom-WtMdBlocks $doc $ctx
    @{ Blocks = $blocks; Headings = $ctx.Headings.ToArray() }
}

# Markdig는 PowerShell 7에 들어 있다 (ConvertFrom-Markdown이 쓰는 것). 처음 한 번 불러와서 쓸 확장만 켠다.
function Get-WtMarkdownPipeline {
    if (-not $script:pipeline) {
        if (-not ('Markdig.Markdown' -as [type])) { [void](ConvertFrom-Markdown -InputObject ' ') }
        $ext = [Markdig.MarkdownExtensions]
        $b = [Markdig.MarkdownPipelineBuilder]::new()
        $b = $ext::UseYamlFrontMatter($b)
        $b = $ext::UsePipeTables($b)
        $b = $ext::UseTaskLists($b)
        $b = $ext::UseEmphasisExtras($b, [Markdig.Extensions.EmphasisExtras.EmphasisExtraOptions]'Strikethrough, Marked')
        $b = $ext::UseAutoLinks($b)
        $b = $ext::UseFootnotes($b)
        $b = $ext::UseMathematics($b)
        $script:pipeline = $b.Build()
    }
    $script:pipeline
}

function ConvertFrom-WtMdBlocks($Container, $Ctx) {
    $out = [System.Collections.Generic.List[object]]::new()
    foreach ($b in $Container) {
        $m = ConvertFrom-WtMdBlock $b $Ctx
        if ($m) { $out.Add($m) }
    }
    , $out.ToArray()
}

function Get-WtMdLines($Block) { ([string]$Block.Lines.ToString()).Replace("`r", '').TrimEnd("`n") }

function ConvertFrom-WtMdBlock($B, $Ctx) {
    $line = $B.Line + 1
    switch ($B.GetType().Name) {
        'HeadingBlock' {
            $inl = ConvertFrom-WtMdInlines $B.Inline $Ctx
            $text = Get-WtMdInlineText $inl
            $Ctx.Headings.Add(@{ Level = $B.Level; Text = $text; Line = $line })
            return @{ T = 'heading'; Level = $B.Level; Inl = $inl; Text = $text; Line = $line }
        }
        'ParagraphBlock' { return @{ T = 'para'; Inl = (ConvertFrom-WtMdInlines $B.Inline $Ctx); Line = $line } }
        'ListBlock' { return ConvertFrom-WtMdList $B $Ctx }
        'QuoteBlock' { return ConvertFrom-WtMdQuote $B $Ctx }
        'MathBlock' { return @{ T = 'math'; Text = (Get-WtMdLines $B); Line = $line } }
        'FencedCodeBlock' { return @{ T = 'code'; Lang = ([string]$B.Info).Trim(); Text = (Get-WtMdLines $B); Line = $line } }
        'CodeBlock' { return @{ T = 'code'; Lang = ''; Text = (Get-WtMdLines $B); Line = $line } }
        'ThematicBreakBlock' { return @{ T = 'hr'; Line = $line } }
        'HtmlBlock' {
            $raw = Get-WtMdLines $B
            if ($raw.TrimStart().StartsWith('<!--')) { return $null }
            # HTML은 그리지 않는다. <br>은 줄바꿈으로, 나머지 태그는 빼고 글자만 보인다
            $text = [System.Net.WebUtility]::HtmlDecode(($raw -replace '<br\s*/?>', "`n" -replace '<[^>]+>', '')).Trim()
            if (-not $text) { return $null }
            return @{ T = 'html'; Text = $text; Line = $line }
        }
        'Table' { return ConvertFrom-WtMdTable $B $Ctx }
        'YamlFrontMatterBlock' { return ConvertFrom-WtMdFrontMatter (Get-WtMdLines $B) }
        'FootnoteGroup' {
            $items = foreach ($fn in $B) { @{ Order = $fn.Order; Blocks = (ConvertFrom-WtMdBlocks $fn $Ctx) } }
            return @{ T = 'footnotes'; Items = @($items); Line = $line }
        }
    }
    $null   # 링크 정의 등 보이지 않는 블록
}

function ConvertFrom-WtMdList($B, $Ctx) {
    $items = [System.Collections.Generic.List[object]]::new()
    foreach ($li in $B) {
        $blocks = ConvertFrom-WtMdBlocks $li $Ctx
        $task = $null
        if ($blocks.Count -and $blocks[0].T -eq 'para' -and $blocks[0].Inl.Count -and $blocks[0].Inl[0].T -eq 'task') {
            $task = [bool]$blocks[0].Inl[0].Checked
            $rest = @($blocks[0].Inl | Select-Object -Skip 1)
            if ($rest.Count -and $rest[0].T -eq 'text') { $rest[0].Text = $rest[0].Text.TrimStart() }
            $blocks[0].Inl = $rest
        }
        $items.Add(@{ Task = $task; Blocks = $blocks; Line = $li.Line + 1 })
    }
    $start = 1
    if ($B.IsOrdered) { [void][int]::TryParse([string]$B.OrderedStart, [ref]$start) }
    @{ T = 'list'; Ordered = [bool]$B.IsOrdered; Start = $start; Loose = [bool]$B.IsLoose; Items = $items.ToArray(); Line = $B.Line + 1 }
}

# 인용 블록. 첫 줄이 [!종류]이면 콜아웃 (Obsidian · GitHub)
function ConvertFrom-WtMdQuote($B, $Ctx) {
    $blocks = ConvertFrom-WtMdBlocks $B $Ctx
    $first = $null
    foreach ($c in $B) { $first = $c; break }
    if (-not $first -or $first.GetType().Name -ne 'ParagraphBlock' -or -not $blocks.Count) {
        return @{ T = 'quote'; Blocks = $blocks; Line = $B.Line + 1 }
    }
    $start = $first.Span.Start
    $end = $Ctx.Src.IndexOf("`n", $start)
    $head = if ($end -lt 0) { $Ctx.Src.Substring($start) } else { $Ctx.Src.Substring($start, $end - $start) }
    $m = [regex]::Match($head, '^\[!([A-Za-z][\w-]*)\]([+-]?)')
    if (-not $m.Success) { return @{ T = 'quote'; Blocks = $blocks; Line = $B.Line + 1 } }

    # 첫 문단을 첫 줄(제목)과 나머지(본문)로 나눈다
    $title = [System.Collections.Generic.List[object]]::new()
    $body = [System.Collections.Generic.List[object]]::new()
    $inBody = $false
    foreach ($x in $blocks[0].Inl) {
        if (-not $inBody -and $x.T -in 'soft', 'br') { $inBody = $true; continue }
        if ($inBody) { $body.Add($x) } else { $title.Add($x) }
    }
    if ($title.Count -and $title[0].T -eq 'text') { $title[0].Text = $title[0].Text -replace '^\s*\[![^\]]*\][+-]?\s*', '' }
    if ($title.Count -and $title[$title.Count - 1].T -eq 'text') { $title[$title.Count - 1].Text = $title[$title.Count - 1].Text.TrimEnd() }
    $title = @($title | Where-Object { -not ($_.T -eq 'text' -and -not $_.Text) })

    $raw = $m.Groups[1].Value.ToLowerInvariant()
    $info = $script:callouts[$raw]
    if (-not $info) { $info = $script:callouts['note'] }
    if (-not $title.Count) {
        $name = [System.Globalization.CultureInfo]::InvariantCulture.TextInfo.ToTitleCase($raw)
        $title = @(@{ T = 'text'; Text = $name })
    }
    $rest = [System.Collections.Generic.List[object]]::new()
    if ($body.Count) { $rest.Add(@{ T = 'para'; Inl = $body.ToArray(); Line = $blocks[0].Line + 1 }) }
    for ($i = 1; $i -lt $blocks.Count; $i++) { $rest.Add($blocks[$i]) }
    @{
        T = 'callout'; Kind = $info.Kind; Raw = $raw; Color = $info.Color; Icon = $info.Icon
        Title = $title; Fold = $m.Groups[2].Value; Blocks = $rest.ToArray(); Line = $B.Line + 1
    }
}

function ConvertFrom-WtMdTable($B, $Ctx) {
    $align = @(foreach ($cd in $B.ColumnDefinitions) {
            switch ([string]$cd.Alignment) { 'Left' { 'left' } 'Center' { 'center' } 'Right' { 'right' } default { '' } }
        })
    $rows = [System.Collections.Generic.List[object]]::new()
    $n = 0
    foreach ($row in $B) {
        $cells = [System.Collections.Generic.List[object]]::new()
        foreach ($cell in $row) { $cells.Add((ConvertFrom-WtMdBlocks $cell $Ctx)) }
        $rows.Add(@{ Header = [bool]$row.IsHeader; Cells = $cells.ToArray() })
        $n = [Math]::Max($n, $cells.Count)
    }
    # Markdig는 끝의 | 뒤에 빈 열 정의를 하나 더 둔다
    if ($align.Count -gt $n) { $align = @($align | Select-Object -First $n) }
    @{ T = 'table'; Align = $align; Rows = $rows.ToArray(); Line = $B.Line + 1 }
}

# 맨 앞 --- 사이의 속성(YAML). 'key: 값', 'key: [a, b]', 그리고 아래에 '- 값'을 늘어놓은 목록만 읽는다.
function ConvertFrom-WtMdFrontMatter([string]$Raw) {
    $items = [System.Collections.Generic.List[object]]::new()
    $cur = $null
    foreach ($l in ($Raw -split "`n")) {
        if ($l -match '^\s*(---|\.\.\.)\s*$' -or -not $l.Trim() -or $l.TrimStart().StartsWith('#')) { continue }
        if ($l -match '^([^\s:#-][^:]*):\s*(.*)$') {
            $key = $Matches[1].Trim()
            $v = $Matches[2].Trim()
            if ($v -match '^\[(.*)\]$') { $v = (($Matches[1] -split ',') | ForEach-Object { $_.Trim().Trim('"', "'") }) -join ', ' }
            $cur = @{ Key = $key; Value = $v.Trim('"', "'") }
            $items.Add($cur)
        } elseif ($cur -and $l -match '^\s+-\s*(.*)$') {
            $v = $Matches[1].Trim().Trim('"', "'")
            $cur.Value = if ($cur.Value) { "$($cur.Value), $v" } else { $v }
        }
    }
    if (-not $items.Count) { return $null }
    @{ T = 'props'; Items = $items.ToArray(); Line = 1 }
}

# 이어진 글자 조각을 모았다가 Obsidian 문법을 읽는다. 줄바꿈 한 번(soft)도 글자 안에 넣어서 %%여러 줄 주석%%도 읽는다.
$script:obsidianRe = [regex]::new(
    '%%[\s\S]*?%%|(?<embed>!)?\[\[(?<target>[^\[\]|\n]+?)(?:\|(?<alias>[^\[\]\n]*))?\]\]|(?<=^|[\s(\[])#(?<tag>[\p{L}\p{N}_/-]*[\p{L}_/-][\p{L}\p{N}_/-]*)',
    'Multiline')

function ConvertFrom-WtMdInlines($Container, $Ctx) {
    $out = [System.Collections.Generic.List[object]]::new()
    if ($null -eq $Container) { return , $out.ToArray() }
    $buf = [System.Text.StringBuilder]::new()
    foreach ($i in $Container) {
        $name = $i.GetType().Name
        if ($name -eq 'LiteralInline') { [void]$buf.Append($i.Content.ToString()); continue }
        if ($name -eq 'HtmlEntityInline') { [void]$buf.Append($i.Transcoded.ToString()); continue }
        if ($name -eq 'LineBreakInline' -and -not $i.IsHard) { [void]$buf.Append("`n"); continue }
        if ($name -eq 'HtmlInline') {
            # <br>만 줄바꿈으로 쓰고, 다른 태그는 빼고 안의 글자만 보인다
            if ([string]$i.Tag -match '^<br\s*/?>$') {
                Add-WtMdText $out $buf
                $out.Add(@{ T = 'br' })
            }
            continue
        }
        Add-WtMdText $out $buf
        $m = ConvertFrom-WtMdInline $i $Ctx
        if ($m) { $out.Add($m) }
    }
    Add-WtMdText $out $buf
    , $out.ToArray()
}

function ConvertFrom-WtMdInline($I, $Ctx) {
    switch ($I.GetType().Name) {
        'EmphasisInline' {
            $kind = switch ("$($I.DelimiterChar)$($I.DelimiterCount)") {
                { $_ -in '*1', '_1' } { 'em' }
                { $_ -in '*2', '_2' } { 'strong' }
                '~2' { 'del' }
                '~1' { 'sub' }
                '=2' { 'mark' }
                '^1' { 'sup' }
                '+2' { 'ins' }
                default { 'span' }
            }
            return @{ T = $kind; Inl = (ConvertFrom-WtMdInlines $I $Ctx) }
        }
        'CodeInline' { return @{ T = 'code'; Text = [string]$I.Content } }
        'LinkInline' {
            $inl = ConvertFrom-WtMdInlines $I $Ctx
            if ($I.IsImage) { return New-WtMdImageModel ([string]$I.Url) (Get-WtMdInlineText $inl) }
            return New-WtMdLinkModel ([string]$I.Url) $inl ([string]$I.Title)
        }
        'AutolinkInline' {
            $url = if ($I.IsEmail) { "mailto:$($I.Url)" } else { [string]$I.Url }
            return New-WtMdLinkModel $url @(@{ T = 'text'; Text = [string]$I.Url })
        }
        'LineBreakInline' { return @{ T = 'br' } }
        'FootnoteLink' { return @{ T = $(if ($I.IsBackLink) { 'fnback' } else { 'fn' }); Order = $I.Footnote.Order } }
        'MathInline' { return @{ T = 'math'; Text = $I.Content.ToString() } }
        'TaskList' { return @{ T = 'task'; Checked = [bool]$I.Checked } }
    }
    if ($I -is [Markdig.Syntax.Inlines.ContainerInline]) { return @{ T = 'span'; Inl = (ConvertFrom-WtMdInlines $I $Ctx) } }
    @{ T = 'text'; Text = $I.ToString() }
}

# 모은 글자를 Obsidian 문법(위키 링크 · 태그 · 주석)과 줄바꿈으로 나눠 넣는다
function Add-WtMdText($Out, [System.Text.StringBuilder]$Buf) {
    if (-not $Buf.Length) { return }
    $text = $Buf.ToString()
    [void]$Buf.Clear()
    $pos = 0
    foreach ($m in $script:obsidianRe.Matches($text)) {
        if ($m.Index -gt $pos) { Add-WtMdPlain $Out $text.Substring($pos, $m.Index - $pos) }
        $pos = $m.Index + $m.Length
        if ($m.Groups['tag'].Success) {
            $Out.Add(@{ T = 'tag'; Text = $m.Groups['tag'].Value })
        } elseif ($m.Groups['target'].Success) {
            $target = $m.Groups['target'].Value.Trim()
            $alias = if ($m.Groups['alias'].Success) { $m.Groups['alias'].Value.Trim() } else { '' }
            $isImage = $script:imageExt -contains [System.IO.Path]::GetExtension($target).ToLowerInvariant()
            if ($m.Groups['embed'].Success -and $isImage) {
                $Out.Add((New-WtMdImageModel $target $(if ($alias) { "$target|$alias" } else { $target })))
                continue
            }
            $show = if ($alias) { $alias } elseif ($target.StartsWith('#')) { $target.Substring(1) } else { $target -replace '#', ' › ' }
            $kind = if ($target.StartsWith('#')) { 'anchor' } else { 'wiki' }
            $Out.Add(@{ T = 'link'; Kind = $kind; Target = $target.TrimStart('#'); Url = $target; Title = ''; Inl = @(@{ T = 'text'; Text = $show }) })
        }
        # %%주석%%은 보이지 않는다
    }
    if ($pos -lt $text.Length) { Add-WtMdPlain $Out $text.Substring($pos) }
}

function Add-WtMdPlain($Out, [string]$Text) {
    $parts = $Text -split "`n"
    for ($i = 0; $i -lt $parts.Count; $i++) {
        if ($i -gt 0) { $Out.Add(@{ T = 'soft' }) }
        if ($parts[$i]) { $Out.Add(@{ T = 'text'; Text = $parts[$i] }) }
    }
}

function Get-WtMdLinkKind([string]$Url) {
    if ($Url.StartsWith('#')) { return 'anchor' }
    if ($Url -match '^[A-Za-z]:[\\/]|^\\\\|^file:') { return 'file' }
    if ($Url -match '^[A-Za-z][A-Za-z0-9+.-]+:') { return 'url' }
    if ($Url -match '^www\.') { return 'url' }
    'file'
}

function New-WtMdLinkModel([string]$Url, $Inl, [string]$Title = '') {
    $kind = Get-WtMdLinkKind $Url
    $target = switch ($kind) {
        'anchor' { [uri]::UnescapeDataString($Url.Substring(1)) }
        'file' { [uri]::UnescapeDataString($Url) }
        default { if ($Url -match '^www\.') { "https://$Url" } else { $Url } }
    }
    if (-not @($Inl).Count) { $Inl = @(@{ T = 'text'; Text = $Url }) }
    @{ T = 'link'; Kind = $kind; Target = $target; Url = $Url; Title = $Title; Inl = @($Inl) }
}

# 그림. Obsidian처럼 대체 글 끝의 |300 또는 |300x200은 너비로 쓴다
function New-WtMdImageModel([string]$Url, [string]$Alt) {
    $width = 0
    if ($Alt -match '^(.*?)\|(\d+)(?:x\d+)?$') { $Alt = $Matches[1]; $width = [int]$Matches[2] }
    @{ T = 'image'; Url = $Url; Alt = $Alt; Width = $width }
}

# 인라인 목록의 글자만 (제목 이름표 · 대체 글 · 표 열 너비 계산용)
function Get-WtMdInlineText($Inl) {
    $sb = [System.Text.StringBuilder]::new()
    foreach ($x in @($Inl)) {
        if ($null -eq $x) { continue }
        switch ($x.T) {
            { $_ -in 'text', 'code', 'math' } { [void]$sb.Append($x.Text) }
            'tag' { [void]$sb.Append('#' + $x.Text) }
            { $_ -in 'soft', 'br' } { [void]$sb.Append(' ') }
            'image' { [void]$sb.Append($x.Alt) }
            default { if ($x.Inl) { [void]$sb.Append((Get-WtMdInlineText $x.Inl)) } }
        }
    }
    $sb.ToString()
}

function Get-WtMdBlocksText($Blocks) {
    (@($Blocks) | ForEach-Object { if ($_.Inl) { Get-WtMdInlineText $_.Inl } elseif ($_.Text) { $_.Text } }) -join ' '
}

# 제목 · 각주를 찾는 열쇠. 대소문자 · 띄어쓰기 · 기호를 무시한다 ([[#제목 이름]], (#제목-이름) 모두 같은 곳)
function Get-WtMdAnchorKey([string]$Text) { ($Text.ToLowerInvariant() -replace '[^\p{L}\p{N}]', '') }

<#
목록 한 줄 요약용: 마크다운 기호를 떼고 줄을 ' · '로 잇는다. 전체를 읽지 않고 줄 단위로 빠르게 처리한다.
#>
function Get-WtMarkdownSnippet([AllowEmptyString()][string]$Text, [int]$MaxLength = 300) {
    if (-not $Text) { return '' }
    $t = $Text.Replace("`r", '')
    $t = $t -replace '\A---\n[\s\S]*?\n---(\n|\z)', ''
    $t = $t -replace '%%[\s\S]*?%%', ''
    $lines = foreach ($line in ($t -split "`n")) {
        $l = $line.Trim()
        if (-not $l) { continue }
        if ($l -match '^(```|~~~)') { continue }
        if ($l -match '^\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?$' -and $l.Contains('-')) { continue }   # 표 구분 줄 · 가로줄
        if ($l -match '^([*_]\s*){3,}$') { continue }
        $l = $l -replace '^(>\s*)+', ''
        $l = $l -replace '^\[![\w-]+\][+-]?\s*', ''
        $l = $l -replace '^#{1,6}\s+', ''
        $l = $l -replace '^\[\^[^\]]+\]:\s*', ''
        $l = $l -replace '^([-*+]|\d+[.)])\s+(\[[ xX]\]\s+)?', ''
        $l = $l -replace '!\[\[[^\]]*\]\]', '' -replace '\[\[([^\]|]+)\|([^\]]+)\]\]', '$2' -replace '\[\[#?([^\]]+)\]\]', '$1'
        $l = $l -replace '!\[([^\]]*)\]\([^)]*\)', '' -replace '\[([^\]]*)\]\([^)]*\)', '$1' -replace '\[\^[^\]]+\]', ''
        $l = $l -replace '(\*\*|__|~~|==)(.+?)\1', '$2' -replace '(?<![\w*])\*(?!\s)(.+?)(?<!\s)\*(?![\w*])', '$1' -replace '`([^`]*)`', '$1'
        $l = [System.Net.WebUtility]::HtmlDecode(($l -replace '<[^>]+>', ' '))
        if ($l.StartsWith('|')) { $l = $l.Trim('|') -replace '\s*\|\s*', '  ' }
        $l = ($l -replace '\s{3,}', '  ').Trim()
        if ($l) { $l }
    }
    $s = @($lines) -join ' · '
    if ($s.Length -gt $MaxLength) { $s = $s.Substring(0, $MaxLength) }
    $s
}

# ---- 모델 → WPF ------------------------------------------------------------------------------

<#
마크다운 보기 요소(읽기 전용 RichTextBox)를 만든다. 글은 Set-WtMarkdownText로 넣는다.
  -Scroll            스스로 스크롤한다 (남는 높이를 채우는 자리). 없으면 글 높이만큼 커진다 (바깥이 스크롤하는 자리)
  -OnLink            { param($link) ... $true } 링크 · 위키 링크 · 태그를 눌렀을 때. 이게 있으면 #태그도 누를 수 있게 그린다
  -BaseDir           상대 경로의 기준 폴더
  -StrictLineBreaks  줄바꿈 한 번은 이어 붙인다 (표준 마크다운). 없으면 Obsidian 기본값처럼 줄을 바꾼다
#>
function New-WtMarkdownView {
    param([switch]$Scroll, [scriptblock]$OnLink, [string]$BaseDir = '', [switch]$StrictLineBreaks)
    $v = [System.Windows.Controls.RichTextBox]::new()
    Set-WtMdStyle $v 'Md.View'
    $v.IsReadOnly = $true
    $v.IsDocumentEnabled = $true
    $v.VerticalScrollBarVisibility = if ($Scroll) { 'Auto' } else { 'Disabled' }
    $v.Tag = @{
        Text = ''; BaseDir = $BaseDir; OnLink = $OnLink; Strict = [bool]$StrictLineBreaks
        Anchors = @{}; Remeasure = $false; Error = $null; ListDepth = 0
    }
    for ($i = $script:views.Count - 1; $i -ge 0; $i--) { if (-not $script:views[$i].IsAlive) { $script:views.RemoveAt($i) } }
    $script:views.Add([System.WeakReference]::new($v))
    $v
}

function Set-WtMarkdownText {
    param([Parameter(Mandatory)]$View, [AllowEmptyString()][string]$Text, [string]$BaseDir)
    $st = $View.Tag
    $st.Text = [string]$Text
    if ($PSBoundParameters.ContainsKey('BaseDir')) { $st.BaseDir = $BaseDir }
    $st.Anchors = @{}
    $st.Remeasure = $false
    $st.Error = $null
    $st.ListDepth = 0
    $doc = [System.Windows.Documents.FlowDocument]::new()
    Set-WtMdStyle $doc 'Md.Doc'
    try {
        $model = ConvertTo-WtMarkdownModel $st.Text
        foreach ($b in $model.Blocks) { Add-WtMdBlock $doc.Blocks $b $st }
        Set-WtMdEdgeMargins $doc.Blocks
    } catch {
        # 그리지 못하면 원문을 그대로 보여 준다. 오류 문구는 복사해서 알려 줄 수 있게 글로 둔다
        $st.Error = $_
        $doc.Blocks.Clear()
        $p = New-WtMdPara 'Md.P'
        $warn = New-WtMdRun "마크다운을 그리지 못해 원문을 보여 줍니다. $($_.Exception.Message)"
        Set-WtResource $warn 'Foreground' 'B.Danger'
        $p.Inlines.Add($warn)
        $doc.Blocks.Add($p)
        $raw = New-WtMdPara 'Md.CodeText'
        $raw.Inlines.Add((New-WtMdRun $st.Text))
        $doc.Blocks.Add($raw)
    }
    $View.Document = $doc
}

# 모델의 제목 목록 (목차 등). @{ Level; Text; Line }
function Get-WtMarkdownHeadings([AllowEmptyString()][string]$Text) { (ConvertTo-WtMarkdownModel $Text).Headings }

# 문서 요소에는 스타일을 바로 넣는다. 리소스 참조로 걸면 문서에 붙기 전에는 스타일이 풀리지 않아서
# 여백을 읽고 고치는 Set-WtMdEdgeMargins가 자동 여백(한 줄 높이)을 읽어 버린다.
function Set-WtMdStyle($Element, [string]$Key) {
    $style = (Get-WtStyleDictionary)[$Key]
    if (-not $style) { throw "Ui.psm1에 스타일 '$Key'가 없습니다." }
    $Element.Style = $style
}

function New-WtMdRun([string]$Text) { [System.Windows.Documents.Run]::new($Text) }

function New-WtMdPara([string]$Style) {
    $p = [System.Windows.Documents.Paragraph]::new()
    Set-WtMdStyle $p $Style
    $p
}

function New-WtMdIconRun([int]$Code) {
    $r = New-WtMdRun ([string][char]$Code)
    Set-WtResource $r 'FontFamily' 'Font.Icon'
    $r
}

# 상자(인용 · 콜아웃 · 목록 항목) 안 맨 위 · 맨 아래 블록의 바깥 여백을 없앤다
function Set-WtMdEdgeMargins($Blocks) {
    $first = $Blocks.FirstBlock
    $last = $Blocks.LastBlock
    if ($first) { $m = $first.Margin; $first.Margin = [System.Windows.Thickness]::new($m.Left, 0, $m.Right, $m.Bottom) }
    if ($last) { $m = $last.Margin; $last.Margin = [System.Windows.Thickness]::new($m.Left, $m.Top, $m.Right, 0) }
}

function Add-WtMdBlock($Blocks, $B, $St, [bool]$Tight = $false) {
    switch ($B.T) {
        'heading' {
            $p = New-WtMdPara "Md.H$([Math]::Min([int]$B.Level, 6))"
            Add-WtMdInlines $p.Inlines $B.Inl $St
            $key = Get-WtMdAnchorKey $B.Text
            if ($key -and -not $St.Anchors.ContainsKey($key)) { $St.Anchors[$key] = $p }
            $Blocks.Add($p)
        }
        'para' {
            $style = if ($Tight) { 'Md.PTight' } else { 'Md.P' }
            # 그림 하나만 있는 문단은 폭에 맞춘 블록 그림으로
            if (@($B.Inl).Count -eq 1 -and $B.Inl[0].T -eq 'image') {
                $img = New-WtMdImage $B.Inl[0] $St
                if ($img) {
                    $c = [System.Windows.Documents.BlockUIContainer]::new($img)
                    Set-WtMdStyle $c $(if ($Tight) { 'Md.FigureTight' } else { 'Md.Figure' })
                    $Blocks.Add($c)
                    return
                }
            }
            $p = New-WtMdPara $style
            Add-WtMdInlines $p.Inlines $B.Inl $St
            $Blocks.Add($p)
        }
        'html' {
            $p = New-WtMdPara $(if ($Tight) { 'Md.PTight' } else { 'Md.P' })
            Add-WtMdInlines $p.Inlines @(Get-WtMdPlainInlines $B.Text) $St
            $Blocks.Add($p)
        }
        'list' { $Blocks.Add((New-WtMdList $B $St $Tight)) }
        'quote' {
            $s = [System.Windows.Documents.Section]::new()
            Set-WtMdStyle $s 'Md.Quote'
            foreach ($c in $B.Blocks) { Add-WtMdBlock $s.Blocks $c $St }
            Set-WtMdEdgeMargins $s.Blocks
            $Blocks.Add($s)
        }
        'callout' { $Blocks.Add((New-WtMdCallout $B $St)) }
        'code' { $Blocks.Add((New-WtMdCode $B $St)) }
        'math' {
            $p = New-WtMdPara 'Md.Math'
            $p.Inlines.Add((New-WtMdRun $B.Text))
            $Blocks.Add($p)
        }
        'hr' {
            $line = [System.Windows.Controls.Border]::new()
            Set-WtMdStyle $line 'Divider'
            $c = [System.Windows.Documents.BlockUIContainer]::new($line)
            Set-WtMdStyle $c 'Md.Hr'
            $Blocks.Add($c)
        }
        'table' { $Blocks.Add((New-WtMdTable $B $St)) }
        'props' { $Blocks.Add((New-WtMdProps $B $St)) }
        'footnotes' {
            $s = [System.Windows.Documents.Section]::new()
            Set-WtMdStyle $s 'Md.Footnotes'
            foreach ($it in $B.Items) {
                $before = $s.Blocks.Count
                foreach ($c in $it.Blocks) { Add-WtMdBlock $s.Blocks $c $St $true }
                $p = $s.Blocks | Select-Object -Skip $before -First 1
                if ($p -is [System.Windows.Documents.Paragraph]) {
                    $n = New-WtMdRun "$($it.Order). "
                    if ($p.Inlines.FirstInline) { $p.Inlines.InsertBefore($p.Inlines.FirstInline, $n) } else { $p.Inlines.Add($n) }
                    $St.Anchors[(Get-WtMdAnchorKey "fn:$($it.Order)")] = $p
                }
            }
            $Blocks.Add($s)
        }
    }
}

function Get-WtMdPlainInlines([string]$Text) {
    $out = [System.Collections.Generic.List[object]]::new()
    Add-WtMdPlain $out $Text
    $out.ToArray()
}

function Add-WtMdInlines($Inlines, $Items, $St) {
    foreach ($x in @($Items)) {
        if ($null -eq $x) { continue }
        $el = New-WtMdInline $x $St
        foreach ($e in @($el)) { if ($e) { $Inlines.Add($e) } }
    }
}

function New-WtMdSpan($Items, $St) {
    $s = [System.Windows.Documents.Span]::new()
    Add-WtMdInlines $s.Inlines $Items $St
    $s
}

function New-WtMdInline($X, $St) {
    switch ($X.T) {
        'text' { return New-WtMdRun $X.Text }
        'soft' { if ($St.Strict) { return New-WtMdRun ' ' } else { return [System.Windows.Documents.LineBreak]::new() } }
        'br' { return [System.Windows.Documents.LineBreak]::new() }
        'em' { $s = New-WtMdSpan $X.Inl $St; $s.FontStyle = 'Italic'; return $s }
        'strong' { $s = New-WtMdSpan $X.Inl $St; $s.FontWeight = 'Bold'; return $s }
        'del' { $s = New-WtMdSpan $X.Inl $St; $s.TextDecorations = [System.Windows.TextDecorations]::Strikethrough; return $s }
        'ins' { $s = New-WtMdSpan $X.Inl $St; $s.TextDecorations = [System.Windows.TextDecorations]::Underline; return $s }
        'mark' { $s = New-WtMdSpan $X.Inl $St; Set-WtMdStyle $s 'Md.Mark'; return $s }
        'sup' { $s = New-WtMdSpan $X.Inl $St; Set-WtMdStyle $s 'Md.Sup'; return $s }
        'sub' { $s = New-WtMdSpan $X.Inl $St; Set-WtMdStyle $s 'Md.Sub'; return $s }
        'span' { return New-WtMdSpan $X.Inl $St }
        'code' { $r = New-WtMdRun $X.Text; Set-WtMdStyle $r 'Md.InlineCode'; return $r }
        'math' { $r = New-WtMdRun $X.Text; Set-WtMdStyle $r 'Md.InlineCode'; return $r }
        'task' { return New-WtMdTaskBox $X.Checked }
        'link' {
            if ($X.Kind -eq 'anchor' -and $X.Target -match '^\^') { return New-WtMdSpan $X.Inl $St }   # 블록 참조는 아직 없음
            $h = New-WtMdHyperlink @{ Kind = $X.Kind; Target = $X.Target; Text = (Get-WtMdInlineText $X.Inl) } $St
            Add-WtMdInlines $h.Inlines $X.Inl $St
            $tip = if ($X.Title) { "$($X.Title)`n$($X.Target)" } else { $X.Target }
            if ($X.Kind -eq 'wiki') { $tip = "문서: $($X.Target)" }
            $h.ToolTip = $tip
            return $h
        }
        'tag' {
            if ($St.OnLink) {
                $h = New-WtMdHyperlink @{ Kind = 'tag'; Target = $X.Text; Text = "#$($X.Text)" } $St
                Set-WtMdStyle $h 'Md.TagLink'
                $h.Inlines.Add((New-WtMdRun "#$($X.Text)"))
                $h.ToolTip = "#$($X.Text) 찾기"
                return $h
            }
            $s = [System.Windows.Documents.Span]::new()
            Set-WtMdStyle $s 'Md.Tag'
            $s.Inlines.Add((New-WtMdRun "#$($X.Text)"))
            return $s
        }
        'image' {
            $img = New-WtMdImage $X $St -Inline
            if ($img) { return [System.Windows.Documents.InlineUIContainer]::new($img) }
            # 불러올 수 없는 그림(웹 주소 · 없는 파일 · SVG)은 이름과 링크로 보여 준다
            $label = if ($X.Alt) { $X.Alt } else { $X.Url }
            $kind = Get-WtMdLinkKind $X.Url
            $h = New-WtMdHyperlink @{ Kind = $(if ($kind -eq 'anchor') { 'wiki' } else { $kind }); Target = $X.Url; Text = $label } $St
            $icon = New-WtMdIconRun 0xE91B
            $h.Inlines.Add($icon)
            $h.Inlines.Add((New-WtMdRun " $label"))
            $h.ToolTip = "그림: $($X.Url)"
            return $h
        }
        'fn' {
            $h = New-WtMdHyperlink @{ Kind = 'anchor'; Target = "fn:$($X.Order)"; Text = "$($X.Order)" } $St
            Set-WtMdStyle $h 'Md.FnRef'
            $h.Inlines.Add((New-WtMdRun "[$($X.Order)]"))
            $key = Get-WtMdAnchorKey "fnref:$($X.Order)"
            if (-not $St.Anchors.ContainsKey($key)) { $St.Anchors[$key] = $h }
            return $h
        }
        'fnback' {
            $h = New-WtMdHyperlink @{ Kind = 'anchor'; Target = "fnref:$($X.Order)"; Text = '↩' } $St
            $h.Inlines.Add((New-WtMdRun ' ↩'))
            $h.ToolTip = '본문으로'
            return $h
        }
    }
    $null
}

function New-WtMdTaskBox([bool]$Checked) {
    $r = New-WtMdIconRun $(if ($Checked) { 0xE73A } else { 0xE739 })
    Set-WtResource $r 'Foreground' $(if ($Checked) { 'B.Accent' } else { 'B.SubText' })
    $r
}

function New-WtMdHyperlink($Link, $St) {
    $h = [System.Windows.Documents.Hyperlink]::new()
    Set-WtMdStyle $h 'Md.Link'
    $h.Tag = @{ Link = $Link; State = $St }
    $h.Add_Click({
            param($s, $e)
            try { Invoke-WtMarkdownLink $s.Tag.State $s.Tag.Link } catch { Show-WtError $_ '링크 열기' }
        })
    $h
}

<#
링크를 눌렀을 때. 같은 문서 안(제목 · 각주)은 그 자리로 스크롤한다.
그 밖에는 -OnLink를 먼저 부르고, 처리하지 않았으면 기본 동작:
  url  http · https · mailto만 기본 브라우저 · 메일 프로그램으로 연다
  file 폴더는 탐색기로, 문서 · 그림은 기본 프로그램으로 연다. 실행 파일 · 스크립트는 열지 않고 탐색기에서 그 파일을 보여 준다
  wiki 연결할 곳을 모른다고 알린다 (메모 · 문서 화면이 -OnLink로 처리한다)
#>
function Invoke-WtMarkdownLink($St, $Link) {
    if ($Link.Kind -eq 'anchor') {
        $el = $St.Anchors[(Get-WtMdAnchorKey $Link.Target)]
        if ($el) { $el.BringIntoView() }
        return
    }
    if ($St.OnLink) {
        $handled = & $St.OnLink $Link | Select-Object -Last 1
        if ($handled -eq $true) { return }
    }
    switch ($Link.Kind) {
        'url' {
            if ($Link.Target -match '^(?i)(https?|mailto):') { Start-Process $Link.Target }
            else { Show-WtMessage "이 주소는 열지 않습니다.`n$($Link.Target)" }
        }
        'file' {
            $path = Resolve-WtMdPath $Link.Target $St.BaseDir
            if (-not $path -or -not (Test-Path -LiteralPath $path)) { Show-WtMessage "파일을 찾지 못했습니다.`n$($Link.Target)"; return }
            if (Test-Path -LiteralPath $path -PathType Container) { Invoke-Item -LiteralPath $path; return }
            if ($script:openable -contains [System.IO.Path]::GetExtension($path).ToLowerInvariant()) { Invoke-Item -LiteralPath $path; return }
            Start-Process explorer.exe -ArgumentList "/select,`"$path`""
        }
        'wiki' { Show-WtMessage "연결된 문서를 찾지 못했습니다: $($Link.Target)" }
    }
}

# 링크 · 그림 경로 → 전체 경로. 상대 경로인데 기준 폴더가 없으면 $null
function Resolve-WtMdPath([string]$Target, [string]$BaseDir) {
    $t = $Target
    if ($t -match '^file:') {
        try { $t = ([uri]$t).LocalPath } catch { return $null }
    }
    $t = $t -replace '#.*$', ''
    if (-not $t) { return $null }
    try {
        if ([System.IO.Path]::IsPathRooted($t)) { return [System.IO.Path]::GetFullPath($t) }
        if ($BaseDir) { return [System.IO.Path]::GetFullPath((Join-Path $BaseDir $t)) }
    } catch { }
    $null
}

# 로컬 그림 파일이면 Image 요소, 아니면 $null (웹 주소의 그림은 내려받지 않는다)
function New-WtMdImage($X, $St, [switch]$Inline) {
    if ((Get-WtMdLinkKind $X.Url) -ne 'file') { return $null }
    $path = Resolve-WtMdPath ([uri]::UnescapeDataString($X.Url)) $St.BaseDir
    if (-not $path -or -not (Test-Path -LiteralPath $path -PathType Leaf)) { return $null }
    if ($script:imageExt -notcontains [System.IO.Path]::GetExtension($path).ToLowerInvariant()) { return $null }
    try {
        $bmp = [System.Windows.Media.Imaging.BitmapImage]::new()
        $bmp.BeginInit()
        $bmp.CacheOption = 'OnLoad'   # 파일을 잡고 있지 않게 바로 읽는다
        $bmp.UriSource = [uri]$path
        $bmp.EndInit()
        $bmp.Freeze()
    } catch { return $null }
    $img = [System.Windows.Controls.Image]::new()
    $img.Source = $bmp
    $img.Stretch = 'Uniform'
    $img.StretchDirection = 'DownOnly'
    $img.HorizontalAlignment = 'Left'
    $img.MaxWidth = if ($X.Width) { $X.Width } else { $bmp.PixelWidth }
    if ($Inline -and -not $X.Width) { $img.MaxHeight = 240 }
    if ($X.Alt) { $img.ToolTip = $X.Alt }
    $img
}

function New-WtMdList($B, $St, [bool]$Nested) {
    $list = [System.Windows.Documents.List]::new()
    Set-WtMdStyle $list $(if ($Nested) { 'Md.ListNested' } else { 'Md.List' })
    $hasTask = @($B.Items | Where-Object { $null -ne $_.Task }).Count -gt 0
    if ($B.Ordered) {
        $list.MarkerStyle = 'Decimal'
        $list.StartIndex = [Math]::Max(1, $B.Start)
    } elseif ($hasTask) {
        $list.MarkerStyle = 'None'
    } else {
        $list.MarkerStyle = @('Disc', 'Circle', 'Square')[$St.ListDepth % 3]
    }
    $St.ListDepth++
    try {
        foreach ($it in $B.Items) {
            $li = [System.Windows.Documents.ListItem]::new()
            Set-WtMdStyle $li 'Md.ListItem'
            foreach ($c in $it.Blocks) { Add-WtMdBlock $li.Blocks $c $St (-not $B.Loose) }
            if (-not $li.Blocks.Count) { $li.Blocks.Add((New-WtMdPara 'Md.PTight')) }
            $p = $li.Blocks.FirstBlock
            if ($null -ne $it.Task -and $p -is [System.Windows.Documents.Paragraph]) {
                # 할 일: 앞에 체크 표시. 끝낸 일은 글만 흐리게 + 취소선 (Obsidian과 같음)
                $text = [System.Windows.Documents.Span]::new()
                foreach ($x in @($p.Inlines)) { [void]$p.Inlines.Remove($x); $text.Inlines.Add($x) }
                if ($it.Task) {
                    Set-WtResource $text 'Foreground' 'B.SubText'
                    $text.TextDecorations = [System.Windows.TextDecorations]::Strikethrough
                }
                $p.Inlines.Add((New-WtMdTaskBox $it.Task))
                $p.Inlines.Add((New-WtMdRun '  '))
                $p.Inlines.Add($text)
            } elseif ($hasTask -and -not $B.Ordered -and $p -is [System.Windows.Documents.Paragraph]) {
                $dot = New-WtMdRun '•   '
                if ($p.Inlines.FirstInline) { $p.Inlines.InsertBefore($p.Inlines.FirstInline, $dot) } else { $p.Inlines.Add($dot) }
            }
            Set-WtMdEdgeMargins $li.Blocks
            $list.ListItems.Add($li)
        }
    } finally { $St.ListDepth-- }
    $list
}

function New-WtMdCallout($B, $St) {
    $sec = [System.Windows.Documents.Section]::new()
    Set-WtMdStyle $sec 'Md.Callout'
    Set-WtResource $sec 'Background' "B.Callout.$($B.Color)Bg"
    Set-WtResource $sec 'BorderBrush' "B.Callout.$($B.Color)"

    $title = New-WtMdPara 'Md.CalloutTitle'
    Set-WtResource $title 'Foreground' "B.Callout.$($B.Color)"
    $body = [System.Windows.Documents.Section]::new()
    foreach ($c in $B.Blocks) { Add-WtMdBlock $body.Blocks $c $St }
    Set-WtMdEdgeMargins $body.Blocks
    $hasBody = $body.Blocks.Count -gt 0

    if ($B.Fold -and $hasBody) {
        # 접는 콜아웃: 제목을 누르면 본문을 넣었다 뺐다 한다
        $open = $B.Fold -eq '+'
        $link = [System.Windows.Documents.Hyperlink]::new()
        Set-WtMdStyle $link 'Md.FoldLink'
        Set-WtResource $link 'Foreground' "B.Callout.$($B.Color)"
        $chev = New-WtMdIconRun $(if ($open) { 0xE70D } else { 0xE76C })
        $link.Inlines.Add((New-WtMdIconRun $B.Icon))
        $link.Inlines.Add((New-WtMdRun "  $(Get-WtMdInlineText $B.Title)  "))
        $link.Inlines.Add($chev)
        $link.ToolTip = '접기 / 펼치기'
        $link.Tag = @{ Section = $sec; Title = $title; Body = $body; Chevron = $chev; Open = $open }
        $link.Add_Click({ param($s, $e) Switch-WtMdFold $s.Tag })
        $title.Inlines.Add($link)
        $sec.Blocks.Add($title)
        if ($open) { $sec.Blocks.Add($body) } else { $title.Margin = [System.Windows.Thickness]::new(0) }
    } else {
        $title.Inlines.Add((New-WtMdIconRun $B.Icon))
        $title.Inlines.Add((New-WtMdRun '  '))
        Add-WtMdInlines $title.Inlines $B.Title $St
        $sec.Blocks.Add($title)
        if ($hasBody) { $sec.Blocks.Add($body) } else { $title.Margin = [System.Windows.Thickness]::new(0) }
    }
    $sec
}

function Switch-WtMdFold($F) {
    $F.Open = -not $F.Open
    $F.Chevron.Text = [string][char]$(if ($F.Open) { 0xE70D } else { 0xE76C })
    if ($F.Open) {
        $F.Section.Blocks.Add($F.Body)
        $F.Title.ClearValue([System.Windows.Documents.Block]::MarginProperty)
    } else {
        [void]$F.Section.Blocks.Remove($F.Body)
        $F.Title.Margin = [System.Windows.Thickness]::new(0)
    }
}

# 코드 블록 머리: 언어 이름(왼쪽) + 복사 버튼(오른쪽)
function New-WtMdCodeHeader([string]$Label, [string]$Code) {
    $d = [System.Windows.Controls.DockPanel]::new()
    $copy = New-WtButton '복사' -Kind Link -Icon 0xE8C8 -ToolTip '코드 복사'
    $copy.Tag = $Code
    $copy.Add_Click({ param($s, $e) Copy-WtMdCode $s })
    [System.Windows.Controls.DockPanel]::SetDock($copy, 'Right')
    [void]$d.Children.Add($copy)
    $name = New-WtText $Label 'T.Sub'
    $name.VerticalAlignment = 'Center'
    [void]$d.Children.Add($name)
    $d
}

function Copy-WtMdCode($Button) {
    $text = [string]$Button.Tag
    $ok = $false
    for ($i = 0; $i -lt 3 -and -not $ok; $i++) {
        try { [System.Windows.Clipboard]::SetText($text); $ok = $true } catch { Start-Sleep -Milliseconds 50 }
    }
    if (-not $ok) { Show-WtMessage '클립보드를 다른 프로그램이 쓰고 있어서 복사하지 못했습니다. 다시 누르세요.' 'WorkTool' 'Warning'; return }
    $Button.Content = New-WtButtonContent '복사됨' 0xE73E
    $t = [System.Windows.Threading.DispatcherTimer]::new()
    $t.Interval = [TimeSpan]::FromMilliseconds(1500)
    $t.Tag = $Button
    $t.Add_Tick({
            param($s, $e)
            $s.Stop()
            $s.Tag.Content = New-WtButtonContent '복사' 0xE8C8
        })
    $t.Start()
}

function New-WtMdCode($B, $St) {
    $lang = (([string]$B.Lang) -split '\s+')[0].ToLowerInvariant()
    $note = $null
    $r = $script:codeRenderers[$lang]
    if ($r) {
        try {
            $el = @(& $r.Render $B.Text | Where-Object { $_ -is [System.Windows.UIElement] }) | Select-Object -Last 1
            if (-not $el) { throw '그림을 만들지 못했습니다.' }
            if ($r.Remeasure) { $St.Remeasure = $true }
            $box = [System.Windows.Controls.Border]::new()
            Set-WtMdStyle $box 'Md.FigureBox'
            $dock = [System.Windows.Controls.DockPanel]::new()
            $head = New-WtMdCodeHeader $r.Label $B.Text
            [System.Windows.Controls.DockPanel]::SetDock($head, 'Top')
            [void]$dock.Children.Add($head)
            [void]$dock.Children.Add($el)
            $box.Child = $dock
            $c = [System.Windows.Documents.BlockUIContainer]::new($box)
            Set-WtMdStyle $c 'Md.Figure'
            return $c
        } catch {
            $note = "그리지 못해 원문을 보여 줍니다. $($_.Exception.Message)"
        }
    }

    $sec = [System.Windows.Documents.Section]::new()
    Set-WtMdStyle $sec 'Md.CodeBlock'
    $label = if ($lang) { $lang } else { '코드' }
    $headC = [System.Windows.Documents.BlockUIContainer]::new((New-WtMdCodeHeader $label $B.Text))
    Set-WtMdStyle $headC 'Md.CodeHead'
    $sec.Blocks.Add($headC)
    if ($note) {
        $np = New-WtMdPara 'Md.CodeNote'
        $np.Inlines.Add((New-WtMdIconRun 0xE7BA))
        $np.Inlines.Add((New-WtMdRun "  $note"))
        $sec.Blocks.Add($np)
    }
    $p = New-WtMdPara 'Md.CodeText'
    foreach ($tok in (Get-WtCodeTokens $B.Text $lang)) {
        $parts = $tok.Text -split "`n"
        for ($i = 0; $i -lt $parts.Count; $i++) {
            if ($i -gt 0) { $p.Inlines.Add([System.Windows.Documents.LineBreak]::new()) }
            if (-not $parts[$i]) { continue }
            $run = New-WtMdRun $parts[$i]
            if ($tok.Kind) { Set-WtResource $run 'Foreground' "B.Code.$($tok.Kind)" }
            $p.Inlines.Add($run)
        }
    }
    $sec.Blocks.Add($p)
    $sec
}

function New-WtMdTable($B, $St) {
    $t = [System.Windows.Documents.Table]::new()
    Set-WtMdStyle $t 'Md.Table'
    $n = @($B.Align).Count
    foreach ($row in $B.Rows) { $n = [Math]::Max($n, @($row.Cells).Count) }
    # 열 너비: FlowDocument 표에는 자동 너비가 없어서 글자 수에 맞춰 비율을 준다
    for ($i = 0; $i -lt $n; $i++) {
        $w = 3
        foreach ($row in $B.Rows) {
            if ($i -lt @($row.Cells).Count) { $w = [Math]::Max($w, (Get-WtMdBlocksText $row.Cells[$i]).Length) }
        }
        $col = [System.Windows.Documents.TableColumn]::new()
        $col.Width = [System.Windows.GridLength]::new([Math]::Min($w, 40), 'Star')
        $t.Columns.Add($col)
    }
    $g = [System.Windows.Documents.TableRowGroup]::new()
    foreach ($row in $B.Rows) {
        $tr = [System.Windows.Documents.TableRow]::new()
        for ($i = 0; $i -lt $n; $i++) {
            $cell = [System.Windows.Documents.TableCell]::new()
            Set-WtMdStyle $cell $(if ($row.Header) { 'Md.Th' } else { 'Md.Td' })
            $a = if ($i -lt @($B.Align).Count) { $B.Align[$i] } else { '' }
            if ($a) { $cell.TextAlignment = switch ($a) { 'center' { 'Center' } 'right' { 'Right' } default { 'Left' } } }
            if ($i -lt @($row.Cells).Count) { foreach ($c in $row.Cells[$i]) { Add-WtMdBlock $cell.Blocks $c $St $true } }
            if (-not $cell.Blocks.Count) { $cell.Blocks.Add((New-WtMdPara 'Md.PTight')) }
            $tr.Cells.Add($cell)
        }
        $g.Rows.Add($tr)
    }
    $t.RowGroups.Add($g)
    $t
}

# 맨 앞 속성(--- 사이): 이름 · 값 두 열 표
function New-WtMdProps($B, $St) {
    $t = [System.Windows.Documents.Table]::new()
    Set-WtMdStyle $t 'Md.Props'
    foreach ($w in 1, 3) {
        $col = [System.Windows.Documents.TableColumn]::new()
        $col.Width = [System.Windows.GridLength]::new($w, 'Star')
        $t.Columns.Add($col)
    }
    $g = [System.Windows.Documents.TableRowGroup]::new()
    foreach ($it in $B.Items) {
        $tr = [System.Windows.Documents.TableRow]::new()
        foreach ($pair in @(@('Md.PropKey', $it.Key), @('Md.PropValue', $it.Value))) {
            $cell = [System.Windows.Documents.TableCell]::new()
            Set-WtMdStyle $cell $pair[0]
            $p = New-WtMdPara 'Md.PTight'
            $p.Inlines.Add((New-WtMdRun $pair[1]))
            $cell.Blocks.Add($p)
            $tr.Cells.Add($cell)
        }
        $g.Rows.Add($tr)
    }
    $t.RowGroups.Add($g)
    $t
}

# 글자 크기를 바꾸면, 크기를 재서 배치한 그림(다이어그램)이 있는 보기를 다시 그린다.
# 크기는 설정에서 바로 읽으므로(Get-WtFontSize) 셸의 테마 갱신 순서와 상관없다.
Register-WtSettingHandler {
    param($key, $value)
    if ($key -ne 'ui.fontScale') { return }
    foreach ($w in @($script:views)) {
        $v = $w.Target
        if ($v -and $v.Tag.Remeasure) { Set-WtMarkdownText $v $v.Tag.Text }
    }
}

Export-ModuleMember -Function Register-WtMarkdownCodeRenderer, Register-WtMarkdownCallout, ConvertTo-WtMarkdownModel,
    Get-WtMarkdownSnippet, Get-WtMarkdownHeadings, New-WtMarkdownView, Set-WtMarkdownText, Invoke-WtMarkdownLink
} | Import-Module -Force -DisableNameChecking
# ---- core\Markdown.psm1 끝 ----

# ---- core\Mermaid.psm1 ----
New-Module -Name 'WorkTool.Mermaid' -ScriptBlock {
# Mermaid 다이어그램. 마크다운의 ```mermaid 코드 블록을 그린다 (Markdown.psm1에 코드 블록 그리기로 등록).
# 이 창에는 브라우저(WebBrowser)를 넣을 수 없어서, 자주 쓰는 문법을 직접 읽고 배치해서 WPF 도형으로 그린다.
#
# 지원
# - flowchart · graph: 방향 TB TD BT LR RL, 노드 모양 14종, 선(실선 · 점선 · 굵은 선 · 안 보이는 선), 끝(화살표 · 원 · X),
#   라벨(-->|글| 와 -- 글 -->), A & B 묶음, subgraph(중첩 포함)
# - sequenceDiagram: participant · actor(as 별칭), 메시지 화살표 10종, Note(left of · right of · over),
#   loop · alt/else · opt · par/and · critical/option · break · rect, autonumber
# - stateDiagram · stateDiagram-v2: 상태, [*] 시작 · 끝, 전이 라벨, state "설명" as X, X : 설명, 복합 상태 { }, choice · fork · join
# 그 밖의 종류(gantt · classDiagram · pie …)는 NotSupportedException을 던진다 → 마크다운 쪽이 원문을 코드 블록으로 보여 준다.
# 색 · 모양 꾸미기(style · classDef · class · linkStyle)와 click은 읽고 무시한다. 색은 Theme.psm1의 Dia.* 만 쓴다.
#
# 1) ConvertFrom-WtMermaid: 글 → 모델                              (WPF 없이 동작)
# 2) Get-WtMermaidLayout: 모델 → 도형 목록(좌표). 글자 크기 재는 함수를 받는다  (WPF 없이 동작)
# 3) New-WtMermaidView: 글 → 그림 요소(Viewbox + Canvas)             (WPF)

$ErrorActionPreference = 'Stop'

# 배치 간격 (epx)
$script:gap = @{
    Margin   = 8    # 그림 가장자리
    PadX     = 16   # 노드 안 글자 좌우 여백
    PadY     = 10   # 노드 안 글자 위아래 여백
    NodeSep  = 32   # 같은 층 노드 사이
    DummySep = 16   # 같은 층에서 선이 지나가는 점 · 라벨 사이
    RankGap  = 24   # 층 사이. 선 하나는 두 층을 건너므로 노드 사이는 48 + 라벨 높이
    GroupPad = 12   # subgraph 상자 안 여백
    Arrow    = 9    # 화살표 길이
}

# ---- 글 → 모델 -------------------------------------------------------------------------------

function New-WtMmdMap { [System.Collections.Specialized.OrderedDictionary]::new([System.StringComparer]::Ordinal) }

function ConvertFrom-WtMermaid([string]$Text) {
    $lines = @(([string]$Text).Replace("`r", '') -split "`n")
    $title = ''
    $i = 0
    while ($i -lt $lines.Count -and -not $lines[$i].Trim()) { $i++ }
    # 맨 앞 --- title: … --- 머리말
    if ($i -lt $lines.Count -and $lines[$i].Trim() -eq '---') {
        $i++
        while ($i -lt $lines.Count -and $lines[$i].Trim() -ne '---') {
            if ($lines[$i] -match '^\s*title:\s*(.+)$') { $title = $Matches[1].Trim().Trim('"', "'") }
            $i++
        }
        $i++
    }
    $stmts = Split-WtMmdStatements @($lines | Select-Object -Skip $i)
    if (-not $stmts.Count) { throw '다이어그램 내용이 없습니다.' }
    $kind = ($stmts[0] -split '\s+')[0].ToLowerInvariant()
    $m = switch -Regex ($kind) {
        '^(graph|flowchart)(-elk)?$' { ConvertFrom-WtMmdFlowchart $stmts }
        '^sequencediagram$' { ConvertFrom-WtMmdSequence $stmts }
        '^statediagram(-v2)?$' { ConvertFrom-WtMmdState $stmts }
        default { throw [System.NotSupportedException]::new("'$kind' 종류는 아직 지원하지 않습니다 (flowchart · sequenceDiagram · stateDiagram만 그립니다).") }
    }
    if ($title -and -not $m.Title) { $m.Title = $title }
    $m
}

# 줄과 ;로 문장을 나눈다. %% 주석 줄은 뺀다. 따옴표 · 괄호 안의 ;는 나누지 않는다.
function Split-WtMmdStatements([string[]]$Lines) {
    $out = [System.Collections.Generic.List[string]]::new()
    foreach ($raw in $Lines) {
        $line = ([string]$raw).Trim()
        if (-not $line -or $line.StartsWith('%%')) { continue }
        $buf = [System.Text.StringBuilder]::new()
        $quote = $false
        $depth = 0
        foreach ($ch in $line.ToCharArray()) {
            if ($ch -eq '"') { $quote = -not $quote }
            elseif (-not $quote) {
                if ('[({'.IndexOf($ch) -ge 0) { $depth++ }
                elseif (')]}'.IndexOf($ch) -ge 0) { if ($depth -gt 0) { $depth-- } }
                elseif ($ch -eq ';' -and $depth -eq 0) {
                    $s = $buf.ToString().Trim()
                    if ($s) { $out.Add($s) }
                    [void]$buf.Clear()
                    continue
                }
            }
            [void]$buf.Append($ch)
        }
        $s = $buf.ToString().Trim()
        if ($s) { $out.Add($s) }
    }
    , $out.ToArray()
}

# 라벨 글: 감싼 따옴표 · `마크다운 글` 기호를 떼고, <br>은 줄바꿈, #quot; 같은 Mermaid 글자 표기를 바꾼다
function ConvertTo-WtMmdText([string]$Text) {
    $t = ([string]$Text).Trim()
    if ($t.Length -ge 2 -and $t.StartsWith('"') -and $t.EndsWith('"')) { $t = $t.Substring(1, $t.Length - 2) }
    if ($t.Length -ge 2 -and $t.StartsWith('`') -and $t.EndsWith('`')) { $t = $t.Substring(1, $t.Length - 2) -replace '\*\*|__', '' }
    $t = $t -replace '<br\s*/?>', "`n"
    $t = $t -replace '#quot;', '"' -replace '#amp;', '&' -replace '#lt;', '<' -replace '#gt;', '>' -replace '#35;', '#' -replace '#59;', ';'
    $t = $t -replace '<[^>]+>', ''
    (@($t -split "`n") | ForEach-Object { $_.Trim() }) -join "`n"
}

function New-WtMmdFlowModel([string]$Dir) {
    @{ Type = 'flowchart'; Dir = $Dir; Nodes = (New-WtMmdMap); Edges = [System.Collections.Generic.List[object]]::new(); Groups = (New-WtMmdMap); Title = '' }
}

# 노드를 더하거나 고친다. 글을 주면 글 · 모양을 바꾸고, subgraph 안에서 나오면 그 subgraph로 옮긴다 (Mermaid와 같음)
function Add-WtMmdNode($M, [string]$Id, $Text, [string]$Shape, $Group, [string]$DefaultShape = 'rect') {
    if (-not $M.Nodes.Contains($Id)) { $M.Nodes[$Id] = @{ Id = $Id; Text = $Id; Shape = $DefaultShape; Group = $null } }
    $n = $M.Nodes[$Id]
    if ($null -ne $Text) { $n.Text = ConvertTo-WtMmdText $Text }
    if ($Shape) { $n.Shape = $Shape }
    if ($Group) { $n.Group = $Group }
    $Id
}

function ConvertFrom-WtMmdFlowchart([string[]]$Stmts) {
    $dir = 'TB'
    if ($Stmts[0] -match '^\S+\s+(TB|TD|BT|LR|RL)\b') { $dir = $Matches[1].ToUpperInvariant() }
    if ($dir -eq 'TD') { $dir = 'TB' }
    $m = New-WtMmdFlowModel $dir
    $stack = [System.Collections.Generic.List[string]]::new()
    for ($i = 1; $i -lt $Stmts.Count; $i++) {
        $s = $Stmts[$i]
        $group = if ($stack.Count) { $stack[$stack.Count - 1] } else { $null }
        if ($s -match '^subgraph\b\s*(.*)$') {
            $g = Read-WtMmdSubgraphHead $Matches[1] $m.Groups.Count
            $g.Parent = $group
            $m.Groups[$g.Id] = $g
            $stack.Add($g.Id)
            continue
        }
        if ($s -match '^end\s*$') { if ($stack.Count) { $stack.RemoveAt($stack.Count - 1) }; continue }
        if ($s -match '^(direction|classDef|class|style|linkStyle|click|accTitle|accDescr)\b') { continue }
        Read-WtMmdChain $m $s $group
    }
    $m
}

function Read-WtMmdSubgraphHead([string]$Rest, [int]$N) {
    $r = $Rest.Trim()
    if ($r -match '^([^\s\[\]"]+)\s*\[(.*)\]$') { return @{ Id = $Matches[1]; Title = (ConvertTo-WtMmdText $Matches[2]) } }
    if ($r -match '^"(.*)"$') { return @{ Id = "subgraph$N"; Title = $Matches[1] } }
    if (-not $r) { return @{ Id = "subgraph$N"; Title = '' } }
    @{ Id = $r; Title = $r }
}

# 노드 모양: 여는 기호, 닫는 기호, 이름. 여는 기호가 긴 것부터 (같은 여는 기호끼리는 붙여 둔다)
$script:shapes = @(
    @('(((', ')))', 'dcircle'), @('([', '])', 'stadium'), @('[[', ']]', 'subroutine'), @('[(', ')]', 'cylinder'),
    @('((', '))', 'circle'), @('{{', '}}', 'hexagon'), @('[/', '/]', 'para'), @('[/', '\]', 'trap'),
    @('[\', '\]', 'paraAlt'), @('[\', '/]', 'trapAlt'), @('>', ']', 'asym'), @('[', ']', 'rect'),
    @('(', ')', 'round'), @('{', '}', 'diamond')
)
$script:idRe = [regex]::new('\G\s*([\p{L}\p{N}_]+(?:[-.][\p{L}\p{N}_]+)*)')
$script:classRe = [regex]::new('\G:::[\w-]+')
$endPat = '(?<h>>|[ox](?![\p{L}\p{N}_]))?'
$script:edgeRes = @(
    [regex]::new('\G\s*(?<s>[<ox])?==\s+(?<t>[^=|>][^=|]*?)\s*(?<e>={2,})' + $endPat)   # == 글 ==>
    [regex]::new('\G\s*(?<s>[<ox])?-\.\s+(?<t>[^|.>-][^|]*?)\s*(?<e>\.-+)' + $endPat)   # -. 글 .->
    [regex]::new('\G\s*(?<s>[<ox])?--\s+(?<t>[^|>-][^|]*?)\s+(?<e>-{2,})' + $endPat)    # -- 글 -->
    [regex]::new('\G\s*(?<s>[<ox])?(?<e>-{2,}|={2,}|-\.+-|~{3,})' + $endPat)            # --> --- -.-> ==> ~~~
)
$script:pipeRe = [regex]::new('\G\s*\|(?<p>[^|]*)\|')

# 한 문장: 노드묶음 (선 노드묶음)*
function Read-WtMmdChain($M, [string]$S, $Group) {
    $pos = [ref]0
    $left = Read-WtMmdNodeGroup $M $S $pos $Group
    if (-not $left) { return }   # 읽지 못하는 문장은 건너뛴다
    while ($pos.Value -lt $S.Length) {
        $edge = Read-WtMmdEdge $S $pos
        if (-not $edge) { return }
        $right = Read-WtMmdNodeGroup $M $S $pos $Group
        if (-not $right) { return }
        foreach ($a in $left) {
            foreach ($b in $right) {
                $M.Edges.Add(@{ From = $a; To = $b; Text = $edge.Text; Style = $edge.Style; Start = $edge.Start; End = $edge.End; Len = $edge.Len })
            }
        }
        $left = $right
    }
}

function Read-WtMmdNodeGroup($M, [string]$S, [ref]$Pos, $Group) {
    $ids = [System.Collections.Generic.List[string]]::new()
    while ($true) {
        $id = Read-WtMmdNode $M $S $Pos $Group
        if (-not $id) { break }
        $ids.Add($id)
        $mm = [regex]::Match($S.Substring($Pos.Value), '^\s*&\s*')
        if (-not $mm.Success) { break }
        $Pos.Value += $mm.Length
    }
    , $ids.ToArray()
}

function Find-WtMmdClose([string]$S, [int]$Start, [string]$Close) {
    $i = $Start
    if ($i -lt $S.Length -and $S[$i] -eq '"') {
        $q = $S.IndexOf('"', $i + 1)
        if ($q -lt 0) { return -1 }
        $i = $q + 1
    }
    $S.IndexOf($Close, $i, [System.StringComparison]::Ordinal)
}

function Read-WtMmdNode($M, [string]$S, [ref]$Pos, $Group) {
    $mm = $script:idRe.Match($S, $Pos.Value)
    if (-not $mm.Success) { return $null }
    $id = $mm.Groups[1].Value
    if ($id -eq 'end') { return $null }
    $p = $mm.Index + $mm.Length
    $rest = $S.Substring($p)
    $best = $null
    foreach ($sh in $script:shapes) {
        if ($best -and $best.Open -ne $sh[0]) { break }
        if (-not $rest.StartsWith($sh[0], [System.StringComparison]::Ordinal)) { continue }
        $start = $p + $sh[0].Length
        $close = Find-WtMmdClose $S $start $sh[1]
        if ($close -lt 0) { continue }
        if (-not $best -or $close -lt $best.At) { $best = @{ Open = $sh[0]; At = $close; Len = $sh[1].Length; Shape = $sh[2]; Start = $start } }
    }
    $text = $null
    $shape = $null
    if ($best) {
        $text = $S.Substring($best.Start, $best.At - $best.Start)
        $shape = $best.Shape
        $p = $best.At + $best.Len
    }
    $cm = $script:classRe.Match($S, $p)
    if ($cm.Success) { $p += $cm.Length }
    $Pos.Value = $p
    Add-WtMmdNode $M $id $text $shape $Group
}

function Get-WtMmdEnd([string]$C) {
    switch -CaseSensitive ($C) { '>' { 'arrow' } '<' { 'arrow' } 'o' { 'circle' } 'x' { 'cross' } default { 'none' } }
}

function Read-WtMmdEdge([string]$S, [ref]$Pos) {
    for ($k = 0; $k -lt $script:edgeRes.Count; $k++) {
        $mm = $script:edgeRes[$k].Match($S, $Pos.Value)
        if (-not $mm.Success) { continue }
        $Pos.Value = $mm.Index + $mm.Length
        $body = $mm.Groups['e'].Value
        $text = if ($mm.Groups['t'].Success) { $mm.Groups['t'].Value } else { '' }
        $pm = $script:pipeRe.Match($S, $Pos.Value)
        if ($pm.Success) { $text = $pm.Groups['p'].Value; $Pos.Value = $pm.Index + $pm.Length }
        $head = $mm.Groups['h'].Value
        $style = if ($k -eq 0 -or $body.Contains('=')) { 'thick' }
        elseif ($k -eq 1 -or $body.Contains('.')) { 'dotted' }
        elseif ($body.StartsWith('~')) { 'invisible' }
        else { 'normal' }
        # 선 길이(-- 개수)는 건너뛰는 층 수
        $len = if ($style -eq 'dotted') { @($body.ToCharArray() | Where-Object { $_ -eq '.' }).Count }
        else { $body.Length - $(if ($head) { 1 } else { 2 }) }
        return @{
            Text = (ConvertTo-WtMmdText $text); Style = $style; Len = [Math]::Max(1, $len)
            Start = (Get-WtMmdEnd $mm.Groups['s'].Value); End = (Get-WtMmdEnd $head)
        }
    }
    $null
}

# 상태 다이어그램 → flowchart 모델 (상태 = 둥근 사각형, [*] = 시작 · 끝 점, 복합 상태 = subgraph)
function ConvertFrom-WtMmdState([string[]]$Stmts) {
    $m = New-WtMmdFlowModel 'TB'
    $m.Kind = 'state'
    $stack = [System.Collections.Generic.List[string]]::new()
    $inNote = $false
    $idp = '[\p{L}\p{N}_.-]+'
    for ($i = 1; $i -lt $Stmts.Count; $i++) {
        $s = $Stmts[$i]
        if ($inNote) { if ($s -match '^end\s+note$') { $inNote = $false }; continue }
        $group = if ($stack.Count) { $stack[$stack.Count - 1] } else { $null }
        $scope = if ($group) { $group } else { 'root' }
        if ($s -match '^direction\s+(TB|TD|BT|LR|RL)$') {
            if (-not $stack.Count) { $m.Dir = $Matches[1].ToUpperInvariant() -replace 'TD', 'TB' }
            continue
        }
        if ($s -match '^note\s') { if (-not $s.Contains(':')) { $inNote = $true }; continue }
        if ($s -eq '}') { if ($stack.Count) { $stack.RemoveAt($stack.Count - 1) }; continue }
        if ($s -eq '--' -or $s -match '^(classDef|class|style|accTitle|accDescr|scale|hide)\b') { continue }
        if ($s -match "^state\s+(?:`"([^`"]*)`"\s+as\s+)?($idp)\s*\{$") {
            $gid = $Matches[2]
            $title = if ($Matches[1]) { $Matches[1] } elseif ($m.Nodes.Contains($gid)) { $m.Nodes[$gid].Text } else { $gid }
            $m.Groups[$gid] = @{ Id = $gid; Title = $title; Parent = $group }
            $stack.Add($gid)
            continue
        }
        if ($s -match "^state\s+($idp)\s*<<(choice|fork|join)>>$") {
            [void](Add-WtMmdNode $m $Matches[1] '' $(if ($Matches[2] -eq 'choice') { 'choice' } else { 'fork' }) $group 'round')
            continue
        }
        if ($s -match "^state\s+`"([^`"]*)`"\s+as\s+($idp)$") { [void](Add-WtMmdNode $m $Matches[2] $Matches[1] $null $group 'round'); continue }
        if ($s -match "^state\s+($idp)$") { [void](Add-WtMmdNode $m $Matches[1] $null $null $group 'round'); continue }
        if ($s -match "^($idp|\[\*\])\s*-->\s*($idp|\[\*\])\s*(?::\s*(.*))?$") {
            $a = $Matches[1]; $b = $Matches[2]; $label = [string]$Matches[3]
            $a = if ($a -eq '[*]') { Add-WtMmdNode $m "[*]start.$scope" '' 'start' $group } else { Add-WtMmdNode $m $a $null $null $group 'round' }
            $b = if ($b -eq '[*]') { Add-WtMmdNode $m "[*]end.$scope" '' 'end' $group } else { Add-WtMmdNode $m $b $null $null $group 'round' }
            $m.Edges.Add(@{ From = $a; To = $b; Text = (ConvertTo-WtMmdText $label); Style = 'normal'; Start = 'none'; End = 'arrow'; Len = 1 })
            continue
        }
        if ($s -match "^($idp)\s*:\s*(.+)$") {
            $id = Add-WtMmdNode $m $Matches[1] $null $null $group 'round'
            $n = $m.Nodes[$id]
            $n.Text = "$($n.Text)`n$(ConvertTo-WtMmdText $Matches[2])"
            continue
        }
        if ($s -match "^($idp)$") { [void](Add-WtMmdNode $m $Matches[1] $null $null $group 'round') }
    }
    # 복합 상태로 가는 선은 그 안의 시작점(없으면 첫 상태)으로, 나오는 선은 끝점(없으면 마지막 상태)에서
    foreach ($gid in @($m.Groups.Keys)) {
        $members = @($m.Nodes.Values | Where-Object { $_.Group -ceq $gid } | ForEach-Object { $_.Id })
        if (-not $members.Count) { continue }
        if ($m.Nodes.Contains($gid)) { $m.Nodes.Remove($gid) }
        $in = if ($m.Nodes.Contains("[*]start.$gid")) { "[*]start.$gid" } else { $members[0] }
        $out = if ($m.Nodes.Contains("[*]end.$gid")) { "[*]end.$gid" } else { $members[$members.Count - 1] }
        foreach ($e in $m.Edges) {
            if ($e.To -ceq $gid) { $e.To = $in }
            if ($e.From -ceq $gid) { $e.From = $out }
        }
    }
    $m
}

$script:msgRe = [regex]::new('^(?<from>[^\s>:+-](?:[^>:]*?[^\s>:-])?)\s*(?<arrow><<-->>|<<->>|-->>|->>|--x|-x|--\)|-\)|-->|->)\s*(?<act>[+-])?\s*(?<to>[^\s:+-][^:]*?)\s*(?::\s*(?<text>.*))?$')

function Add-WtMmdParticipant($M, [string]$Id, [string]$Label, [bool]$Actor) {
    if (-not $M.Participants.Contains($Id)) { $M.Participants[$Id] = @{ Id = $Id; Label = $Id; Actor = $false } }
    $p = $M.Participants[$Id]
    if ($Label) { $p.Label = ConvertTo-WtMmdText $Label }
    if ($Actor) { $p.Actor = $true }
}

function ConvertFrom-WtMmdSequence([string[]]$Stmts) {
    $m = @{ Type = 'sequence'; Participants = (New-WtMmdMap); Items = [System.Collections.Generic.List[object]]::new(); AutoNumber = $false; Title = '' }
    $open = [System.Collections.Generic.List[string]]::new()   # 열린 블록 (box는 그리지 않는다)
    for ($i = 1; $i -lt $Stmts.Count; $i++) {
        $s = $Stmts[$i] -replace '^create\s+', ''
        if ($s -match '^(participant|actor)\s+(.+?)(?:\s+as\s+(.+))?$') {
            Add-WtMmdParticipant $m $Matches[2].Trim() ([string]$Matches[3]) ($Matches[1] -eq 'actor')
            continue
        }
        if ($s -match '^autonumber\b') { $m.AutoNumber = $true; continue }
        if ($s -match '^title\s*:?\s*(.*)$') { $m.Title = $Matches[1]; continue }
        if ($s -match '^(activate|deactivate|destroy|links?|properties|details)\b') { continue }
        if ($s -match '^box\b') { $open.Add('box'); continue }
        if ($s -match '^(loop|alt|opt|par_over|par|critical|break|rect)\b\s*(.*)$') {
            $kind = $Matches[1]
            $open.Add($kind)
            $m.Items.Add(@{ K = 'start'; Kind = $kind; Text = $(if ($kind -eq 'rect') { '' } else { ConvertTo-WtMmdText $Matches[2] }) })
            continue
        }
        if ($s -match '^(else|and|option)\b\s*(.*)$') { $m.Items.Add(@{ K = 'else'; Text = (ConvertTo-WtMmdText $Matches[2]) }); continue }
        if ($s -match '^end$') {
            if ($open.Count) {
                $k = $open[$open.Count - 1]
                $open.RemoveAt($open.Count - 1)
                if ($k -ne 'box') { $m.Items.Add(@{ K = 'end' }) }
            }
            continue
        }
        if ($s -match '^note\s+(left of|right of|over)\s+([^:]+?)\s*:\s*(.*)$') {
            $pos = ($Matches[1] -split ' ')[0]
            $text = ConvertTo-WtMmdText $Matches[3]
            $ids = @($Matches[2] -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            foreach ($id in $ids) { Add-WtMmdParticipant $m $id '' $false }
            $m.Items.Add(@{ K = 'note'; Pos = $pos; Ids = $ids; Text = $text })
            continue
        }
        $mm = $script:msgRe.Match($s)
        if ($mm.Success) {
            $from = $mm.Groups['from'].Value.Trim()
            $to = $mm.Groups['to'].Value.Trim()
            Add-WtMmdParticipant $m $from '' $false
            Add-WtMmdParticipant $m $to '' $false
            $a = $mm.Groups['arrow'].Value
            $head = if ($a.EndsWith('>>')) { 'arrow' } elseif ($a.EndsWith('x')) { 'cross' } elseif ($a.EndsWith(')')) { 'async' } else { 'none' }
            $m.Items.Add(@{
                    K = 'msg'; From = $from; To = $to; Text = (ConvertTo-WtMmdText $mm.Groups['text'].Value)
                    Line = $(if ($a.Contains('--')) { 'dashed' } else { 'normal' }); Head = $head; Both = $a.StartsWith('<<')
                })
        }
    }
    if (-not $m.Participants.Count) { throw '참여자(participant)나 메시지가 없습니다.' }
    $m
}

# ---- 모델 → 도형 목록 (WPF 없이) ---------------------------------------------------------------

<#
도형 목록 (좌표는 epx, 왼쪽 위가 0,0)
  node   Shape X Y W H Text TH [Fill Line]   노드 · 참여자 · 노트 (글은 가운데)
  rect   X Y W H Fill Line Dash Radius       subgraph · 블록 틀
  text   X Y W Text Size Align Bold Color Bg  라벨 · 제목
  line   Pts(x,y,x,y…) Style Start End Smooth 선 (끝: arrow · circle · cross · async · none)
  actor  X Y W H                             사람 모양
#>
function Get-WtMermaidLayout($Model, [scriptblock]$Measure) {
    if (-not $Measure) { $Measure = { param($t, $s) Measure-WtMmdTextEstimate $t $s } }
    switch ($Model.Type) {
        'flowchart' { return Get-WtMmdFlowLayout $Model $Measure }
        'sequence' { return Get-WtMmdSeqLayout $Model $Measure }
    }
    throw "알 수 없는 다이어그램: $($Model.Type)"
}

# 글자 크기 어림 (WPF가 없을 때: 테스트용). 한글 1em, 영문 약 0.58em
function Measure-WtMmdTextEstimate([string]$Text, [string]$Size) {
    $em = Get-WtFontSize $Size
    $w = 0.0
    $lines = @(([string]$Text) -split "`n")
    foreach ($l in $lines) {
        $lw = 0.0
        foreach ($ch in $l.ToCharArray()) {
            $lw += if ([int]$ch -ge 0x1100) { $em } elseif ('iljtf.,:;|!''I '.IndexOf($ch) -ge 0) { $em * 0.32 } elseif ('mwMW@'.IndexOf($ch) -ge 0) { $em * 0.85 } else { $em * 0.58 }
        }
        $w = [Math]::Max($w, $lw)
    }
    @{ W = [Math]::Ceiling($w); H = [Math]::Ceiling($lines.Count * $em * 1.35) }
}

function Get-WtMmdNodeSize($N, $Measure) {
    $g = $script:gap
    switch ($N.Shape) {
        'start' { return @{ W = 16; H = 16; TH = 0 } }
        'end' { return @{ W = 20; H = 20; TH = 0 } }
        'fork' { return @{ W = 72; H = 8; TH = 0 } }
        'choice' { return @{ W = 28; H = 28; TH = 0 } }
    }
    $t = & $Measure $N.Text 'Body'
    $tw = $t.W
    $th = $t.H
    $w = $tw + 2 * $g.PadX
    $h = $th + 2 * $g.PadY
    switch ($N.Shape) {
        'diamond' { $w = $tw * 1.6 + 24; $h = $th * 2.4 + 12 }
        'circle' { $w = [Math]::Sqrt($tw * $tw + $th * $th) + 16; $h = $w }
        'dcircle' { $w = [Math]::Sqrt($tw * $tw + $th * $th) + 28; $h = $w }
        'hexagon' { $w += $h * 0.5 }
        { $_ -in 'para', 'paraAlt', 'trap', 'trapAlt' } { $w += $h * 0.8 }
        'asym' { $w += $h * 0.4 }
        'stadium' { $w += $h * 0.3 }
        'cylinder' { $h += 16 }
        'subroutine' { $w += 16 }
    }
    @{ W = [Math]::Max([Math]::Ceiling($w), 40); H = [Math]::Ceiling($h); TH = $th }
}

# 노드 가장자리와 (중심 → 목표점) 선이 만나는 점
function Get-WtMmdClip($Vx, [double]$Cx, [double]$Cy, [double]$Tx, [double]$Ty) {
    if ($Vx.Dummy) { return @($Cx, $Cy) }
    $dx = $Tx - $Cx
    $dy = $Ty - $Cy
    if ([Math]::Abs($dx) -lt 0.01 -and [Math]::Abs($dy) -lt 0.01) { return @($Cx, $Cy) }
    $hw = $Vx.W / 2
    $hh = $Vx.H / 2
    $t = switch ($Vx.Shape) {
        { $_ -in 'circle', 'dcircle', 'start', 'end' } { $hw / [Math]::Sqrt($dx * $dx + $dy * $dy) }
        { $_ -in 'diamond', 'choice' } { 1 / ([Math]::Abs($dx) / $hw + [Math]::Abs($dy) / $hh) }
        default {
            $fx = if ([Math]::Abs($dx) -gt 0.001) { $hw / [Math]::Abs($dx) } else { [double]::MaxValue }
            $fy = if ([Math]::Abs($dy) -gt 0.001) { $hh / [Math]::Abs($dy) } else { [double]::MaxValue }
            [Math]::Min($fx, $fy)
        }
    }
    if ($t -gt 1) { $t = 1 }
    @(($Cx + $dx * $t), ($Cy + $dy * $t))
}

<#
flowchart 배치 (층 나누기 방식)
1. 순환을 끊는다: 깊이 우선으로 돌며 거꾸로 가는 선은 배치할 때만 뒤집는다
2. 층: 가장 긴 경로. 선 하나는 두 층 이상 건너고, 가운데 층에 라벨 자리를 둔다
3. 두 층 넘게 건너는 선은 층마다 지나가는 점을 둔다
4. 층 안 순서: 이웃의 평균 위치로 위아래 번갈아 정렬. 같은 subgraph는 붙여 둔다
5. 좌표: 이웃 평균 쪽으로 당기면서 간격을 지킨다
#>
function Get-WtMmdFlowLayout($M, $Measure) {
    $g = $script:gap
    $horiz = $M.Dir -in 'LR', 'RL'
    $verts = [System.Collections.Generic.List[hashtable]]::new()
    $index = [System.Collections.Generic.Dictionary[string, int]]::new([System.StringComparer]::Ordinal)
    foreach ($id in $M.Nodes.Keys) {
        $n = $M.Nodes[$id]
        $sz = Get-WtMmdNodeSize $n $Measure
        if ($horiz -and $n.Shape -eq 'fork') { $sz = @{ W = $sz.H; H = $sz.W; TH = 0 } }
        $index[$id] = $verts.Count
        $verts.Add(@{ Id = $id; Text = $n.Text; Shape = $n.Shape; W = $sz.W; H = $sz.H; TH = $sz.TH; Dummy = $false; Group = $n.Group; Rank = 0; Extra = 0.0; Self = $null })
    }
    $real = $verts.Count
    if (-not $real) { throw '노드가 없습니다.' }

    $edges = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($e in $M.Edges) {
        if (-not $index.ContainsKey($e.From) -or -not $index.ContainsKey($e.To)) { continue }
        $a = $index[$e.From]
        $b = $index[$e.To]
        if ($a -eq $b) {
            # 자기 자신으로 가는 선: 노드 옆에 고리로 그린다
            $lw = if ($e.Text) { (& $Measure $e.Text 'Small').W + 8 } else { 0 }
            $verts[$a].Self = $e
            $verts[$a].Extra = [Math]::Max($verts[$a].Extra, 30 + $lw)
            continue
        }
        $edges.Add(@{ A = $a; B = $b; E = $e; Rev = $false })
    }

    # 1. 순환 끊기
    $outE = @(for ($i = 0; $i -lt $real; $i++) { , [System.Collections.Generic.List[int]]::new() })
    for ($k = 0; $k -lt $edges.Count; $k++) { $outE[$edges[$k].A].Add($k) }
    $state = [int[]]::new($real)
    for ($r = 0; $r -lt $real; $r++) {
        if ($state[$r]) { continue }
        $st = [System.Collections.Generic.Stack[int[]]]::new()
        $st.Push([int[]]@($r, 0))
        $state[$r] = 1
        while ($st.Count) {
            $top = $st.Peek()
            $v = $top[0]
            if ($top[1] -lt $outE[$v].Count) {
                $ei = $outE[$v][$top[1]]
                $top[1]++
                $w = $edges[$ei].B
                if ($state[$w] -eq 1) { $edges[$ei].Rev = $true }
                elseif ($state[$w] -eq 0) { $state[$w] = 1; $st.Push([int[]]@($w, 0)) }
            } else {
                $state[$v] = 2
                [void]$st.Pop()
            }
        }
    }
    foreach ($e in $edges) { if ($e.Rev) { $t = $e.A; $e.A = $e.B; $e.B = $t } }

    # 2. 층 (가장 긴 경로, 선 길이 x2)
    $outE = @(for ($i = 0; $i -lt $real; $i++) { , [System.Collections.Generic.List[int]]::new() })
    $inCount = [int[]]::new($real)
    for ($k = 0; $k -lt $edges.Count; $k++) { $outE[$edges[$k].A].Add($k); $inCount[$edges[$k].B]++ }
    $rank = [int[]]::new($real)
    $deg = [int[]]$inCount.Clone()
    $queue = [System.Collections.Generic.Queue[int]]::new()
    for ($i = 0; $i -lt $real; $i++) { if (-not $deg[$i]) { $queue.Enqueue($i) } }
    $topo = [System.Collections.Generic.List[int]]::new()
    while ($queue.Count) {
        $v = $queue.Dequeue()
        $topo.Add($v)
        foreach ($k in $outE[$v]) {
            $w = $edges[$k].B
            $rank[$w] = [Math]::Max($rank[$w], $rank[$v] + 2 * $edges[$k].E.Len)
            $deg[$w]--
            if (-not $deg[$w]) { $queue.Enqueue($w) }
        }
    }
    # 들어오는 선이 없는 노드는 나가는 선 쪽으로 붙인다 (맨 위에 혼자 떨어져 있지 않게)
    for ($j = $topo.Count - 1; $j -ge 0; $j--) {
        $v = $topo[$j]
        if ($inCount[$v] -or -not $outE[$v].Count) { continue }
        $best = [int]::MaxValue
        foreach ($k in $outE[$v]) { $best = [Math]::Min($best, $rank[$edges[$k].B] - 2 * $edges[$k].E.Len) }
        if ($best -gt $rank[$v]) { $rank[$v] = $best }
    }
    for ($i = 0; $i -lt $real; $i++) { $verts[$i].Rank = $rank[$i] }

    # 3. 지나가는 점 · 라벨 점
    $chains = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($e in $edges) {
        $ra = $verts[$e.A].Rank
        $rb = $verts[$e.B].Rank
        $mid = $ra + [int][Math]::Floor(($rb - $ra) / 2)
        $grp = if ($verts[$e.A].Group -and $verts[$e.A].Group -ceq $verts[$e.B].Group) { $verts[$e.A].Group } else { $null }
        $chain = [System.Collections.Generic.List[int]]::new()
        $chain.Add($e.A)
        $labelAt = -1
        for ($r = $ra + 1; $r -lt $rb; $r++) {
            $d = @{ Id = ''; Dummy = $true; W = 0.0; H = 0.0; Group = $grp; Rank = $r; Extra = 0.0; Shape = '' }
            if ($r -eq $mid -and $e.E.Text) {
                $t = & $Measure $e.E.Text 'Small'
                $d.W = $t.W + 8
                $d.H = $t.H + 4
                $d.Label = $e.E.Text
                $labelAt = $chain.Count
            }
            $chain.Add($verts.Count)
            $verts.Add($d)
        }
        $chain.Add($e.B)
        $chains.Add(@{ Chain = $chain.ToArray(); E = $e.E; Rev = $e.Rev; LabelAt = $labelAt })
    }

    # 4. 층 안 순서
    $rankCount = 0
    foreach ($x in $verts) { $rankCount = [Math]::Max($rankCount, $x.Rank + 1) }
    $layers = @(for ($r = 0; $r -lt $rankCount; $r++) { , [System.Collections.Generic.List[int]]::new() })
    for ($i = 0; $i -lt $verts.Count; $i++) { $layers[$verts[$i].Rank].Add($i) }
    $up = @(for ($i = 0; $i -lt $verts.Count; $i++) { , [System.Collections.Generic.List[int]]::new() })
    $down = @(for ($i = 0; $i -lt $verts.Count; $i++) { , [System.Collections.Generic.List[int]]::new() })
    foreach ($c in $chains) {
        for ($k = 0; $k -lt $c.Chain.Count - 1; $k++) { $down[$c.Chain[$k]].Add($c.Chain[$k + 1]); $up[$c.Chain[$k + 1]].Add($c.Chain[$k]) }
    }
    $order = [double[]]::new($verts.Count)
    foreach ($l in $layers) { for ($i = 0; $i -lt $l.Count; $i++) { $order[$l[$i]] = $i } }
    for ($iter = 0; $iter -lt 6; $iter++) {
        for ($r = 1; $r -lt $rankCount; $r++) { Sort-WtMmdLayer $layers[$r] $up $order $verts $M }
        for ($r = $rankCount - 2; $r -ge 0; $r--) { Sort-WtMmdLayer $layers[$r] $down $order $verts $M }
    }

    # 5. 좌표 (cs: 층 안 방향 크기, ms: 층 방향 크기)
    $cs = [double[]]::new($verts.Count)
    $ms = [double[]]::new($verts.Count)
    for ($i = 0; $i -lt $verts.Count; $i++) {
        $cs[$i] = if ($horiz) { $verts[$i].H } else { $verts[$i].W }
        $ms[$i] = if ($horiz) { $verts[$i].W } else { $verts[$i].H }
    }
    $titleH = (& $Measure 'Ag' 'Small').H + 6
    $gaps = @(foreach ($l in $layers) {
            $arr = [double[]]::new([Math]::Max(0, $l.Count - 1))
            for ($i = 0; $i -lt $l.Count - 1; $i++) {
                $a = $verts[$l[$i]]
                $b = $verts[$l[$i + 1]]
                $sep = if ($a.Dummy -or $b.Dummy) { $g.DummySep } else { $g.NodeSep }
                if (($a.Group -or $b.Group) -and -not ($a.Group -ceq $b.Group)) { $sep += 2 * $g.GroupPad + $(if ($horiz) { $titleH } else { 0 }) }
                $extra = if ($horiz) { 0 } else { $a.Extra }
                $arr[$i] = ($cs[$l[$i]] + $cs[$l[$i + 1]]) / 2 + $sep + $extra
            }
            , $arr
        })
    $x = [double[]]::new($verts.Count)
    for ($r = 0; $r -lt $rankCount; $r++) {
        $l = $layers[$r]
        if (-not $l.Count) { continue }
        $x[$l[0]] = 0
        for ($i = 1; $i -lt $l.Count; $i++) { $x[$l[$i]] = $x[$l[$i - 1]] + $gaps[$r][$i - 1] }
        $shift = ($x[$l[0]] + $x[$l[$l.Count - 1]]) / 2
        foreach ($v in $l) { $x[$v] -= $shift }
    }
    for ($iter = 0; $iter -lt 8; $iter++) {
        for ($r = 1; $r -lt $rankCount; $r++) { Set-WtMmdLayerX $layers[$r] $gaps[$r] $up $x }
        for ($r = $rankCount - 2; $r -ge 0; $r--) { Set-WtMmdLayerX $layers[$r] $gaps[$r] $down $x }
    }

    # 층 위치
    $band = [double[]]::new($rankCount)
    for ($i = 0; $i -lt $verts.Count; $i++) { $band[$verts[$i].Rank] = [Math]::Max($band[$verts[$i].Rank], $ms[$i]) }
    $mainPos = [double[]]::new($rankCount)
    $cursor = 0.0
    for ($r = 0; $r -lt $rankCount; $r++) { $mainPos[$r] = $cursor + $band[$r] / 2; $cursor += $band[$r] + $g.RankGap }
    $total = $cursor - $g.RankGap
    $cx = [double[]]::new($verts.Count)
    $cy = [double[]]::new($verts.Count)
    for ($i = 0; $i -lt $verts.Count; $i++) {
        $mp = $mainPos[$verts[$i].Rank]
        if ($M.Dir -in 'BT', 'RL') { $mp = $total - $mp }
        if ($horiz) { $cx[$i] = $mp; $cy[$i] = $x[$i] } else { $cx[$i] = $x[$i]; $cy[$i] = $mp }
    }

    # subgraph 상자 (안쪽 것부터)
    $depth = @{}
    foreach ($gid in $M.Groups.Keys) {
        $d = 0; $p = $M.Groups[$gid].Parent
        while ($p -and $d -lt 50) { $d++; $p = $M.Groups[$p].Parent }
        $depth[$gid] = $d
    }
    $boxes = New-WtMmdMap
    foreach ($gid in @($M.Groups.Keys | Sort-Object { - $depth[$_] })) {
        $l = [double]::MaxValue; $t = [double]::MaxValue; $rr = [double]::MinValue; $b = [double]::MinValue
        for ($i = 0; $i -lt $real; $i++) {
            if (-not ($verts[$i].Group -ceq $gid)) { continue }
            $l = [Math]::Min($l, $cx[$i] - $verts[$i].W / 2); $rr = [Math]::Max($rr, $cx[$i] + $verts[$i].W / 2 + $(if ($horiz) { 0 } else { $verts[$i].Extra }))
            $t = [Math]::Min($t, $cy[$i] - $verts[$i].H / 2); $b = [Math]::Max($b, $cy[$i] + $verts[$i].H / 2 + $(if ($horiz) { $verts[$i].Extra } else { 0 }))
        }
        foreach ($cid in $boxes.Keys) {
            if (-not ($M.Groups[$cid].Parent -ceq $gid)) { continue }
            $cb = $boxes[$cid]
            $l = [Math]::Min($l, $cb.X); $t = [Math]::Min($t, $cb.Y); $rr = [Math]::Max($rr, $cb.X + $cb.W); $b = [Math]::Max($b, $cb.Y + $cb.H)
        }
        if ($l -eq [double]::MaxValue) { continue }
        $title = [string]$M.Groups[$gid].Title
        $tsz = if ($title) { & $Measure $title 'Small' } else { @{ W = 0; H = 0 } }
        $l -= $g.GroupPad; $rr += $g.GroupPad; $b += $g.GroupPad
        $t -= $g.GroupPad + $(if ($title) { $tsz.H + 4 } else { 0 })
        if ($rr - $l -lt $tsz.W + 2 * $g.GroupPad) { $rr = $l + $tsz.W + 2 * $g.GroupPad }
        $boxes[$gid] = @{ X = $l; Y = $t; W = $rr - $l; H = $b - $t; Title = $title }
    }

    # 선 점
    $lines = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($c in $chains) {
        $ids = $c.Chain
        $pts = [System.Collections.Generic.List[double]]::new()
        foreach ($v in $ids) { $pts.Add($cx[$v]); $pts.Add($cy[$v]) }
        $n = $ids.Count
        $p0 = Get-WtMmdClip $verts[$ids[0]] $pts[0] $pts[1] $pts[2] $pts[3]
        $pn = Get-WtMmdClip $verts[$ids[$n - 1]] $pts[2 * $n - 2] $pts[2 * $n - 1] $pts[2 * $n - 4] $pts[2 * $n - 3]
        $pts[0] = $p0[0]; $pts[1] = $p0[1]; $pts[2 * $n - 2] = $pn[0]; $pts[2 * $n - 1] = $pn[1]
        $arr = $pts.ToArray()
        if ($c.Rev) {
            $rev = [double[]]::new($arr.Length)
            for ($k = 0; $k -lt $n; $k++) { $rev[2 * $k] = $arr[2 * ($n - 1 - $k)]; $rev[2 * $k + 1] = $arr[2 * ($n - 1 - $k) + 1] }
            $arr = $rev
        }
        $lbl = $null
        if ($c.LabelAt -ge 0) {
            $d = $verts[$ids[$c.LabelAt]]
            $lbl = @{ X = $cx[$ids[$c.LabelAt]] - $d.W / 2; Y = $cy[$ids[$c.LabelAt]] - $d.H / 2; W = $d.W; H = $d.H; Text = $d.Label }
        }
        $lines.Add(@{ Pts = $arr; E = $c.E; Label = $lbl })
    }
    # 자기 자신으로 가는 고리
    for ($i = 0; $i -lt $real; $i++) {
        $se = $verts[$i].Self
        if (-not $se) { continue }
        $hw = $verts[$i].W / 2; $hh = $verts[$i].H / 2
        if ($horiz) {
            $pts = [double[]]@(($cx[$i] - 8), ($cy[$i] + $hh), ($cx[$i] - 12), ($cy[$i] + $hh + 24), ($cx[$i] + 12), ($cy[$i] + $hh + 24), ($cx[$i] + 8), ($cy[$i] + $hh))
            $lpos = @(($cx[$i] + 16), ($cy[$i] + $hh + 14))
        } else {
            $pts = [double[]]@(($cx[$i] + $hw), ($cy[$i] - 8), ($cx[$i] + $hw + 24), ($cy[$i] - 12), ($cx[$i] + $hw + 24), ($cy[$i] + 12), ($cx[$i] + $hw), ($cy[$i] + 8))
            $lpos = @(($cx[$i] + $hw + 30), $cy[$i])
        }
        $lbl = $null
        if ($se.Text) {
            $t = & $Measure $se.Text 'Small'
            $lbl = @{ X = $lpos[0]; Y = $lpos[1] - ($t.H + 4) / 2; W = $t.W + 8; H = $t.H + 4; Text = $se.Text }
        }
        $lines.Add(@{ Pts = $pts; E = $se; Label = $lbl })
    }

    # 왼쪽 위를 여백에 맞춘다
    $minX = [double]::MaxValue; $minY = [double]::MaxValue; $maxX = [double]::MinValue; $maxY = [double]::MinValue
    for ($i = 0; $i -lt $verts.Count; $i++) {
        $minX = [Math]::Min($minX, $cx[$i] - $verts[$i].W / 2); $maxX = [Math]::Max($maxX, $cx[$i] + $verts[$i].W / 2)
        $minY = [Math]::Min($minY, $cy[$i] - $verts[$i].H / 2); $maxY = [Math]::Max($maxY, $cy[$i] + $verts[$i].H / 2)
    }
    foreach ($bx in $boxes.Values) {
        $minX = [Math]::Min($minX, $bx.X); $maxX = [Math]::Max($maxX, $bx.X + $bx.W)
        $minY = [Math]::Min($minY, $bx.Y); $maxY = [Math]::Max($maxY, $bx.Y + $bx.H)
    }
    foreach ($ln in $lines) {
        for ($k = 0; $k -lt $ln.Pts.Length; $k += 2) {
            $minX = [Math]::Min($minX, $ln.Pts[$k]); $maxX = [Math]::Max($maxX, $ln.Pts[$k])
            $minY = [Math]::Min($minY, $ln.Pts[$k + 1]); $maxY = [Math]::Max($maxY, $ln.Pts[$k + 1])
        }
        if ($ln.Label) { $maxX = [Math]::Max($maxX, $ln.Label.X + $ln.Label.W); $maxY = [Math]::Max($maxY, $ln.Label.Y + $ln.Label.H) }
    }
    $items = [System.Collections.Generic.List[hashtable]]::new()
    $top = $g.Margin
    if ($M.Title) {
        $ts = & $Measure $M.Title 'Body'
        $items.Add(@{ K = 'text'; X = $g.Margin; Y = $g.Margin; W = [Math]::Max($ts.W, $maxX - $minX); Text = $M.Title; Size = 'Body'; Align = 'Center'; Bold = $true; Color = 'Text' })
        $top += $ts.H + 12
    }
    $ox = $g.Margin - $minX
    $oy = $top - $minY

    foreach ($gid in @($boxes.Keys | Sort-Object { $depth[$_] })) {
        $bx = $boxes[$gid]
        $items.Add(@{ K = 'rect'; X = $bx.X + $ox; Y = $bx.Y + $oy; W = $bx.W; H = $bx.H; Fill = 'Dia.GroupBg'; Line = 'Dia.GroupLine'; Radius = 6 })
        if ($bx.Title) {
            $items.Add(@{ K = 'text'; X = $bx.X + $ox + $g.GroupPad; Y = $bx.Y + $oy + 4; W = $bx.W - 2 * $g.GroupPad; Text = $bx.Title; Size = 'Small'; Align = 'Left'; Color = 'SubText' })
        }
    }
    foreach ($ln in $lines) {
        if ($ln.E.Style -eq 'invisible') { continue }
        $pts = [double[]]::new($ln.Pts.Length)
        for ($k = 0; $k -lt $pts.Length; $k += 2) { $pts[$k] = $ln.Pts[$k] + $ox; $pts[$k + 1] = $ln.Pts[$k + 1] + $oy }
        $items.Add(@{ K = 'line'; Pts = $pts; Style = $ln.E.Style; Start = $ln.E.Start; End = $ln.E.End; Smooth = $true })
    }
    for ($i = 0; $i -lt $real; $i++) {
        $v = $verts[$i]
        $items.Add(@{ K = 'node'; Shape = $v.Shape; X = $cx[$i] - $v.W / 2 + $ox; Y = $cy[$i] - $v.H / 2 + $oy; W = $v.W; H = $v.H; Text = $v.Text; TH = $v.TH })
    }
    foreach ($ln in $lines) {
        if (-not $ln.Label) { continue }
        $lb = $ln.Label
        $items.Add(@{ K = 'text'; X = $lb.X + $ox; Y = $lb.Y + $oy; W = $lb.W; H = $lb.H; Text = $lb.Text; Size = 'Small'; Align = 'Center'; Color = 'Text'; Bg = 'Card' })
    }
    @{ Width = [Math]::Ceiling($maxX - $minX + 2 * $g.Margin); Height = [Math]::Ceiling($maxY - $minY + $top + $g.Margin); Items = $items.ToArray() }
}

# 층 하나를 이웃(위 또는 아래 층)의 평균 순번으로 정렬한다. 같은 subgraph는 그 평균으로 묶는다
function Sort-WtMmdLayer($Layer, $Nbr, [double[]]$Order, $verts, $M) {
    if ($Layer.Count -lt 2) { return }
    $key = @{}
    foreach ($v in $Layer) {
        $ns = $Nbr[$v]
        if ($ns.Count) {
            $sum = 0.0
            foreach ($u in $ns) { $sum += $Order[$u] }
            $key[$v] = $sum / $ns.Count
        } else { $key[$v] = $Order[$v] }
    }
    # 가장 바깥 subgraph와 자기 subgraph의 평균
    $outer = @{}
    $groupSum = @{}
    $groupCnt = @{}
    foreach ($v in $Layer) {
        $gname = $verts[$v].Group
        $top = $gname
        $guard = 0
        while ($top -and $M.Groups -and $M.Groups.Contains($top) -and $M.Groups[$top].Parent -and $guard -lt 50) { $top = $M.Groups[$top].Parent; $guard++ }
        $outer[$v] = $top
        foreach ($k in @($gname, "^$top")) {
            if (-not $k -or $k -eq '^') { continue }
            $groupSum[$k] = [double]$groupSum[$k] + $key[$v]
            $groupCnt[$k] = [int]$groupCnt[$k] + 1
        }
    }
    $rows = foreach ($v in $Layer) {
        $k1 = if ($outer[$v]) { $groupSum["^$($outer[$v])"] / $groupCnt["^$($outer[$v])"] } else { $key[$v] }
        $k2 = if ($verts[$v].Group) { $groupSum[$verts[$v].Group] / $groupCnt[$verts[$v].Group] } else { $key[$v] }
        [pscustomobject]@{ V = $v; K1 = $k1; K2 = $k2; K3 = $key[$v]; K4 = $Order[$v] }
    }
    $sorted = @($rows | Sort-Object K1, K2, K3, K4)
    $Layer.Clear()
    for ($i = 0; $i -lt $sorted.Count; $i++) { $Layer.Add($sorted[$i].V); $Order[$sorted[$i].V] = $i }
}

# 층 하나의 좌표: 이웃 평균 위치로 당기되 순서와 간격은 지킨다 (왼쪽에서 · 오른쪽에서 한 번씩 놓고 평균)
function Set-WtMmdLayerX($Layer, [double[]]$Gaps, $Nbr, [double[]]$X) {
    $n = $Layer.Count
    if (-not $n) { return }
    $d = [double[]]::new($n)
    for ($i = 0; $i -lt $n; $i++) {
        $v = $Layer[$i]
        $ns = $Nbr[$v]
        if ($ns.Count) {
            $sum = 0.0
            foreach ($u in $ns) { $sum += $X[$u] }
            $d[$i] = $sum / $ns.Count
        } else { $d[$i] = $X[$v] }
    }
    $a = [double[]]::new($n)
    $b = [double[]]::new($n)
    $a[0] = $d[0]
    for ($i = 1; $i -lt $n; $i++) { $a[$i] = [Math]::Max($d[$i], $a[$i - 1] + $Gaps[$i - 1]) }
    $b[$n - 1] = $d[$n - 1]
    for ($i = $n - 2; $i -ge 0; $i--) { $b[$i] = [Math]::Min($d[$i], $b[$i + 1] - $Gaps[$i]) }
    $prev = 0.0
    for ($i = 0; $i -lt $n; $i++) {
        $val = ($a[$i] + $b[$i]) / 2
        if ($i -gt 0) { $val = [Math]::Max($val, $prev + $Gaps[$i - 1]) }
        $X[$Layer[$i]] = $val
        $prev = $val
    }
}

function Get-WtMmdSeqLayout($M, $Measure) {
    $g = $script:gap
    $items = [System.Collections.Generic.List[hashtable]]::new()
    $parts = @($M.Participants.Values)
    $n = $parts.Count
    $idx = [System.Collections.Generic.Dictionary[string, int]]::new([System.StringComparer]::Ordinal)
    for ($i = 0; $i -lt $n; $i++) { $idx[$parts[$i].Id] = $i }
    $boxW = [double[]]::new($n)
    $boxTH = [double[]]::new($n)
    for ($i = 0; $i -lt $n; $i++) {
        $t = & $Measure $parts[$i].Label 'Body'
        $boxW[$i] = [Math]::Max($t.W + 2 * $g.PadX, 72)
        $boxTH[$i] = $t.H
    }
    $maxTH = ($boxTH | Measure-Object -Maximum).Maximum
    $boxH = $maxTH + 2 * $g.PadY
    $hasActor = @($parts | Where-Object { $_.Actor }).Count -gt 0
    $headH = if ($hasActor) { [Math]::Max($boxH, 36 + $maxTH) } else { $boxH }
    $gapX = [double[]]::new([Math]::Max(0, $n - 1))
    for ($i = 0; $i -lt $n - 1; $i++) { $gapX[$i] = ($boxW[$i] + $boxW[$i + 1]) / 2 + 28 }
    $leftExtra = 0.0
    $rightExtra = 0.0

    # 글 크기를 먼저 재서 가로 간격을 넓힌다
    $info = [System.Collections.Generic.List[hashtable]]::new()
    $num = 0
    foreach ($it in $M.Items) {
        $x = @{ It = $it }
        if ($it.K -eq 'msg') {
            $text = $it.Text
            if ($M.AutoNumber) { $num++; $text = if ($text) { "$num. $text" } else { "$num." } }
            $x.Text = $text
            $t = if ($text) { & $Measure $text 'Small' } else { @{ W = 0; H = 0 } }
            $x.TW = $t.W; $x.TH = $t.H
            $a = $idx[$it.From]; $b = $idx[$it.To]
            if ($a -eq $b) {
                $need = 40 + $t.W + 12
                if ($a -lt $n - 1) { $gapX[$a] = [Math]::Max($gapX[$a], $need) } else { $rightExtra = [Math]::Max($rightExtra, $need - $boxW[$a] / 2) }
            } else {
                $lo = [Math]::Min($a, $b); $hi = [Math]::Max($a, $b)
                $need = $t.W + 24
                $sum = 0.0
                for ($k = $lo; $k -lt $hi; $k++) { $sum += $gapX[$k] }
                if ($sum -lt $need) { for ($k = $lo; $k -lt $hi; $k++) { $gapX[$k] += ($need - $sum) / ($hi - $lo) } }
            }
        } elseif ($it.K -eq 'note') {
            $t = & $Measure $it.Text 'Small'
            $x.TW = $t.W; $x.TH = $t.H
            $nw = $t.W + 20
            $ids = @($it.Ids | ForEach-Object { $idx[$_] })
            $i0 = ($ids | Measure-Object -Minimum).Minimum
            $i1 = ($ids | Measure-Object -Maximum).Maximum
            switch ($it.Pos) {
                'left' { if ($i0 -eq 0) { $leftExtra = [Math]::Max($leftExtra, $nw + 10 - $boxW[0] / 2) } else { $gapX[$i0 - 1] = [Math]::Max($gapX[$i0 - 1], $nw + 20) } }
                'right' { if ($i1 -eq $n - 1) { $rightExtra = [Math]::Max($rightExtra, $nw + 10 - $boxW[$i1] / 2) } else { $gapX[$i1] = [Math]::Max($gapX[$i1], $nw + 20) } }
                default {
                    if ($i0 -eq 0) { $leftExtra = [Math]::Max($leftExtra, $nw / 2 - $boxW[0] / 2) }
                    if ($i1 -eq $n - 1) { $rightExtra = [Math]::Max($rightExtra, $nw / 2 - $boxW[$i1] / 2) }
                }
            }
        } elseif ($it.K -in 'start', 'else') {
            $t = & $Measure $(if ($it.Text) { "[$($it.Text)]" } else { 'Ag' }) 'Small'
            $x.TW = $t.W; $x.TH = $t.H
        }
        $info.Add($x)
    }
    $cx = [double[]]::new($n)
    $cx[0] = $g.Margin + [Math]::Max(0, $leftExtra) + $boxW[0] / 2
    for ($i = 1; $i -lt $n; $i++) { $cx[$i] = $cx[$i - 1] + $gapX[$i - 1] }

    $top = $g.Margin
    if ($M.Title) {
        $ts = & $Measure $M.Title 'Body'
        $items.Add(@{ K = 'text'; X = $g.Margin; Y = $top; W = [Math]::Max($ts.W, $cx[$n - 1] + $boxW[$n - 1] / 2 - $g.Margin); Text = $M.Title; Size = 'Body'; Align = 'Center'; Bold = $true; Color = 'Text' })
        $top += $ts.H + 12
    }
    $y = $top + $headH + 20
    $frames = [System.Collections.Generic.List[hashtable]]::new()
    $frameItems = [System.Collections.Generic.List[hashtable]]::new()
    $body = [System.Collections.Generic.List[hashtable]]::new()
    $maxRight = $cx[$n - 1] + $boxW[$n - 1] / 2 + [Math]::Max(0, $rightExtra)
    foreach ($x in $info) {
        $it = $x.It
        switch ($it.K) {
            'msg' {
                $a = $idx[$it.From]; $b = $idx[$it.To]
                if ($a -eq $b) {
                    $loopH = [Math]::Max(22, $x.TH)
                    $pts = [double[]]@($cx[$a], $y, ($cx[$a] + 30), $y, ($cx[$a] + 30), ($y + $loopH), $cx[$a], ($y + $loopH))
                    $body.Add(@{ K = 'line'; Pts = $pts; Style = $it.Line; Start = 'none'; End = $it.Head; Smooth = $false })
                    if ($x.Text) { $body.Add(@{ K = 'text'; X = $cx[$a] + 38; Y = $y + $loopH / 2 - $x.TH / 2; W = $x.TW; Text = $x.Text; Size = 'Small'; Align = 'Left'; Color = 'Text' }) }
                    Add-WtMmdFrameSpan $frames $a $a ($cx[$a] + 38 + $x.TW)
                    $y += $loopH + 18
                } else {
                    if ($x.Text) {
                        $mid = ($cx[$a] + $cx[$b]) / 2
                        $body.Add(@{ K = 'text'; X = $mid - $x.TW / 2; Y = $y; W = $x.TW; Text = $x.Text; Size = 'Small'; Align = 'Center'; Color = 'Text' })
                    }
                    $ly = $y + $x.TH + 4
                    $body.Add(@{ K = 'line'; Pts = [double[]]@($cx[$a], $ly, $cx[$b], $ly); Style = $it.Line; Start = $(if ($it.Both) { $it.Head } else { 'none' }); End = $it.Head; Smooth = $false })
                    Add-WtMmdFrameSpan $frames ([Math]::Min($a, $b)) ([Math]::Max($a, $b)) 0
                    $y = $ly + 18
                }
            }
            'note' {
                $ids = @($it.Ids | ForEach-Object { $idx[$_] })
                $i0 = ($ids | Measure-Object -Minimum).Minimum
                $i1 = ($ids | Measure-Object -Maximum).Maximum
                $nw = $x.TW + 20
                $nh = $x.TH + 12
                switch ($it.Pos) {
                    'left' { $nx = $cx[$i0] - 10 - $nw }
                    'right' { $nx = $cx[$i1] + 10 }
                    default {
                        if ($i0 -ne $i1) { $nw = [Math]::Max($nw, $cx[$i1] - $cx[$i0] + 40) }
                        $nx = ($cx[$i0] + $cx[$i1]) / 2 - $nw / 2
                    }
                }
                $body.Add(@{ K = 'node'; Shape = 'rect'; X = $nx; Y = $y; W = $nw; H = $nh; Text = $it.Text; TH = $x.TH; Size = 'Small'; Fill = 'Dia.NoteBg'; Line = 'Dia.NoteLine' })
                Add-WtMmdFrameSpan $frames $i0 $i1 ($nx + $nw)
                $maxRight = [Math]::Max($maxRight, $nx + $nw)
                $y += $nh + 14
            }
            'start' {
                $frames.Add(@{ Kind = $it.Kind; Text = $it.Text; Y = $y; Lo = -1; Hi = -1; MaxX = 0.0; Inner = 0; Elses = [System.Collections.Generic.List[hashtable]]::new(); TH = $x.TH })
                $y += $x.TH + 6 + 12
            }
            'else' {
                if ($frames.Count) { $frames[$frames.Count - 1].Elses.Add(@{ Y = $y; Text = $it.Text; TW = $x.TW }) }
                $y += $x.TH + 14
            }
            'end' {
                if (-not $frames.Count) { break }
                $f = $frames[$frames.Count - 1]
                $frames.RemoveAt($frames.Count - 1)
                $lo = if ($f.Lo -ge 0) { $f.Lo } else { 0 }
                $hi = if ($f.Hi -ge 0) { $f.Hi } else { $n - 1 }
                $left = $cx[$lo] - $boxW[$lo] / 2 - 6 - $f.Inner * 6
                $right = [Math]::Max($cx[$hi] + $boxW[$hi] / 2 + 6 + $f.Inner * 6, $f.MaxX + 10)
                $bottom = $y + 2
                $frameItems.Add(@{ K = 'rect'; X = $left; Y = $f.Y; W = $right - $left; H = $bottom - $f.Y; Fill = $null; Line = 'Dia.GroupLine'; Radius = 2 })
                $kt = & $Measure $f.Kind 'Small'
                $tabW = $kt.W + 16
                $frameItems.Add(@{ K = 'rect'; X = $left; Y = $f.Y; W = $tabW; H = $f.TH + 6; Fill = 'Dia.GroupBg'; Line = 'Dia.GroupLine'; Radius = 2 })
                $frameItems.Add(@{ K = 'text'; X = $left; Y = $f.Y + 3; W = $tabW; Text = $f.Kind; Size = 'Small'; Align = 'Center'; Bold = $true; Color = 'Text' })
                if ($f.Text) { $frameItems.Add(@{ K = 'text'; X = $left + $tabW + 8; Y = $f.Y + 3; W = [Math]::Max(10, $right - $left - $tabW - 12); Text = "[$($f.Text)]"; Size = 'Small'; Align = 'Left'; Color = 'SubText' }) }
                foreach ($el in $f.Elses) {
                    $frameItems.Add(@{ K = 'line'; Pts = [double[]]@($left, $el.Y, $right, $el.Y); Style = 'dashed'; Start = 'none'; End = 'none'; Smooth = $false; Color = 'Dia.GroupLine' })
                    if ($el.Text) { $frameItems.Add(@{ K = 'text'; X = $left + 8; Y = $el.Y + 3; W = $el.TW + 4; Text = "[$($el.Text)]"; Size = 'Small'; Align = 'Left'; Color = 'SubText' }) }
                }
                if ($frames.Count) {
                    $p = $frames[$frames.Count - 1]
                    $p.Inner = [Math]::Max($p.Inner, $f.Inner + 1)
                    $p.MaxX = [Math]::Max($p.MaxX, $right)
                    if ($p.Lo -lt 0 -or $lo -lt $p.Lo) { $p.Lo = $lo }
                    if ($hi -gt $p.Hi) { $p.Hi = $hi }
                }
                $maxRight = [Math]::Max($maxRight, $right)
                $y += 14
            }
        }
    }
    $bottomY = $y + 4
    $boxTop = $top
    foreach ($row in @($boxTop, $bottomY)) {
        for ($i = 0; $i -lt $n; $i++) {
            $p = $parts[$i]
            if ($p.Actor) {
                $items.Add(@{ K = 'actor'; X = $cx[$i] - 12; Y = $row; W = 24; H = 30 })
                $items.Add(@{ K = 'text'; X = $cx[$i] - $boxW[$i] / 2; Y = $row + 34; W = $boxW[$i]; Text = $p.Label; Size = 'Body'; Align = 'Center'; Color = 'Text' })
            } else {
                $items.Add(@{ K = 'node'; Shape = 'rect'; X = $cx[$i] - $boxW[$i] / 2; Y = $row + ($headH - $boxH); W = $boxW[$i]; H = $boxH; Text = $p.Label; TH = $boxTH[$i] })
            }
        }
    }
    $lifelines = for ($i = 0; $i -lt $n; $i++) {
        @{ K = 'line'; Pts = [double[]]@($cx[$i], ($boxTop + $headH), $cx[$i], $bottomY); Style = 'dashed'; Start = 'none'; End = 'none'; Smooth = $false; Color = 'Dia.GroupLine' }
    }
    $all = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($x in $frameItems) { if ($x.K -eq 'rect' -and -not $x.Fill) { $all.Add($x) } }
    foreach ($x in @($lifelines)) { $all.Add($x) }
    foreach ($x in $frameItems) { if (-not ($x.K -eq 'rect' -and -not $x.Fill)) { $all.Add($x) } }
    foreach ($x in $items) { $all.Add($x) }
    foreach ($x in $body) { $all.Add($x) }
    @{ Width = [Math]::Ceiling($maxRight + $g.Margin); Height = [Math]::Ceiling($bottomY + $headH + $g.Margin); Items = $all.ToArray() }
}

# 열린 블록 틀(loop 등) 모두에 이 메시지 · 노트가 걸친 참여자 범위를 넣는다
function Add-WtMmdFrameSpan($Frames, [int]$Lo, [int]$Hi, [double]$Right) {
    foreach ($f in $Frames) {
        if ($f.Lo -lt 0 -or $Lo -lt $f.Lo) { $f.Lo = $Lo }
        if ($Hi -gt $f.Hi) { $f.Hi = $Hi }
        if ($Right -gt $f.MaxX) { $f.MaxX = $Right }
    }
}

# ---- 도형 목록 → WPF -------------------------------------------------------------------------

$script:typeface = $null

# 지금 글자 크기 설정으로 글 크기를 잰다 (Get-WtFontSize: 설정에서 바로 읽음)
function Measure-WtMmdText([string]$Text, [string]$Size) {
    if (-not $script:typeface) {
        $script:typeface = [System.Windows.Media.Typeface]::new([System.Windows.Media.FontFamily]::new((Get-WtFontName 'UI')),
            [System.Windows.FontStyles]::Normal, [System.Windows.FontWeights]::Normal, [System.Windows.FontStretches]::Normal)
    }
    $em = Get-WtFontSize $Size
    $w = 0.0
    $h = 0.0
    foreach ($l in (([string]$Text) -split "`n")) {
        $ft = [System.Windows.Media.FormattedText]::new($(if ($l) { $l } else { ' ' }), [System.Globalization.CultureInfo]::CurrentUICulture,
            [System.Windows.FlowDirection]::LeftToRight, $script:typeface, $em, [System.Windows.Media.Brushes]::Black, 1.0)
        $w = [Math]::Max($w, $ft.WidthIncludingTrailingWhitespace)
        $h += $ft.Height
    }
    @{ W = [Math]::Ceiling($w) + 4; H = [Math]::Ceiling($h) }
}

# Mermaid 글 → 그림 요소. 폭이 모자라면 줄여서 보여 준다 (Viewbox, 키우지는 않음)
function New-WtMermaidView([string]$Code) {
    $model = ConvertFrom-WtMermaid $Code
    $layout = Get-WtMermaidLayout $model { param($t, $s) Measure-WtMmdText $t $s }
    $canvas = [System.Windows.Controls.Canvas]::new()
    $canvas.Width = $layout.Width
    $canvas.Height = $layout.Height
    $canvas.SnapsToDevicePixels = $true
    foreach ($it in $layout.Items) {
        foreach ($el in @(New-WtMmdElement $it)) { if ($el) { [void]$canvas.Children.Add($el) } }
    }
    $vb = [System.Windows.Controls.Viewbox]::new()
    $vb.Stretch = 'Uniform'
    $vb.StretchDirection = 'DownOnly'
    $vb.HorizontalAlignment = 'Center'
    $vb.Margin = [System.Windows.Thickness]::new(0, 4, 0, 0)
    $vb.Child = $canvas
    $vb
}

function Set-WtMmdPos($El, [double]$X, [double]$Y) {
    [System.Windows.Controls.Canvas]::SetLeft($El, $X)
    [System.Windows.Controls.Canvas]::SetTop($El, $Y)
    $El
}

function New-WtMmdPoints([double[]]$P) {
    $c = [System.Windows.Media.PointCollection]::new()
    for ($i = 0; $i -lt $P.Length; $i += 2) { $c.Add([System.Windows.Point]::new($P[$i], $P[$i + 1])) }
    , $c   # 모음 그대로 돌려준다 (그냥 돌려주면 PowerShell이 낱개로 풀어 버린다)
}

function New-WtMmdTextBlock([string]$Text, [string]$Size, [string]$Align, [double]$boxW, [bool]$Bold, [string]$Color) {
    $t = [System.Windows.Controls.TextBlock]::new()
    $t.Text = $Text
    Set-WtResource $t 'FontFamily' 'Font.UI'
    Set-WtResource $t 'FontSize' "F.$Size"
    Set-WtResource $t 'Foreground' "B.$Color"
    $t.TextAlignment = $Align
    $t.Width = $boxW
    if ($Bold) { $t.FontWeight = 'SemiBold' }
    $t
}

function New-WtMmdElement($It) {
    switch ($It.K) {
        'rect' {
            $r = [System.Windows.Shapes.Rectangle]::new()
            $r.Width = $It.W; $r.Height = $It.H
            $r.RadiusX = [double]$It.Radius; $r.RadiusY = [double]$It.Radius
            if ($It.Fill) { Set-WtResource $r 'Fill' "B.$($It.Fill)" }
            Set-WtResource $r 'Stroke' "B.$($It.Line)"
            $r.StrokeThickness = 1
            return Set-WtMmdPos $r $It.X $It.Y
        }
        'text' {
            $size = if ($It.Size) { $It.Size } else { 'Body' }
            $t = New-WtMmdTextBlock $It.Text $size $It.Align $It.W ([bool]$It.Bold) $(if ($It.Color) { $It.Color } else { 'Text' })
            if ($It.Bg) {
                $b = [System.Windows.Controls.Border]::new()
                Set-WtResource $b 'Background' "B.$($It.Bg)"
                $b.CornerRadius = [System.Windows.CornerRadius]::new(3)
                $b.Width = $It.W
                $b.Height = $It.H
                $t.Width = [double]::NaN
                $t.VerticalAlignment = 'Center'
                $t.HorizontalAlignment = 'Center'
                $b.Child = $t
                return Set-WtMmdPos $b $It.X $It.Y
            }
            return Set-WtMmdPos $t $It.X $It.Y
        }
        'node' { return New-WtMmdNode $It }
        'line' { return New-WtMmdLine $It }
        'actor' {
            $p = [System.Windows.Shapes.Path]::new()
            $w = $It.W; $hh = $It.H
            $geo = "M $($w/2),11 L $($w/2),$($hh*0.68) M 1,$($hh*0.45) L $($w-1),$($hh*0.45) M $($w/2),$($hh*0.68) L 2,$hh M $($w/2),$($hh*0.68) L $($w-2),$hh"
            $p.Data = [System.Windows.Media.Geometry]::Parse($geo)
            Set-WtResource $p 'Stroke' 'B.Dia.Line'
            $p.StrokeThickness = 1.5
            $head = [System.Windows.Shapes.Ellipse]::new()
            $head.Width = 11; $head.Height = 11
            Set-WtResource $head 'Fill' 'B.Dia.NodeBg'
            Set-WtResource $head 'Stroke' 'B.Dia.Line'
            $head.StrokeThickness = 1.5
            return @((Set-WtMmdPos $p $It.X $It.Y), (Set-WtMmdPos $head ($It.X + $w / 2 - 5.5) $It.Y))
        }
    }
    $null
}

function New-WtMmdNode($It) {
    $w = [double]$It.W
    $hh = [double]$It.H
    $fill = if ($It.Fill) { "B.$($It.Fill)" } else { 'B.Dia.NodeBg' }
    $line = if ($It.Line) { "B.$($It.Line)" } else { 'B.Dia.NodeLine' }
    $shapes = [System.Collections.Generic.List[object]]::new()
    $k = [Math]::Min($hh / 2, $w / 4)
    switch ($It.Shape) {
        { $_ -in 'rect', 'round', 'stadium', 'subroutine' } {
            $r = [System.Windows.Shapes.Rectangle]::new()
            $rad = switch ($It.Shape) { 'round' { 10 } 'stadium' { $hh / 2 } default { 3 } }
            $r.RadiusX = $rad; $r.RadiusY = $rad
            $shapes.Add($r)
            if ($It.Shape -eq 'subroutine') {
                foreach ($lx in 8, ($w - 8)) {
                    $l = [System.Windows.Shapes.Line]::new()
                    $l.X1 = $lx; $l.X2 = $lx; $l.Y1 = 0; $l.Y2 = $hh
                    $shapes.Add($l)
                }
            }
        }
        { $_ -in 'circle', 'dcircle', 'start', 'end' } {
            $e = [System.Windows.Shapes.Ellipse]::new()
            $shapes.Add($e)
            if ($It.Shape -in 'dcircle', 'end') {
                $in = [System.Windows.Shapes.Ellipse]::new()
                $inset = if ($It.Shape -eq 'end') { 4 } else { 5 }
                $in.Width = $w - 2 * $inset; $in.Height = $hh - 2 * $inset
                $in.Margin = [System.Windows.Thickness]::new($inset)
                $shapes.Add($in)
            }
        }
        'fork' { $shapes.Add([System.Windows.Shapes.Rectangle]::new()) }
        'cylinder' {
            $ry = 7
            $p = [System.Windows.Shapes.Path]::new()
            $ci = [System.Globalization.CultureInfo]::InvariantCulture
            $data = [string]::Format($ci, 'M0,{0} A{1},{0} 0 0 1 {2},{0} L{2},{3} A{1},{0} 0 0 1 0,{3} Z M0,{0} A{1},{0} 0 0 0 {2},{0}', $ry, ($w / 2), $w, ($hh - $ry))
            $p.Data = [System.Windows.Media.Geometry]::Parse($data)
            $shapes.Add($p)
        }
        default {
            $pts = switch ($It.Shape) {
                { $_ -in 'diamond', 'choice' } { @(($w / 2), 0, $w, ($hh / 2), ($w / 2), $hh, 0, ($hh / 2)) }
                'hexagon' { @($k, 0, ($w - $k), 0, $w, ($hh / 2), ($w - $k), $hh, $k, $hh, 0, ($hh / 2)) }
                'para' { $s = $hh * 0.4; @($s, 0, $w, 0, ($w - $s), $hh, 0, $hh) }
                'paraAlt' { $s = $hh * 0.4; @(0, 0, ($w - $s), 0, $w, $hh, $s, $hh) }
                'trap' { $s = $hh * 0.4; @($s, 0, ($w - $s), 0, $w, $hh, 0, $hh) }
                'trapAlt' { $s = $hh * 0.4; @(0, 0, $w, 0, ($w - $s), $hh, $s, $hh) }
                'asym' { @(0, 0, $w, 0, $w, $hh, 0, $hh, ($hh / 2), ($hh / 2)) }
                default { @(0, 0, $w, 0, $w, $hh, 0, $hh) }
            }
            $pg = [System.Windows.Shapes.Polygon]::new()
            $pg.Points = New-WtMmdPoints ([double[]]$pts)
            $shapes.Add($pg)
        }
    }
    $out = [System.Collections.Generic.List[object]]::new()
    $solid = $It.Shape -in 'start', 'fork'
    for ($i = 0; $i -lt $shapes.Count; $i++) {
        $s = $shapes[$i]
        if ($s -is [System.Windows.Shapes.Rectangle] -or $s -is [System.Windows.Shapes.Ellipse]) {
            if ([double]::IsNaN($s.Width)) { $s.Width = $w; $s.Height = $hh }
        }
        if ($s -isnot [System.Windows.Shapes.Line]) {
            $f = if ($solid -or ($It.Shape -eq 'end' -and $i -eq 1)) { 'B.Dia.Line' } elseif ($It.Shape -eq 'end') { 'B.Card' } else { $fill }
            Set-WtResource $s 'Fill' $f
        }
        Set-WtResource $s 'Stroke' $(if ($solid) { 'B.Dia.Line' } else { $line })
        $s.StrokeThickness = 1
        $out.Add((Set-WtMmdPos $s $It.X $It.Y))
    }
    if ($It.Text -and $It.Shape -notin 'start', 'end', 'fork', 'choice') {
        $size = if ($It.Size) { $It.Size } else { 'Body' }
        $t = New-WtMmdTextBlock $It.Text $size 'Center' $w $false 'Text'
        $out.Add((Set-WtMmdPos $t $It.X ($It.Y + ($hh - [double]$It.TH) / 2)))
    }
    $out.ToArray()
}

function New-WtMmdLine($It) {
    $p = [double[]]$It.Pts.Clone()
    $n = $p.Length / 2
    if ($n -lt 2) { return $null }
    $g = $script:gap
    $color = if ($It.Color) { "B.$($It.Color)" } else { 'B.Dia.Line' }
    $thick = switch ($It.Style) { 'thick' { 3.0 } default { 1.5 } }
    $out = [System.Collections.Generic.List[object]]::new()
    # 끝 모양 자리만큼 선을 줄이고 끝 모양을 그린다
    foreach ($end in @(@{ Kind = $It.End; I = $n - 1; J = $n - 2 }, @{ Kind = $It.Start; I = 0; J = 1 })) {
        if (-not $end.Kind -or $end.Kind -eq 'none') { continue }
        $tx = $p[2 * $end.I]; $ty = $p[2 * $end.I + 1]
        $dx = $tx - $p[2 * $end.J]; $dy = $ty - $p[2 * $end.J + 1]
        $len = [Math]::Sqrt($dx * $dx + $dy * $dy)
        if ($len -lt 0.01) { continue }
        $ux = $dx / $len; $uy = $dy / $len
        $back = switch ($end.Kind) { 'arrow' { $g.Arrow } 'circle' { 8 } default { 0 } }
        $back = [Math]::Min($back, $len - 1)
        $p[2 * $end.I] = $tx - $ux * $back
        $p[2 * $end.I + 1] = $ty - $uy * $back
        $out.Add((New-WtMmdMarker $end.Kind $tx $ty $ux $uy $color))
    }
    $path = [System.Windows.Shapes.Path]::new()
    $path.Data = New-WtMmdGeometry $p ([bool]$It.Smooth)
    Set-WtResource $path 'Stroke' $color
    $path.StrokeThickness = $thick
    switch ($It.Style) {
        'dotted' { $path.StrokeDashArray = [System.Windows.Media.DoubleCollection]::new([double[]]@(2, 2)) }
        'dashed' { $path.StrokeDashArray = [System.Windows.Media.DoubleCollection]::new([double[]]@(4, 3)) }
    }
    $out.Insert(0, $path)
    $out.ToArray()
}

# 점들을 잇는 선. 부드럽게(Catmull-Rom → 베지어) 또는 꺾은선
function New-WtMmdGeometry([double[]]$P, [bool]$Smooth) {
    $n = $P.Length / 2
    $fig = [System.Windows.Media.PathFigure]::new()
    $fig.StartPoint = [System.Windows.Point]::new($P[0], $P[1])
    $fig.IsFilled = $false
    if ($Smooth -and $n -gt 2) {
        for ($i = 0; $i -lt $n - 1; $i++) {
            $i0 = [Math]::Max($i - 1, 0); $i3 = [Math]::Min($i + 2, $n - 1)
            $c1 = [System.Windows.Point]::new($P[2 * $i] + ($P[2 * ($i + 1)] - $P[2 * $i0]) / 6, $P[2 * $i + 1] + ($P[2 * ($i + 1) + 1] - $P[2 * $i0 + 1]) / 6)
            $c2 = [System.Windows.Point]::new($P[2 * ($i + 1)] - ($P[2 * $i3] - $P[2 * $i]) / 6, $P[2 * ($i + 1) + 1] - ($P[2 * $i3 + 1] - $P[2 * $i + 1]) / 6)
            $fig.Segments.Add([System.Windows.Media.BezierSegment]::new($c1, $c2, [System.Windows.Point]::new($P[2 * ($i + 1)], $P[2 * ($i + 1) + 1]), $true))
        }
    } else {
        for ($i = 1; $i -lt $n; $i++) { $fig.Segments.Add([System.Windows.Media.LineSegment]::new([System.Windows.Point]::new($P[2 * $i], $P[2 * $i + 1]), $true)) }
    }
    $geo = [System.Windows.Media.PathGeometry]::new()
    $geo.Figures.Add($fig)
    $geo
}

function New-WtMmdMarker([string]$Kind, [double]$Tx, [double]$Ty, [double]$Ux, [double]$Uy, [string]$Color) {
    $arrowLen = $script:gap.Arrow
    $px = -$Uy; $py = $Ux   # 선에 수직
    switch ($Kind) {
        'arrow' {
            $pg = [System.Windows.Shapes.Polygon]::new()
            $bx = $Tx - $Ux * $arrowLen; $by = $Ty - $Uy * $arrowLen
            $pg.Points = New-WtMmdPoints ([double[]]@($Tx, $Ty, ($bx + $px * 4.5), ($by + $py * 4.5), ($bx - $px * 4.5), ($by - $py * 4.5)))
            Set-WtResource $pg 'Fill' $Color
            return $pg
        }
        'async' {
            $pl = [System.Windows.Shapes.Polyline]::new()
            $bx = $Tx - $Ux * $arrowLen; $by = $Ty - $Uy * $arrowLen
            $pl.Points = New-WtMmdPoints ([double[]]@(($bx + $px * 4.5), ($by + $py * 4.5), $Tx, $Ty, ($bx - $px * 4.5), ($by - $py * 4.5)))
            Set-WtResource $pl 'Stroke' $Color
            $pl.StrokeThickness = 1.5
            return $pl
        }
        'circle' {
            $e = [System.Windows.Shapes.Ellipse]::new()
            $e.Width = 8; $e.Height = 8
            Set-WtResource $e 'Fill' 'B.Card'
            Set-WtResource $e 'Stroke' $Color
            $e.StrokeThickness = 1.5
            return Set-WtMmdPos $e ($Tx - $Ux * 4 - 4) ($Ty - $Uy * 4 - 4)
        }
        'cross' {
            $cx = $Tx - $Ux * 5; $cy = $Ty - $Uy * 5
            $pth = [System.Windows.Shapes.Path]::new()
            $ci = [System.Globalization.CultureInfo]::InvariantCulture
            $d = [string]::Format($ci, 'M{0},{1} L{2},{3} M{4},{5} L{6},{7}',
                ($cx + ($Ux + $px) * 4), ($cy + ($Uy + $py) * 4), ($cx - ($Ux + $px) * 4), ($cy - ($Uy + $py) * 4),
                ($cx + ($Ux - $px) * 4), ($cy + ($Uy - $py) * 4), ($cx - ($Ux - $px) * 4), ($cy - ($Uy - $py) * 4))
            $pth.Data = [System.Windows.Media.Geometry]::Parse($d)
            Set-WtResource $pth 'Stroke' $Color
            $pth.StrokeThickness = 1.5
            return $pth
        }
    }
    $null
}

Register-WtMarkdownCodeRenderer -Language 'mermaid' -Label 'mermaid' -Remeasure -Render { param($code) New-WtMermaidView $code }

Export-ModuleMember -Function ConvertFrom-WtMermaid, Get-WtMermaidLayout, New-WtMermaidView
} | Import-Module -Force -DisableNameChecking
# ---- core\Mermaid.psm1 끝 ----

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
$script:helpMd = $null
$script:helpRaw = $null

# 도움말의 마크다운 예시. 원문 · 보기를 바꿔 가며 문법을 본다
$script:markdownSample = @'
# 제목은 `#`, 소제목은 `##`

**굵게**, *기울임*, ~~취소선~~, ==형광펜==, `코드`, [링크](https://github.com)
[[메모 제목]]을 누르면 그 메모로, #태그 를 누르면 그 태그로 찾습니다.
줄을 한 번만 바꿔도 줄이 바뀝니다 (Obsidian과 같음).

- [x] 끝낸 일
- [ ] 할 일
- 목록
    1. 안쪽 번호 목록

> [!tip] 콜아웃
> `> [!종류] 제목`으로 씁니다. 종류: note · info · tip · success · question · warning · danger · bug · example · quote

> [!warning]- 접는 콜아웃 (눌러서 펼치기)
> `-`는 접힌 채로, `+`는 펼친 채로 시작합니다.

| 서버 | IP | 상태 |
|:--|:--|:-:|
| SRV-01 | 10.0.0.10 | 정상 |
| SRV-02 | 10.0.0.11 | 점검 |

```powershell
Get-Service | Where-Object Status -eq 'Running'   # 코드 블록은 [복사]로 한 번에 복사
```

```mermaid
flowchart LR
    A[요청] --> B{승인?}
    B -->|예| C[작업]
    B -->|아니오| D([반려])
```
'@

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
    $seg = New-WtSegmented @(@('s', '작게'), @('m', '보통'), @('l', '크게')) ([string](Get-WtSetting 'ui.fontScale')) 'WtFont' {
        param($value) Set-WtSetting 'ui.fontScale' $value
    }
    [void](Add-WtChild $col (New-WtSettingRow '글자 크기' $seg))

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
    [void](Add-WtChild $page.Panel (New-WtMarkdownHelpCard))

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

function New-WtMarkdownHelpCard {
    $card = New-WtCard
    $col = New-WtStack Vertical 8
    [void](Add-WtChild $col (New-WtCardHeader '마크다운'))
    [void](Add-WtChild $col (New-WtText '메모 본문은 마크다운으로 보입니다. 원문을 누르면 아래 보기를 만든 글이 나옵니다.' 'T.Sub'))
    $mode = New-WtSegmented @(@('preview', '보기'), @('raw', '원문')) 'preview' 'WtHelpMd' {
        param($value)
        $script:helpMd.Visibility = if ($value -eq 'raw') { 'Collapsed' } else { 'Visible' }
        $script:helpRaw.Visibility = if ($value -eq 'raw') { 'Visible' } else { 'Collapsed' }
    }
    [void](Add-WtChild $col $mode)
    # 도움말에서는 [[링크]] · #태그가 갈 곳이 없으므로 무엇을 하는지만 알려 준다
    $script:helpMd = New-WtMarkdownView -OnLink {
        param($link)
        if ($link.Kind -eq 'wiki') { Show-WtMessage '메모에서 누르면 제목이 같은 메모를 엽니다.'; return $true }
        if ($link.Kind -eq 'tag') { Show-WtMessage '메모에서 누르면 이 태그로 찾습니다.'; return $true }
        $false
    }
    $script:helpMd.Margin = [System.Windows.Thickness]::new(0, 8, 0, 0)
    Set-WtMarkdownText $script:helpMd $script:markdownSample
    $script:helpRaw = [System.Windows.Controls.TextBox]::new()
    Set-WtStyle $script:helpRaw 'Reader'
    Set-WtResource $script:helpRaw 'FontFamily' 'Font.Mono'
    $script:helpRaw.Text = $script:markdownSample
    $script:helpRaw.Visibility = 'Collapsed'
    $script:helpRaw.Margin = [System.Windows.Thickness]::new(0, 8, 0, 0)
    [void](Add-WtChild $col $script:helpMd)
    [void](Add-WtChild $col $script:helpRaw)
    $card.Child = $col
    $card
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
