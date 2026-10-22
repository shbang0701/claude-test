<#
.SYNOPSIS
    WorkTool을 파일 하나(dist\WorkTool.ps1)로 묶는다. 업무 PC에 복사 한 번으로 가져가기 위한 것이다.

.DESCRIPTION
    - Start-WorkTool.ps1의 Import-Module 줄 자리에 해당 모듈 소스를 New-Module 블록으로 그대로 넣는다.
      모듈 코드는 문자열로 실행하지 않고 스크립트 블록 안에 그대로 들어간다 (Invoke-Expression 없음).
    - Storage의 데이터 폴더는 묶음 파일 옆 data\ 로 바꾼다. 원래 구조의 WorkTool\data 와 같은 위치다.
    - 바꿀 줄을 정확히 한 번 찾지 못하면 멈춘다. 원본 구조가 바뀌면 이 스크립트도 같이 고친다.
    - 결과는 원본이 같으면 항상 같다 (날짜·커밋을 넣지 않음). 다시 만들어서 차이가 없으면 최신이다.

.EXAMPLE
    pwsh ./apps/WorkTool/build/Build-Bundle.ps1
    apps/WorkTool/dist/WorkTool.ps1 을 새로 만든다.
#>
#Requires -Version 7.0
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$app = Split-Path $PSScriptRoot -Parent
$moduleDir = Join-Path $app 'modules'
$outPath = Join-Path $app 'dist' 'WorkTool.ps1'

function Get-SourceText([string]$Path) {
    [System.IO.File]::ReadAllText($Path).Replace("`r`n", "`n").TrimEnd("`n")
}

# $Old가 정확히 한 번 있을 때만 바꾼다
function Set-Once([string]$Text, [string]$Old, [string]$New, [string]$What) {
    $i = $Text.IndexOf($Old, [System.StringComparison]::Ordinal)
    if ($i -lt 0 -or $Text.IndexOf($Old, $i + 1, [System.StringComparison]::Ordinal) -ge 0) {
        throw "${What}: 바꿀 부분을 정확히 한 번 찾지 못했습니다 → $Old"
    }
    $Text.Substring(0, $i) + $New + $Text.Substring($i + $Old.Length)
}

function Get-ModuleBlock([string]$Name) {
    $src = Get-SourceText (Join-Path $moduleDir "$Name.psm1")
    if ($Name -eq 'Storage') {
        $src = Set-Once $src "Join-Path (Split-Path `$PSScriptRoot -Parent) 'data'" "Join-Path `$PSScriptRoot 'data'" 'Storage 데이터 폴더'
    }
    "# ---- modules\$Name.psm1 ----`nNew-Module -Name $Name -ScriptBlock {`n$src`n} | Import-Module -Force`n# ---- modules\$Name.psm1 끝 ----"
}

$text = Get-SourceText (Join-Path $app 'Start-WorkTool.ps1')

# 공통 모듈: Import-Module 줄을 모듈 블록으로
foreach ($name in 'Storage', 'UiHelper') {
    $text = Set-Once $text "Import-Module (Join-Path `$root 'modules\$name.psm1') -Force" (Get-ModuleBlock $name) "공통 모듈 $name"
}

# 기능 모듈: 반복문 안의 Import-Module을 빼고, 반복문 앞에 모든 기능 모듈 블록을 넣는다
$text = Set-Once $text "`n    Import-Module (Join-Path `$root `"modules\`$(`$f.Module).psm1`") -Force" '' '기능 모듈 불러오기'
$features = Get-ChildItem $moduleDir -Filter '*.psm1' |
    Where-Object { $_.BaseName -notin 'Storage', 'UiHelper' } |
    Sort-Object Name |
    ForEach-Object { Get-ModuleBlock $_.BaseName }
$text = Set-Once $text "foreach (`$f in `$features) {" (($features -join "`n`n") + "`n`nforeach (`$f in `$features) {") '기능 반복문'

$header = @(
    '# WorkTool — 파일 하나로 묶은 버전 (자동 생성)'
    '# 이 파일은 직접 고치지 않는다. apps/WorkTool의 원본을 고친 뒤 build/Build-Bundle.ps1로 다시 만든다.'
    '# 데이터는 이 파일 옆 data 폴더에 저장된다. 기존 WorkTool 폴더에 이 파일을 두면 기존 데이터를 그대로 쓴다.'
    ''
) -join "`n"

New-Item -ItemType Directory -Path (Split-Path $outPath) -Force | Out-Null
[System.IO.File]::WriteAllText($outPath, $header + $text + "`n", [System.Text.UTF8Encoding]::new($false))
Write-Host "만듦: $outPath"
