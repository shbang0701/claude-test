# 저장 담당. 다른 모듈은 파일을 직접 읽고 쓰지 않고 이 모듈만 거친다.
# 지금은 data\<이름>.csv에 저장한다. 나중에 저장 방식을 바꿀 때는 이 파일만 고친다.

# 읽기 실패를 빈 결과로 넘기면 다음 저장에서 기존 데이터를 덮어쓰므로, 실패는 항상 예외로 올린다.
$ErrorActionPreference = 'Stop'
$dataDir = Join-Path (Split-Path $PSScriptRoot -Parent) 'data'
# Excel에서 열어도 한글이 깨지지 않게 BOM을 붙인다.
$encoding = [System.Text.UTF8Encoding]::new($true)

# 핸들을 앱이 끝날 때까지 유지한다. 비정상 종료되어도 OS가 잠금을 풀어 준다.
function Open-WorkToolDataLock {
    [void][System.IO.Directory]::CreateDirectory($dataDir)
    try {
        [System.IO.File]::Open((Join-Path $dataDir '.worktool.lock'), 'OpenOrCreate', 'ReadWrite', 'None')
    } catch {
        throw '같은 데이터 폴더를 사용하는 WorkTool이 실행 중이거나 폴더를 잠글 수 없습니다. 기존 창·트레이를 확인하세요.'
    }
}

function Get-DataPath {
    param([Parameter(Mandatory)][ValidatePattern('^[A-Za-z0-9_-]+$')][string]$Name)
    if (-not (Test-Path $dataDir)) { New-Item -ItemType Directory -Path $dataDir | Out-Null }
    Join-Path $dataDir "$Name.csv"
}

# 기록을 읽는다. 파일이 없으면 아무것도 돌려주지 않으므로 호출하는 쪽에서 @()로 감싼다.
# Import-Csv는 모든 값을 문자열로 읽는다. 날짜·숫자는 호출하는 쪽에서 변환한다.
function Get-Records {
    param([Parameter(Mandatory)][string]$Name)
    $path = Get-DataPath $Name
    if (Test-Path $path) { Import-Csv -Path $path -Encoding utf8 }
}

# 기록 전체를 저장한다. 임시 파일에 먼저 쓰고 바꿔치기해서, 저장 중에 꺼져도 원래 파일은 남는다.
# 직전 파일은 <이름>.csv.bak으로 하나 남긴다.
function Save-Records {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Records
    )
    $path = Get-DataPath $Name
    $tmp = "$path.tmp"
    $text = if ($Records.Count) { ($Records | ConvertTo-Csv -NoTypeInformation) -join "`r`n" } else { '' }
    [System.IO.File]::WriteAllText($tmp, $text, $encoding)
    if (Test-Path $path) {
        [System.IO.File]::Replace($tmp, $path, "$path.bak")
    } else {
        Move-Item -Path $tmp -Destination $path
    }
}

# 새 기록에 붙일 짧은 ID
function New-RecordId {
    [guid]::NewGuid().ToString('N').Substring(0, 12)
}

Export-ModuleMember -Function Get-Records, Save-Records, New-RecordId, Open-WorkToolDataLock
