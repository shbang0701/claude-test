<#
.SYNOPSIS
    (한 줄 설명) 예: 결과를 새 Excel 파일로 저장한다.

.DESCRIPTION
    - 무엇을 읽고 무엇을 만드는지 적는다.
    - 원본은 바꾸지 않는다. 결과는 출력 폴더에 날짜가 붙은 새 파일로 저장한다.
    - Excel이 필요 없는 스크립트면 Invoke-Excel 함수와 본문 예시를 지운다.

.EXAMPLE
    .\Verb-Noun.ps1
    설정 블록의 값으로 실행한다.

.EXAMPLE
    .\Verb-Noun.ps1 -WhatIf
    파일을 저장하지 않고 할 일만 보여 준다.
#>
#Requires -Version 7.0
[CmdletBinding(SupportsShouldProcess)]
param()

# ===== 설정 (여기만 고친다) =====
$Config = @{
    OutputFolder = Join-Path $PSScriptRoot 'output'   # 결과를 저장할 폴더
    OutputName   = '결과'                              # 결과 파일 이름 (뒤에 날짜가 붙는다)
}
# =================================

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# 결과 파일 경로: <이름>_yyyyMMdd_HHmm<확장자>. 같은 이름이 있어도 덮어쓰지 않게 날짜를 붙인다.
function Get-OutputPath([string]$BaseName, [string]$Extension) {
    if (-not (Test-Path $Config.OutputFolder)) {
        New-Item -ItemType Directory -Path $Config.OutputFolder | Out-Null
    }
    Join-Path $Config.OutputFolder ('{0}_{1:yyyyMMdd_HHmm}{2}' -f $BaseName, (Get-Date), $Extension)
}

# Excel을 띄우지 않고 열어 작업한 뒤, 오류가 나도 반드시 닫는다 (EXCEL.EXE가 남지 않게).
# 사용: Invoke-Excel { param($excel) ... }
function Invoke-Excel([scriptblock]$Work) {
    $excel = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $false
        & $Work $excel
    } finally {
        if ($excel) {
            $excel.Quit()
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

# ===== 본문 =====
Invoke-Excel {
    param($excel)
    $book = $excel.Workbooks.Add()
    try {
        $sheet = $book.Worksheets.Item(1)

        # 값은 2차원 배열로 한 번에 쓴다. 셀을 하나씩 쓰면 아주 느리다.
        $rows = @(
            , @('항목', '값')
            , @('예시', 1)
        )
        $colCount = $rows[0].Count
        $data = [object[,]]::new($rows.Count, $colCount)
        for ($r = 0; $r -lt $rows.Count; $r++) {
            for ($c = 0; $c -lt $colCount; $c++) { $data[$r, $c] = $rows[$r][$c] }
        }
        $range = $sheet.Range($sheet.Cells.Item(1, 1), $sheet.Cells.Item($rows.Count, $colCount))
        $range.Value2 = $data

        $path = Get-OutputPath $Config.OutputName '.xlsx'
        if ($PSCmdlet.ShouldProcess($path, '저장')) {
            $book.SaveAs($path, 51)   # 51 = .xlsx
            Write-Host "저장: $path"
        }
    } finally {
        $book.Close($false)
    }
}
