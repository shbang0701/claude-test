<#
.SYNOPSIS
    근태 탭을 임시 데이터로 검사한다: 예전 데이터 가져오기 · 한 달 요약 · 시간 표시.

.DESCRIPTION
    tests\Test-Core.ps1이 찾아서 돌린다. 혼자 돌려도 된다. 실제 데이터 폴더(문서\WorkTool)는 쓰지 않는다.
    근태 표시 색의 대비는 테마 값이라 tests\Test-Core.ps1이 검사한다.

.EXAMPLE
    pwsh -NoProfile -File apps/WorkTool/wpf/addons/attendance/Test-Attendance.ps1
#>
#Requires -Version 7.0
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) 'tests/TestKit.ps1')
Import-WtTestCore
Import-Module (Join-Path $PSScriptRoot 'Attendance.psm1') -Force -DisableNameChecking
$test = New-WtTestData 'attendance'

try {
    Write-WtTestLegacyCsv $test 'attendance' @'
"Id","Type","StartDate","EndDate","StartTime","EndTime","Place","Detail","Updated"
"t1","출근","2026-10-21","2026-10-21","08:47","","","",""
"t2","출근","2026-10-21","2026-10-21","09:00","","","",""
"t3","업무외시간 근무","2026-10-20","2026-10-20","19:00","21:30","","패치",""
"t4","업무외시간 근무","2026-10-13","2026-10-13","23:00","01:00","","야간",""
"t5","출장","2026-10-16","2026-10-17","","","데이터센터","점검",""
"t6","휴가","2026-10-17","2026-10-17","","","","오후 반차",""
"t7","휴가","2026-09-30","2026-10-02","","","","연차",""
"t8","출근","2026-09-30","2026-09-30","08:50","","","",""
'@
    $r = Import-WtLegacyData
    Assert-WtTest '근태: 예전 데이터를 가져온다' ($r.Count -eq 8 -and $r.Tabs.Count -eq 1) "($($r.Count) / $($r.Tabs -join ', '))"

    $d = { param($t) [datetime]::ParseExact($t, 'yyyy-MM-dd', $null) }
    $sum = Get-WtAttendanceSummary (& $d '2026-10-01') (Get-WtAttendanceItems)
    Assert-WtTest '근태: 출근 일수(같은 날 두 번은 하루)' ($sum.Work -eq 1) $sum.Work
    Assert-WtTest '근태: 업무외 시간(자정 넘김 포함)' ($sum.OverMinutes -eq 270) $sum.OverMinutes
    Assert-WtTest '근태: 출장 건수' ($sum.Trip -eq 1) $sum.Trip
    Assert-WtTest '근태: 휴가 일수(반차 0.5, 달 경계 자름)' ($sum.Leave -eq 2.5) $sum.Leave
    Assert-WtTest '근태: 시간 표시' ((Format-WtHours 270) -eq '4.5' -and (Format-WtHours 120) -eq '2')
} finally {
    Remove-WtTestData $test
}
