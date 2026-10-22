<#
.SYNOPSIS
    WorkTool 새 버전에서 창 없이 확인할 수 있는 부분을 검사한다.

.DESCRIPTION
    저장 · 설정 · 예전 데이터 가져오기 · 일정 반복 · 근태 요약 · 메모 검색 · 테마 색 대비.
    임시 폴더만 쓰고, 문서\WorkTool 같은 실제 데이터 폴더는 건드리지 않는다. Windows · Linux 모두에서 돈다.

.EXAMPLE
    pwsh ./apps/WorkTool/wpf/tests/Test-Core.ps1
#>
#Requires -Version 7.0
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
foreach ($n in 'Storage', 'Settings', 'Theme', 'Ui', 'Addons', 'Pages', 'Shell') {
    Import-Module (Join-Path $root "core/$n.psm1") -Force -DisableNameChecking
}
foreach ($f in (Get-ChildItem (Join-Path $root 'addons') -Filter '*.psm1' -Recurse | Sort-Object FullName)) {
    Import-Module $f.FullName -Force -DisableNameChecking
}

$script:failed = 0
function Assert-Wt([string]$Name, [bool]$Ok, $Detail = '') {
    if ($Ok) { Write-Host "PASS $Name" } else { Write-Host "FAIL $Name $Detail" -ForegroundColor Red; $script:failed++ }
}

$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('wt-test-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
$app = (New-Item -ItemType Directory -Path (Join-Path $tmp 'app')).FullName
$dataDir = Join-Path $tmp 'data'
Set-Content -LiteralPath (Join-Path $app 'datapath.txt') -Value $dataDir -Encoding utf8
Set-WtAppRoot $app

try {
    # ---- 저장 ----
    Assert-Wt '데이터 폴더는 datapath.txt를 따른다' ((Get-WtDataRoot) -eq (Resolve-Path $dataDir).ProviderPath)

    Write-WtStore -Name 'sample' -Items @([ordered]@{ id = 'x1'; title = '한글 제목'; tags = @() })
    $one = Get-WtStoreItems -Name 'sample'
    Assert-Wt '항목 하나도 배열로 읽힌다' ($one -is [array] -and $one.Count -eq 1 -and $one[0].title -eq '한글 제목')
    Assert-Wt 'JSON에 한글이 그대로 저장된다' ((Get-Content (Get-WtStorePath 'sample') -Raw -Encoding utf8) -match '한글 제목')
    Write-WtStore -Name 'sample' -Items @()
    Assert-Wt '빈 목록 저장 · 읽기' ((Get-WtStoreItems -Name 'sample').Count -eq 0)
    Assert-Wt '덮어쓸 때 .bak을 남긴다' (Test-Path ((Get-WtStorePath 'sample') + '.bak'))
    Assert-Wt '없는 저장소는 빈 배열' ((Get-WtStoreItems -Name 'nothing').Count -eq 0)

    Write-WtStore -Name 'mig' -Schema 1 -Items @([ordered]@{ id = 'm1' })
    $doc = Read-WtStore -Name 'mig' -Schema 2 -Migrate {
        param($d, $from)
        foreach ($i in $d.items) { $i | Add-Member -NotePropertyName extra -NotePropertyValue "from$from" -Force }
        $d
    }
    Assert-Wt '형식 변환: schema 1 → 2' ($doc.schema -eq 2 -and $doc.items[0].extra -eq 'from1')
    Assert-Wt '형식 변환 전 원본을 남긴다' (Test-Path ((Get-WtStorePath 'mig') + '.schema1.bak'))
    $threw = $false
    try { [void](Read-WtStore -Name 'mig' -Schema 1) } catch { $threw = $true }
    Assert-Wt '더 새 형식 파일은 읽지 않고 멈춘다' $threw

    # ---- 설정 ----
    Initialize-WtSettings
    Assert-Wt '설정 기본값' ((Get-WtSetting 'theme.color') -eq 'sky' -and (Get-WtSetting 'ui.opacity') -eq 100)
    Register-WtSettingHandler { param($k, $v) $script:seenKey = $k }
    Set-WtSetting 'theme.color' 'pink'
    Set-WtSetting 'addons.disabled' @('memo')
    Assert-Wt '설정이 바뀌면 처리기를 부른다' ($script:seenKey -eq 'addons.disabled')
    $saved = Get-Content (Get-WtSettingsPath) -Raw | ConvertFrom-Json
    Assert-Wt '바꾼 값만 저장한다' (@($saved.values.PSObject.Properties.Name).Count -eq 2 -and $saved.values.'theme.color' -eq 'pink')
    Set-WtSetting 'theme.color' 'sky'
    $saved = Get-Content (Get-WtSettingsPath) -Raw | ConvertFrom-Json
    Assert-Wt '기본값으로 돌리면 파일에서 빠진다' (-not ($saved.values.PSObject.Properties.Name -contains 'theme.color'))
    Initialize-WtSettings
    Assert-Wt '배열 설정을 다시 읽는다' (@(Get-WtSetting 'addons.disabled') -contains 'memo')
    Assert-Wt '끈 탭은 메뉴 목록에서 빠진다' (-not (@(Get-WtAddon | ForEach-Object Id) -contains 'memo') -and (@(Get-WtAddon -All | ForEach-Object Id) -contains 'memo'))
    Set-WtSetting 'addons.disabled' @()
    Assert-Wt '메뉴 순서: 홈 · 일정 · 근태 · 메모' ((@(Get-WtAddon | ForEach-Object Id) -join ',') -eq 'home,schedule,attendance,memo')

    # ---- 예전 데이터 ----
    $legacy = (New-Item -ItemType Directory -Path (Join-Path $app 'data')).FullName
    $bom = [System.Text.UTF8Encoding]::new($true)
    [System.IO.File]::WriteAllText((Join-Path $legacy 'memo.csv'), @'
"Id","Title","Body","Created","Updated"
"a1","점검","첫 줄
둘째 줄","2026-10-01 09:00:00","2026-10-21 11:20:00"
"a2","회의 메모","장비 목록 정리","2026-10-02 10:00:00","2026-10-22 08:00:00"
'@, $bom)
    [System.IO.File]::WriteAllText((Join-Path $legacy 'schedule.csv'), @'
"Id","Date","Time","Title","Note","Alarm","Repeat","LastAlarm","Updated"
"s1","2026-10-01","09:30","주간 회의","","10","weekly","",""
"s2","2026-10-01","14:00","백업 확인","","","weekdays","",""
"s3","2026-10-23","","보안 교육","","","","",""
"s4","2026-10-20","08:00","매일 점검","","0","daily","",""
'@, $bom)
    [System.IO.File]::WriteAllText((Join-Path $legacy 'attendance.csv'), @'
"Id","Type","StartDate","EndDate","StartTime","EndTime","Place","Detail","Updated"
"t1","출근","2026-10-21","2026-10-21","08:47","","","",""
"t2","출근","2026-10-21","2026-10-21","09:00","","","",""
"t3","업무외시간 근무","2026-10-20","2026-10-20","19:00","21:30","","패치",""
"t4","업무외시간 근무","2026-10-13","2026-10-13","23:00","01:00","","야간",""
"t5","출장","2026-10-16","2026-10-17","","","데이터센터","점검",""
"t6","휴가","2026-10-17","2026-10-17","","","","오후 반차",""
"t7","휴가","2026-09-30","2026-10-02","","","","연차",""
"t8","출근","2026-09-30","2026-09-30","08:50","","","",""
'@, $bom)

    Assert-Wt '예전 data 폴더를 찾는다' ((Find-WtLegacyFolder) -eq $legacy)
    $r = Import-WtLegacyData
    Assert-Wt '처음 실행: 세 탭을 가져온다' ($r.Count -eq 14 -and $r.Tabs.Count -eq 3) "($($r.Count) / $($r.Tabs -join ', '))"
    $r2 = Import-WtLegacyData
    Assert-Wt '두 번째 실행: 이미 가져온 탭은 건너뛴다' ($r2.Count -eq 0)
    $r3 = Import-WtLegacyData -Force
    Assert-Wt '다시 가져오기는 덮어쓴다' ($r3.Count -eq 14 -and (Test-Path ((Get-WtStorePath 'memo') + '.bak')))
    Assert-Wt '예전 CSV는 그대로 둔다' ((Get-Content (Join-Path $legacy 'memo.csv') -Raw) -match '"a1","점검"')

    $memos = Get-WtMemoItems
    $a1 = $memos | Where-Object id -eq 'a1'
    Assert-Wt '메모: 여러 줄 본문 유지' ($a1.body -eq "첫 줄`n둘째 줄" -or $a1.body -eq "첫 줄`r`n둘째 줄")
    Assert-Wt '메모: 최근 고친 순서' ((@(Get-WtMemoSorted $memos)[0]).id -eq 'a2')
    Assert-Wt '메모 검색: 낱말 모두 포함' (@(Find-WtMemo $memos '장비 정리').Count -eq 1 -and @(Find-WtMemo $memos '').Count -eq 2)

    # ---- 일정 ----
    $s = Get-WtScheduleItems
    $byId = @{}; foreach ($i in $s) { $byId[$i.id] = $i }
    $d = { param($t) [datetime]::ParseExact($t, 'yyyy-MM-dd', $null) }
    Assert-Wt '일정: 매주(목) 반복' ((Test-WtScheduleOn $byId.s1 (& $d '2026-10-22')) -and -not (Test-WtScheduleOn $byId.s1 (& $d '2026-10-21')))
    Assert-Wt '일정: 평일 반복' (-not (Test-WtScheduleOn $byId.s2 (& $d '2026-10-24')) -and (Test-WtScheduleOn $byId.s2 (& $d '2026-10-26')))
    Assert-Wt '일정: 반복 없음' ((Test-WtScheduleOn $byId.s3 (& $d '2026-10-23')) -and -not (Test-WtScheduleOn $byId.s3 (& $d '2026-10-24')))
    Assert-Wt '일정: 매일 반복은 시작일부터' (-not (Test-WtScheduleOn $byId.s4 (& $d '2026-10-19')) -and (Test-WtScheduleOn $byId.s4 (& $d '2026-10-22')))
    $order = (@(Get-WtScheduleOn (& $d '2026-10-22') $s) | ForEach-Object id) -join ','
    Assert-Wt '일정: 하루 일정은 시각 순서' ($order -eq 's4,s1,s2') $order
    Assert-Wt '일정: 알람 표시' ((Format-WtAlarm 0) -eq '정시' -and (Format-WtAlarm 10) -eq '10분 전' -and (Format-WtAlarm 60) -eq '1시간 전' -and (Format-WtAlarm 1440) -eq '하루 전' -and $null -eq (Format-WtAlarm ''))
    Assert-Wt '일정: 기간 일정 칸이 있다' ($byId.s1.endDate -eq '2026-10-01' -and $byId.s1.PSObject.Properties.Name -contains 'tags')

    # ---- 근태 ----
    $sum = Get-WtAttendanceSummary (& $d '2026-10-01') (Get-WtAttendanceItems)
    Assert-Wt '근태: 출근 일수(같은 날 두 번은 하루)' ($sum.Work -eq 1) $sum.Work
    Assert-Wt '근태: 업무외 시간(자정 넘김 포함)' ($sum.OverMinutes -eq 270) $sum.OverMinutes
    Assert-Wt '근태: 출장 건수' ($sum.Trip -eq 1) $sum.Trip
    Assert-Wt '근태: 휴가 일수(반차 0.5, 달 경계 자름)' ($sum.Leave -eq 2.5) $sum.Leave
    Assert-Wt '근태: 시간 표시' ((Format-WtHours 270) -eq '4.5' -and (Format-WtHours 120) -eq '2')

    # ---- 테마 색 대비 ----
    $pairs = @(
        @('OnAccent', 'Accent', 4.5), @('AccentStrong', 'AccentSoft', 4.5),
        @('Text', 'WindowBg', 4.5), @('Text', 'Sidebar', 4.5), @('Text', 'AccentSoft', 4.5), @('Text', 'Hover', 4.5),
        @('SubText', 'WindowBg', 4.5), @('SubText', 'Sidebar', 4.5), @('SubText', 'Card', 4.5),
        @('NavText', 'Sidebar', 4.5), @('NavText', 'Hover', 4.5),
        @('InputBorder', 'Card', 3.0), @('Accent', 'Card', 3.0))
    foreach ($name in (Get-WtPalette).Keys) {
        $c = Get-WtThemeColors $name
        $bad = foreach ($p in $pairs) {
            $ratio = Get-WtContrast $c[$p[0]] $c[$p[1]]
            if ($ratio -lt $p[2]) { '{0}/{1} {2:N2}' -f $p[0], $p[1], $ratio }
        }
        Assert-Wt "테마 '$name' 색 대비" (-not $bad) ($bad -join ', ')
    }
    foreach ($k in 'Work', 'Over', 'Trip', 'Leave') {
        $c = Get-WtThemeColors 'sky'
        Assert-Wt "근태 표시 '$k' 색 대비" ((Get-WtContrast $c["Att.$k"] $c["Att.${k}Bg"]) -ge 4.5)
    }
} finally {
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}

if ($script:failed) { Write-Host "실패 $script:failed 건" -ForegroundColor Red; exit 1 }
Write-Host '모두 통과' -ForegroundColor Green
