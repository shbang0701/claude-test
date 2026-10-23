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
foreach ($n in 'Storage', 'Settings', 'Theme', 'Ui', 'Highlight', 'Markdown', 'Mermaid', 'Addons', 'Pages', 'Shell') {
    Import-Module (Join-Path $root "core/$n.psm1") -Force -DisableNameChecking
}
foreach ($f in (Get-ChildItem (Join-Path $root 'addons') -Filter '*.psm1' -Recurse | Sort-Object FullName)) {
    Import-Module $f.FullName -Force -DisableNameChecking
}

$script:failed = 0
function Assert-Wt([string]$Name, [bool]$Ok, $Detail = '') {
    if ($Ok) { Write-Host "PASS $Name" } else { Write-Host "FAIL $Name $Detail" -ForegroundColor Red; $script:failed++ }
}

function Wait-WtDocTest {
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $docsModule = Get-Module Docs
    do {
        & $docsModule { Complete-WtDocJobs; Complete-WtDocTreeBatch }
        if ($watch.Elapsed.TotalSeconds -gt 15) { throw '문서 테스트 조회 시간 초과' }
        Start-Sleep -Milliseconds 10
    } while (& $docsModule { $script:jobs.Count -or $script:pending.Count -or $script:treeBatches.Count })
}

$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('wt-test-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
$app = (New-Item -ItemType Directory -Path (Join-Path $tmp 'app')).FullName
$dataDir = Join-Path $tmp 'data'
Set-Content -LiteralPath (Join-Path $app 'datapath.txt') -Value $dataDir -Encoding utf8
Set-WtAppRoot $app

try {
    # ---- 코드 검사 ----
    # PowerShell 변수는 대소문자를 가리지 않는다. 함수 안에서 $title과 [string]$Title처럼 대소문자만 다른 이름을 쓰면 같은 변수가 된다
    $clash = foreach ($f in (Get-ChildItem $root -Recurse -Include *.psm1, *.ps1)) {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$null, [ref]$null)
        foreach ($fn in $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
            $names = @($fn.Body.FindAll({ $args[0] -is [System.Management.Automation.Language.VariableExpressionAst] }, $false) |
                    ForEach-Object { $_.VariablePath.UserPath } | Where-Object { $_ -notmatch ':' })
            $plist = if ($fn.Parameters) { $fn.Parameters } elseif ($fn.Body.ParamBlock) { $fn.Body.ParamBlock.Parameters } else { @() }
            $names += @($plist | ForEach-Object { $_.Name.VariablePath.UserPath })
            $names | Group-Object { $_.ToLowerInvariant() } | Where-Object { @($_.Group | Sort-Object -Unique -CaseSensitive).Count -gt 1 } |
                ForEach-Object { "$($f.Name) $($fn.Name): $(@($_.Group | Sort-Object -Unique -CaseSensitive) -join '/')" }
        }
    }
    Assert-Wt '코드 검사: 함수 안에 대소문자만 다른 변수 이름 없음' (-not $clash) ($clash -join '; ')

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
    Assert-Wt '메뉴 순서: 홈 · 일정 · 근태 · 메모 · 문서 · 활성 작업' ((@(Get-WtAddon | ForEach-Object Id) -join ',') -eq 'home,schedule,attendance,memo,docs,active-tasks')

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

    # ---- 메모 쓰기 ----
    $when = [datetime]'2026-10-22 11:20:05'
    $nm = New-WtMemo $when
    Assert-Wt '메모: 새 메모' ($nm.id.Length -eq 12 -and $nm.created -eq '2026-10-22 11:20:05' -and $nm.body -eq '')
    [void](Set-WtMemoContent $nm '  ' @'
점검 #서버 #서버
```
#코드안
```
- 할 일 #주간
'@ $when)
    Assert-Wt '메모: 빈 제목은 저장 시각' ($nm.title -eq '2026-10-22 11:20')
    Assert-Wt '메모: 본문 #태그를 모은다 (코드 안 · 중복 제외)' ((@($nm.tags) -join ',') -eq '서버,주간') (@($nm.tags) -join ',')
    $legacyMemo = $memos | Where-Object id -eq 'a2'
    $legacyMemo.PSObject.Properties.Remove('tags')
    [void](Set-WtMemoContent $legacyMemo '회의 메모' '고침 #회의' $when)
    Save-WtMemoItems @($memos + $nm)
    $re = Get-WtMemoItems
    Assert-Wt '메모: 고친 것 저장 · 다시 읽기 (tags 없던 예전 항목 포함)' ($re.Count -eq 3 -and ($re | Where-Object id -eq 'a2').body -eq '고침 #회의' -and @(($re | Where-Object id -eq 'a2').tags)[0] -eq '회의')

    # ---- 문서 ----
    $docRoot = (New-Item -ItemType Directory -Path (Join-Path $tmp '문서 루트')).FullName
    New-Item -ItemType Directory -Path (Join-Path $docRoot '운영\점검') -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $docRoot '가나다') | Out-Null
    Set-Content -LiteralPath (Join-Path $docRoot '계정 목록.xlsx') -Value 'x'
    Set-Content -LiteralPath (Join-Path $docRoot 'readme.md') -Value ('a' * 3000)
    $hid = New-Item -ItemType File -Path (Join-Path $docRoot 'secret.tmp')
    $hid.Attributes = 'Hidden'
    $de = Get-WtDocEntries $docRoot
    Assert-Wt '문서: 폴더 먼저 · 이름 순 · 숨김 빼기' ((($de | ForEach-Object Name) -join ',') -eq '가나다,운영,계정 목록.xlsx,readme.md') (($de | ForEach-Object Name) -join ',')
    Assert-Wt '문서: 종류 · 크기' (($de | Where-Object Name -eq '계정 목록.xlsx').Kind -eq 'Excel' -and (Format-WtDocSize 3002) -eq '3 KB' -and (Format-WtDocSize 5MB) -eq '5.0 MB' -and (Get-WtDocKind 'a.abc' $false) -eq 'ABC')
    Assert-Wt '문서: 하위 폴더' ((@(Get-WtDocSubfolders $docRoot) | ForEach-Object { Split-Path $_ -Leaf }) -join ',' -eq '가나다,운영')
    $chain = Get-WtDocChain @('D:\x', $docRoot) (Join-Path $docRoot '운영\점검')
    Assert-Wt '문서: 맨 위 폴더부터 경로 차례' ($chain.Count -eq 3 -and $chain[0] -eq $docRoot -and $chain[2] -like '*\운영\점검') ($chain -join ' > ')
    Assert-Wt '문서: 맨 위 폴더 밖은 빈 차례' ((Get-WtDocChain @($docRoot) 'Z:\other').Count -eq 0)
    Assert-Wt '문서: 경로 표시' ((Format-WtDocPathLabel @($docRoot) (Join-Path $docRoot '운영\점검')) -eq '문서 루트 › 운영 › 점검' -and (Format-WtDocPathLabel @($docRoot) 'Z:\a') -eq 'Z:\a')
    Assert-Wt '문서: 이름 거르기' (@(Find-WtDocEntry $de '계정 목록').Count -eq 1 -and @(Find-WtDocEntry $de '').Count -eq 4)
    $xl = Join-Path $docRoot '계정 목록.xlsx'
    Assert-Wt '문서: 즐겨찾기 넣기 · 빼기' ((Switch-WtDocPin $xl) -and (Test-WtDocPinned $xl) -and -not (Switch-WtDocPin $xl) -and -not (Test-WtDocPinned $xl))
    [void](Switch-WtDocPin 'Z:\없는\파일.xlsx')
    Assert-Wt '문서: 없어진 즐겨찾기도 보여 준다' (@(Get-WtDocPinEntries | Where-Object Missing).Count -eq 1)
    Set-WtSetting 'docs.pins' @()

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
    $c = Get-WtThemeColors 'sky'
    $bad = foreach ($k in 'Blue', 'Teal', 'Green', 'Orange', 'Red', 'Purple', 'Gray') {
        foreach ($fg in "Callout.$k", 'Text') {
            $ratio = Get-WtContrast $c[$fg] $c["Callout.${k}Bg"]
            if ($ratio -lt 4.5) { '{0}/{1} {2:N2}' -f $fg, $k, $ratio }
        }
    }
    Assert-Wt '콜아웃 색 대비' (-not $bad) ($bad -join ', ')
    $bad = foreach ($k in 'Comment', 'String', 'Keyword', 'Number', 'Command', 'Variable', 'Type', 'Attr', 'Inserted', 'Deleted', 'Text') {
        $fg = if ($k -eq 'Text') { 'Text' } else { "Code.$k" }
        $ratio = Get-WtContrast $c[$fg] $c['Code.Bg']
        if ($ratio -lt 4.5) { '{0} {1:N2}' -f $fg, $ratio }
    }
    Assert-Wt '코드 색 대비 (코드 바탕 위)' (-not $bad) ($bad -join ', ')
    Assert-Wt '마크다운 · 다이어그램 색 대비' ((Get-WtContrast $c.Text $c['Md.Mark']) -ge 4.5 -and (Get-WtContrast $c.Text $c['Dia.NodeBg']) -ge 4.5 -and
        (Get-WtContrast $c.Text $c['Dia.NoteBg']) -ge 4.5 -and (Get-WtContrast $c.SubText $c['Dia.GroupBg']) -ge 4.5 -and
        (Get-WtContrast $c['Dia.Line'] $c.Card) -ge 3 -and (Get-WtContrast $c['Dia.NodeLine'] $c.Card) -ge 3)

    # ---- 마크다운 ----
    $md = @'
---
tags: [a, b]
---
# 제목 **굵게**

본문 [[다른 메모|별칭]] #태그 #123 %%숨김%% [링크](https://example.com) [파일](docs/a%20b.md)
둘째 줄

> [!warning]- 조심 **해**
> 본문

> [!NOTE]
> 깃허브 알림

> 그냥 인용

- [x] 끝남
- [ ] 할 일

| a | b |
|:--|--:|
| 1 | 2 |

```ps1
Get-Item C:\
```

각주[^1]

[^1]: 내용
'@
    $mm = ConvertTo-WtMarkdownModel $md
    $types = (@($mm.Blocks) | ForEach-Object T) -join ','
    Assert-Wt '마크다운: 블록 종류' ($types -eq 'props,heading,para,callout,callout,quote,list,table,code,para,footnotes') $types
    Assert-Wt '마크다운: 속성(--- 사이)' ($mm.Blocks[0].Items[0].Key -eq 'tags' -and $mm.Blocks[0].Items[0].Value -eq 'a, b')
    $p = $mm.Blocks[2].Inl
    $wiki = $p | Where-Object { $_.T -eq 'link' -and $_.Kind -eq 'wiki' }
    Assert-Wt '마크다운: 위키 링크' ($wiki.Target -eq '다른 메모' -and $wiki.Inl[0].Text -eq '별칭')
    Assert-Wt '마크다운: 태그 (숫자만은 태그 아님)' (@($p | Where-Object T -eq 'tag').Count -eq 1 -and ($p | Where-Object T -eq 'tag').Text -eq '태그')
    Assert-Wt '마크다운: %%주석%%은 빠진다' (-not ((Get-WtMarkdownSnippet '가 %%숨김%% 나') -match '숨김') -and -not (@($p | Where-Object { $_.T -eq 'text' -and $_.Text -match '숨김' }).Count))
    $links = @($p | Where-Object { $_.T -eq 'link' -and $_.Kind -ne 'wiki' })
    Assert-Wt '마크다운: 링크 종류 (주소 · 파일)' ($links[0].Kind -eq 'url' -and $links[1].Kind -eq 'file' -and $links[1].Target -eq 'docs/a b.md')
    Assert-Wt '마크다운: 줄바꿈 한 번' (@($p | Where-Object T -eq 'soft').Count -eq 1)
    $co = $mm.Blocks[3]
    Assert-Wt '마크다운: 콜아웃 종류 · 접기 · 제목' ($co.Kind -eq 'warning' -and $co.Fold -eq '-' -and (@($co.Title | ForEach-Object { if ($_.Inl) { $_.Inl[0].Text } else { $_.Text } }) -join '') -eq '조심 해' -and $co.Blocks.Count -eq 1)
    Assert-Wt '마크다운: 제목 없는 콜아웃은 종류 이름' ($mm.Blocks[4].Kind -eq 'note' -and $mm.Blocks[4].Title[0].Text -eq 'Note')
    Assert-Wt '마크다운: 할 일 목록' ($mm.Blocks[6].Items[0].Task -eq $true -and $mm.Blocks[6].Items[1].Task -eq $false -and $mm.Blocks[6].Items[0].Blocks[0].Inl[0].Text -eq '끝남')
    Assert-Wt '마크다운: 표 정렬' (($mm.Blocks[7].Align -join ',') -eq 'left,right' -and $mm.Blocks[7].Rows[0].Header)
    Assert-Wt '마크다운: 코드 블록' ($mm.Blocks[8].Lang -eq 'ps1' -and $mm.Blocks[8].Text -eq 'Get-Item C:\')
    Assert-Wt '마크다운: 각주' ($mm.Blocks[9].Inl[1].T -eq 'fn' -and $mm.Blocks[10].Items[0].Order -eq 1)
    Assert-Wt '마크다운: 제목 목록' ($mm.Headings.Count -eq 1 -and $mm.Headings[0].Text -eq '제목 굵게')
    $snip = Get-WtMarkdownSnippet "# 점검`n- [ ] 디스크`n- **로그** 확인`n| a | b |`n|--|--|`n[[메모|별칭]]"
    Assert-Wt '마크다운: 목록용 한 줄 요약' ($snip -eq '점검 · 디스크 · 로그 확인 · a  b · 별칭') $snip
    Assert-Wt '마크다운: 빈 글' ((ConvertTo-WtMarkdownModel '').Blocks.Count -eq 0 -and (Get-WtMarkdownSnippet '') -eq '')

    # ---- 코드 색칠 ----
    $tok = Get-WtCodeTokens 'Get-ChildItem -Path $env:TEMP | Where-Object { $_.Length -gt 1KB } # 주석' 'powershell'
    $kinds = (@($tok) | Where-Object Kind | ForEach-Object { "$($_.Kind):$($_.Text)" }) -join ' '
    Assert-Wt '코드 색칠: PowerShell' ($kinds -eq 'Command:Get-ChildItem Attr:-Path Variable:$env:TEMP Command:Where-Object Variable:$_ Keyword:-gt Number:1KB Comment:# 주석') $kinds
    Assert-Wt '코드 색칠: 조각을 이으면 원래 글' ((@($tok) | ForEach-Object Text) -join '' -eq 'Get-ChildItem -Path $env:TEMP | Where-Object { $_.Length -gt 1KB } # 주석')
    Assert-Wt '코드 색칠: 별칭 · 모르는 언어' ((Get-WtCodeLanguage 'PS1') -eq 'powershell' -and (Get-WtCodeLanguage 'sh') -eq 'bash' -and $null -eq (Get-WtCodeLanguage 'nope') -and @(Get-WtCodeTokens 'x' 'nope')[0].Kind -eq '')
    $str = @(Get-WtCodeTokens 'echo "a # b" # c' 'bash') | Where-Object Kind
    Assert-Wt '코드 색칠: 문자열 안의 #은 주석이 아님' (($str | ForEach-Object Kind) -join ',' -eq 'Command,String,Comment')

    # ---- Mermaid ----
    $fc = ConvertFrom-WtMermaid @'
flowchart LR
    A[시작] --> B{확인?}
    B -->|예| C[(DB)]
    B -- 아니오 --> D([끝])
    C -.-> D & A
    subgraph g1 [묶음]
      C
    end
'@
    Assert-Wt 'Mermaid: flowchart 노드 · 모양' ($fc.Dir -eq 'LR' -and (($fc.Nodes.Values | ForEach-Object Shape) -join ',') -eq 'rect,diamond,cylinder,stadium' -and $fc.Nodes['C'].Group -eq 'g1')
    $el = ($fc.Edges | ForEach-Object { "$($_.From)$($_.To):$($_.Style):$($_.Text)" }) -join ' '
    Assert-Wt 'Mermaid: flowchart 선 · 라벨 · & 묶음' ($el -eq 'AB:normal: BC:normal:예 BD:normal:아니오 CD:dotted: CA:dotted:') $el
    $lay = Get-WtMermaidLayout $fc
    $nodes = @($lay.Items | Where-Object K -eq 'node')
    $overlap = $false
    for ($i = 0; $i -lt $nodes.Count; $i++) {
        for ($j = $i + 1; $j -lt $nodes.Count; $j++) {
            $a = $nodes[$i]; $b = $nodes[$j]
            if ($a.X -lt $b.X + $b.W -and $b.X -lt $a.X + $a.W -and $a.Y -lt $b.Y + $b.H -and $b.Y -lt $a.Y + $a.H) { $overlap = $true }
        }
    }
    Assert-Wt 'Mermaid: 배치 (노드 4개, 겹치지 않음, 그림 안)' ($nodes.Count -eq 4 -and -not $overlap -and -not @($nodes | Where-Object { $_.X -lt 0 -or $_.Y -lt 0 -or $_.X + $_.W -gt $lay.Width -or $_.Y + $_.H -gt $lay.Height }).Count)
    $ax = ($nodes | Where-Object Text -eq '시작').X; $bx = ($nodes | Where-Object Text -eq '확인?').X
    Assert-Wt 'Mermaid: LR은 왼쪽에서 오른쪽' ($ax -lt $bx)
    $sq = ConvertFrom-WtMermaid "sequenceDiagram`n participant A as 서버`n actor U`n U->>A: 요청`n A-->>U: 응답`n loop 매일`n A->>A: 점검`n end`n Note right of A: 메모"
    Assert-Wt 'Mermaid: sequence' ((($sq.Participants.Values | ForEach-Object Label) -join ',') -eq '서버,U' -and $sq.Participants['U'].Actor -and
        (($sq.Items | ForEach-Object K) -join ',') -eq 'msg,msg,start,msg,end,note' -and $sq.Items[1].Line -eq 'dashed')
    Assert-Wt 'Mermaid: sequence 배치' (@((Get-WtMermaidLayout $sq).Items | Where-Object K -eq 'line').Count -ge 5)
    $stm = ConvertFrom-WtMermaid "stateDiagram-v2`n [*] --> 대기`n 대기 --> 실행 : 시작`n 실행 --> [*]"
    Assert-Wt 'Mermaid: state ([*] · 라벨)' ($stm.Nodes.Count -eq 4 -and (($stm.Edges | ForEach-Object Text) -join ',') -eq ',시작,')
    $threw = $null
    try { [void](ConvertFrom-WtMermaid "gantt`n title x") } catch { $threw = $_.Exception }
    Assert-Wt 'Mermaid: 모르는 종류는 NotSupported' ($threw -is [System.NotSupportedException])

    # ---- WPF로 그리기 (Windows · STA에서만) ----
    if ($IsWindows -and [System.Threading.Thread]::CurrentThread.GetApartmentState() -eq 'STA') {
        Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
        $sample = $md + @'


```mermaid
flowchart TD
    A --> B{x}
    B -->|y| C
```

```mermaid
sequenceDiagram
    A->>B: hi
```

```mermaid
pie
```
'@
        $view = New-WtMarkdownView -Scroll -OnLink { param($l) $true }
        Set-WtMarkdownText $view $sample
        Assert-Wt 'WPF: 마크다운을 오류 없이 그린다' (-not $view.Tag.Error -and $view.Document.Blocks.Count -ge 12) "$($view.Tag.Error)"
        $figs = @($view.Document.Blocks | Where-Object { $_ -is [System.Windows.Documents.BlockUIContainer] -and $_.Child -is [System.Windows.Controls.Border] })
        Assert-Wt 'WPF: Mermaid 2개는 그림, 모르는 종류는 코드 블록' ($figs.Count -eq 2 -and $view.Tag.Remeasure)
        Assert-Wt 'WPF: 제목 · 각주 자리' ($view.Tag.Anchors.ContainsKey('제목굵게') -and $view.Tag.Anchors.ContainsKey('fn1'))

        # 메모 화면: 새 메모 → 입력 → 저장 → 삭제 → 실행 취소
        $styles = Get-WtStyleDictionary
        $memoAddon = Get-WtAddon -Id 'memo'
        $mv = @(& $memoAddon.NewView | Where-Object { $_ -is [System.Windows.UIElement] })[-1]
        $mv.Resources.MergedDictionaries.Add((New-WtThemeDictionary)); $mv.Resources.MergedDictionaries.Add($styles)
        [void](& $memoAddon.Refresh $mv)
        $mod = Get-Module Memo
        $before = (Get-WtMemoItems).Count
        & $mod { Add-WtMemo; $script:ui.TitleBox.Text = '새 제목'; $script:ui.BodyBox.Text = "본문 #새태그"; Save-WtMemoNow }
        $saved = Get-WtMemoItems
        $new = $saved | Where-Object title -eq '새 제목'
        Assert-Wt 'WPF 메모: 새 메모를 쓰고 저장' ($saved.Count -eq $before + 1 -and $new.body -eq '본문 #새태그' -and @($new.tags)[0] -eq '새태그')
        Assert-Wt 'WPF 메모: 저장 상태 표시 · 목록 맨 위' ((& $mod { $script:ui.Status.Visibility }) -eq 'Visible' -and (& $mod { [string]$script:ui.List.Items[0].Id }) -eq $new.id)
        & $mod { Remove-WtMemoCurrent }
        Assert-Wt 'WPF 메모: 삭제' ((Get-WtMemoItems).Count -eq $before -and (& $mod { $script:ui.UndoBar.Visibility }) -eq 'Visible')
        & $mod { Restore-WtMemoDeleted }
        Assert-Wt 'WPF 메모: 삭제 실행 취소' ((Get-WtMemoItems).Count -eq $before + 1 -and (& $mod { $script:currentId }) -eq $new.id)
        & $mod { $script:ui.BodyBox.Text = '고치는 중' }
        & $memoAddon.OnExit
        # 문서 화면: 맨 위 폴더 → 하위 폴더로 들어가기 → 즐겨찾기
        Set-WtSetting 'docs.roots' @($docRoot)
        Set-WtSetting 'docs.last' ''
        $docAddon = Get-WtAddon -Id 'docs'
        $dv = @(& $docAddon.NewView | Where-Object { $_ -is [System.Windows.UIElement] })[-1]
        $dv.Resources.MergedDictionaries.Add((New-WtThemeDictionary)); $dv.Resources.MergedDictionaries.Add($styles)
        [void](& $docAddon.Refresh $dv)
        $dmod = Get-Module Docs
        Wait-WtDocTest
        Assert-Wt 'WPF 문서: 맨 위 폴더를 연다' ((& $dmod { $script:current }) -eq $docRoot -and (& $dmod { $script:ui.List.Items.Count }) -eq 4) "(cur=$(& $dmod { $script:current }) err=$(& $dmod { $script:ui.Error }))"
        & $dmod { Open-WtDocEntry ($script:entries | Where-Object Name -eq '운영') }
        Wait-WtDocTest
        Assert-Wt 'WPF 문서: 폴더로 들어가면 트리도 따라간다' ((& $dmod { $script:current }) -like '*\운영' -and (& $dmod { $script:ui.Tree.SelectedItem.Tag.Path }) -like '*\운영' -and (& $dmod { $script:ui.PathText.Text }) -eq '문서 루트 › 운영')
        & $dmod { $script:ui.List.SelectedIndex = 0; $script:ui.Pin.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) }
        Assert-Wt 'WPF 문서: 즐겨찾기 버튼' (@(Get-WtSetting 'docs.pins').Count -eq 1 -and @(Get-WtSetting 'docs.pins')[0] -like '*\운영\점검')
        Set-WtSetting 'docs.pins' @(); Set-WtSetting 'docs.roots' @()

        $after = Get-WtMemoItems
        Assert-Wt 'WPF 메모: 끝내기 전에 저장' (($after | Where-Object id -eq $new.id).body -eq '고치는 중')
        & $dmod { Stop-WtDocs }
    } else {
        Write-Host 'SKIP WPF로 그리기 (Windows STA에서만)'
    }
} finally {
    $safe = [IO.Path]::GetFullPath($tmp)
    if ((Split-Path $safe -Leaf) -notmatch '^wt-test-[a-f0-9]{8}$' -or -not $safe.StartsWith([IO.Path]::GetFullPath([IO.Path]::GetTempPath()), [StringComparison]::OrdinalIgnoreCase)) { throw '임시 경로 검증 실패' }
    Remove-Item -LiteralPath $safe -Recurse -Force -ErrorAction SilentlyContinue
}

if ($script:failed) { Write-Host "실패 $script:failed 건" -ForegroundColor Red; exit 1 }
& (Join-Path $PSScriptRoot 'Test-ActiveTasks.ps1')
& (Join-Path $PSScriptRoot 'Test-Reliability.ps1')
& (Join-Path $PSScriptRoot 'Test-Launcher.ps1')
Write-Host '모두 통과' -ForegroundColor Green
