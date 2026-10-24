<#
.SYNOPSIS
    메모 탭을 임시 데이터로 검사한다: 예전 데이터 가져오기 · 검색 · 쓰기 · 화면 · 저장 실패.

.DESCRIPTION
    tests\Test-Core.ps1이 찾아서 돌린다. 혼자 돌려도 된다. 실제 데이터 폴더(문서\WorkTool)는 쓰지 않는다.
    화면 · 저장 실패 검사는 Windows STA에서만 한다.

.EXAMPLE
    pwsh -NoProfile -STA -File apps/WorkTool/wpf/addons/memo/Test-Memo.ps1
#>
#Requires -Version 7.0
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) 'tests/TestKit.ps1')
Import-WtTestCore
Import-Module (Join-Path $PSScriptRoot 'Memo.psm1') -Force -DisableNameChecking
$memoModule = Get-Module Memo; $shellModule = Get-Module Shell
$test = New-WtTestData 'memo'
$held = $null

try {
    # ---- 예전 데이터 ----
    Write-WtTestLegacyCsv $test 'memo' @'
"Id","Title","Body","Created","Updated"
"a1","점검","첫 줄
둘째 줄","2026-10-01 09:00:00","2026-10-21 11:20:00"
"a2","회의 메모","장비 목록 정리","2026-10-02 10:00:00","2026-10-22 08:00:00"
'@
    $r = Import-WtLegacyData
    Assert-WtTest '메모: 예전 데이터를 가져온다' ($r.Count -eq 2 -and $r.Tabs.Count -eq 1) "($($r.Count) / $($r.Tabs -join ', '))"
    Assert-WtTest '메모: 예전 CSV는 그대로 둔다' ((Get-Content (Join-Path $test.Legacy 'memo.csv') -Raw) -match '"a1","점검"')
    $memos = Get-WtMemoItems
    $a1 = $memos | Where-Object id -eq 'a1'
    Assert-WtTest '메모: 여러 줄 본문 유지' ($a1.body -eq "첫 줄`n둘째 줄" -or $a1.body -eq "첫 줄`r`n둘째 줄")
    Assert-WtTest '메모: 최근 고친 순서' ((@(Get-WtMemoSorted $memos)[0]).id -eq 'a2')
    Assert-WtTest '메모 검색: 낱말 모두 포함' (@(Find-WtMemo $memos '장비 정리').Count -eq 1 -and @(Find-WtMemo $memos '').Count -eq 2)

    # ---- 쓰기 ----
    $when = [datetime]'2026-10-22 11:20:05'
    $nm = New-WtMemo $when
    Assert-WtTest '메모: 새 메모' ($nm.id.Length -eq 12 -and $nm.created -eq '2026-10-22 11:20:05' -and $nm.body -eq '')
    [void](Set-WtMemoContent $nm '  ' @'
점검 #서버 #서버
```
#코드안
```
- 할 일 #주간
'@ $when)
    Assert-WtTest '메모: 빈 제목은 저장 시각' ($nm.title -eq '2026-10-22 11:20')
    Assert-WtTest '메모: 본문 #태그를 모은다 (코드 안 · 중복 제외)' ((@($nm.tags) -join ',') -eq '서버,주간') (@($nm.tags) -join ',')
    $legacyMemo = $memos | Where-Object id -eq 'a2'
    $legacyMemo.PSObject.Properties.Remove('tags')
    [void](Set-WtMemoContent $legacyMemo '회의 메모' '고침 #회의' $when)
    Save-WtMemoItems @($memos + $nm)
    $re = Get-WtMemoItems
    Assert-WtTest '메모: 고친 것 저장 · 다시 읽기 (tags 없던 예전 항목 포함)' ($re.Count -eq 3 -and ($re | Where-Object id -eq 'a2').body -eq '고침 #회의' -and @(($re | Where-Object id -eq 'a2').tags)[0] -eq '회의')

    if (-not (Test-WtTestWpf)) { Write-Host 'SKIP 메모 화면 · 저장 실패 (Windows STA에서만)'; return }
    $addon = Get-WtAddon -Id 'memo'

    # ---- 화면: 새 메모 → 입력 → 저장 → 삭제 → 실행 취소 ----
    $view = New-WtTestView 'memo'
    [void](& $addon.Refresh $view)
    $before = (Get-WtMemoItems).Count
    [void](& $memoModule { Add-WtMemo; $script:ui.TitleBox.Text = '새 제목'; $script:ui.BodyBox.Text = "본문 #새태그"; Save-WtMemoNow })
    $saved = Get-WtMemoItems
    $new = $saved | Where-Object title -eq '새 제목'
    Assert-WtTest 'WPF 메모: 새 메모를 쓰고 저장' ($saved.Count -eq $before + 1 -and $new.body -eq '본문 #새태그' -and @($new.tags)[0] -eq '새태그')
    Assert-WtTest 'WPF 메모: 저장 상태 표시 · 목록 맨 위' ((& $memoModule { $script:ui.Status.Visibility }) -eq 'Visible' -and (& $memoModule { [string]$script:ui.List.Items[0].Id }) -eq $new.id)
    & $memoModule { Remove-WtMemoCurrent }
    Assert-WtTest 'WPF 메모: 삭제' ((Get-WtMemoItems).Count -eq $before -and (& $memoModule { $script:ui.UndoBar.Visibility }) -eq 'Visible')
    & $memoModule { Restore-WtMemoDeleted }
    Assert-WtTest 'WPF 메모: 삭제 실행 취소' ((Get-WtMemoItems).Count -eq $before + 1 -and (& $memoModule { $script:currentId }) -eq $new.id)
    & $memoModule { $script:ui.BodyBox.Text = '고치는 중' }
    & $addon.OnExit
    Assert-WtTest 'WPF 메모: 끝내기 전에 저장' (((Get-WtMemoItems) | Where-Object id -eq $new.id).body -eq '고치는 중')

    # ---- 저장 실패: 작성 중인 글을 버리지 않는다 ----
    $one = New-WtMemo; $one.title = 'A'; $one.body = 'saved A'
    $two = New-WtMemo; $two.title = 'B'; $two.body = 'saved B'
    Save-WtMemoItems @($one, $two)
    $view = New-WtTestView 'memo'
    & $addon.Refresh $view
    & $memoModule { param($id) Select-WtMemoItem $id; $script:ui.BodyBox.Text = 'unsaved draft' } $one.id
    # UI 오류창 · 확인창만 바꾼다. 저장 · 전환 로직은 실제 함수를 실행한다. 확인창의 대답은 차례로 정해 두고, 없으면 [취소]
    & $memoModule { function script:Show-WtError { param($ErrorRecord, $Action) $script:lastTestError = $Action } }
    & $shellModule {
        function script:Show-WtMessage { param($Text, $Title, $Kind) }
        function script:Show-WtError { param($ErrorRecord, $Action) }
        $script:testChoices = [System.Collections.Generic.Queue[string]]::new(); $script:testAsked = @()
        function script:Show-WtPendingChoice { param($Purpose, $Title) $script:testAsked += $Purpose; if ($script:testChoices.Count) { $script:testChoices.Dequeue() } else { 'cancel' } }
    }
    $memoPath = Get-WtStorePath 'memo'
    $original = [System.IO.File]::ReadAllText($memoPath)
    $held = [System.IO.File]::Open($memoPath, 'Open', 'Read', 'Read')
    Assert-WtTest '메모 안정성: 잠긴 파일 저장은 실패를 돌려준다' (-not (& $memoModule { Save-WtMemoNow }))
    & $memoModule { Update-WtMemoView }
    Assert-WtTest '메모 안정성: 실패 후 새로 고침해도 작성 중인 본문·dirty를 유지한다' ((& $memoModule { $script:ui.BodyBox.Text }) -eq 'unsaved draft' -and (& $memoModule { $script:dirty }))
    & $memoModule { param($id) Select-WtMemoItem $id } $two.id
    Assert-WtTest '메모 안정성: 실패 후 다른 메모 선택을 취소한다' ((& $memoModule { $script:currentId }) -eq $one.id -and (& $memoModule { $script:ui.List.SelectedItem.Id }) -eq $one.id)
    & $memoModule { $script:ui.Search.Text = 'B'; Add-WtMemo; Remove-WtMemoCurrent }
    Assert-WtTest '메모 안정성: 실패 후 검색·추가·삭제가 초안을 버리지 않는다' ((& $memoModule { $script:ui.BodyBox.Text }) -eq 'unsaved draft' -and (Get-WtMemoItems).Count -eq 2)

    # 셸과 맞물리는 곳: 저장 실패를 셸이 알아듣고 묻는다 (묻는 흐름 자체는 tests\Test-Reliability.ps1)
    $asked = & $shellModule { $script:testAsked = @(); [void](Save-WtPendingChanges -Purpose exit); $script:testAsked -join ',' }
    Assert-WtTest '메모 안정성: 저장 실패를 셸이 끝내기 전에 묻는다' ($asked -eq 'exit')
    $moved = & $shellModule { $script:testChoices.Enqueue('proceed'); Save-WtPendingChanges -Purpose switch -Only 'memo' }
    Assert-WtTest '메모 안정성: [이동]을 골라도 저장하지 못한 초안을 메모 화면에 둔다' ($moved -and (& $memoModule { $script:ui.BodyBox.Text }) -eq 'unsaved draft' -and (& $memoModule { $script:dirty }))
    Assert-WtTest '메모 안정성: 저장 실패로 원본 파일을 바꾸지 않는다' ([System.IO.File]::ReadAllText($memoPath) -ceq $original)

    $held.Dispose(); $held = $null
    Assert-WtTest '메모 안정성: 잠금 해제 후 같은 초안을 저장할 수 있다' (& $memoModule { Save-WtMemoNow })
    Assert-WtTest '메모 안정성: 재시도한 본문을 다시 읽을 수 있다' (((Get-WtMemoItems) | Where-Object id -eq $one.id).body -eq 'unsaved draft')
    & $memoModule { param($id) $script:ui.Search.Text = ''; $script:ui.BodyBox.Text = 'next draft'; Select-WtMemoItem $id } $two.id
    Assert-WtTest '메모 안정성: 자동 저장 뒤 목록 선택과 편집기 메모가 일치한다' ((& $memoModule { $script:currentId }) -eq $two.id -and (& $memoModule { $script:ui.List.SelectedItem.Id }) -eq $two.id)
    & $memoModule { param($id) Select-WtMemoItem $id } $one.id

    # 실패한 삭제 · 실행 취소를 여러 번 해도 배열이 먼저 바뀌지 않는다
    & $memoModule { $script:ui.Search.Text = '' }
    $held = [System.IO.File]::Open($memoPath, 'Open', 'Read', 'Read')
    & $memoModule { Remove-WtMemoCurrent; Remove-WtMemoCurrent; Add-WtMemo }
    Assert-WtTest '메모 안정성: 삭제·추가 실패 시 메모리 배열을 유지한다' ((& $memoModule { $script:items.Count }) -eq 2)
    $held.Dispose(); $held = $null
    & $memoModule { Remove-WtMemoCurrent }
    $held = [System.IO.File]::Open($memoPath, 'Open', 'Read', 'Read')
    & $memoModule { Restore-WtMemoDeleted; Restore-WtMemoDeleted }
    Assert-WtTest '메모 안정성: 실행 취소 실패를 재시도해도 중복 항목을 만들지 않는다' ((& $memoModule { $script:items.Count }) -eq 1)
    $held.Dispose(); $held = $null
    & $memoModule { Restore-WtMemoDeleted }
    Assert-WtTest '메모 안정성: 실행 취소 성공 뒤 ID가 모두 고유하다' ((Get-WtMemoItems).Count -eq 2 -and @((Get-WtMemoItems).id | Sort-Object -Unique).Count -eq 2)
} finally {
    if ($held) { $held.Dispose() }
    & $memoModule { if ($script:saveTimer) { $script:saveTimer.Stop() }; if ($script:undoTimer) { $script:undoTimer.Stop() } }
    Remove-WtTestData $test
}
