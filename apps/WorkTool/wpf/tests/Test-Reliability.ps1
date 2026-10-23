<#
.SYNOPSIS
    저장 실패·중복 실행·문서 조회 취소를 임시 데이터로 검사한다.
.EXAMPLE
    pwsh -NoProfile -STA -File apps/WorkTool/wpf/tests/Test-Reliability.ps1
#>
#Requires -Version 7.0
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
foreach ($name in 'Storage', 'Settings', 'Theme', 'Ui', 'Highlight', 'Markdown', 'Mermaid', 'Addons', 'Pages', 'Shell') {
    Import-Module (Join-Path $root "core/$name.psm1") -Force -DisableNameChecking
}
foreach ($name in 'memo/Memo', 'docs/Docs') { Import-Module (Join-Path $root "addons/$name.psm1") -Force -DisableNameChecking }
$tmp = Join-Path ([IO.Path]::GetTempPath()) ('wt-reliability-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($tmp)
Set-Content -LiteralPath (Join-Path $tmp 'datapath.txt') -Value (Join-Path $tmp 'data') -Encoding utf8
Set-WtAppRoot $tmp
Initialize-WtSettings
$held = $null; $instance = $null; $legacyForm = $null
$memoModule = Get-Module Memo; $docsModule = Get-Module Docs; $shellModule = Get-Module Shell

function Assert-Reliable([string]$Name, [bool]$Ok) {
    if (-not $Ok) { throw "FAIL 안정성: $Name" }
    Write-Host "PASS 안정성: $Name"
}

function Wait-DocJobs {
    $watch = [Diagnostics.Stopwatch]::StartNew()
    while (& $docsModule { $script:jobs.Count -or $script:pending.Count -or $script:treeBatches.Count }) {
        & $docsModule { Complete-WtDocJobs; Complete-WtDocTreeBatch }
        if ($watch.Elapsed.TotalSeconds -gt 15) { throw '문서 테스트 시간 초과' }
        Start-Sleep -Milliseconds 10
    }
}

try {
    # 다른 프로세스도 같은 데이터 폴더를 열 수 없어야 한다.
    $child = Join-Path $tmp 'lock-child.ps1'
    @'
param($Source, $App)
Import-Module $Source -Force
Set-WtAppRoot $App
try { $lease = Open-WtDataLock; $lease.Dispose(); exit 0 } catch { exit 23 }
'@ | Set-Content -LiteralPath $child -Encoding utf8
    $instance = Open-WtDataLock
    & pwsh -NoProfile -NonInteractive -File $child (Join-Path $root 'core/Storage.psm1') $tmp
    Assert-Reliable '다른 프로세스의 같은 데이터 폴더 실행을 막는다' ($LASTEXITCODE -eq 23)
    $instance.Dispose(); $instance = $null
    & pwsh -NoProfile -NonInteractive -File $child (Join-Path $root 'core/Storage.psm1') $tmp
    Assert-Reliable '잠금 파일이 남아도 종료 뒤 다시 실행할 수 있다' ($LASTEXITCODE -eq 0)

    Set-WtSetting 'memo.mode' 'edit'
    $settingsPath = Get-WtSettingsPath
    $held = [IO.File]::Open($settingsPath, 'Open', 'Read', 'Read')
    $threw = $false
    try { Set-WtSetting 'memo.mode' 'preview' } catch { $threw = $true }
    Assert-Reliable '설정 저장 실패 시 메모리 값도 유지한다' ($threw -and (Get-WtSetting 'memo.mode') -eq 'edit')
    $held.Dispose(); $held = $null

    if (-not $IsWindows -or [Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
        Write-Host 'SKIP 안정성 WPF/WinForms (Windows STA에서만)'
        return
    }
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml, System.Windows.Forms, System.Drawing
    $one = New-WtMemo; $one.title = 'A'; $one.body = 'saved A'
    $two = New-WtMemo; $two.title = 'B'; $two.body = 'saved B'
    Save-WtMemoItems @($one, $two)
    $addon = Get-WtAddon -Id 'memo'
    $view = @(& $addon.NewView | Where-Object { $_ -is [Windows.UIElement] })[-1]
    $view.Resources.MergedDictionaries.Add((New-WtThemeDictionary))
    $view.Resources.MergedDictionaries.Add((Get-WtStyleDictionary))
    & $addon.Refresh $view
    & $memoModule { param($id) Select-WtMemoItem $id; $script:ui.BodyBox.Text = 'unsaved draft' } $one.id
    # UI 오류창만 캡처한다. 저장·전환·종료 로직은 실제 함수를 실행한다.
    & $memoModule { function script:Show-WtError { param($ErrorRecord, $Action) $script:lastTestError = $Action } }
    & $shellModule { function script:Show-WtMessage { param($Text, $Title, $Kind) }; function script:Show-WtError { param($ErrorRecord, $Action) } }
    # 저장 실패 확인창의 대답은 차례로 정해 둔다. 정해 둔 대답이 없으면 [취소].
    & $shellModule {
        $script:testChoices = [System.Collections.Generic.Queue[string]]::new(); $script:testAsked = @()
        function script:Show-WtPendingChoice { param($Purpose, $Title) $script:testAsked += $Purpose; if ($script:testChoices.Count) { $script:testChoices.Dequeue() } else { 'cancel' } }
    }
    $memoPath = Get-WtStorePath 'memo'
    $before = [IO.File]::ReadAllText($memoPath)
    $held = [IO.File]::Open($memoPath, 'Open', 'Read', 'Read')
    Assert-Reliable '잠긴 파일 저장은 실패를 돌려준다' (-not (& $memoModule { Save-WtMemoNow }))
    & $memoModule { Update-WtMemoView }
    Assert-Reliable '실패 후 새로 고침해도 작성 중인 본문·dirty를 유지한다' ((& $memoModule { $script:ui.BodyBox.Text }) -eq 'unsaved draft' -and (& $memoModule { $script:dirty }))
    & $memoModule { param($id) Select-WtMemoItem $id } $two.id
    Assert-Reliable '실패 후 다른 메모 선택을 취소한다' ((& $memoModule { $script:currentId }) -eq $one.id -and (& $memoModule { $script:ui.List.SelectedItem.Id }) -eq $one.id)
    & $memoModule { $script:ui.Search.Text = 'B'; Add-WtMemo; Remove-WtMemoCurrent }
    Assert-Reliable '실패 후 검색·추가·삭제가 초안을 버리지 않는다' ((& $memoModule { $script:ui.BodyBox.Text }) -eq 'unsaved draft' -and (Get-WtMemoItems).Count -eq 2)
    $exitCounter = [pscustomobject]@{ Count = 0 }
    Register-WtAddon -Id 'test-cleanup' -Title 'test' -Order 1 -NewView { } -OnExit { $exitCounter.Count++ }.GetNewClosure()
    $allowed = & $shellModule { $script:frame = [pscustomobject]@{ Continue = $true }; Stop-WtApp }
    Assert-Reliable '저장 실패는 끝낼지 묻고, [취소]면 정리 처리 전에 종료를 멈춘다' (-not $allowed -and $exitCounter.Count -eq 0 -and (& $shellModule { $script:frame.Continue }) -and (& $shellModule { $script:testAsked -join ',' }) -eq 'exit')
    $allowed = & $shellModule { $script:testAsked = @(); $script:testChoices.Enqueue('retry'); $script:testChoices.Enqueue('cancel'); Stop-WtApp }
    Assert-Reliable '[다시 시도]는 다시 저장해 보고 실패하면 또 묻는다' (-not $allowed -and (& $shellModule { $script:testAsked.Count }) -eq 2 -and $exitCounter.Count -eq 0)
    & $shellModule {
        $script:current = 'memo'; $script:testAsked = @()
        $script:navButtons = @{ memo = @{Button=[Windows.Controls.RadioButton]::new()}; docs = @{Button=[Windows.Controls.RadioButton]::new()} }
        $script:navButtons.docs.Button.IsChecked = $true
    }
    Select-WtPage '_help'
    Assert-Reliable '저장 실패는 이동할지 묻고, [취소]면 페이지 전환을 멈춘다' ((& $shellModule { $script:current }) -ne '_help' -and (& $shellModule { $script:testAsked -join ',' }) -eq 'switch')
    Assert-Reliable '전환을 취소하면 메뉴 선택도 원래 화면으로 돌아간다' ((& $shellModule { $script:navButtons.memo.Button.IsChecked }) -and -not (& $shellModule { $script:navButtons.docs.Button.IsChecked }))
    & $shellModule { $script:ui = @{ PageHost = [Windows.Controls.ContentControl]::new() }; $script:testChoices.Enqueue('proceed') }
    Select-WtPage '_help'
    Assert-Reliable '[이동]을 고르면 저장하지 못한 초안을 메모 화면에 둔 채 옮긴다' ((& $shellModule { $script:current }) -eq '_help' -and (& $memoModule { $script:ui.BodyBox.Text }) -eq 'unsaved draft' -and (& $memoModule { $script:dirty }))
    $askedBefore = & $shellModule { $script:testAsked.Count }
    Assert-Reliable '저장할 것이 없는 화면을 떠날 때는 묻지 않는다' ((& $shellModule { Save-WtPendingChanges -Purpose switch -Only '_help' }) -and (& $shellModule { $script:testAsked.Count }) -eq $askedBefore)
    & $shellModule { $script:current = 'memo'; $script:ui = $null }
    Assert-Reliable '저장 실패로 원본 파일을 바꾸지 않는다' ([IO.File]::ReadAllText($memoPath) -ceq $before)
    $held.Dispose(); $held = $null
    Assert-Reliable '잠금 해제 후 같은 초안을 저장할 수 있다' (& $memoModule { Save-WtMemoNow })
    Assert-Reliable '재시도한 본문을 다시 읽을 수 있다' (((Get-WtMemoItems) | Where-Object id -eq $one.id).body -eq 'unsaved draft')
    & $memoModule { param($id) $script:ui.Search.Text = ''; $script:ui.BodyBox.Text = 'next draft'; Select-WtMemoItem $id } $two.id
    Assert-Reliable '자동 저장 뒤 목록 선택과 편집기 메모가 일치한다' ((& $memoModule { $script:currentId }) -eq $two.id -and (& $memoModule { $script:ui.List.SelectedItem.Id }) -eq $two.id)
    & $memoModule { param($id) Select-WtMemoItem $id } $one.id

    # 실패한 삭제/실행 취소를 여러 번 해도 배열이 먼저 바뀌지 않는다.
    & $memoModule { $script:ui.Search.Text = '' }
    $held = [IO.File]::Open($memoPath, 'Open', 'Read', 'Read')
    & $memoModule { Remove-WtMemoCurrent; Remove-WtMemoCurrent; Add-WtMemo }
    Assert-Reliable '삭제·추가 실패 시 메모리 배열을 유지한다' ((& $memoModule { $script:items.Count }) -eq 2)
    $held.Dispose(); $held = $null
    & $memoModule { Remove-WtMemoCurrent }
    $held = [IO.File]::Open($memoPath, 'Open', 'Read', 'Read')
    & $memoModule { Restore-WtMemoDeleted; Restore-WtMemoDeleted }
    Assert-Reliable '실행 취소 실패를 재시도해도 중복 항목을 만들지 않는다' ((& $memoModule { $script:items.Count }) -eq 1)
    $held.Dispose(); $held = $null
    & $memoModule { Restore-WtMemoDeleted }
    Assert-Reliable '실행 취소 성공 뒤 ID가 모두 고유하다' ((Get-WtMemoItems).Count -eq 2 -and @((Get-WtMemoItems).id | Sort-Object -Unique).Count -eq 2)
    Assert-Reliable '성공한 종료에서만 정리 처리를 실행한다' ((& $shellModule { Stop-WtApp }) -and $exitCounter.Count -eq 1)

    # 실제 가상화 패널: 데이터 5천 개여도 보이는 컨테이너만 만든다.
    $docsAddon = Get-WtAddon -Id 'docs'
    $docView = @(& $docsAddon.NewView | Where-Object { $_ -is [Windows.UIElement] })[-1]
    $docView.Resources.MergedDictionaries.Add((New-WtThemeDictionary))
    $docView.Resources.MergedDictionaries.Add((Get-WtStyleDictionary))
    $docList = & $docsModule { $script:ui.List }
    $rows = @(for ($i=0; $i -lt 5000; $i++) { [pscustomobject]@{ Name="sample-$i.txt"; Path="X:\synthetic\sample-$i.txt"; IsDir=$false; Kind='텍스트'; SizeText='1 KB'; ModifiedText='2026-01-01'; IconText=[string][char]0xE8A5; Missing=$false } })
    & $docsModule { param($items) $script:entries = $items; Update-WtDocList } $rows
    $docList.Measure([Windows.Size]::new(900,500)); $docList.Arrange([Windows.Rect]::new(0,0,900,500)); $docList.UpdateLayout()
    $realized = 0
    for ($i=0; $i -lt 5000; $i++) { if ($docList.ItemContainerGenerator.ContainerFromIndex($i)) { $realized++ } }
    Assert-Reliable '문서 5천 항목을 모두 보관하고 화면 행은 100개 미만으로 만든다' ($docList.Items.Count -eq 5000 -and $realized -gt 0 -and $realized -lt 100)

    # 예측 가능한 느린 작업을 주입해 이전 조회 결과와 실행 수 제한을 검사한다.
    $originalWorker = & $docsModule { $script:worker }
    & $docsModule {
        $script:worker = { param($request, $kinds, $icons) if ($request.Path -eq 'slow') { Start-Sleep -Milliseconds 300 }; @{ Items = @([pscustomobject]@{ Name=$request.Path; Path=$request.Path; IsDir=$false; Missing=$false }); Error=$null } }
        Show-WtDocFolder 'slow'
        Show-WtDocFolder 'latest'
    }
    Wait-DocJobs
    Assert-Reliable '늦은 이전 조회가 새 폴더 결과를 덮지 않는다' ((& $docsModule { $script:entries[0].Name }) -eq 'latest')
    & $docsModule { for ($i=0; $i -lt 20; $i++) { Show-WtDocFolder 'slow' }; Cancel-WtDocFolder }
    Assert-Reliable '빠른 연속 탐색도 실행 중 조회는 8개 이하이고 취소 뒤 살아 있는 조회가 없다' ((& $docsModule { $script:jobs.Count }) -le 8 -and -not (& $docsModule { @($script:jobs | Where-Object { -not $_.Request.Cancel.Value }).Count }))
    Wait-DocJobs
    Assert-Reliable '취소된 조회가 나중에 목록을 채우지 않는다' ((& $docsModule { $script:entries.Count }) -eq 0 -and -not (& $docsModule { $script:ui.Loading }))
    # 많은 하위 폴더도 한 차례에 모두 생성하지 않는다.
    $batchState = & $docsModule {
        $node = New-WtDocTreeItem 'synthetic' 'X:\synthetic' 0xE8B7
        $paths = @(for ($i=0; $i -lt 200; $i++) { "X:\synthetic\folder-$i" })
        $script:treeTarget = $null; $script:treeWaitNode = $null
        Set-WtDocResult @{ Key='tree:test'; Generation=$script:treeGeneration; Context=$node } @{ Items=$paths; Error=$null }
        Complete-WtDocTreeBatch
        @{ Node=$node; FirstCount=@($node.Items | Where-Object { $_ -is [Windows.Controls.TreeViewItem] }).Count; Remaining=$script:treeBatches.Count }
    }
    Assert-Reliable '하위 폴더 트리를 작은 묶음으로 나눠 반영한다' ($batchState.FirstCount -gt 0 -and $batchState.FirstCount -le 32 -and $batchState.Remaining -eq 1)
    Wait-DocJobs
    Assert-Reliable '나눠 반영한 하위 폴더를 빠짐없이 보여 준다' ($batchState.Node.Items.Count -eq 200 -and $batchState.Node.Tag.Loaded)
    Set-WtSetting 'docs.roots' @('X:\tree-a', 'X:\tree-b')
    & $docsModule {
        Update-WtDocTree
        Show-WtDocFolder 'X:\tree-a' -SelectTree
        $script:ui.Tree.Items[2].IsSelected = $true
        Set-WtDocResult @{Key='tree:old'; Generation=$script:treeGeneration; Context=$script:ui.Tree.Items[1]} @{Items=@(); Error=$null}
        Complete-WtDocTreeBatch
    }
    Assert-Reliable '늦은 트리 결과가 새로 고른 폴더 선택을 되돌리지 않는다' ((& $docsModule { $script:current }) -eq 'X:\tree-b' -and (& $docsModule { $script:ui.Tree.SelectedItem.Tag.Path }) -eq 'X:\tree-b')
    Wait-DocJobs
    # 트리 대기열이 많아도 현재 폴더 요청의 자리를 남긴다.
    & $docsModule {
        $script:worker = { param($request, $kinds, $icons) Start-Sleep -Milliseconds 200; @{ Items=@(); Error=$null } }
        [void](Start-WtDocRequest 'home' 'pins')
        [void](Start-WtDocRequest 'folder' 'folder' 'busy')
        for ($i=0; $i -lt 32; $i++) { $script:pending["tree:queued-$i"] = @{ Key="tree:queued-$i"; Request=@{ Action='tree'; Path="X:\queued-$i"; Cancel=[hashtable]::Synchronized(@{ Value = $false }) }; Generation=-1 } }
        $script:current = 'new'
        [void](Start-WtDocRequest 'folder' 'folder' 'new')
    }
    Assert-Reliable '트리 대기열이 꽉 차도 새 폴더 조회는 바로 시작한다' (& $docsModule { @($script:jobs | Where-Object { $_.Key -eq 'folder' -and $_.Request.Path -eq 'new' -and -not $_.Request.Cancel.Value }).Count -eq 1 })
    & $docsModule { Stop-WtDocs }
    Wait-DocJobs
    # 응답 없는 폴더 조회를 취소하면 그 조회가 끝나기를 기다리지 않고 새 폴더를 연다.
    & $docsModule {
        $script:worker = { param($request, $kinds, $icons) if ($request.Path -like 'stuck*') { Start-Sleep -Milliseconds 1500 }; @{ Items = @([pscustomobject]@{ Name=$request.Path; Path=$request.Path; IsDir=$false; Missing=$false }); Error=$null } }
        Show-WtDocFolder 'stuck-1'; Show-WtDocFolder 'stuck-2'; Cancel-WtDocFolder
        Show-WtDocFolder 'fresh'
    }
    Assert-Reliable '응답 없는 조회 두 개를 취소해도 새 폴더 조회가 바로 시작한다' (& $docsModule { @($script:jobs | Where-Object { $_.Key -eq 'folder' -and $_.Request.Path -eq 'fresh' }).Count -eq 1 })
    Wait-DocJobs
    Assert-Reliable '늦게 끝난 취소 조회는 새 폴더 목록을 덮지 않는다' ((& $docsModule { $script:entries[0].Name }) -eq 'fresh')
    # 다른 폴더를 펼쳐도 지금 폴더 쪽으로 펼침 · 스크롤을 되돌리지 않는다.
    & $docsModule {
        $script:worker = {
            param($request, $kinds, $icons)
            $tree = @{ 'X:\jump' = @('X:\jump\a', 'X:\jump\b'); 'X:\jump\a' = @('X:\jump\a\deep'); 'X:\jump\b' = @('X:\jump\b\x') }
            if ($request.Action -eq 'tree') { return @{ Items = @($tree[$request.Path] | Where-Object { $_ }); Error = $null } }
            @{ Items = @(); Error = $null }
        }
    }
    Set-WtSetting 'docs.roots' @('X:\jump')
    & $docsModule { Update-WtDocTree; Show-WtDocFolder 'X:\jump\a\deep' -SelectTree }
    Wait-DocJobs
    $jump = & $docsModule { $rootNode = $script:ui.Tree.Items[1]; @{ A = $rootNode.Items[0]; B = $rootNode.Items[1]; Selected = $script:ui.Tree.SelectedItem } }
    Assert-Reliable '깊은 폴더를 열면 트리를 차례로 펼쳐 찾아간다' ($jump.Selected.Tag.Path -eq 'X:\jump\a\deep' -and $jump.A.IsExpanded)
    # 다시 찾아가기(선택 · 스크롤 이동)가 불리는지 센다. 화면 없이 스크롤 위치를 볼 수 없어서다.
    & $docsModule {
        $script:testSelectCalls = 0
        $script:originalSelect = (Get-Command Select-WtDocTreePath).ScriptBlock
        Set-Item Function:script:Select-WtDocTreePath -Value { param($Path) $script:testSelectCalls++; & $script:originalSelect $Path }
    }
    $jump.A.IsExpanded = $false
    $jump.B.IsExpanded = $true
    Wait-DocJobs
    $jumpCalls = & $docsModule { $script:testSelectCalls; Set-Item Function:script:Select-WtDocTreePath -Value $script:originalSelect }
    Assert-Reliable '다른 폴더를 펼쳐도 지금 폴더로 되돌아가거나 접어 둔 폴더를 펼치지 않는다' (-not $jump.A.IsExpanded -and $jump.B.Tag.Loaded -and $jumpCalls -eq 0)
    # 마지막으로 본 폴더가 없어졌으면 첫 폴더를 연다.
    & $docsModule {
        $script:worker = {
            param($request, $kinds, $icons)
            if ($request.Action -eq 'tree') { return @{ Items = @(); Error = $null } }
            if ($request.Path -eq 'X:\fb\gone') { return @{ Items = @(); Error = '경로를 찾을 수 없습니다.' } }
            @{ Items = @([pscustomobject]@{ Name='ok.txt'; Path="$($request.Path)\ok.txt"; IsDir=$false; Missing=$false }); Error = $null }
        }
    }
    Set-WtSetting 'docs.roots' @('X:\fb'); Set-WtSetting 'docs.last' 'X:\fb\gone'
    & $docsModule { Update-WtDocsView }
    Wait-DocJobs
    Assert-Reliable '마지막 폴더를 못 열면 오류 대신 첫 폴더를 연다' ((& $docsModule { $script:current }) -eq 'X:\fb' -and -not (& $docsModule { $script:ui.Error }) -and (& $docsModule { $script:entries.Count }) -eq 1)
    Set-WtSetting 'docs.roots' @(); Set-WtSetting 'docs.last' ''
    & $docsModule { param($worker) $script:worker = $worker; Stop-WtDocs } $originalWorker
    Assert-Reliable '완료된 조회의 실행 자원과 타이머를 정리한다' ((& $docsModule { $script:jobs.Count }) -eq 0 -and -not (& $docsModule { $script:pollTimer.IsEnabled }))

    # 저장하지 못해도 [끝내기]를 고르면 끝낼 수 있다. 끌 수 없는 상태를 만들지 않는다.
    & $memoModule { $script:ui.BodyBox.Text = 'lost draft' }
    $held = [IO.File]::Open($memoPath, 'Open', 'Read', 'Read')
    $allowed = & $shellModule { $script:exiting = $false; $script:frame = [pscustomobject]@{ Continue = $true }; $script:testAsked = @(); $script:testChoices.Enqueue('proceed'); Stop-WtApp }
    Assert-Reliable '저장하지 못해도 [끝내기]를 고르면 정리하고 끝낸다' ($allowed -and $exitCounter.Count -eq 2 -and -not (& $shellModule { $script:frame.Continue }) -and (& $shellModule { $script:testAsked -join ',' }) -eq 'exit')
    $held.Dispose(); $held = $null

    # 기존 WinForms도 실제 FormClosing에서 저장 실패를 처리해야 한다: 끝낼지 묻고, [아니요]면 취소한다.
    $legacyRoot = Split-Path $root -Parent
    $legacyStorage = Import-Module (Join-Path $legacyRoot 'modules/Storage.psm1') -Force -PassThru
    & $legacyStorage { param($path) $script:dataDir = $path } (Join-Path $tmp 'legacy-data')
    $instance = Open-WorkToolDataLock
    $blocked = $false
    try { $extra = Open-WorkToolDataLock; $extra.Dispose() } catch { $blocked = $true }
    Assert-Reliable '기존 버전도 같은 데이터 폴더의 중복 실행을 막는다' $blocked
    Import-Module (Join-Path $legacyRoot 'modules/UiHelper.psm1') -Force
    $legacyMemo = Import-Module (Join-Path $legacyRoot 'modules/Memo.psm1') -Force -PassThru
    Save-Records memo @([pscustomobject]@{Id='old';Title='old';Body='original';Created='2026-01-01';Updated='2026-01-01'})
    $legacyForm = [Windows.Forms.Form]::new()
    $legacyView = New-MemoView
    $legacyForm.Controls.Add($legacyView)
    [void]$legacyForm.Handle; [void]$legacyView.Handle
    & $legacyMemo {
        function script:Show-MemoMessage { param($Text) $script:testShown++ }
        function script:Confirm-MemoDiscard { param($Reason) $script:testAsked++; $script:testAnswer }
        $script:testShown = 0; $script:testAsked = 0; $script:testAnswer = $false
        Show-Memo $script:memos[0]; $script:bodyBox.Text = 'legacy draft'
    }
    $legacyPath = Join-Path $tmp 'legacy-data/memo.csv'
    $held = [IO.File]::Open($legacyPath, 'Open', 'Read', 'Read')
    $closeArgs = [Windows.Forms.FormClosingEventArgs]::new('UserClosing', $false)
    $closing = [Windows.Forms.Form].GetMethod('OnFormClosing', [Reflection.BindingFlags]'Instance,NonPublic')
    [void]$closing.Invoke($legacyForm, @($closeArgs))
    Assert-Reliable '기존 버전 저장 실패 시 끝낼지 묻고 [아니요]면 종료를 취소한다' ($closeArgs.Cancel -and (& $legacyMemo { $script:dirty }) -and (& $legacyMemo { $script:testAsked }) -eq 1)
    & $legacyMemo { $script:testAnswer = $true }
    $closeArgs = [Windows.Forms.FormClosingEventArgs]::new('WindowsShutDown', $false)
    [void]$closing.Invoke($legacyForm, @($closeArgs))
    Assert-Reliable '기존 버전도 [예]면 저장하지 못해도 끝낸다 (Windows 종료 포함)' (-not $closeArgs.Cancel -and (& $legacyMemo { $script:testAsked }) -eq 2 -and @(Get-Records memo)[0].Body -eq 'original')
    # 트레이로 숨기는 중(메인 창이 이미 취소함)에는 묻지 않고 창을 다시 띄워 알린다.
    $closeArgs = [Windows.Forms.FormClosingEventArgs]::new('UserClosing', $true)
    [void]$closing.Invoke($legacyForm, @($closeArgs))
    Assert-Reliable '기존 버전 트레이 숨김 중 저장 실패는 묻지 않고 창을 다시 띄워 알린다' ($closeArgs.Cancel -and $legacyForm.Visible -and (& $legacyMemo { $script:testAsked }) -eq 2 -and (& $legacyMemo { $script:testShown }) -eq 1)
    $legacyForm.Hide()
    $held.Dispose(); $held = $null
    $closeArgs = [Windows.Forms.FormClosingEventArgs]::new('UserClosing', $false)
    [void]$closing.Invoke($legacyForm, @($closeArgs))
    Assert-Reliable '기존 버전도 잠금 해제 후 종료 저장을 재시도한다' (-not $closeArgs.Cancel -and @(Get-Records memo)[0].Body -eq 'legacy draft')
} finally {
    if ($held) { $held.Dispose() }
    if ($instance) { $instance.Dispose() }
    if ($legacyForm) { $legacyForm.Dispose() }
    & $docsModule { Stop-WtDocs }
    & $memoModule { if ($script:saveTimer) { $script:saveTimer.Stop() }; if ($script:undoTimer) { $script:undoTimer.Stop() } }
    $safe = [IO.Path]::GetFullPath($tmp)
    if ((Split-Path $safe -Leaf) -notmatch '^wt-reliability-[a-f0-9]{32}$' -or -not $safe.StartsWith([IO.Path]::GetFullPath([IO.Path]::GetTempPath()), [StringComparison]::OrdinalIgnoreCase)) { throw '임시 경로 검증 실패' }
    Remove-Item -LiteralPath $safe -Recurse -Force
}
