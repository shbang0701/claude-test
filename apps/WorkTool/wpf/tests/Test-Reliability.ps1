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
    Assert-Reliable '저장 실패는 정리 처리 전에 종료를 취소한다' (-not $allowed -and $exitCounter.Count -eq 0 -and (& $shellModule { $script:frame.Continue }))
    & $shellModule {
        $script:current = 'memo'
        $script:navButtons = @{ memo = @{Button=[Windows.Controls.RadioButton]::new()}; docs = @{Button=[Windows.Controls.RadioButton]::new()} }
        $script:navButtons.docs.Button.IsChecked = $true
    }
    Select-WtPage '_help'
    Assert-Reliable '저장 실패는 페이지 전환을 취소한다' ((& $shellModule { $script:current }) -ne '_help')
    Assert-Reliable '전환을 취소하면 메뉴 선택도 원래 화면으로 돌아간다' ((& $shellModule { $script:navButtons.memo.Button.IsChecked }) -and -not (& $shellModule { $script:navButtons.docs.Button.IsChecked }))
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
    Assert-Reliable '빠른 연속 탐색도 실행 중 조회는 두 개 이하이다' ((& $docsModule { $script:jobs.Count }) -le 2)
    Wait-DocJobs
    Assert-Reliable '취소된 조회가 나중에 목록을 채우지 않는다' ((& $docsModule { $script:entries.Count }) -eq 0 -and -not (& $docsModule { $script:ui.Loading }))
    # 많은 하위 폴더도 한 차례에 모두 생성하지 않는다.
    $batchState = & $docsModule {
        $node = New-WtDocTreeItem 'synthetic' 'X:\synthetic' 0xE8B7
        $paths = @(for ($i=0; $i -lt 200; $i++) { "X:\synthetic\folder-$i" })
        $script:treeSelectionPath = $null
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
        for ($i=0; $i -lt 32; $i++) { $script:pending["tree:queued-$i"] = @{ Key="tree:queued-$i"; Request=@{Action='tree'} } }
        $script:current = 'new'
        [void](Start-WtDocRequest 'folder' 'folder' 'new')
    }
    Assert-Reliable '트리 대기열이 꽉 차도 새 폴더 요청을 남긴다' (& $docsModule { $script:pending.Contains('folder') })
    & $docsModule { Stop-WtDocs }
    Wait-DocJobs
    & $docsModule { param($worker) $script:worker = $worker; Stop-WtDocs } $originalWorker
    Assert-Reliable '완료된 조회의 실행 자원과 타이머를 정리한다' ((& $docsModule { $script:jobs.Count }) -eq 0 -and -not (& $docsModule { $script:pollTimer.IsEnabled }))

    # 기존 WinForms도 실제 FormClosing에서 저장 실패를 취소해야 한다.
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
    & $legacyMemo { function script:Show-MemoMessage { param($Text) }; Show-Memo $script:memos[0]; $script:bodyBox.Text = 'legacy draft' }
    $legacyPath = Join-Path $tmp 'legacy-data/memo.csv'
    $held = [IO.File]::Open($legacyPath, 'Open', 'Read', 'Read')
    $closeArgs = [Windows.Forms.FormClosingEventArgs]::new('UserClosing', $false)
    $closing = [Windows.Forms.Form].GetMethod('OnFormClosing', [Reflection.BindingFlags]'Instance,NonPublic')
    [void]$closing.Invoke($legacyForm, @($closeArgs))
    Assert-Reliable '기존 버전 저장 실패 시 실제 종료 이벤트를 취소한다' ($closeArgs.Cancel -and (& $legacyMemo { $script:dirty }))
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
