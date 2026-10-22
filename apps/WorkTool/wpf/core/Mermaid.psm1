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
