# 코드 색칠. 마크다운 코드 블록 등에서 쓴다. WPF 없이 동작한다 (글을 조각 목록으로 나누기만 한다).
# 언어 하나 = 규칙 목록 하나 (Register-WtCodeLanguage). 새 언어는 이 파일 아래쪽 목록에 하나 더 등록하면 된다.
#
# 규칙: @{ Kind = '색 이름'; Pattern = '정규식' } 또는 @{ Kind = …; Words = '낱말 낱말 …' }
# - 같은 자리에서는 위에 있는 규칙이 먼저다. 주석 · 문자열을 맨 위에 둔다.
# - Kind는 Theme.psm1의 Code.* 색 이름이다: Comment · String · Keyword · Number · Command · Variable · Type · Attr · Inserted · Deleted
# - 정규식은 여러 줄 모드(^ $가 줄 단위)다. 규칙 안에서는 이름 붙은 그룹을 쓰지 않는다.

$ErrorActionPreference = 'Stop'

$script:languages = @{}   # 이름 → @{ Name; Regex; Kinds }
$script:aliases = @{}     # 별칭 → 이름

function Register-WtCodeLanguage {
    param(
        [Parameter(Mandatory)][string]$Name,
        [string[]]$Aliases = @(),
        [Parameter(Mandatory)][object[]]$Rules,
        [switch]$IgnoreCase
    )
    $parts = for ($i = 0; $i -lt $Rules.Count; $i++) {
        $r = $Rules[$i]
        $pattern = if ($r.Words) {
            $words = @($r.Words -split '\s+' | Where-Object { $_ } | ForEach-Object { [regex]::Escape($_) })
            '(?<![\w$-])(?:' + ($words -join '|') + ')(?![\w-])'
        } else { $r.Pattern }
        "(?<k$i>$pattern)"
    }
    $options = [System.Text.RegularExpressions.RegexOptions]'Multiline, CultureInvariant'
    if ($IgnoreCase) { $options = $options -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase }
    $script:languages[$Name] = @{
        Name  = $Name
        Regex = [regex]::new(($parts -join '|'), $options)
        Kinds = @($Rules | ForEach-Object { $_.Kind })
    }
    foreach ($a in @($Name) + $Aliases) { $script:aliases[$a.ToLowerInvariant()] = $Name }
}

# 코드 블록 이름표(```powershell 의 powershell)로 언어를 찾는다. 없으면 $null
function Get-WtCodeLanguage([string]$Name) {
    if (-not $Name) { return $null }
    $key = $Name.Trim().ToLowerInvariant()
    if ($script:aliases.ContainsKey($key)) { return $script:aliases[$key] }
    $null
}

function Get-WtCodeLanguageNames { @($script:languages.Keys | Sort-Object) }

# 글을 @{ Kind; Text } 조각들로 나눈다. 색칠하지 않는 부분은 Kind가 ''다. 조각을 이으면 원래 글이 된다.
function Get-WtCodeTokens([string]$Code, [string]$Language) {
    $out = [System.Collections.Generic.List[object]]::new()
    if (-not $Code) { return }
    $name = Get-WtCodeLanguage $Language
    if (-not $name) { return @{ Kind = ''; Text = $Code } }
    $lang = $script:languages[$name]
    $pos = 0
    foreach ($m in $lang.Regex.Matches($Code)) {
        if ($m.Length -eq 0) { continue }
        if ($m.Index -gt $pos) { $out.Add(@{ Kind = ''; Text = $Code.Substring($pos, $m.Index - $pos) }) }
        $kind = ''
        for ($i = 0; $i -lt $lang.Kinds.Count; $i++) {
            if ($m.Groups["k$i"].Success) { $kind = $lang.Kinds[$i]; break }
        }
        $out.Add(@{ Kind = $kind; Text = $m.Value })
        $pos = $m.Index + $m.Length
    }
    if ($pos -lt $Code.Length) { $out.Add(@{ Kind = ''; Text = $Code.Substring($pos) }) }
    $out.ToArray()
}

# ---- 언어 목록 ---------------------------------------------------------------------------------

$num = '\b(?:0x[0-9a-fA-F]+|\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)\b'
$dq = '"(?:\\.|[^"\\\n])*"'
$sq = '''(?:\\.|[^''\\\n])*'''
$cBlock = '/\*[\s\S]*?(?:\*/|\z)'

Register-WtCodeLanguage -Name 'powershell' -Aliases 'ps1', 'ps', 'pwsh', 'posh', 'psm1' -IgnoreCase -Rules @(
    @{ Kind = 'Comment'; Pattern = '<#[\s\S]*?(?:#>|\z)|(?<![\w`])#.*' }
    @{ Kind = 'String'; Pattern = '@"[\s\S]*?^"@|@''[\s\S]*?^''@|"(?:`[\s\S]|""|[^"`])*"|''(?:''''|[^''])*''' }
    @{ Kind = 'Variable'; Pattern = '\$(?:\{[^}\n]*\}|(?:[A-Za-z]+:)?[\w?]+|\$|\^)' }
    @{ Kind = 'Type'; Pattern = '\[[A-Za-z][\w.]*(?:\[\])?(?:,\s*[\w.]+)*\](?!\s*\()' }
    @{ Kind = 'Command'; Pattern = '\b[A-Za-z]+-[A-Za-z][A-Za-z0-9]*\b' }
    @{ Kind = 'Keyword'; Pattern = '(?<![\w-])-(?:c|i)?(?:eq|ne|gt|ge|lt|le|like|notlike|match|notmatch|contains|notcontains|in|notin|replace|split)\b|(?<![\w-])-(?:and|or|not|xor|band|bor|bxor|bnot|shl|shr|is|isnot|as|join|f)\b' }
    @{ Kind = 'Keyword'; Words = 'begin break catch class clean continue data default do dynamicparam else elseif end enum exit filter finally for foreach function hidden if in param process return static switch throw trap try until using while' }
    @{ Kind = 'Attr'; Pattern = '(?<![\w-])-[A-Za-z][\w]*' }
    @{ Kind = 'Number'; Pattern = '\b\d+(?:\.\d+)?(?:[kmgtp]b)?\b|\b0x[0-9a-f]+\b' }
)

Register-WtCodeLanguage -Name 'bash' -Aliases 'sh', 'shell', 'zsh', 'console', 'shellscript' -Rules @(
    @{ Kind = 'Comment'; Pattern = '(?<=^|[ \t;])#.*' }
    @{ Kind = 'String'; Pattern = "$dq|'[^']*'" }
    @{ Kind = 'Variable'; Pattern = '\$(?:\{[^}\n]*\}|\w+|[@*#?$!0-9-])' }
    @{ Kind = 'Keyword'; Words = 'if then else elif fi for while until do done case esac function in return break continue local export readonly declare select time' }
    @{ Kind = 'Command'; Words = 'echo printf cd ls cat grep egrep sed awk find xargs sort uniq head tail wc cut tr tee sudo su chmod chown chgrp cp mv rm mkdir rmdir ln touch tar gzip gunzip zip unzip curl wget ssh scp rsync systemctl journalctl service kill killall ps top htop df du free mount umount lsblk blkid fdisk parted mkfs ip ifconfig ping traceroute ss netstat nmcli firewall-cmd iptables yum dnf apt apt-get rpm dpkg pip source set unset exit test read sleep date hostname uname whoami id useradd usermod passwd crontab nohup git docker kubectl podman vi vim nano less more which env alias' }
    @{ Kind = 'Attr'; Pattern = '(?<=\s)--?[A-Za-z][\w-]*' }
    @{ Kind = 'Number'; Pattern = '\b\d+(?:\.\d+)?\b' }
)

Register-WtCodeLanguage -Name 'python' -Aliases 'py', 'python3' -Rules @(
    @{ Kind = 'Comment'; Pattern = '#.*' }
    @{ Kind = 'String'; Pattern = '(?i:[rbfu]{0,2})(?:"""[\s\S]*?(?:"""|\z)|''''''[\s\S]*?(?:''''''|\z)|' + $dq + '|' + $sq + ')' }
    @{ Kind = 'Attr'; Pattern = '^[ \t]*@[\w.]+' }
    @{ Kind = 'Keyword'; Words = 'False None True and as assert async await break class continue def del elif else except finally for from global if import in is lambda nonlocal not or pass raise return try while with yield match case' }
    @{ Kind = 'Command'; Pattern = '(?<=\b(?:def|class)\s+)\w+|\b(?:print|len|range|open|str|int|float|bool|list|dict|set|tuple|type|isinstance|enumerate|zip|map|filter|sorted|sum|min|max|abs|any|all|super|input|format|round|repr|getattr|setattr|hasattr)\b(?=\s*\()' }
    @{ Kind = 'Variable'; Words = 'self cls' }
    @{ Kind = 'Number'; Pattern = $num }
)

Register-WtCodeLanguage -Name 'javascript' -Aliases 'js', 'jsx', 'mjs', 'cjs', 'typescript', 'ts', 'tsx' -Rules @(
    @{ Kind = 'Comment'; Pattern = "//.*|$cBlock" }
    @{ Kind = 'String'; Pattern = "$dq|$sq|``(?:\\[\s\S]|[^``\\])*``" }
    @{ Kind = 'Keyword'; Words = 'break case catch class const continue debugger default delete do else export extends finally for function if import in instanceof let new return super switch this throw try typeof var void while with yield async await of static get set interface type enum implements private public protected readonly as from null undefined true false' }
    @{ Kind = 'Type'; Pattern = '\b[A-Z][A-Za-z0-9_]*\b' }
    @{ Kind = 'Command'; Pattern = '\b[A-Za-z_$][\w$]*(?=\s*\()' }
    @{ Kind = 'Number'; Pattern = $num }
)

Register-WtCodeLanguage -Name 'json' -Aliases 'jsonc', 'json5' -Rules @(
    @{ Kind = 'Comment'; Pattern = "//.*|$cBlock" }
    @{ Kind = 'Attr'; Pattern = '"(?:\\.|[^"\\\n])*"(?=\s*:)' }
    @{ Kind = 'String'; Pattern = $dq }
    @{ Kind = 'Keyword'; Words = 'true false null' }
    @{ Kind = 'Number'; Pattern = '-?\b\d+(?:\.\d+)?(?:[eE][+-]?\d+)?\b' }
)

Register-WtCodeLanguage -Name 'yaml' -Aliases 'yml' -Rules @(
    @{ Kind = 'Comment'; Pattern = '(?<=^|\s)#.*' }
    @{ Kind = 'Attr'; Pattern = '(?<=^[ \t]*(?:-[ \t]+)?)[^\s#:''"\-][^:#\n]*?(?=[ \t]*:(?:[ \t]|$))' }
    @{ Kind = 'String'; Pattern = "$dq|'(?:''|[^'\n])*'" }
    @{ Kind = 'Type'; Pattern = '^(?:---|\.\.\.)[ \t]*$|!!?[\w/]+' }
    @{ Kind = 'Variable'; Pattern = '[&*][\w-]+' }
    @{ Kind = 'Keyword'; Pattern = '(?i)(?<=:[ \t]+|-[ \t]+|^)(?:true|false|null|yes|no|on|off|~)(?=[ \t]*(?:#|$))' }
    @{ Kind = 'Number'; Pattern = '(?<=:[ \t]+|-[ \t]+)-?\d+(?:\.\d+)?(?=[ \t]*(?:#|$))' }
)

Register-WtCodeLanguage -Name 'xml' -Aliases 'html', 'htm', 'xaml', 'svg', 'xhtml', 'plist' -IgnoreCase -Rules @(
    @{ Kind = 'Comment'; Pattern = '<!--[\s\S]*?(?:-->|\z)' }
    @{ Kind = 'Keyword'; Pattern = '<!\[CDATA\[|\]\]>|<!DOCTYPE[^>]*>|<\?[\s\S]*?\?>' }
    @{ Kind = 'Type'; Pattern = '(?<=</?)[\w:.-]+' }
    @{ Kind = 'String'; Pattern = '(?<==\s*)(?:"[^"]*"|''[^'']*'')' }
    @{ Kind = 'Attr'; Pattern = '(?<=<[^<>]*\s)[\w:.-]+(?=\s*=)' }
    @{ Kind = 'Variable'; Pattern = '&[#\w]+;' }
)

Register-WtCodeLanguage -Name 'sql' -Aliases 'mysql', 'tsql', 'plsql', 'postgresql', 'psql', 'sqlite' -IgnoreCase -Rules @(
    @{ Kind = 'Comment'; Pattern = "--.*|$cBlock" }
    @{ Kind = 'String'; Pattern = "'(?:''|[^'])*'" }
    @{ Kind = 'Type'; Words = 'int integer bigint smallint tinyint varchar nvarchar char nchar text ntext date datetime datetime2 time timestamp decimal numeric float real bit boolean money uniqueidentifier blob clob' }
    @{ Kind = 'Keyword'; Words = 'select from where and or not insert into values update set delete create alter drop truncate table index view database schema join left right inner outer full cross on group by order having limit offset top distinct as is null like in between exists case when then else end union all primary key foreign references default constraint unique check begin commit rollback transaction declare exec execute procedure function returns return if while grant revoke with desc asc use go' }
    @{ Kind = 'Command'; Pattern = '\b(?:count|sum|avg|min|max|coalesce|isnull|ifnull|nullif|cast|convert|concat|substring|len|length|upper|lower|trim|getdate|now|row_number|rank|dense_rank)\b(?=\s*\()' }
    @{ Kind = 'Variable'; Pattern = '@@?\w+' }
    @{ Kind = 'Number'; Pattern = '\b\d+(?:\.\d+)?\b' }
)

Register-WtCodeLanguage -Name 'csharp' -Aliases 'cs', 'c#' -Rules @(
    @{ Kind = 'Comment'; Pattern = "//.*|$cBlock" }
    @{ Kind = 'String'; Pattern = '@"(?:""|[^"])*"|\$?' + $dq + "|'(?:\\.|[^'\\\n])'" }
    @{ Kind = 'Keyword'; Words = 'abstract as base bool break byte case catch char checked class const continue decimal default delegate do double else enum event explicit extern false finally fixed float for foreach goto if implicit in int interface internal is lock long namespace new null object operator out override params private protected public readonly ref return sbyte sealed short sizeof stackalloc static string struct switch this throw true try typeof uint ulong unchecked unsafe ushort using virtual void volatile while var async await get set value record init' }
    @{ Kind = 'Type'; Pattern = '\b[A-Z]\w*\b' }
    @{ Kind = 'Command'; Pattern = '\b[a-z_]\w*(?=\s*\()' }
    @{ Kind = 'Number'; Pattern = $num }
)

Register-WtCodeLanguage -Name 'ini' -Aliases 'toml', 'conf', 'cfg', 'properties', 'env', 'dotenv', 'editorconfig' -Rules @(
    @{ Kind = 'Comment'; Pattern = '^[ \t]*[;#].*' }
    @{ Kind = 'Type'; Pattern = '^[ \t]*\[[^\]\n]*\]' }
    @{ Kind = 'Attr'; Pattern = '^[ \t]*[^=:\s\[;#][^=:\n]*?(?=[ \t]*[=:])' }
    @{ Kind = 'String'; Pattern = "$dq|'[^'\n]*'" }
    @{ Kind = 'Keyword'; Pattern = '(?i)(?<=[=:][ \t]*)(?:true|false|yes|no|on|off)\b' }
    @{ Kind = 'Number'; Pattern = '(?<=[=:][ \t]*)-?\d+(?:\.\d+)?\b' }
)

Register-WtCodeLanguage -Name 'diff' -Aliases 'patch', 'udiff' -Rules @(
    @{ Kind = 'Type'; Pattern = '^(?:@@.*|diff .*|index .*|---(?: .*)?|\+\+\+(?: .*)?)$' }
    @{ Kind = 'Inserted'; Pattern = '^[+>].*' }
    @{ Kind = 'Deleted'; Pattern = '^[-<].*' }
)

Register-WtCodeLanguage -Name 'css' -Aliases 'scss', 'less' -Rules @(
    @{ Kind = 'Comment'; Pattern = "$cBlock|(?<=^|\s)//.*" }
    @{ Kind = 'String'; Pattern = "$dq|$sq" }
    @{ Kind = 'Keyword'; Pattern = '@[\w-]+|!important' }
    @{ Kind = 'Attr'; Pattern = '(?<=[{;\s])[\w-]+(?=\s*:[^{}\n]*[;}\n])' }
    @{ Kind = 'Number'; Pattern = '#[0-9a-fA-F]{3,8}\b|-?\b\d+(?:\.\d+)?(?:px|em|rem|%|vh|vw|s|ms|pt|deg)?' }
    @{ Kind = 'Type'; Pattern = '[.#][A-Za-z_][\w-]*' }
)

Register-WtCodeLanguage -Name 'dockerfile' -Aliases 'docker', 'containerfile' -IgnoreCase -Rules @(
    @{ Kind = 'Comment'; Pattern = '^[ \t]*#.*' }
    @{ Kind = 'Keyword'; Pattern = '^[ \t]*(?:FROM|RUN|CMD|LABEL|MAINTAINER|EXPOSE|ENV|ADD|COPY|ENTRYPOINT|VOLUME|USER|WORKDIR|ARG|ONBUILD|STOPSIGNAL|HEALTHCHECK|SHELL)\b|\bAS\b' }
    @{ Kind = 'String'; Pattern = "$dq|'[^'\n]*'" }
    @{ Kind = 'Variable'; Pattern = '\$(?:\{[^}\n]*\}|\w+)' }
    @{ Kind = 'Attr'; Pattern = '(?<=\s)--?[A-Za-z][\w-]*' }
)

Register-WtCodeLanguage -Name 'batch' -Aliases 'bat', 'cmd' -IgnoreCase -Rules @(
    @{ Kind = 'Comment'; Pattern = '^[ \t@]*(?:rem(?:\s.*)?|::.*)$' }
    @{ Kind = 'String'; Pattern = '"[^"\n]*"' }
    @{ Kind = 'Variable'; Pattern = '%~?[\w]*%|%%~?\w|%~?\d|![\w]+!' }
    @{ Kind = 'Type'; Pattern = '^[ \t]*:\w+' }
    @{ Kind = 'Keyword'; Words = 'if else for in do goto call set setlocal endlocal echo exit not exist defined errorlevel equ neq lss leq gtr geq pause shift start cd pushd popd title' }
    @{ Kind = 'Attr'; Pattern = '(?<=\s)/[\w?]+' }
    @{ Kind = 'Number'; Pattern = '\b\d+\b' }
)

Export-ModuleMember -Function Register-WtCodeLanguage, Get-WtCodeLanguage, Get-WtCodeLanguageNames, Get-WtCodeTokens
