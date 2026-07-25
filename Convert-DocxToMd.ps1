<#
.SYNOPSIS
    Converts .docx files to clean Markdown: strips hyperlink URLs, removes images,
    and applies case-aware keyword replacements.

.DESCRIPTION
    Owns a fixed folder structure next to the script:

        input\    drop .docx files here (a file picker offers to copy them in if empty)
        output\   .md files land here, never overwritten
        logs\     one timestamped log per run

    Conversion is done by pandoc. Images are dropped entirely (no media files are
    written to disk), hyperlinks keep their text but lose the URL, tracked changes
    are accepted, and comments, footnotes and document metadata are removed.

.PARAMETER Root
    Workspace folder. Defaults to the folder containing this script.

.PARAMETER ReplacementsPath
    Path to the keyword mapping file. Defaults to <Root>\replacements.json.

.PARAMETER NoPrompt
    Never show dialogs. Use for scheduled/unattended runs.

.PARAMETER KeepFootnotes
    Keep footnotes and endnotes instead of removing them.

.PARAMETER StripBareUrls
    Also remove bare URLs typed as plain text (off by default).

.PARAMETER PandocTimeoutSec
    Maximum seconds one document may spend inside pandoc before it is stopped and the file
    marked as failed (default 120). The rest of the batch continues.

.PARAMETER Version
    Print the script version and exit.

.EXAMPLE
    .\Convert-DocxToMd.ps1

.EXAMPLE
    .\Convert-DocxToMd.ps1 -NoPrompt -Verbose

.NOTES
    Version history
      1.4.0  Hardening: 5s match timeout on every regex (including replacements.json patterns)
             and a configurable timeout on pandoc itself - a hostile or corrupt document now
             fails alone instead of hanging the batch. Code-fence tracking pairs open/close
             markers correctly. Log folder keeps the newest 30 runs. Performance: compiled
             regexes reused per table cell, non-HTML lines skip the tag sweep.
      1.3.0  Leftover inline HTML is unwrapped: <span class="mark">, <u>, <sub>, <sup>,
             small caps and custom character styles keep their words but lose the markup.
             Fenced code blocks and inline code are left untouched. Keyword replacement no
             longer corrupts bare URLs.
      1.2.0  Replacements are written exactly as spelled in replacements.json. Casing of the
             matched text is only mirrored when an entry opts in with "preserveCase": true.
      1.1.0  Raw HTML tables converted to Markdown pipe tables; Word bookmark anchor markup
             removed; empty table headers filled from the first row; -Version switch added.
      1.0.0  Initial release.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string] $Root,
    [string] $ReplacementsPath,
    [switch] $NoPrompt,
    [switch] $KeepFootnotes,
    [switch] $StripBareUrls,
    [ValidateRange(1, 3600)]
    [int] $PandocTimeoutSec = 120,
    [switch] $Version
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:ScriptVersion = '1.4.0'

if ($Version) {
    Write-Host ('Convert-DocxToMd {0}' -f $script:ScriptVersion)
    exit 0
}

#region Regex patterns

# Every pattern gets a match timeout: document text is untrusted input, and replacements.json
# may carry user-authored patterns (regex: true). A runaway match fails that one file instead
# of hanging the batch.
$script:RegexTimeout = [TimeSpan]::FromSeconds(5)

function New-CompiledRegex {
    param(
        [Parameter(Mandatory)][string] $Pattern,
        [string] $Options = 'None'
    )
    [regex]::new($Pattern,
        [System.Text.RegularExpressions.RegexOptions]("Compiled, $Options"),
        $script:RegexTimeout)
}

# Balancing groups keep these correct when link text or URLs contain nested brackets.
# The optional trailing {...} is a pandoc attribute block. The whitespace before it is only
# consumed when such a block is actually present, so ordinary spacing between words survives.
$script:RxImage = New-CompiledRegex '!\[(?<alt>(?>[^\[\]]|(?<o>\[)|(?<-o>\]))*(?(o)(?!)))\]\((?>[^()]|(?<p>\()|(?<-p>\)))*(?(p)(?!))\)(?:[ \t]*\{[^{}]*\})?'
$script:RxHtmlImage = New-CompiledRegex '<img\b[^>]*>' 'IgnoreCase'
# Word images arrive as <figure><img/><figcaption>..</figcaption></figure>; the caption belongs
# to the image, so the whole block goes.
$script:RxFigure = New-CompiledRegex '<figure\b[^>]*>.*?</figure>' 'IgnoreCase, Singleline'
$script:RxLink = New-CompiledRegex '(?<!!)\[(?<text>(?>[^\[\]]|(?<o>\[)|(?<-o>\]))*(?(o)(?!)))\]\((?>[^()]|(?<p>\()|(?<-p>\)))*(?(p)(?!))\)(?:[ \t]*\{[^{}]*\})?'
# Pandoc emits <https://...> when the link text is the URL itself; keep the text, drop the markup.
$script:RxAutolink  = New-CompiledRegex '<((?:https?|ftp|mailto):[^>\s]+)>' 'IgnoreCase'
$script:RxBareUrl   = New-CompiledRegex '(?<![\w/])((?:https?|ftp)://|www\.)[^\s<>()\[\]"'']+' 'IgnoreCase'
$script:RxFootnote  = New-CompiledRegex '\[\^[^\]\s]+\]'
$script:RxFootnoteDef = New-CompiledRegex '^\[\^[^\]\s]+\]:'

# Word bookmarks (TOC targets, cross-reference targets) survive as empty anchor markup.
$script:RxAnchorSpan   = New-CompiledRegex '<span\b[^>]*>\s*</span>' 'IgnoreCase'
$script:RxAnchorTag    = New-CompiledRegex '<a\b(?=[^>]*\b(?:id|name)\s*=)[^>]*>\s*</a>' 'IgnoreCase'
# Same thing in pandoc's attribute syntax, emitted when raw HTML is unavailable: []{#_Toc123 .anchor}
$script:RxAnchorAttr   = New-CompiledRegex '\[\s*\]\{[^{}]*\}'

# Pandoc falls back to raw HTML for any table GFM cannot express (merged cells, block content).
$script:RxHtmlTable    = New-CompiledRegex '<table\b[^>]*>.*?</table>' 'IgnoreCase, Singleline'
$script:RxTableCaption = New-CompiledRegex '<caption\b[^>]*>(?<text>.*?)</caption>' 'IgnoreCase, Singleline'
$script:RxTableRow     = New-CompiledRegex '<tr\b[^>]*>(?<body>.*?)</tr>' 'IgnoreCase, Singleline'
$script:RxTableCell    = New-CompiledRegex '<(?<tag>t[hd])\b(?<attrs>[^>]*)>(?<body>.*?)</\k<tag>>' 'IgnoreCase, Singleline'
$script:RxTheadRow     = New-CompiledRegex '<thead\b[^>]*>.*?</thead>' 'IgnoreCase, Singleline'
$script:RxColspan      = New-CompiledRegex '\bcolspan\s*=\s*["'']?(\d+)' 'IgnoreCase'
$script:RxRowspan      = New-CompiledRegex '\browspan\s*=\s*["'']?(\d+)' 'IgnoreCase'
# Used by ConvertFrom-HtmlCell; hoisted so they are not rebuilt for every table cell.
$script:RxParaBreak    = New-CompiledRegex '</p>\s*<p\b[^>]*>' 'IgnoreCase'
$script:RxAnyTag       = New-CompiledRegex '<[^>]+>'

# Character formatting Word has but Markdown does not: highlights, underline, small caps,
# sub/superscript and custom character styles all arrive as raw inline HTML. The text is kept,
# the markup is not. Code is protected separately so genuine samples survive untouched.
$script:RxCodeFence  = New-CompiledRegex '^\s*(```|~~~)'
$script:RxInlineCode = New-CompiledRegex '`+[^`\r\n]*`+'
$script:RxUnwrapTag  = New-CompiledRegex '</?(?:span|u|ins|del|mark|small|abbr|cite|font|bdi|bdo|q|sub|sup|a|div|section|kbd|samp|var)\b[^>]*>' 'IgnoreCase'
$script:RxStrongTag  = New-CompiledRegex '</?(?:strong|b)\b[^>]*>' 'IgnoreCase'
$script:RxEmTag      = New-CompiledRegex '</?(?:em|i)\b[^>]*>' 'IgnoreCase'
$script:RxCodeTag    = New-CompiledRegex '</?code\b[^>]*>' 'IgnoreCase'
$script:RxBrTag      = New-CompiledRegex '<br\s*/?>' 'IgnoreCase'
# Pandoc's attribute syntax for the same thing, e.g. [highlighted]{.mark}
$script:RxAttrSpan   = New-CompiledRegex '\[(?<text>[^\[\]]*)\]\{[.#][^{}]*\}'

#endregion

#region Helpers

function Write-Log {
    param(
        [Parameter(Mandatory)][string] $Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'OK')][string] $Level = 'INFO',
        [switch] $NoConsole
    )
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level.PadRight(5), $Message
    if ($script:LogPath) {
        # Logging and temp cleanup are bookkeeping, not the operation -WhatIf is previewing.
        Add-Content -LiteralPath $script:LogPath -Value $line -Encoding utf8 -WhatIf:$false
    }
    if (-not $NoConsole) {
        $color = switch ($Level) {
            'OK'    { 'Green' }
            'WARN'  { 'Yellow' }
            'ERROR' { 'Red' }
            default { 'Gray' }
        }
        Write-Host $Message -ForegroundColor $color
    }
}

function Initialize-Workspace {
    param([Parameter(Mandatory)][string] $Path)

    $folders = [ordered]@{
        Root   = $Path
        Input  = Join-Path $Path 'input'
        Output = Join-Path $Path 'output'
        Logs   = Join-Path $Path 'logs'
    }
    foreach ($key in @('Root', 'Input', 'Output', 'Logs')) {
        if (-not (Test-Path -LiteralPath $folders[$key])) {
            New-Item -ItemType Directory -Path $folders[$key] -Force | Out-Null
        }
    }
    return $folders
}

function Get-PandocPath {
    $cmd = Get-Command pandoc -CommandType Application -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if (-not $cmd) {
        throw "pandoc was not found on PATH. Install it with:`n  winget install --id JohnMacFarlane.Pandoc`nor  choco install pandoc"
    }
    return $cmd.Source
}

function Get-InputDocuments {
    param([Parameter(Mandatory)][string] $InputFolder)
    Get-ChildItem -LiteralPath $InputFolder -Filter '*.docx' -File |
        Where-Object { $_.Name -notlike '~$*' } |
        Sort-Object Name
}

<#
    Returns a path that does not exist yet: name.md, then name-1.md, name-2.md, ...
    Picks the next number above the highest existing suffix so re-runs stay ordered.
#>
function Get-AvailablePath {
    param(
        [Parameter(Mandatory)][string] $Folder,
        [Parameter(Mandatory)][string] $BaseName,
        [Parameter(Mandatory)][string] $Extension
    )
    $candidate = Join-Path $Folder ($BaseName + $Extension)
    if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }

    $escaped = [regex]::Escape($BaseName)
    $rx = [regex]::new(('^{0}-(\d+)$' -f $escaped), 'IgnoreCase')
    $highest = 0
    foreach ($existing in Get-ChildItem -LiteralPath $Folder -Filter ($BaseName + '-*' + $Extension) -File) {
        $m = $rx.Match([System.IO.Path]::GetFileNameWithoutExtension($existing.Name))
        if ($m.Success) {
            $n = [int]$m.Groups[1].Value
            if ($n -gt $highest) { $highest = $n }
        }
    }
    return Join-Path $Folder ('{0}-{1}{2}' -f $BaseName, ($highest + 1), $Extension)
}

function Show-InfoDialog {
    param(
        [Parameter(Mandatory)][string] $Message,
        [string] $Title = 'Convert docx to Markdown'
    )
    [System.Windows.Forms.MessageBox]::Show(
        $Message, $Title,
        [System.Windows.Forms.MessageBoxButtons]::OKCancel,
        [System.Windows.Forms.MessageBoxIcon]::Information)
}

<#
    Input folder is empty: explain, then let the user pick files. Selected files are
    COPIED into input\ - the originals are left untouched.
#>
function Request-InputFiles {
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)][string] $InputFolder)

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $intro = @"
The input folder is empty:

    $InputFolder

Click OK to pick the Word documents you want to convert.

The files you select will be COPIED into the input folder - they are NOT moved,
and your originals stay exactly where they are.

Conversion starts as soon as the copy is finished.
"@
    if ((Show-InfoDialog -Message $intro) -ne [System.Windows.Forms.DialogResult]::OK) {
        return @()
    }

    $dialog = [System.Windows.Forms.OpenFileDialog]::new()
    try {
        $dialog.Title = 'Select the .docx files to convert (they will be copied)'
        $dialog.Filter = 'Word documents (*.docx)|*.docx|All files (*.*)|*.*'
        $dialog.Multiselect = $true
        $dialog.InitialDirectory = [Environment]::GetFolderPath('MyDocuments')
        if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
            return @()
        }
        $selected = @($dialog.FileNames)
    }
    finally {
        $dialog.Dispose()
    }

    $copied = [System.Collections.Generic.List[string]]::new()
    foreach ($file in $selected) {
        $base = [System.IO.Path]::GetFileNameWithoutExtension($file)
        $target = Get-AvailablePath -Folder $InputFolder -BaseName $base -Extension '.docx'
        if ($PSCmdlet.ShouldProcess($target, 'Copy source document into input folder')) {
            Copy-Item -LiteralPath $file -Destination $target
            $copied.Add($target)
            Write-Log -Level INFO ('Copied "{0}" -> input\{1}' -f $file, (Split-Path $target -Leaf))
        }
    }
    return $copied.ToArray()
}

#endregion

#region Keyword replacement

function ConvertTo-TitleCase {
    param([Parameter(Mandatory)][AllowEmptyString()][string] $Text)
    $sb = [System.Text.StringBuilder]::new()
    $startOfWord = $true
    foreach ($ch in $Text.ToCharArray()) {
        if ([char]::IsLetter($ch)) {
            [void]$sb.Append($(if ($startOfWord) { [char]::ToUpperInvariant($ch) } else { [char]::ToLowerInvariant($ch) }))
            $startOfWord = $false
        }
        else {
            [void]$sb.Append($ch)
            if (-not [char]::IsDigit($ch) -and $ch -ne '''') { $startOfWord = $true }
        }
    }
    return $sb.ToString()
}

function Test-IsTitleCase {
    param([Parameter(Mandatory)][string] $Text)
    $startOfWord = $true
    $sawUpperStart = $false
    foreach ($ch in $Text.ToCharArray()) {
        if ([char]::IsLetter($ch)) {
            if ($startOfWord) {
                if ([char]::IsLower($ch)) { return $false }
                $sawUpperStart = $true
            }
            elseif ([char]::IsUpper($ch)) { return $false }
            $startOfWord = $false
        }
        else {
            if (-not [char]::IsDigit($ch) -and $ch -ne '''') { $startOfWord = $true }
        }
    }
    return $sawUpperStart
}

<#
    True when the replacement carries capitals that are not simply all-lower, all-upper or plain
    title case - "ConsultantCompanyXYZ", "iPhone", "Company ABC". Reshaping those would destroy
    the exact spelling the author asked for.
#>
function Test-HasDeliberateCasing {
    param([Parameter(Mandatory)][AllowEmptyString()][string] $Text)
    $letters = @($Text.ToCharArray() | Where-Object { [char]::IsLetter($_) })
    if ($letters.Count -eq 0) { return $false }

    $hasLower = @($letters | Where-Object { [char]::IsLower($_) }).Count -gt 0
    $hasUpper = @($letters | Where-Object { [char]::IsUpper($_) }).Count -gt 0
    if (-not $hasUpper -or -not $hasLower) { return $false }
    return -not (Test-IsTitleCase -Text $Text)
}

<#
    Shapes the replacement to match the casing of the text that was actually found,
    so "CONTOSO" -> "FABRIKAM" while "Contoso" -> "Fabrikam" from one lower-case rule.

    A replacement with deliberate internal capitals is used verbatim, so a rule such as
    Contoso -> ConsultantCompanyXYZ never degrades into "Consultantcompanyxyz". ALL CAPS in
    the source still wins, because that is usually a heading or an emphasised mention.
#>
function Get-CaseShapedReplacement {
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string] $Matched,
        [Parameter(Mandatory)][AllowEmptyString()][string] $Replacement
    )
    $letters = @($Matched.ToCharArray() | Where-Object { [char]::IsLetter($_) })
    if ($letters.Count -eq 0) { return $Replacement }

    $hasLower = @($letters | Where-Object { [char]::IsLower($_) }).Count -gt 0
    $hasUpper = @($letters | Where-Object { [char]::IsUpper($_) }).Count -gt 0

    if (-not $hasLower -and $letters.Count -gt 1) { return $Replacement.ToUpperInvariant() }
    if (Test-HasDeliberateCasing -Text $Replacement) { return $Replacement }
    if (-not $hasUpper) { return $Replacement.ToLowerInvariant() }
    if (Test-IsTitleCase -Text $Matched) { return ConvertTo-TitleCase -Text $Replacement }
    return $Replacement
}

function Import-ReplacementRule {
    param([Parameter(Mandatory)][string] $Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Log -Level WARN "No replacements file at '$Path' - no keyword replacement will be done."
        return @()
    }

    $raw = Get-Content -LiteralPath $Path -Raw -Encoding utf8
    if ([string]::IsNullOrWhiteSpace($raw)) { return @() }

    try { $entries = @($raw | ConvertFrom-Json) }
    catch { throw "'$Path' is not valid JSON: $($_.Exception.Message)" }

    $rules = [System.Collections.Generic.List[psobject]]::new()
    $index = 0
    foreach ($entry in $entries) {
        $index++
        $props = $entry.PSObject.Properties.Name
        if ('find' -notin $props -or [string]::IsNullOrEmpty($entry.find)) {
            Write-Log -Level WARN "Entry $index in replacements.json has no 'find' value - skipped."
            continue
        }
        if ('replace' -notin $props) {
            Write-Log -Level WARN "Entry $index ('$($entry.find)') has no 'replace' value - skipped."
            continue
        }

        # Optional flags, all defaulting to the documented behaviour.
        # preserveCase is opt-in: by default the replacement is written exactly as spelled here.
        $isRegex     = if ('regex' -in $props)          { [bool]$entry.regex }          else { $false }
        $wholeWord   = if ('wholeWord' -in $props)      { [bool]$entry.wholeWord }      else { $true }
        $ignoreCase  = if ('caseInsensitive' -in $props){ [bool]$entry.caseInsensitive }else { $true }
        $preserve    = if ('preserveCase' -in $props)   { [bool]$entry.preserveCase }   else { $false }

        $pattern = if ($isRegex) { $entry.find } else { [regex]::Escape($entry.find) }
        if ($wholeWord -and -not $isRegex) {
            # Lookarounds rather than \b, so rules that start or end with punctuation still work.
            $pattern = '(?<!\w){0}(?!\w)' -f $pattern
        }
        $options = [System.Text.RegularExpressions.RegexOptions]::None
        if ($ignoreCase) { $options = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase }

        # Same match timeout as the built-in patterns: a pathological user pattern fails the
        # file it hangs on instead of freezing the whole run.
        try { $rx = [regex]::new($pattern, $options, $script:RegexTimeout) }
        catch { throw "Entry $index ('$($entry.find)') is not a valid pattern: $($_.Exception.Message)" }

        $rules.Add([pscustomobject]@{
            Find         = $entry.find
            Replace      = $entry.replace
            Regex        = $rx
            PreserveCase = $preserve
        })
    }
    return $rules.ToArray()
}

function Invoke-KeywordReplacement {
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string] $Text,
        [Parameter(Mandatory)][AllowEmptyCollection()][psobject[]] $Rules
    )
    $counts = [ordered]@{}
    foreach ($rule in $Rules) {
        # Rebuilt by hand rather than with a MatchEvaluator: a scriptblock passed to .NET as a
        # delegate cannot call back into this script's functions.
        $found = $rule.Regex.Matches($Text)
        if ($found.Count -eq 0) { continue }

        # Bare URLs are kept verbatim, so replacing inside one would only corrupt it -
        # "contoso.com" must not become "Company ABC.com". Recomputed per rule because each
        # pass shifts offsets.
        $urls = $script:RxBareUrl.Matches($Text)

        $sb = [System.Text.StringBuilder]::new()
        $position = 0
        $hits = 0
        foreach ($m in $found) {
            [void]$sb.Append($Text.Substring($position, $m.Index - $position))

            $insideUrl = $false
            foreach ($u in $urls) {
                if ($m.Index -ge $u.Index -and $m.Index -lt ($u.Index + $u.Length)) { $insideUrl = $true; break }
            }

            if ($insideUrl) {
                [void]$sb.Append($m.Value)
            }
            elseif ($rule.PreserveCase) {
                [void]$sb.Append((Get-CaseShapedReplacement -Matched $m.Value -Replacement $rule.Replace))
                $hits++
            }
            else {
                [void]$sb.Append($rule.Replace)
                $hits++
            }
            $position = $m.Index + $m.Length
        }
        [void]$sb.Append($Text.Substring($position))

        $Text = $sb.ToString()
        if ($hits -gt 0) { $counts[$rule.Find] = $hits }
    }
    return [pscustomobject]@{ Text = $Text; Counts = $counts }
}

#endregion

#region Raw HTML handling

<#
    Removes the empty anchor markup Word bookmarks leave behind, e.g.
    <span id="_Toc232673667" class="anchor"></span> or []{#_Toc232673667 .anchor}
#>
function Remove-HtmlAnchors {
    param([Parameter(Mandatory)][AllowEmptyString()][string] $Text)

    $count = $script:RxAnchorSpan.Matches($Text).Count +
             $script:RxAnchorTag.Matches($Text).Count +
             $script:RxAnchorAttr.Matches($Text).Count
    if ($count -eq 0) { return [pscustomobject]@{ Text = $Text; Count = 0 } }

    $Text = $script:RxAnchorSpan.Replace($Text, '')
    $Text = $script:RxAnchorTag.Replace($Text, '')
    $Text = $script:RxAnchorAttr.Replace($Text, '')
    return [pscustomobject]@{ Text = $Text; Count = $count }
}

<#
    Flattens the HTML inside one table cell to inline Markdown. <a> keeps its text and loses
    its href, matching how Markdown links are treated elsewhere.
#>
function ConvertFrom-HtmlCell {
    param([Parameter(Mandatory)][AllowEmptyString()][string] $Html)

    $t = $Html
    $t = $script:RxBrTag.Replace($t, ' ')
    $t = $script:RxParaBreak.Replace($t, ' ')
    $t = $script:RxStrongTag.Replace($t, '**')
    $t = $script:RxEmTag.Replace($t, '*')
    $t = $script:RxCodeTag.Replace($t, '`')
    $t = $script:RxHtmlImage.Replace($t, '')
    $t = $script:RxAnyTag.Replace($t, '')                       # drop every remaining tag
    $t = [System.Net.WebUtility]::HtmlDecode($t)
    $t = $t -replace [char]0x00A0, ' '
    $t = $t -replace '\|', '\|'                                 # a literal pipe would split the cell
    $t = [regex]::Replace($t, '\s+', ' ')
    return $t.Trim()
}

function Get-HtmlAttributeInt {
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string] $Attributes,
        [Parameter(Mandatory)][regex] $Pattern
    )
    $m = $Pattern.Match($Attributes)
    if ($m.Success) { return [Math]::Max(1, [int]$m.Groups[1].Value) }
    return 1
}

<#
    Turns one raw <table> block into a GFM pipe table. Merged cells are flattened: the content
    stays in the first cell of the span and the cells it covered are left blank.
#>
function ConvertFrom-HtmlTable {
    param([Parameter(Mandatory)][string] $Html)

    $headerHtml = ''
    $theadMatch = $script:RxTheadRow.Match($Html)
    if ($theadMatch.Success) { $headerHtml = $theadMatch.Value }

    $grid = [System.Collections.Generic.List[System.Collections.Generic.List[string]]]::new()
    $headerRowCount = 0
    $carry = @{}   # column index -> number of further rows still covered by a rowspan

    foreach ($rowMatch in $script:RxTableRow.Matches($Html)) {
        $cells = [System.Collections.Generic.List[string]]::new()
        $col = 0

        foreach ($cellMatch in $script:RxTableCell.Matches($rowMatch.Groups['body'].Value)) {
            while ($carry.ContainsKey($col) -and $carry[$col] -gt 0) {
                $cells.Add(''); $carry[$col]--; $col++
            }
            $colspan = Get-HtmlAttributeInt -Attributes $cellMatch.Groups['attrs'].Value -Pattern $script:RxColspan
            $rowspan = Get-HtmlAttributeInt -Attributes $cellMatch.Groups['attrs'].Value -Pattern $script:RxRowspan

            $cells.Add((ConvertFrom-HtmlCell -Html $cellMatch.Groups['body'].Value))
            if ($rowspan -gt 1) { $carry[$col] = $rowspan - 1 }
            $col++
            for ($i = 1; $i -lt $colspan; $i++) {
                $cells.Add('')
                if ($rowspan -gt 1) { $carry[$col] = $rowspan - 1 }
                $col++
            }
        }
        while ($carry.ContainsKey($col) -and $carry[$col] -gt 0) {
            $cells.Add(''); $carry[$col]--; $col++
        }

        if ($cells.Count -eq 0) { continue }
        $grid.Add($cells)
        if ($headerHtml -and $headerHtml.Contains($rowMatch.Value)) { $headerRowCount++ }
    }

    if ($grid.Count -eq 0) { return '' }

    # No <thead>: treat an all-<th> first row as the header.
    if ($headerRowCount -eq 0) {
        $firstRowBody = $script:RxTableRow.Match($Html).Groups['body'].Value
        $tags = @($script:RxTableCell.Matches($firstRowBody) | ForEach-Object { $_.Groups['tag'].Value.ToLowerInvariant() })
        if ($tags.Count -gt 0 -and ($tags | Where-Object { $_ -ne 'th' }).Count -eq 0) { $headerRowCount = 1 }
    }

    $width = 0
    foreach ($row in $grid) { if ($row.Count -gt $width) { $width = $row.Count } }
    foreach ($row in $grid) { while ($row.Count -lt $width) { $row.Add('') } }

    # Word tables rarely mark a header row, but their first row almost always is one.
    if ($headerRowCount -eq 0) { $headerRowCount = 1 }

    $header = $grid[0]
    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add('| ' + ($header -join ' | ') + ' |')
    $lines.Add('|' + (@(1..$width | ForEach-Object { '---' }) -join '|') + '|')
    for ($r = $headerRowCount; $r -lt $grid.Count; $r++) {
        $lines.Add('| ' + ($grid[$r] -join ' | ') + ' |')
    }

    $caption = $script:RxTableCaption.Match($Html)
    $result = $lines -join "`n"
    if ($caption.Success) {
        $text = ConvertFrom-HtmlCell -Html $caption.Groups['text'].Value
        if ($text) { $result = "**$text**`n`n" + $result }
    }
    return "`n" + $result + "`n"
}

<#
    Cleans one stretch of text that is known not to be code. Formatting that Markdown can express
    is converted; everything else is unwrapped so the words survive without the markup.
#>
function Convert-InlineHtmlChunk {
    param([Parameter(Mandatory)][AllowEmptyString()][string] $Chunk)

    # Fast path: most lines contain no candidate markup at all. Every pattern below needs a
    # '<' or a '[', so two IndexOf calls skip the six regex scans for the common case.
    if ($Chunk.IndexOf('<') -lt 0 -and $Chunk.IndexOf('[') -lt 0) {
        return [pscustomobject]@{ Text = $Chunk; Count = 0 }
    }

    $count = $script:RxUnwrapTag.Matches($Chunk).Count +
             $script:RxStrongTag.Matches($Chunk).Count +
             $script:RxEmTag.Matches($Chunk).Count +
             $script:RxCodeTag.Matches($Chunk).Count +
             $script:RxBrTag.Matches($Chunk).Count +
             $script:RxAttrSpan.Matches($Chunk).Count
    if ($count -eq 0) { return [pscustomobject]@{ Text = $Chunk; Count = 0 } }

    $Chunk = $script:RxStrongTag.Replace($Chunk, '**')
    $Chunk = $script:RxEmTag.Replace($Chunk, '*')
    $Chunk = $script:RxCodeTag.Replace($Chunk, '`')
    $Chunk = $script:RxBrTag.Replace($Chunk, ' ')
    $Chunk = $script:RxUnwrapTag.Replace($Chunk, '')
    $Chunk = $script:RxAttrSpan.Replace($Chunk, '${text}')
    return [pscustomobject]@{ Text = $Chunk; Count = $count }
}

<#
    Removes leftover inline HTML from the whole document, skipping fenced code blocks and inline
    code spans so code samples containing angle brackets are never touched.
#>
function Remove-InlineHtml {
    param([Parameter(Mandatory)][AllowEmptyString()][string] $Text)

    $out = [System.Collections.Generic.List[string]]::new()
    $count = 0
    $fenceMarker = ''   # '' = not inside a fence; otherwise the marker that opened it

    foreach ($line in ($Text -split "`r?`n")) {
        $fence = $script:RxCodeFence.Match($line)
        if ($fence.Success) {
            if (-not $fenceMarker) {
                $fenceMarker = $fence.Groups[1].Value        # opening ``` or ~~~
            }
            elseif ($fence.Groups[1].Value -eq $fenceMarker) {
                $fenceMarker = ''                            # only its own kind closes it
            }
            $out.Add($line)
            continue
        }
        if ($fenceMarker) { $out.Add($line); continue }

        $sb = [System.Text.StringBuilder]::new()
        $position = 0
        foreach ($m in $script:RxInlineCode.Matches($line)) {
            $clean = Convert-InlineHtmlChunk -Chunk $line.Substring($position, $m.Index - $position)
            [void]$sb.Append($clean.Text); $count += $clean.Count
            [void]$sb.Append($m.Value)          # inline code passes through verbatim
            $position = $m.Index + $m.Length
        }
        $clean = Convert-InlineHtmlChunk -Chunk $line.Substring($position)
        [void]$sb.Append($clean.Text); $count += $clean.Count

        $out.Add($sb.ToString())
    }
    return [pscustomobject]@{ Text = ($out -join "`n"); Count = $count }
}

function ConvertTo-MarkdownTables {
    param([Parameter(Mandatory)][AllowEmptyString()][string] $Text)

    $found = $script:RxHtmlTable.Matches($Text)
    if ($found.Count -eq 0) { return [pscustomobject]@{ Text = $Text; Count = 0 } }

    $sb = [System.Text.StringBuilder]::new()
    $position = 0
    foreach ($m in $found) {
        [void]$sb.Append($Text.Substring($position, $m.Index - $position))
        [void]$sb.Append((ConvertFrom-HtmlTable -Html $m.Value))
        $position = $m.Index + $m.Length
    }
    [void]$sb.Append($Text.Substring($position))
    return [pscustomobject]@{ Text = $sb.ToString(); Count = $found.Count }
}

<#
    Pandoc emits an empty header row for Word tables, which renders as a blank strip. When the
    header is entirely empty, promote the first body row into it.
#>
function Optimize-MarkdownTables {
    param([Parameter(Mandatory)][AllowEmptyString()][string] $Text)

    $lines = @($Text -split "`r?`n")
    $out = [System.Collections.Generic.List[string]]::new()
    $promoted = 0
    $i = 0

    while ($i -lt $lines.Count) {
        $isHeader = $lines[$i] -match '^\s*\|.*\|\s*$'
        $hasRule = ($i + 2) -lt $lines.Count -and $lines[$i + 1] -match '^\s*\|[\s\-:|]+\|\s*$'
        $emptyHeader = $isHeader -and ($lines[$i] -replace '[\s|]', '') -eq ''
        $nextIsRow = $hasRule -and $lines[$i + 2] -match '^\s*\|.*\|\s*$'

        if ($isHeader -and $hasRule -and $emptyHeader -and $nextIsRow) {
            $out.Add($lines[$i + 2])   # first body row becomes the header
            $out.Add($lines[$i + 1])
            $i += 3
            $promoted++
            continue
        }
        $out.Add($lines[$i])
        $i++
    }
    return [pscustomobject]@{ Text = ($out -join "`n"); Count = $promoted }
}

#endregion

#region Markdown post-processing

function Remove-MarkdownImages {
    param([Parameter(Mandatory)][AllowEmptyString()][string] $Text)

    $figures = $script:RxFigure.Matches($Text).Count
    if ($figures -gt 0) { $Text = $script:RxFigure.Replace($Text, '') }

    $count = $figures + $script:RxImage.Matches($Text).Count + $script:RxHtmlImage.Matches($Text).Count
    if ($count -eq 0) { return [pscustomobject]@{ Text = $Text; Count = 0 } }

    $lines = $Text -split "`r?`n"
    $kept = [System.Collections.Generic.List[string]]::new()
    foreach ($line in $lines) {
        $hadContent = -not [string]::IsNullOrWhiteSpace($line)
        $stripped = $script:RxImage.Replace($line, '')
        $stripped = $script:RxHtmlImage.Replace($stripped, '')
        # A line that only ever held an image disappears instead of leaving a blank gap.
        if ($hadContent -and [string]::IsNullOrWhiteSpace($stripped) -and $stripped -ne $line) { continue }
        $kept.Add($stripped)
    }
    return [pscustomobject]@{ Text = ($kept -join "`n"); Count = $count }
}

function Remove-MarkdownLinkUrls {
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string] $Text,
        [switch] $AlsoStripBareUrls
    )
    $count = $script:RxLink.Matches($Text).Count + $script:RxAutolink.Matches($Text).Count

    $Text = $script:RxLink.Replace($Text, '${text}')
    $Text = $script:RxAutolink.Replace($Text, '$1')

    if ($AlsoStripBareUrls) {
        $bare = $script:RxBareUrl.Matches($Text).Count
        if ($bare -gt 0) {
            $count += $bare
            $Text = $script:RxBareUrl.Replace($Text, '')
        }
    }
    return [pscustomobject]@{ Text = $Text; Count = $count }
}

function Remove-MarkdownFootnotes {
    param([Parameter(Mandatory)][AllowEmptyString()][string] $Text)

    $lines = $Text -split "`r?`n"
    $kept = [System.Collections.Generic.List[string]]::new()
    $count = 0
    $inDefinition = $false

    foreach ($line in $lines) {
        if ($script:RxFootnoteDef.IsMatch($line)) {
            $inDefinition = $true
            $count++
            continue
        }
        if ($inDefinition) {
            # Continuation lines of a definition are indented; a blank line may still be inside it.
            if ([string]::IsNullOrWhiteSpace($line) -or $line -match '^[ \t]') { continue }
            $inDefinition = $false
        }
        $kept.Add($script:RxFootnote.Replace($line, ''))
    }
    return [pscustomobject]@{ Text = ($kept -join "`n"); Count = $count }
}

function Optimize-MarkdownWhitespace {
    param([Parameter(Mandatory)][AllowEmptyString()][string] $Text)

    $Text = $Text -replace "`r`n", "`n"
    $Text = $Text -replace [char]0x00A0, ' '           # non-breaking spaces Word loves to emit
    $Text = ($Text -split "`n" | ForEach-Object { $_.TrimEnd() }) -join "`n"
    $Text = $Text -replace "`n{3,}", "`n`n"
    return $Text.Trim() + "`n"
}

#endregion

#region Conversion

function Convert-OneDocument {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][System.IO.FileInfo] $File,
        [Parameter(Mandatory)][string] $OutputFolder,
        [Parameter(Mandatory)][AllowEmptyCollection()][psobject[]] $Rules
    )

    $temp = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "docx2md-$([guid]::NewGuid()).md")
    try {
        # No --standalone => no YAML metadata block. No --extract-media => no image files on disk.
        $pandocArgs = @(
            '--from=docx'
            '--to=gfm'
            '--wrap=none'
            '--track-changes=accept'
            '--output', $temp
            '--', $File.FullName
        )

        # Run pandoc with a hard timeout so a document that hangs it fails alone instead of
        # freezing the batch. ProcessStartInfo.ArgumentList quotes each argument correctly.
        $psi = [System.Diagnostics.ProcessStartInfo]::new($script:PandocPath)
        foreach ($a in $pandocArgs) { $psi.ArgumentList.Add($a) }
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true
        $psi.RedirectStandardError = $true

        $proc = [System.Diagnostics.Process]::Start($psi)
        $stderrTask = $proc.StandardError.ReadToEndAsync()   # async read prevents pipe deadlock
        if (-not $proc.WaitForExit($PandocTimeoutSec * 1000)) {
            try { $proc.Kill($true) } catch { }
            [void]$proc.WaitForExit(5000)
            throw "pandoc did not finish within ${PandocTimeoutSec}s and was stopped (see -PandocTimeoutSec)"
        }
        $stderr = $stderrTask.Result
        if ($proc.ExitCode -ne 0) {
            throw "pandoc exited with code $($proc.ExitCode): $($stderr.Trim())"
        }
        if ($stderr.Trim()) {
            Write-Log -Level WARN -NoConsole ('pandoc: {0}' -f $stderr.Trim())
        }

        $markdown = [System.IO.File]::ReadAllText($temp, [System.Text.UTF8Encoding]::new($false))
    }
    finally {
        if (Test-Path -LiteralPath $temp) {
            Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue -WhatIf:$false
        }
    }

    # Anchors first, so leftover bookmark markup never ends up inside a converted table cell.
    $anchors = Remove-HtmlAnchors -Text $markdown
    $markdown = $anchors.Text

    # Then raw HTML tables -> pipe tables, before link/image stripping runs over the result.
    $tables = ConvertTo-MarkdownTables -Text $markdown
    $markdown = $tables.Text

    $images = Remove-MarkdownImages -Text $markdown
    $markdown = $images.Text

    $links = Remove-MarkdownLinkUrls -Text $markdown -AlsoStripBareUrls:$StripBareUrls
    $markdown = $links.Text

    # Safety net: anything still carrying raw inline HTML loses the markup but keeps its words.
    $html = Remove-InlineHtml -Text $markdown
    $markdown = $html.Text

    $footnoteCount = 0
    if (-not $KeepFootnotes) {
        $footnotes = Remove-MarkdownFootnotes -Text $markdown
        $markdown = $footnotes.Text
        $footnoteCount = $footnotes.Count
    }

    $headers = Optimize-MarkdownTables -Text $markdown
    $markdown = $headers.Text

    $keywords = Invoke-KeywordReplacement -Text $markdown -Rules $Rules
    $markdown = Optimize-MarkdownWhitespace -Text $keywords.Text

    $target = Get-AvailablePath -Folder $OutputFolder -BaseName $File.BaseName -Extension '.md'
    if ($PSCmdlet.ShouldProcess($target, 'Write Markdown')) {
        [System.IO.File]::WriteAllText($target, $markdown, [System.Text.UTF8Encoding]::new($false))
    }

    return [pscustomobject]@{
        Source        = $File.FullName
        Target        = $target
        Images        = $images.Count
        Links         = $links.Count
        Footnotes     = $footnoteCount
        Anchors       = $anchors.Count
        Tables        = $tables.Count
        Headers       = $headers.Count
        Html          = $html.Count
        KeywordCounts = $keywords.Counts
    }
}

#endregion

#region Main

$script:LogPath = $null

if (-not $Root) {
    $Root = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
}
$resolved = Resolve-Path -LiteralPath $Root -ErrorAction SilentlyContinue
if ($resolved) { $Root = $resolved.Path }

$folders = Initialize-Workspace -Path $Root
$script:LogPath = Join-Path $folders.Logs ('convert-{0}.log' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))

# Retention: keep the newest 30 run logs (names sort chronologically thanks to the timestamp).
Get-ChildItem -LiteralPath $folders.Logs -Filter 'convert-*.log' -File |
    Sort-Object Name -Descending |
    Select-Object -Skip 29 |
    Remove-Item -Force -WhatIf:$false -Confirm:$false -ErrorAction SilentlyContinue

Write-Host ''
Write-Host ('Convert-DocxToMd v{0}' -f $script:ScriptVersion) -ForegroundColor Cyan
Write-Host ('-' * 60) -ForegroundColor DarkGray
Write-Log -Level INFO -NoConsole ('Version   : {0}' -f $script:ScriptVersion)
Write-Log -Level INFO -NoConsole ('PowerShell: {0}' -f $PSVersionTable.PSVersion)
Write-Log -Level INFO -NoConsole ("Workspace : {0}" -f $folders.Root)

try {
    $script:PandocPath = Get-PandocPath
    Write-Log -Level INFO -NoConsole ('pandoc    : {0}' -f $script:PandocPath)

    if (-not $ReplacementsPath) { $ReplacementsPath = Join-Path $folders.Root 'replacements.json' }
    $rules = @(Import-ReplacementRule -Path $ReplacementsPath)
    Write-Host ('Rules     : {0} keyword replacement(s)' -f $rules.Count) -ForegroundColor Gray
    Write-Log -Level INFO -NoConsole ('Rules     : {0} from {1}' -f $rules.Count, $ReplacementsPath)

    $documents = @(Get-InputDocuments -InputFolder $folders.Input)

    if ($documents.Count -eq 0) {
        if ($NoPrompt) {
            Write-Log -Level WARN ('No .docx files found in "{0}". Add files and run again.' -f $folders.Input)
            exit 0
        }
        $copied = @(Request-InputFiles -InputFolder $folders.Input)
        if ($copied.Count -eq 0) {
            Write-Log -Level WARN ('Nothing selected. Put your .docx files in "{0}" and run again.' -f $folders.Input)
            exit 0
        }
        $documents = @(Get-InputDocuments -InputFolder $folders.Input)
    }

    Write-Host ('Documents : {0}' -f $documents.Count) -ForegroundColor Gray
    Write-Host ('-' * 60) -ForegroundColor DarkGray

    $ok = 0; $failed = 0
    $totals = [ordered]@{ Images = 0; Links = 0; Footnotes = 0; Keywords = 0; Anchors = 0; Tables = 0; Html = 0 }

    foreach ($doc in $documents) {
        try {
            $result = Convert-OneDocument -File $doc -OutputFolder $folders.Output -Rules $rules

            $keywordHits = 0
            foreach ($v in $result.KeywordCounts.Values) { $keywordHits += $v }
            $totals.Images    += $result.Images
            $totals.Links     += $result.Links
            $totals.Footnotes += $result.Footnotes
            $totals.Keywords  += $keywordHits
            $totals.Anchors   += $result.Anchors
            $totals.Tables    += $result.Tables
            $totals.Html      += $result.Html
            $ok++

            $detail = 'links {0}, images {1}, footnotes {2}, anchors {3}, html tables {4}, html tags {5}, keywords {6}' -f
                $result.Links, $result.Images, $result.Footnotes, $result.Anchors, $result.Tables, $result.Html, $keywordHits
            Write-Host ('  OK    {0}' -f $doc.Name) -ForegroundColor Green -NoNewline
            Write-Host ('  ->  output\{0}  ({1})' -f (Split-Path $result.Target -Leaf), $detail) -ForegroundColor DarkGray

            Write-Log -Level OK -NoConsole ('{0} -> {1} | {2}' -f $doc.Name, (Split-Path $result.Target -Leaf), $detail)
            foreach ($pair in $result.KeywordCounts.GetEnumerator()) {
                Write-Log -Level INFO -NoConsole ('    keyword "{0}" replaced {1}x' -f $pair.Key, $pair.Value)
            }
        }
        catch {
            $failed++
            Write-Host ('  FAIL  {0}' -f $doc.Name) -ForegroundColor Red -NoNewline
            Write-Host ('  {0}' -f $_.Exception.Message) -ForegroundColor DarkGray
            Write-Log -Level ERROR -NoConsole ('{0} : {1}' -f $doc.Name, $_.Exception.Message)
        }
    }

    Write-Host ('-' * 60) -ForegroundColor DarkGray
    Write-Host ('Converted {0} of {1} document(s){2}' -f $ok, $documents.Count,
        $(if ($failed) { " - $failed failed" } else { '' })) -ForegroundColor $(if ($failed) { 'Yellow' } else { 'Green' })
    Write-Host ('Removed {0} link URL(s), {1} image(s), {2} footnote(s), {3} anchor(s)' -f
        $totals.Links, $totals.Images, $totals.Footnotes, $totals.Anchors) -ForegroundColor Gray
    Write-Host ('Converted {0} HTML table(s), unwrapped {1} HTML tag(s); replaced {2} keyword occurrence(s)' -f
        $totals.Tables, $totals.Html, $totals.Keywords) -ForegroundColor Gray
    Write-Host ('Output    : {0}' -f $folders.Output) -ForegroundColor Gray
    Write-Host ('Log       : {0}' -f $script:LogPath) -ForegroundColor DarkGray
    Write-Host ''

    Write-Log -Level INFO -NoConsole ('Done. ok={0} failed={1} links={2} images={3} footnotes={4} anchors={5} tables={6} html={7} keywords={8}' -f
        $ok, $failed, $totals.Links, $totals.Images, $totals.Footnotes, $totals.Anchors, $totals.Tables, $totals.Html, $totals.Keywords)

    exit $(if ($failed) { 1 } else { 0 })
}
catch {
    Write-Host ''
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Log -Level ERROR -NoConsole ('Fatal: {0}' -f $_.Exception.ToString())
    exit 1
}

#endregion
