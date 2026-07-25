<#
.SYNOPSIS
    Converts .pdf files to clean Markdown: strips hyperlink URLs, removes images,
    and applies case-aware keyword replacements.

.DESCRIPTION
    Companion to Convert-DocxToMd.ps1. Shares the same folder structure, the same
    replacements.json and the same Markdown cleanup pipeline:

        input\    drop .pdf files here (a file picker offers to copy them in if empty)
        output\   .md files land here, never overwritten
        logs\     one timestamped log per run (pdfconvert-*.log)

    pandoc cannot read PDF, so this script needs an extraction front-end. Two engines:

      Word  Microsoft Word opens the PDF and reflows it into a real .docx, which pandoc
            then converts. Headings, tables, lists and bold survive, and the entire
            docx cleanup pipeline applies unchanged. This is the default and by far the
            better result. Word must be installed and NOT already open (see below).

      Text  Xpdf pdftotext extracts plain text, which is then tidied into Markdown:
            page furniture removed, hyphenation joined, numbered headings detected.
            Always available, but there are no tables and no bold or italic.

    -Engine Auto (the default) tries Word and falls back to Text.

    IMPORTANT - close Word before running. Word is a single-instance COM server, so
    automating it while you have documents open would attach to YOUR session. To protect
    your unsaved work this script refuses to use the Word engine whenever any WINWORD
    process is already running, and falls back to Text instead.

    There is no OCR on this machine, so scanned or image-only PDFs cannot be converted.
    They are detected up front and reported as failures rather than producing an empty file.

.PARAMETER Root
    Workspace folder. Defaults to the folder containing this script.

.PARAMETER ReplacementsPath
    Path to the keyword mapping file. Defaults to <Root>\replacements.json.

.PARAMETER Engine
    Auto (default), Word or Text. See DESCRIPTION.

.PARAMETER NoPrompt
    Never show dialogs. Use for scheduled/unattended runs.

.PARAMETER KeepFootnotes
    Keep footnotes and endnotes instead of removing them.

.PARAMETER StripBareUrls
    Also remove bare URLs typed as plain text (off by default).

.PARAMETER PandocTimeoutSec
    Maximum seconds one document may spend inside pandoc (default 120).

.PARAMETER WordTimeoutSec
    Maximum seconds one document may spend inside Word (default 300). Reflowing a large
    PDF is slow; on timeout the Word process is killed and that file fails alone.

.PARAMETER TextTimeoutSec
    Maximum seconds one document may spend inside pdftotext (default 120).

.PARAMETER TextLayout
    pdftotext extraction mode: Layout (default), Table, Simple or Raw.

.PARAMETER PdfToTextPath
    Full path to pdftotext.exe. Use when it is not on PATH. Note this is the command line tool
    from the "Xpdf command line tools" download, not xpdf.exe (the XpdfReader GUI), which cannot
    extract text to a file. When omitted, PATH is searched first, then the usual install folders.

.PARAMETER MinCharsPerPage
    Below this many extractable characters per page a PDF is treated as scanned and
    fails with a clear message (default 50). Use 0 to disable the check.

.PARAMETER KeepRunningHeaders
    Keep page headers and footers that repeat across pages (Text engine).

.PARAMETER KeepPageNumbers
    Keep lines that contain nothing but a page number (Text engine).

.PARAMETER KeepLineHyphens
    Keep words hyphenated across a line break instead of rejoining them (Text engine).

.PARAMETER ReflowParagraphs
    Join wrapped lines into single paragraphs (Text engine). Off by default: Markdown
    already renders consecutive lines as one paragraph, so reflowing can only do harm.

.PARAMETER DetectCapsHeadings
    Treat short ALL-CAPS lines as headings (Text engine). Off by default - high
    false-positive rate on acronym-heavy documents.

.PARAMETER FenceLayoutBlocks
    Wrap column-aligned blocks (tables, ASCII art) in code fences so their alignment
    survives rendering (Text engine).

.PARAMETER TextStrictPipeline
    Run the full docx cleanup chain on Text-engine output too. Off by default, because
    those stages expect pandoc-generated markup and would match literal document text.

.PARAMETER AllowRunningWord
    DANGEROUS. Use the Word engine even when Word is already running. Automation would
    then attach to your own Word session and closing it discards unsaved work.

.PARAMETER WordWorker
    Internal. Re-entrant worker role; requires the PDF2MD_* environment variables.

.PARAMETER Version
    Print the script version and exit.

.EXAMPLE
    .\Convert-PdfToMd.ps1

.EXAMPLE
    .\Convert-PdfToMd.ps1 -Engine Text -FenceLayoutBlocks -Verbose

.NOTES
    Version history
      1.0.0  Initial release. Word-reflow and pdftotext engines feeding the Markdown
             cleanup pipeline of Convert-DocxToMd 1.4.0.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string] $Root,
    [string] $ReplacementsPath,

    [ValidateSet('Auto', 'Word', 'Text')]
    [string] $Engine = 'Auto',

    [switch] $NoPrompt,
    [switch] $KeepFootnotes,
    [switch] $StripBareUrls,

    [ValidateRange(1, 3600)][int] $PandocTimeoutSec = 120,
    [ValidateRange(1, 3600)][int] $WordTimeoutSec = 300,
    [ValidateRange(1, 3600)][int] $TextTimeoutSec = 120,

    [ValidateSet('Layout', 'Table', 'Simple', 'Raw')]
    [string] $TextLayout = 'Layout',

    [string] $PdfToTextPath,

    [ValidateRange(0, 10000)][int] $MinCharsPerPage = 50,

    [switch] $KeepRunningHeaders,
    [switch] $KeepPageNumbers,
    [switch] $KeepLineHyphens,
    [switch] $ReflowParagraphs,
    [switch] $DetectCapsHeadings,
    [switch] $FenceLayoutBlocks,
    [switch] $TextStrictPipeline,

    [switch] $AllowRunningWord,
    [switch] $WordWorker,
    [switch] $Version
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:ScriptVersion = '1.0.0'

if ($Version) {
    Write-Host ('Convert-PdfToMd {0}' -f $script:ScriptVersion)
    exit 0
}

#region Word COM worker (child-process role)

<#
    Runs ONLY when this script re-launches itself with -WordWorker. Nothing is passed on
    the command line: every value arrives through the environment block, so no file name
    can ever be parsed as an option or interpolated into source.

        PDF2MD_ROLE               'word-worker'
        PDF2MD_STAGE              staging folder holding in.pdf; receives out.docx
        PDF2MD_ALLOW_RUNNING_WORD '1' to skip the "Word is already open" refusal

    Exit codes: 0 ok | 10 bad invocation | 11 Word unavailable | 12 Word in use
                13 open failed | 14 save failed | 15 reflow produced no text
#>
function Invoke-WordWorker {
    $stage = $env:PDF2MD_STAGE
    if ($env:PDF2MD_ROLE -ne 'word-worker' -or
        [string]::IsNullOrWhiteSpace($stage) -or
        -not (Test-Path -LiteralPath $stage)) {
        [Console]::Error.WriteLine('worker: bad invocation')
        return 10
    }

    $inPdf   = Join-Path $stage 'in.pdf'
    $outDocx = Join-Path $stage 'out.docx'
    $pidTmp  = Join-Path $stage 'word.pid.tmp'
    $pidFile = Join-Path $stage 'word.pid'

    $wdDoNotSaveChanges  = 0
    $wdAlertsNone        = 0
    $msoSecForceDisable  = 3
    $wdOpenFormatAuto    = 0
    $wdFormatXMLDocument = 12
    $wdLeftToRight       = 0

    $word = $null
    $doc = $null
    $ownWord = $false

    try {
        # Word is a SINGLE-INSTANCE COM server: if the user has Word open, CreateObject
        # binds to THEIR process, and Quit() would then discard their unsaved work
        # silently because DisplayAlerts is off. Prove ownership by process-id diff
        # before writing a single property.
        $before = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
        if ($before.Count -gt 0 -and $env:PDF2MD_ALLOW_RUNNING_WORD -ne '1') {
            [Console]::Error.WriteLine('worker: Word is already running')
            return 12
        }

        try { $word = New-Object -ComObject Word.Application }
        catch {
            [Console]::Error.WriteLine('worker: ' + $_.Exception.Message)
            return 11
        }

        $after = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
        $new = @($after | Where-Object { $_ -notin $before })
        if ($new.Count -ne 1) {
            # Bound to a process we did not start. Release without Quit, touch nothing.
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch { }
            $word = $null
            [Console]::Error.WriteLine('worker: bound to a Word instance we did not start')
            return 12
        }
        $ownWord = $true

        # Written atomically so the parent never reads a half-written file.
        [System.IO.File]::WriteAllText($pidTmp, [string]$new[0])
        [System.IO.File]::Move($pidTmp, $pidFile, $true)

        # Per-instance properties only. Application.Options.* are PERSISTED user settings:
        # writing them would permanently change the user's interactive Word, and a kill on
        # timeout means the restore in finally never runs.
        $word.Visible            = $false
        $word.DisplayAlerts      = $wdAlertsNone
        $word.ScreenUpdating     = $false
        $word.AutomationSecurity = $msoSecForceDisable

        # A junk PasswordDocument turns an encrypted PDF into a thrown error instead of a
        # modal password prompt; it is ignored for unencrypted files.
        #
        # The trailing XMLTransform parameter is deliberately not passed. PowerShell 7's COM
        # binder cannot fill it - supplying [Type]::Missing there fails the whole call with
        # "Missing parameter does not have a default value" - and we have no use for it.
        # Encoding is passed as $null for the same reason; a PDF has no text encoding to pick.
        $junk = [guid]::NewGuid().ToString('N')
        $doc = $word.Documents.Open(
            $inPdf,               # FileName
            $false,               # ConfirmConversions - suppresses the "Convert File" dialog
            $true,                # ReadOnly
            $false,               # AddToRecentFiles   - keeps the user's MRU clean
            $junk,                # PasswordDocument
            $junk,                # PasswordTemplate
            $false,               # Revert
            '',                   # WritePasswordDocument
            '',                   # WritePasswordTemplate
            $wdOpenFormatAuto,    # Format
            $null,                # Encoding
            $false,               # Visible
            $false,               # OpenAndRepair - repair prompts, and can run for minutes
            $wdLeftToRight,       # DocumentDirection
            $true)                # NoEncodingDialog

        if ($doc.Content.End -le 1) {
            [Console]::Error.WriteLine('worker: reflow produced no text')
            return 15
        }

        $doc.SaveAs2($outDocx, $wdFormatXMLDocument, $false, '', $false)
        $doc.Saved = $true
        return 0
    }
    catch {
        [Console]::Error.WriteLine('worker: ' + $_.Exception.Message)
        if ($doc) { return 14 }
        return 13
    }
    finally {
        if ($doc) {
            try { $doc.Close($wdDoNotSaveChanges) } catch { }
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc) } catch { }
            $doc = $null
        }
        if ($word) {
            if ($ownWord) {
                try { $word.NormalTemplate.Saved = $true } catch { }   # never rewrite Normal.dotm
                try { $word.Quit($wdDoNotSaveChanges) } catch { }
            }
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) } catch { }
            $word = $null
        }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

# Dispatched before any workspace setup, so the worker never creates logs\, writes a log
# file or runs the retention sweep.
if ($WordWorker) {
    $codes = @(Invoke-WordWorker)
    if ($codes.Count -gt 0) { exit ([int]$codes[-1]) }
    exit 10
}

#endregion

#region Regex patterns
# Verbatim from Convert-DocxToMd.ps1 v1.4.0 - keep in sync.

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

#region Regex patterns - Text engine only

# A run of 3+ spaces between two glyphs means pdftotext -layout is holding columns apart.
# Lines like that are tables, ASCII art or multi-column text: destructive stages skip them.
$script:RxLayoutGap    = New-CompiledRegex '\S {3,}\S'
# Non-ASCII characters are written as \u escapes so this file stays pure ASCII and cannot be
# corrupted by a host that guesses the wrong encoding. \u2013 and \u2014 are en and em dashes.
$script:RxPageNumber   = New-CompiledRegex '^\s*(?:[-\u2013\u2014\[(]\s*)?(?:page\s+)?\d{1,4}(?:\s*(?:of|/)\s*\d{1,4})?\s*(?:[-\u2013\u2014\])]\s*)?$' 'IgnoreCase'
$script:RxRomanNumeral = New-CompiledRegex '^\s*[ivxlcdm]{1,7}\s*$' 'IgnoreCase'
$script:RxBulletGlyph  = New-CompiledRegex '^(\s*)[\u2022\u25AA\u25CF\u25E6\u2023\u2219\u00B7\u25A0\u25AB]\s+'
$script:RxOrderedParen = New-CompiledRegex '^(\s*)(\d{1,3})\)\s+'
$script:RxNumHeadMulti = New-CompiledRegex '^(\d{1,3}(?:\.\d{1,3})+)\.?\s+(\S.*)$'
$script:RxNumHeadOne   = New-CompiledRegex '^(\d{1,3})\.\s+(\S.*)$'
$script:RxListItem     = New-CompiledRegex '^\s*(?:[-*+]|\d{1,3}[.)])\s+'
$script:RxDigitRun     = New-CompiledRegex '\d+'
$script:RxHyphenEnd    = New-CompiledRegex '(\p{Ll}{2,})-$'

#endregion

#region Helpers
# Write-Log, Initialize-Workspace, Get-PandocPath and Get-AvailablePath are verbatim from
# Convert-DocxToMd.ps1 v1.4.0 - keep in sync.

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

<#
    Folders the Xpdf command line tools commonly unpack into. Searched only when pdftotext is not
    already on PATH, so a working setup is never second-guessed.
#>
function Get-PdfToTextCandidate {
    $roots = @(
        $env:ProgramFiles
        ${env:ProgramFiles(x86)}
        $env:LOCALAPPDATA
        'C:\'
    ) | Where-Object { $_ }

    $patterns = @(
        'xpdf-tools-win64\bin64'
        'xpdf-tools-win32\bin32'
        'Xpdf\bin64'
        'Xpdf\bin32'
        'Glyph & Cog\xpdf-tools-win64\bin64'
        'Glyph & Cog\XpdfCommandLineTools\bin64'
        'Glyph & Cog\XpdfReader-win64'
    )

    $found = [System.Collections.Generic.List[string]]::new()
    foreach ($root in $roots) {
        foreach ($pattern in $patterns) {
            $candidate = Join-Path (Join-Path $root $pattern) 'pdftotext.exe'
            if (Test-Path -LiteralPath $candidate) { $found.Add($candidate) }
        }
    }
    return ,$found
}

<#
    Locates pdftotext and verifies it really is the Xpdf build.

    Two different programs get confused here. pdftotext.exe is the command line extractor this
    script drives; xpdf.exe is the XpdfReader GUI, which cannot write text to a file and is no use
    for conversion. They ship as separate downloads, so the error says which one is missing.

    The vendor check matters as well: poppler ships its own pdftotext with the same name but
    different options and different exit codes, so accepting one silently would produce wrong
    output rather than an obvious failure.
#>
function Get-PdfToTextPath {
    param([string] $Preferred)

    $source = $null
    if ($Preferred) {
        $resolved = Resolve-Path -LiteralPath $Preferred -ErrorAction SilentlyContinue
        if (-not $resolved) { throw "-PdfToTextPath '$Preferred' does not exist." }
        $source = $resolved.Path
        $leaf = [System.IO.Path]::GetFileNameWithoutExtension($source)
        if ($leaf -eq 'xpdf') {
            throw "-PdfToTextPath points at '$source', which is the XpdfReader GUI, not the command line extractor. This script needs pdftotext.exe from the 'Xpdf command line tools' download at https://www.xpdfreader.com/download.html - the reader cannot write text to a file."
        }
        if ($leaf -ne 'pdftotext') {
            throw "-PdfToTextPath points at '$source'. It must be pdftotext.exe."
        }
    }

    if (-not $source) {
        $cmd = Get-Command pdftotext -CommandType Application -ErrorAction SilentlyContinue |
            Select-Object -First 1
        if ($cmd) { $source = $cmd.Source }
    }

    if (-not $source) {
        $candidates = @(Get-PdfToTextCandidate)
        if ($candidates.Count -gt 0) { $source = $candidates[0] }
    }

    if (-not $source) {
        throw "pdftotext.exe was not found on PATH or in the usual install folders.`nIt ships with Git for Windows, or download 'Xpdf command line tools' from https://www.xpdfreader.com/download.html and unpack it (the XpdfReader GUI does not include it).`nThen add its bin64 folder to PATH, or pass -PdfToTextPath <path to pdftotext.exe>."
    }

    $psi = [System.Diagnostics.ProcessStartInfo]::new($source)
    $psi.ArgumentList.Add('-v')
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true

    $banner = ''
    $proc = [System.Diagnostics.Process]::Start($psi)
    try {
        $outTask = $proc.StandardOutput.ReadToEndAsync()
        $errTask = $proc.StandardError.ReadToEndAsync()
        [void]$proc.WaitForExit(15000)
        if ($outTask.Wait(2000)) { $banner += $outTask.Result }
        if ($errTask.Wait(2000)) { $banner += $errTask.Result }
    }
    finally {
        if (-not $proc.HasExited) { try { $proc.Kill($true) } catch { } }
        $proc.Dispose()
    }

    if ($banner -notmatch 'xpdfreader\.com') {
        throw "'$source' is not the Xpdf build of pdftotext (it is probably poppler's). Its options and exit codes differ, so using it would produce wrong output rather than an obvious error. Put the Xpdf build earlier on PATH, or pass -PdfToTextPath <path to the Xpdf pdftotext.exe>."
    }
    return $source
}

<#
    Word's LocalServer32 CLSID must be registered AND the session must be interactive:
    Office automation is unsupported in a non-interactive session (KB 257757) and hangs
    there, which would burn the full timeout on every file.
#>
function Test-WordAvailable {
    $clsid = 'Registry::HKEY_CLASSES_ROOT\CLSID\{000209FF-0000-0000-C000-000000000046}\LocalServer32'
    if (-not (Test-Path -LiteralPath $clsid)) {
        return [pscustomobject]@{ Available = $false; Reason = 'Microsoft Word is not installed' }
    }
    if (-not [Environment]::UserInteractive) {
        return [pscustomobject]@{ Available = $false; Reason = 'session is not interactive, where Word automation is unsupported' }
    }
    if ([string]::IsNullOrWhiteSpace($PSCommandPath)) {
        return [pscustomobject]@{ Available = $false; Reason = 'script path is unknown, so the Word worker cannot be launched' }
    }
    return [pscustomobject]@{ Available = $true; Reason = '' }
}

<#
    Unlike '*.docx', the filter '*.pdf' also matches 8.3 short names - 'report.pdfx' has the
    short name 'REPORT~1.PDF' - so the extension is re-checked explicitly.
#>
function Get-InputDocuments {
    param([Parameter(Mandatory)][string] $InputFolder)
    Get-ChildItem -LiteralPath $InputFolder -Filter '*.pdf' -File |
        Where-Object { $_.Extension -eq '.pdf' -and $_.Name -notlike '~$*' } |
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

<#
    Both scripts write into the same output\ folder, so Get-AvailablePath's Test-Path is only
    advisory - two runs can pick the same name and one would overwrite the other. CreateNew
    makes "never overwrites" atomic: losing the race raises IOException and we pick again.
#>
function Write-MarkdownAtomic {
    param(
        [Parameter(Mandatory)][string] $Folder,
        [Parameter(Mandatory)][string] $BaseName,
        [Parameter(Mandatory)][AllowEmptyString()][string] $Text
    )
    $encoding = [System.Text.UTF8Encoding]::new($false)
    for ($try = 0; $try -lt 50; $try++) {
        $target = Get-AvailablePath -Folder $Folder -BaseName $BaseName -Extension '.md'
        try {
            $stream = [System.IO.File]::Open($target, 'CreateNew', 'Write', 'None')
            try {
                $bytes = $encoding.GetBytes($Text)
                $stream.Write($bytes, 0, $bytes.Length)
            }
            finally { $stream.Dispose() }
            return $target
        }
        catch [System.IO.IOException] { continue }
    }
    throw "could not create a unique output file for '$BaseName'"
}

function Show-InfoDialog {
    param(
        [Parameter(Mandatory)][string] $Message,
        [string] $Title = 'Convert PDF to Markdown'
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

Click OK to pick the PDF files you want to convert.

The files you select will be COPIED into the input folder - they are NOT moved,
and your originals stay exactly where they are.

Conversion starts as soon as the copy is finished.
"@
    if ((Show-InfoDialog -Message $intro) -ne [System.Windows.Forms.DialogResult]::OK) {
        return @()
    }

    $dialog = [System.Windows.Forms.OpenFileDialog]::new()
    try {
        $dialog.Title = 'Select the .pdf files to convert (they will be copied)'
        $dialog.Filter = 'PDF documents (*.pdf)|*.pdf|All files (*.*)|*.*'
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
        $target = Get-AvailablePath -Folder $InputFolder -BaseName $base -Extension '.pdf'
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
# Verbatim from Convert-DocxToMd.ps1 v1.4.0 - keep in sync.

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
# Verbatim from Convert-DocxToMd.ps1 v1.4.0 - keep in sync.

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
# Verbatim from Convert-DocxToMd.ps1 v1.4.0 - keep in sync.

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

<#
    Text-engine variant. RxLink is deliberately not applied: plain text from a PDF has no
    pandoc-generated links, so every [text](url) match would be literal document text -
    "see clause [3](a)" and "Smith [2019](p. 4)" would silently collapse to "3" and "2019".
#>
function Remove-TextLinkUrls {
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string] $Text,
        [switch] $AlsoStripBareUrls
    )
    $count = $script:RxAutolink.Matches($Text).Count
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

#region Staging

<#
    Each document gets its own folder under %TEMP%, holding in.pdf, out.docx, out.txt and
    word.pid. Nothing temporary is ever written into input\ or output\, where the docx script
    would pick it up on its next run.
#>
function New-StageFolder {
    $path = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "pdf2md-$([guid]::NewGuid())")
    New-Item -ItemType Directory -Path $path -Force -WhatIf:$false | Out-Null
    return $path
}

function Remove-StageFolder {
    param([Parameter(Mandatory)][string] $Path)
    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue -WhatIf:$false
    }
}

<#
    A run killed mid-document cannot clean up after itself, so sweep anything older than a day.
#>
function Clear-StaleStage {
    $temp = [System.IO.Path]::GetTempPath()
    $cutoff = (Get-Date).AddHours(-24)
    $stale = @(Get-ChildItem -LiteralPath $temp -Filter 'pdf2md-*' -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt $cutoff })
    foreach ($folder in $stale) {
        Remove-Item -LiteralPath $folder.FullName -Recurse -Force -ErrorAction SilentlyContinue -WhatIf:$false
    }
    return $stale.Count
}

<#
    Copies only the default data stream, so the staged file carries no Zone.Identifier and Word
    will not open it in Protected View. Copy-Item would NOT do - CopyFileEx duplicates alternate
    data streams, Mark-of-the-Web included.

    The source is opened read-only and shared, and neither Word nor pdftotext ever learns the
    original path: no ~$ owner file appears beside it, it is never added to the MRU, and it is
    never locked. The ASCII name also sidesteps Xpdf's unreliable Unicode path handling and its
    lack of a '--' option terminator.
#>
function Copy-PdfToStage {
    param(
        [Parameter(Mandatory)][string] $Source,
        [Parameter(Mandatory)][string] $StageDir
    )
    $target = Join-Path $StageDir 'in.pdf'
    $src = [System.IO.File]::Open($Source, 'Open', 'Read', 'Read')
    try {
        $dst = [System.IO.File]::Create($target)
        try { $src.CopyTo($dst, 1MB) }
        finally { $dst.Dispose() }
    }
    finally { $src.Dispose() }
    return $target
}

#endregion

#region PDF text extraction and probing

$script:PdfToTextModes = @{
    Layout = '-layout'
    Table  = '-table'
    Simple = '-simple'
    Raw    = '-raw'
}

function Get-PdfToTextMessage {
    param([Parameter(Mandatory)][int] $ExitCode, [AllowEmptyString()][string] $Stderr)

    $detail = if ([string]::IsNullOrWhiteSpace($Stderr)) { '' } else { ": $($Stderr.Trim())" }
    switch ($ExitCode) {
        1  { return "the PDF could not be opened - it is corrupt, or encrypted with a password$detail" }
        2  { return "pdftotext could not write its output file$detail" }
        3  { return "the PDF's permissions forbid text extraction$detail" }
        99 { return "pdftotext failed$detail" }
        default { return "pdftotext exited with code $ExitCode$detail" }
    }
}

<#
    Runs pdftotext against the staged in.pdf. Page breaks are deliberately kept (no -nopgbrk):
    the text layer needs them to find running headers, footers and page numbers.
#>
function Invoke-PdfToText {
    param(
        [Parameter(Mandatory)][string] $StageDir,
        [Parameter(Mandatory)][string] $Mode,
        [Parameter(Mandatory)][int] $TimeoutSec
    )

    $outFile = Join-Path $StageDir 'out.txt'
    if (Test-Path -LiteralPath $outFile) {
        Remove-Item -LiteralPath $outFile -Force -WhatIf:$false
    }

    $psi = [System.Diagnostics.ProcessStartInfo]::new($script:PdfToTextExe)
    foreach ($a in @($script:PdfToTextModes[$Mode], '-enc', 'UTF-8', '-eol', 'unix', 'in.pdf', 'out.txt')) {
        $psi.ArgumentList.Add($a)
    }
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardError = $true
    # Relative names plus a working directory: pdftotext never sees a user-controlled path,
    # which matters because Xpdf has no '--' terminator to stop option parsing.
    $psi.WorkingDirectory = $StageDir

    $proc = [System.Diagnostics.Process]::Start($psi)
    try {
        $errTask = $proc.StandardError.ReadToEndAsync()   # async read prevents pipe deadlock
        if (-not $proc.WaitForExit($TimeoutSec * 1000)) {
            try { $proc.Kill($true) } catch { }
            [void]$proc.WaitForExit(5000)
            throw "pdftotext did not finish within ${TimeoutSec}s and was stopped (see -TextTimeoutSec)"
        }
        $stderr = if ($errTask.Wait(5000)) { $errTask.Result } else { '' }
        if ($proc.ExitCode -ne 0) {
            throw (Get-PdfToTextMessage -ExitCode $proc.ExitCode -Stderr $stderr)
        }
        if ($stderr.Trim()) {
            Write-Log -Level WARN -NoConsole ('pdftotext: {0}' -f $stderr.Trim())
        }
    }
    finally {
        if (-not $proc.HasExited) { try { $proc.Kill($true) } catch { } }
        $proc.Dispose()
    }

    if (-not (Test-Path -LiteralPath $outFile)) { return '' }
    return [System.IO.File]::ReadAllText($outFile, [System.Text.UTF8Encoding]::new($false))
}

<#
    One pdftotext run answers three questions at once: is the PDF readable, how much text does
    it actually contain, and - if the Text engine is chosen - here is its input already. That
    is why the probe runs before engine selection rather than after Word has burned five
    minutes producing an empty document.
#>
function Get-PdfTextProfile {
    param([Parameter(Mandatory)][AllowEmptyString()][string] $Text)

    $pages = @($Text -split "`f")
    # A trailing form feed leaves an empty last element that is not a real page.
    if ($pages.Count -gt 1 -and [string]::IsNullOrWhiteSpace($pages[$pages.Count - 1])) {
        $pages = @($pages[0..($pages.Count - 2)])
    }
    $pageCount = [Math]::Max(1, $pages.Count)

    $perPage = [System.Collections.Generic.List[int]]::new()
    $total = 0
    foreach ($page in $pages) {
        $n = ($page -replace '\s', '').Length
        $perPage.Add($n)
        $total += $n
    }

    return [pscustomobject]@{
        Pages      = $pageCount
        TotalChars = $total
        PerPage    = $perPage
        CharsPerPage = [Math]::Round($total / $pageCount, 1)
    }
}

#endregion

#region Word engine (parent side)

<#
    DCOM starts WINWORD.EXE from the RPCSS service, not as a child of the worker, so killing
    the worker process tree does not reap it. Identification is layered:

      1. the pid file the worker wrote (exact, but only exists once CreateObject returned)
      2. a command-line scan for /Automation or -Embedding, which the SCM adds when it
         activates an out-of-process server - an interactive Word never carries them. This
         covers a hang inside CreateObject itself, before the pid file exists.

    Nothing is killed without re-checking the process name and start time: PIDs recorded up to
    WordTimeoutSec ago may since have been recycled onto an unrelated process.
#>
function Stop-OrphanWord {
    param(
        [Parameter(Mandatory)][string] $StageDir,
        [Parameter(Mandatory)][AllowEmptyCollection()][int[]] $Before,
        [Parameter(Mandatory)][datetime] $RunStart
    )

    $killed = 0
    $candidates = [System.Collections.Generic.List[int]]::new()

    $pidFile = Join-Path $StageDir 'word.pid'
    if (Test-Path -LiteralPath $pidFile) {
        $raw = ''
        try { $raw = [System.IO.File]::ReadAllText($pidFile) } catch { }
        if ($raw -match '^\s*(\d{1,10})\s*$') { $candidates.Add([int]$Matches[1]) }
    }

    try {
        $running = @(Get-CimInstance -ClassName Win32_Process -Filter "Name='WINWORD.EXE'" -ErrorAction SilentlyContinue)
        foreach ($p in $running) {
            $cmdline = if ($p.CommandLine) { $p.CommandLine } else { '' }
            if ($cmdline -notmatch '(?i)(/|-)Automation|(/|-)Embedding') { continue }
            if ([int]$p.ProcessId -in $Before) { continue }
            if ($p.CreationDate -and $p.CreationDate -lt $RunStart) { continue }
            if ([int]$p.ProcessId -notin $candidates) { $candidates.Add([int]$p.ProcessId) }
        }
    }
    catch { }

    foreach ($id in $candidates) {
        if ($id -in $Before) { continue }
        $proc = Get-Process -Id $id -ErrorAction SilentlyContinue
        if (-not $proc) { continue }
        if ($proc.ProcessName -ne 'WINWORD') { continue }
        try { if ($proc.StartTime -lt $RunStart) { continue } } catch { continue }
        try { $proc.Kill(); $killed++ } catch { }
    }
    return $killed
}

function Get-WordWorkerMessage {
    param([Parameter(Mandatory)][int] $ExitCode, [AllowEmptyString()][string] $Stderr)

    $detail = if ([string]::IsNullOrWhiteSpace($Stderr)) { '' } else { " ($($Stderr.Trim()))" }
    switch ($ExitCode) {
        10 { return "the Word worker was invoked incorrectly$detail" }
        11 { return "Word could not be started$detail" }
        12 { return 'Word is already running - close it, or use -AllowRunningWord (which risks your unsaved documents)' }
        13 { return "Word could not open the PDF - it may be encrypted or malformed$detail" }
        14 { return "Word could not save the converted document$detail" }
        15 { return "Word reflowed the PDF to an empty document$detail" }
        default { return "the Word worker exited with code $ExitCode$detail" }
    }
}

<#
    Word automation blocks and cannot be cancelled, so it runs in a child process the parent
    can kill - the same guarantee the pandoc timeout gives: one bad document fails alone.

    The child is this same script re-launched with -WordWorker. Self-relaunch rather than
    -EncodedCommand: base64 command lines are a first-tier EDR heuristic and would get the
    script quarantined on a managed machine.
#>
function Invoke-WordReflow {
    param(
        [Parameter(Mandatory)][string] $StageDir,
        [Parameter(Mandatory)][int] $TimeoutSec
    )

    $runStart = Get-Date
    $before = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)

    $psi = [System.Diagnostics.ProcessStartInfo]::new(
        [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
    foreach ($a in @('-NoLogo', '-NoProfile', '-NonInteractive', '-STA', '-File', $PSCommandPath, '-WordWorker')) {
        $psi.ArgumentList.Add($a)
    }
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardInput = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.WorkingDirectory = $StageDir
    # INVARIANT: values travel in the environment block, never on the command line and never
    # interpolated into the child's source. A file name containing a quote must never be able
    # to become an argument, let alone a statement.
    $psi.Environment['PDF2MD_ROLE'] = 'word-worker'
    $psi.Environment['PDF2MD_STAGE'] = $StageDir
    if ($AllowRunningWord) { $psi.Environment['PDF2MD_ALLOW_RUNNING_WORD'] = '1' }

    $exitCode = -1
    $stderr = ''
    $timedOut = $false

    $proc = [System.Diagnostics.Process]::Start($psi)
    try {
        $proc.StandardInput.Close()          # anything prompting gets EOF instead of hanging
        $outTask = $proc.StandardOutput.ReadToEndAsync()
        $errTask = $proc.StandardError.ReadToEndAsync()
        if (-not $proc.WaitForExit($TimeoutSec * 1000)) {
            $timedOut = $true
            try { $proc.Kill($true) } catch { }
            [void]$proc.WaitForExit(5000)
        }
        else {
            $exitCode = $proc.ExitCode
        }
        [void]$outTask.Wait(5000)
        if ($errTask.Wait(5000)) { $stderr = $errTask.Result }
    }
    finally {
        if (-not $proc.HasExited) { try { $proc.Kill($true) } catch { } }
        $proc.Dispose()
        # Runs on the success path too: a worker that exited cleanly may still have left Word
        # behind if the COM release did not take.
        $killed = Stop-OrphanWord -StageDir $StageDir -Before $before -RunStart $runStart
        if ($killed -gt 0) {
            Write-Log -Level WARN -NoConsole ('Stopped {0} orphaned Word process(es).' -f $killed)
        }
    }

    if ($timedOut) {
        throw "Word did not finish within ${TimeoutSec}s and was stopped (see -WordTimeoutSec)"
    }
    if ($exitCode -ne 0) {
        throw (Get-WordWorkerMessage -ExitCode $exitCode -Stderr $stderr)
    }

    $docx = Join-Path $StageDir 'out.docx'
    if (-not (Test-Path -LiteralPath $docx)) {
        throw 'Word reported success but produced no document'
    }
    return $docx
}

<#
    A Word left behind by an earlier crashed run would otherwise trip the "Word is already
    running" refusal forever. It is distinguishable from an interactive Word by its command
    line, so say so instead of leaving the user to guess.
#>
function Write-OrphanWordHint {
    try {
        $running = @(Get-CimInstance -ClassName Win32_Process -Filter "Name='WINWORD.EXE'" -ErrorAction SilentlyContinue)
        foreach ($p in $running) {
            $cmdline = if ($p.CommandLine) { $p.CommandLine } else { '' }
            if ($cmdline -match '(?i)(/|-)Automation|(/|-)Embedding') {
                Write-Log -Level WARN ('  Word process {0} looks like an automation leftover, not a window you opened. Ending it in Task Manager re-enables the Word engine.' -f $p.ProcessId)
            }
        }
    }
    catch { }
}

#endregion

#region Text engine: plain text to Markdown

# Note on the [AllowEmptyString()] attributes below: a mandatory parameter validates the elements
# of a collection as well as the collection itself, so a List[string] holding a blank line is
# rejected with "Cannot bind argument ... because it is an empty string". Every page of a document
# has blank lines, so removing those attributes breaks the whole text engine.

# Ligatures survive PDF extraction as single code points. Left alone, keyword replacement and
# ordinary search silently miss the words containing them.
$script:TextNormalMap = [ordered]@{
    ([char]0x00A0) = ' '        # non-breaking space
    ([char]0x2007) = ' '        # figure space
    ([char]0x202F) = ' '        # narrow no-break space
    ([char]0x00AD) = ''         # soft hyphen
    ([char]0x200B) = ''         # zero-width space
    ([char]0xFEFF) = ''         # byte-order mark
}
$script:LigatureMap = [ordered]@{
    ([string][char]0xFB00) = 'ff'
    ([string][char]0xFB01) = 'fi'
    ([string][char]0xFB02) = 'fl'
    ([string][char]0xFB03) = 'ffi'
    ([string][char]0xFB04) = 'ffl'
    ([string][char]0xFB05) = 'st'
    ([string][char]0xFB06) = 'st'
}

function ConvertTo-NormalizedText {
    param([Parameter(Mandatory)][AllowEmptyString()][string] $Text)

    $Text = $Text -replace "`r`n", "`n"
    $Text = $Text -replace "`r", "`n"
    foreach ($pair in $script:TextNormalMap.GetEnumerator()) {
        $Text = $Text.Replace([string]$pair.Key, [string]$pair.Value)
    }
    foreach ($pair in $script:LigatureMap.GetEnumerator()) {
        $Text = $Text.Replace([string]$pair.Key, [string]$pair.Value)
    }
    return $Text
}

function Split-PdfPage {
    param([Parameter(Mandatory)][AllowEmptyString()][string] $Text)

    $pages = [System.Collections.Generic.List[System.Collections.Generic.List[string]]]::new()
    foreach ($chunk in ($Text -split "`f")) {
        $lines = [System.Collections.Generic.List[string]]::new()
        foreach ($line in ($chunk -split "`n")) { $lines.Add($line.TrimEnd()) }
        $pages.Add($lines)
    }
    # Drop a trailing empty page left by a final form feed.
    if ($pages.Count -gt 1) {
        $last = $pages[$pages.Count - 1]
        $hasText = $false
        foreach ($line in $last) { if (-not [string]::IsNullOrWhiteSpace($line)) { $hasText = $true; break } }
        if (-not $hasText) { $pages.RemoveAt($pages.Count - 1) }
    }
    # The leading comma stops PowerShell unrolling the list into loose elements on return.
    return ,$pages
}

<#
    pdftotext -layout pads every line by the page's left margin. Left in place those leading
    spaces make GFM render ordinary paragraphs as indented code blocks, so the common prefix
    comes off per page - which preserves relative indentation while removing the artefact.
#>
function Remove-CommonLeftMargin {
    param([Parameter(Mandatory)][AllowEmptyString()][AllowEmptyCollection()][System.Collections.Generic.List[string]] $Lines)

    $margin = [int]::MaxValue
    foreach ($line in $Lines) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $indent = $line.Length - $line.TrimStart(' ').Length
        if ($indent -lt $margin) { $margin = $indent }
    }
    if ($margin -eq [int]::MaxValue -or $margin -eq 0) { return ,$Lines }

    $out = [System.Collections.Generic.List[string]]::new()
    foreach ($line in $Lines) {
        if ([string]::IsNullOrWhiteSpace($line)) { $out.Add(''); continue }
        $out.Add($line.Substring($margin))
    }
    return ,$out
}

<#
    -layout interleaves the columns of a two-column page onto the same physical line. Nothing
    downstream can un-scramble that, so the only useful response is to say so and point at the
    engine that handles columns properly.
#>
function Test-MultiColumnPage {
    param([Parameter(Mandatory)][AllowEmptyString()][AllowEmptyCollection()][System.Collections.Generic.List[string]] $Lines)

    $body = @($Lines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($body.Count -lt 6) { return $false }

    $width = 0
    foreach ($line in $body) { if ($line.Length -gt $width) { $width = $line.Length } }
    if ($width -lt 40) { return $false }

    $low = [int]($width * 0.25)
    $high = [int]($width * 0.75)
    $hits = 0
    foreach ($line in $body) {
        foreach ($m in [regex]::Matches($line, ' {6,}')) {
            if ($m.Index -ge $low -and $m.Index -le $high -and ($m.Index + $m.Length) -lt $line.Length) {
                $hits++
                break
            }
        }
    }
    return ($hits / $body.Count) -gt 0.4
}

function Get-NormalizedFurnitureKey {
    param([Parameter(Mandatory)][AllowEmptyString()][string] $Line)
    $key = $Line.Trim().ToLowerInvariant()
    $key = [regex]::Replace($key, '\s+', ' ')
    return $script:RxDigitRun.Replace($key, '#')
}

<#
    Running headers and footers repeat at the same position on most pages. Keying on position
    as well as text is what stops a title page's title - which appears once, at the top of page
    one - from being mistaken for one.

    Digit runs are normalised to '#' first, so "Page 3 of 12" and "Page 4 of 12" collapse to
    the same key.
#>
function Remove-RunningHeader {
    param([Parameter(Mandatory)][AllowEmptyString()][AllowEmptyCollection()][System.Collections.Generic.List[System.Collections.Generic.List[string]]] $Pages)

    if ($Pages.Count -lt 3) { return 0 }

    $tally = @{}
    foreach ($page in $Pages) {
        $body = @($page | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        if ($body.Count -lt 3) { continue }

        $slots = [ordered]@{}
        for ($i = 0; $i -lt 3 -and $i -lt $body.Count; $i++) { $slots["h$i"] = $body[$i] }
        for ($i = 1; $i -le 3 -and $i -le $body.Count; $i++) { $slots["f$i"] = $body[$body.Count - $i] }

        foreach ($slot in $slots.GetEnumerator()) {
            if ($slot.Value.Length -gt 120) { continue }     # that length is body text
            $key = '{0}|{1}' -f $slot.Key, (Get-NormalizedFurnitureKey -Line $slot.Value)
            if (-not $tally.ContainsKey($key)) { $tally[$key] = 0 }
            $tally[$key]++
        }
    }

    $threshold = [Math]::Max(3, [int][Math]::Ceiling(0.6 * $Pages.Count))
    $furniture = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($entry in $tally.GetEnumerator()) {
        if ($entry.Value -ge $threshold) { [void]$furniture.Add($entry.Key) }
    }
    if ($furniture.Count -eq 0) { return 0 }

    $removed = 0
    for ($p = 0; $p -lt $Pages.Count; $p++) {
        $page = $Pages[$p]

        $bodyIndex = [System.Collections.Generic.List[int]]::new()
        for ($i = 0; $i -lt $page.Count; $i++) {
            if (-not [string]::IsNullOrWhiteSpace($page[$i])) { $bodyIndex.Add($i) }
        }
        if ($bodyIndex.Count -lt 3) { continue }

        # Slots hold line indices rather than text, so only the edge line that matched is
        # dropped - a page whose body happens to repeat the header's wording keeps it.
        $slots = [ordered]@{}
        for ($i = 0; $i -lt 3 -and $i -lt $bodyIndex.Count; $i++) { $slots["h$i"] = $bodyIndex[$i] }
        for ($i = 1; $i -le 3 -and $i -le $bodyIndex.Count; $i++) { $slots["f$i"] = $bodyIndex[$bodyIndex.Count - $i] }

        $drop = [System.Collections.Generic.HashSet[int]]::new()
        foreach ($slot in $slots.GetEnumerator()) {
            $key = '{0}|{1}' -f $slot.Key, (Get-NormalizedFurnitureKey -Line $page[$slot.Value])
            if ($furniture.Contains($key)) { [void]$drop.Add([int]$slot.Value) }
        }
        if ($drop.Count -eq 0) { continue }

        # Never strip a page down to nothing: if the removal would leave fewer than two lines,
        # what matched was body text, not furniture.
        if (($bodyIndex.Count - $drop.Count) -lt 2) { continue }

        $kept = [System.Collections.Generic.List[string]]::new()
        for ($i = 0; $i -lt $page.Count; $i++) {
            if ($drop.Contains($i)) { $removed++; continue }
            $kept.Add($page[$i])
        }
        $Pages[$p] = $kept
    }
    return $removed
}

<#
    Only the outermost two lines of a page are considered: a bare number in the middle of a page
    is data, not a page number.
#>
function Remove-PageNumberLine {
    param([Parameter(Mandatory)][AllowEmptyString()][AllowEmptyCollection()][System.Collections.Generic.List[System.Collections.Generic.List[string]]] $Pages)

    $removed = 0
    for ($p = 0; $p -lt $Pages.Count; $p++) {
        $page = $Pages[$p]
        $indexes = [System.Collections.Generic.List[int]]::new()
        for ($i = 0; $i -lt $page.Count; $i++) {
            if (-not [string]::IsNullOrWhiteSpace($page[$i])) { $indexes.Add($i) }
        }
        if ($indexes.Count -lt 3) { continue }

        $edges = [System.Collections.Generic.HashSet[int]]::new()
        for ($i = 0; $i -lt 2 -and $i -lt $indexes.Count; $i++) { [void]$edges.Add($indexes[$i]) }
        for ($i = 1; $i -le 2 -and $i -le $indexes.Count; $i++) { [void]$edges.Add($indexes[$indexes.Count - $i]) }

        $kept = [System.Collections.Generic.List[string]]::new()
        for ($i = 0; $i -lt $page.Count; $i++) {
            if ($edges.Contains($i) -and
                ($script:RxPageNumber.IsMatch($page[$i]) -or $script:RxRomanNumeral.IsMatch($page[$i]))) {
                $removed++
                continue
            }
            $kept.Add($page[$i])
        }
        $Pages[$p] = $kept
    }
    return $removed
}

<#
    Marks the lines that -layout is holding in columns. Every destructive stage consults this
    map, which is the single reason tables and multi-column text survive the text engine
    unmangled. The lines immediately above and below a run are included, because a table's
    caption or its first row often has no wide gap of its own.
#>
function Get-PreformattedMap {
    param([Parameter(Mandatory)][AllowEmptyString()][AllowEmptyCollection()][System.Collections.Generic.List[string]] $Lines)

    $map = [bool[]]::new($Lines.Count)
    for ($i = 0; $i -lt $Lines.Count; $i++) {
        $map[$i] = $script:RxLayoutGap.IsMatch($Lines[$i])
    }

    $expanded = [bool[]]::new($Lines.Count)
    for ($i = 0; $i -lt $Lines.Count; $i++) {
        if (-not $map[$i]) { continue }
        $expanded[$i] = $true
        if ($i -gt 0 -and -not [string]::IsNullOrWhiteSpace($Lines[$i - 1])) { $expanded[$i - 1] = $true }
        if ($i -lt ($Lines.Count - 1) -and -not [string]::IsNullOrWhiteSpace($Lines[$i + 1])) { $expanded[$i + 1] = $true }
    }
    return ,$expanded
}

<#
    Rejoins a word split across a line break. Restricted to lowercase-to-lowercase so genuine
    compounds keep their hyphen: "pre-" followed by "Production" is one word in the source, not
    an artefact of wrapping.
#>
function Join-HyphenatedWord {
    param(
        [Parameter(Mandatory)][AllowEmptyString()][AllowEmptyCollection()][System.Collections.Generic.List[string]] $Lines,
        [Parameter(Mandatory)][AllowEmptyCollection()][bool[]] $Preformatted
    )

    $outLines = [System.Collections.Generic.List[string]]::new()
    $outMap = [System.Collections.Generic.List[bool]]::new()
    $joined = 0
    $i = 0

    while ($i -lt $Lines.Count) {
        $line = $Lines[$i]
        $canJoin = -not $Preformatted[$i] -and
                   ($i + 1) -lt $Lines.Count -and
                   -not $Preformatted[$i + 1] -and
                   $script:RxHyphenEnd.IsMatch($line) -and
                   $Lines[$i + 1] -cmatch '^\s*\p{Ll}'

        if ($canJoin) {
            $head = $script:RxHyphenEnd.Replace($line, '$1')
            $outLines.Add($head + $Lines[$i + 1].TrimStart())
            $outMap.Add($false)
            $joined++
            $i += 2
            continue
        }
        $outLines.Add($line)
        $outMap.Add($Preformatted[$i])
        $i++
    }
    return [pscustomobject]@{ Lines = $outLines; Preformatted = $outMap.ToArray(); Count = $joined }
}

function ConvertTo-MarkdownList {
    param(
        [Parameter(Mandatory)][AllowEmptyString()][AllowEmptyCollection()][System.Collections.Generic.List[string]] $Lines,
        [Parameter(Mandatory)][AllowEmptyCollection()][bool[]] $Preformatted
    )
    $count = 0
    for ($i = 0; $i -lt $Lines.Count; $i++) {
        if ($Preformatted[$i]) { continue }
        $line = $Lines[$i]
        if ($script:RxBulletGlyph.IsMatch($line)) {
            $Lines[$i] = $script:RxBulletGlyph.Replace($line, '$1- ')
            $count++
        }
        elseif ($script:RxOrderedParen.IsMatch($line)) {
            $Lines[$i] = $script:RxOrderedParen.Replace($line, '$1$2. ')
            $count++
        }
    }
    return $count
}

function Test-LooksLikeHeadingText {
    param([Parameter(Mandatory)][AllowEmptyString()][string] $Text)
    if ($Text.Length -gt 110) { return $false }
    if ($Text.TrimEnd().EndsWith('.')) { return $false }
    return $true
}

<#
    Multi-component numbering ("2.3 Scope") is unambiguous - an ordered list never nests its
    numbering into the text like that - so it is converted by default. Single-component numbering
    ("2. Scope") is indistinguishable from an ordered list on its own, so it is only converted
    when the document has already proved it numbers its headings this way.
#>
function ConvertTo-MarkdownHeading {
    param(
        [Parameter(Mandatory)][AllowEmptyString()][AllowEmptyCollection()][System.Collections.Generic.List[string]] $Lines,
        [Parameter(Mandatory)][AllowEmptyCollection()][bool[]] $Preformatted,
        [switch] $DetectCaps
    )

    $count = 0
    $sawMulti = $false

    for ($i = 0; $i -lt $Lines.Count; $i++) {
        if ($Preformatted[$i]) { continue }
        $m = $script:RxNumHeadMulti.Match($Lines[$i])
        if (-not $m.Success) { continue }
        if (-not (Test-LooksLikeHeadingText -Text $m.Groups[2].Value)) { continue }
        $parts = $m.Groups[1].Value.Split('.')
        if ([int]$parts[0] -gt 40) { continue }
        $sawMulti = $true
        $level = [Math]::Min(6, $parts.Count)
        $Lines[$i] = '{0} {1} {2}' -f ('#' * $level), $m.Groups[1].Value, $m.Groups[2].Value
        $count++
    }

    if ($sawMulti) {
        for ($i = 0; $i -lt $Lines.Count; $i++) {
            if ($Preformatted[$i]) { continue }
            $m = $script:RxNumHeadOne.Match($Lines[$i])
            if (-not $m.Success) { continue }
            if (-not (Test-LooksLikeHeadingText -Text $m.Groups[2].Value)) { continue }
            if ([int]$m.Groups[1].Value -gt 40) { continue }
            $Lines[$i] = '# {0} {1}' -f $m.Groups[1].Value, $m.Groups[2].Value
            $count++
        }
    }

    if ($DetectCaps) {
        for ($i = 0; $i -lt $Lines.Count; $i++) {
            if ($Preformatted[$i]) { continue }
            $line = $Lines[$i].Trim()
            if ($line.Length -lt 3 -or $line.Length -gt 110) { continue }
            if ($line.StartsWith('#')) { continue }
            if ($line.EndsWith('.') -or $line.EndsWith(':')) { continue }
            $letters = @($line.ToCharArray() | Where-Object { [char]::IsLetter($_) })
            if ($letters.Count -lt 2) { continue }
            if (@($letters | Where-Object { [char]::IsLower($_) }).Count -gt 0) { continue }
            $Lines[$i] = '## ' + $line
            $count++
        }
    }
    return $count
}

<#
    Four or more leading spaces is an indented code block in GFM. After the per-page margin
    strip only genuinely nested content is still indented, and clamping it to three spaces keeps
    the nesting visible without turning prose into code.
#>
function Set-MarkdownIndentSafe {
    param(
        [Parameter(Mandatory)][AllowEmptyString()][AllowEmptyCollection()][System.Collections.Generic.List[string]] $Lines,
        [Parameter(Mandatory)][AllowEmptyCollection()][bool[]] $Preformatted
    )
    $count = 0
    for ($i = 0; $i -lt $Lines.Count; $i++) {
        if ($Preformatted[$i]) { continue }
        $line = $Lines[$i]
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $indent = $line.Length - $line.TrimStart(' ').Length
        if ($indent -lt 4) { continue }
        $Lines[$i] = ('   ' + $line.TrimStart(' '))
        $count++
    }
    return $count
}

function Test-BreaksParagraph {
    param([Parameter(Mandatory)][AllowEmptyString()][string] $Line)
    return [string]::IsNullOrWhiteSpace($Line) -or
           $Line.TrimStart().StartsWith('#') -or
           $script:RxListItem.IsMatch($Line)
}

<#
    Opt-in only. Markdown already renders consecutive lines as one paragraph, so leaving the
    line breaks alone produces correct output from untidy source, while joining the wrong pair
    of lines produces wrong output.
#>
function Join-WrappedLine {
    param(
        [Parameter(Mandatory)][AllowEmptyString()][AllowEmptyCollection()][System.Collections.Generic.List[string]] $Lines,
        [Parameter(Mandatory)][AllowEmptyCollection()][bool[]] $Preformatted
    )

    $out = [System.Collections.Generic.List[string]]::new()
    $outMap = [System.Collections.Generic.List[bool]]::new()
    $buffer = ''

    for ($i = 0; $i -lt $Lines.Count; $i++) {
        $line = $Lines[$i]
        if ($Preformatted[$i] -or (Test-BreaksParagraph -Line $line)) {
            if ($buffer) { $out.Add($buffer); $outMap.Add($false); $buffer = '' }
            $out.Add($line)
            $outMap.Add($Preformatted[$i])
            continue
        }
        $buffer = if ($buffer) { $buffer + ' ' + $line.Trim() } else { $line.TrimEnd() }
    }
    if ($buffer) { $out.Add($buffer); $outMap.Add($false) }

    return [pscustomobject]@{ Lines = $out; Preformatted = $outMap.ToArray() }
}

function Add-LayoutFence {
    param(
        [Parameter(Mandatory)][AllowEmptyString()][AllowEmptyCollection()][System.Collections.Generic.List[string]] $Lines,
        [Parameter(Mandatory)][AllowEmptyCollection()][bool[]] $Preformatted
    )
    $out = [System.Collections.Generic.List[string]]::new()
    $inBlock = $false
    $fenced = 0

    for ($i = 0; $i -lt $Lines.Count; $i++) {
        if ($Preformatted[$i] -and -not $inBlock) {
            $out.Add(''); $out.Add('```text'); $inBlock = $true; $fenced++
        }
        elseif (-not $Preformatted[$i] -and $inBlock) {
            $out.Add('```'); $out.Add(''); $inBlock = $false
        }
        $out.Add($Lines[$i])
    }
    if ($inBlock) { $out.Add('```') }
    return [pscustomobject]@{ Lines = $out; Count = $fenced }
}

<#
    Orchestrates the plain text layer. Order matters: page-level furniture is removed first
    (while page boundaries still exist), then the preformatted map is built over what remains,
    then every destructive stage consults it.
#>
function ConvertFrom-PdfText {
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string] $Text,
        [Parameter(Mandatory)][string] $SourceName
    )

    $Text = ConvertTo-NormalizedText -Text $Text
    $pages = Split-PdfPage -Text $Text

    $multiColumn = [System.Collections.Generic.List[int]]::new()
    for ($p = 0; $p -lt $pages.Count; $p++) {
        $pages[$p] = Remove-CommonLeftMargin -Lines $pages[$p]
        if (Test-MultiColumnPage -Lines $pages[$p]) { $multiColumn.Add($p + 1) }
    }

    $headersRemoved = 0
    if (-not $KeepRunningHeaders) { $headersRemoved = Remove-RunningHeader -Pages $pages }

    $numbersRemoved = 0
    if (-not $KeepPageNumbers) { $numbersRemoved = Remove-PageNumberLine -Pages $pages }

    $lines = [System.Collections.Generic.List[string]]::new()
    for ($p = 0; $p -lt $pages.Count; $p++) {
        if ($p -gt 0) { $lines.Add('') }
        foreach ($line in $pages[$p]) { $lines.Add($line) }
    }

    $preformatted = Get-PreformattedMap -Lines $lines

    $hyphensJoined = 0
    if (-not $KeepLineHyphens) {
        $dehyphenated = Join-HyphenatedWord -Lines $lines -Preformatted $preformatted
        $lines = $dehyphenated.Lines
        $preformatted = $dehyphenated.Preformatted
        $hyphensJoined = $dehyphenated.Count
    }

    $listsConverted = ConvertTo-MarkdownList -Lines $lines -Preformatted $preformatted
    $headingsFound = ConvertTo-MarkdownHeading -Lines $lines -Preformatted $preformatted -DetectCaps:$DetectCapsHeadings
    [void](Set-MarkdownIndentSafe -Lines $lines -Preformatted $preformatted)

    if ($ReflowParagraphs) {
        $reflowed = Join-WrappedLine -Lines $lines -Preformatted $preformatted
        $lines = $reflowed.Lines
        $preformatted = $reflowed.Preformatted
    }

    if ($FenceLayoutBlocks) {
        $fenced = Add-LayoutFence -Lines $lines -Preformatted $preformatted
        $lines = $fenced.Lines
    }

    foreach ($page in $multiColumn) {
        Write-Log -Level WARN ('  {0}: page {1} looks like multi-column text, which pdftotext interleaves. Try -Engine Word for that document.' -f $SourceName, $page)
    }

    return [pscustomobject]@{
        Text           = ($lines -join "`n")
        Pages          = $pages.Count
        HeadersRemoved = $headersRemoved
        NumbersRemoved = $numbersRemoved
        HyphensJoined  = $hyphensJoined
        ListsConverted = $listsConverted
        HeadingsFound  = $headingsFound
        MultiColumn    = $multiColumn.Count
    }
}

#endregion

#region Conversion

<#
    Verbatim from Convert-DocxToMd.ps1 v1.4.0 - keep in sync.
    No --standalone => no YAML metadata block. No --extract-media => no image files on disk.
#>
function Invoke-Pandoc {
    param(
        [Parameter(Mandatory)][string] $DocxPath,
        [Parameter(Mandatory)][string] $OutFile,
        [Parameter(Mandatory)][int] $TimeoutSec
    )

    $pandocArgs = @(
        '--from=docx'
        '--to=gfm'
        '--wrap=none'
        '--track-changes=accept'
        '--output', $OutFile
        '--', $DocxPath
    )

    # Run pandoc with a hard timeout so a document that hangs it fails alone instead of
    # freezing the batch. ProcessStartInfo.ArgumentList quotes each argument correctly.
    $psi = [System.Diagnostics.ProcessStartInfo]::new($script:PandocPath)
    foreach ($a in $pandocArgs) { $psi.ArgumentList.Add($a) }
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardError = $true

    $proc = [System.Diagnostics.Process]::Start($psi)
    try {
        $stderrTask = $proc.StandardError.ReadToEndAsync()   # async read prevents pipe deadlock
        if (-not $proc.WaitForExit($TimeoutSec * 1000)) {
            try { $proc.Kill($true) } catch { }
            [void]$proc.WaitForExit(5000)
            throw "pandoc did not finish within ${TimeoutSec}s and was stopped (see -PandocTimeoutSec)"
        }
        $stderr = if ($stderrTask.Wait(5000)) { $stderrTask.Result } else { '' }
        if ($proc.ExitCode -ne 0) {
            throw "pandoc exited with code $($proc.ExitCode): $($stderr.Trim())"
        }
        if ($stderr.Trim()) {
            Write-Log -Level WARN -NoConsole ('pandoc: {0}' -f $stderr.Trim())
        }
    }
    finally {
        if (-not $proc.HasExited) { try { $proc.Kill($true) } catch { } }
        $proc.Dispose()
    }

    return [System.IO.File]::ReadAllText($OutFile, [System.Text.UTF8Encoding]::new($false))
}

<#
    The one cleanup chain, called by both engines so they cannot drift apart.

    -TextMode skips two stages, on a principle rather than a whim: on the Word path they consume
    pandoc-generated markup, but plain text from a PDF has no generated markup, so anything they
    match is literal document text.
      Remove-InlineHtml    would eat the <span> and <div> in a PDF that documents HTML, and
                           cannot protect them because plain text has no code fences.
      Remove-MarkdownLinkUrls's RxLink would collapse "see clause [3](a)" to "3".
    -TextStrictPipeline forces the full chain anyway.
#>
function Invoke-MarkdownPipeline {
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string] $Markdown,
        [Parameter(Mandatory)][AllowEmptyCollection()][psobject[]] $Rules,
        [switch] $TextMode
    )

    $strict = (-not $TextMode) -or $TextStrictPipeline

    # Anchors first, so leftover bookmark markup never ends up inside a converted table cell.
    $anchors = Remove-HtmlAnchors -Text $Markdown
    $Markdown = $anchors.Text

    # Then raw HTML tables -> pipe tables, before link/image stripping runs over the result.
    $tables = ConvertTo-MarkdownTables -Text $Markdown
    $Markdown = $tables.Text

    $images = Remove-MarkdownImages -Text $Markdown
    $Markdown = $images.Text

    $links = if ($strict) {
        Remove-MarkdownLinkUrls -Text $Markdown -AlsoStripBareUrls:$StripBareUrls
    }
    else {
        Remove-TextLinkUrls -Text $Markdown -AlsoStripBareUrls:$StripBareUrls
    }
    $Markdown = $links.Text

    # Safety net: anything still carrying raw inline HTML loses the markup but keeps its words.
    $htmlCount = 0
    if ($strict) {
        $html = Remove-InlineHtml -Text $Markdown
        $Markdown = $html.Text
        $htmlCount = $html.Count
    }

    $footnoteCount = 0
    if (-not $KeepFootnotes) {
        $footnotes = Remove-MarkdownFootnotes -Text $Markdown
        $Markdown = $footnotes.Text
        $footnoteCount = $footnotes.Count
    }

    $headers = Optimize-MarkdownTables -Text $Markdown
    $Markdown = $headers.Text

    $keywords = Invoke-KeywordReplacement -Text $Markdown -Rules $Rules
    $Markdown = Optimize-MarkdownWhitespace -Text $keywords.Text

    return [pscustomobject]@{
        Text          = $Markdown
        Images        = $images.Count
        Links         = $links.Count
        Footnotes     = $footnoteCount
        Anchors       = $anchors.Count
        Tables        = $tables.Count
        Headers       = $headers.Count
        Html          = $htmlCount
        KeywordCounts = $keywords.Counts
    }
}

function Convert-OnePdf {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][System.IO.FileInfo] $File,
        [Parameter(Mandatory)][string] $OutputFolder,
        [Parameter(Mandatory)][AllowEmptyCollection()][psobject[]] $Rules
    )

    $stage = New-StageFolder
    try {
        [void](Copy-PdfToStage -Source $File.FullName -StageDir $stage)

        # The probe doubles as the Text engine's input, so a PDF is only ever extracted once.
        $probeText = Invoke-PdfToText -StageDir $stage -Mode $TextLayout -TimeoutSec $TextTimeoutSec
        $textProfile = Get-PdfTextProfile -Text $probeText

        $looksScanned = $MinCharsPerPage -gt 0 -and
                        ($textProfile.CharsPerPage -lt $MinCharsPerPage -or $textProfile.TotalChars -lt 100)
        if ($looksScanned -and $Engine -ne 'Word') {
            throw ('no extractable text ({0} page(s), {1} non-whitespace characters). This looks like a scanned or image-only PDF, and no OCR engine is available on this machine.' -f
                $textProfile.Pages, $textProfile.TotalChars)
        }
        if ($looksScanned) {
            Write-Log -Level WARN ('  {0}: only {1} extractable character(s); trying Word anyway because -Engine Word was requested.' -f $File.Name, $textProfile.TotalChars)
        }

        $chosen = $Engine
        $wordError = ''
        $markdown = $null
        $usedEngine = ''

        if ($chosen -ne 'Text') {
            $available = Test-WordAvailable
            if (-not $available.Available) {
                $wordError = $available.Reason
            }
            elseif (-not $PSCmdlet.ShouldProcess($File.Name, 'Reflow PDF with Microsoft Word')) {
                # -WhatIf must not start Word: a preview run is not worth minutes of automation.
                $wordError = '-WhatIf'
            }
            else {
                try {
                    $docx = Invoke-WordReflow -StageDir $stage -TimeoutSec $WordTimeoutSec
                    $markdown = Invoke-Pandoc -DocxPath $docx -OutFile (Join-Path $stage 'out.md') -TimeoutSec $PandocTimeoutSec
                    $usedEngine = 'word'
                }
                catch {
                    $wordError = $_.Exception.Message
                    $markdown = $null
                }
            }

            if ($null -eq $markdown) {
                if ($chosen -eq 'Word') { throw $wordError }
                if ($wordError -eq '-WhatIf') {
                    Write-Log -Level INFO ('  {0}: previewing with the text engine; -WhatIf never starts Word.' -f $File.Name)
                }
                else {
                    Write-Log -Level WARN ('  {0}: Word engine unavailable ({1}). Falling back to the text engine.' -f $File.Name, $wordError)
                    if ($wordError -match 'already running') { Write-OrphanWordHint }
                }
            }
        }

        # pandoc's docx reader ignores text boxes, and Word's reflow puts floating elements in
        # them. The probe already counted the real characters, so a large shortfall is
        # detectable rather than silent.
        $textBoxLoss = $false
        if ($usedEngine -eq 'word' -and $textProfile.TotalChars -gt 0) {
            $produced = ($markdown -replace '\s', '').Length
            if ($produced -lt (0.6 * $textProfile.TotalChars)) {
                $textBoxLoss = $true
                Write-Log -Level WARN ('  {0}: Word produced {1} of {2} characters, so content was probably lost in text boxes.' -f
                    $File.Name, $produced, $textProfile.TotalChars)
                if ($Engine -eq 'Auto') {
                    Write-Log -Level WARN ('  {0}: using the text engine instead.' -f $File.Name)
                    $markdown = $null
                    $usedEngine = ''
                }
            }
        }

        $textStats = $null
        if ($null -eq $markdown) {
            $textStats = ConvertFrom-PdfText -Text $probeText -SourceName $File.Name
            $markdown = $textStats.Text
            $usedEngine = 'text'
        }

        $result = Invoke-MarkdownPipeline -Markdown $markdown -Rules $Rules -TextMode:($usedEngine -eq 'text')

        # The actual "never a silently empty .md" guarantee, covering both engines and any
        # pipeline stage that might have eaten everything.
        if (($result.Text -replace '\s', '').Length -lt 20) {
            throw 'conversion produced no text'
        }

        $target = Join-Path $OutputFolder ($File.BaseName + '.md')
        if ($PSCmdlet.ShouldProcess($target, 'Write Markdown')) {
            $target = Write-MarkdownAtomic -Folder $OutputFolder -BaseName $File.BaseName -Text $result.Text
        }

        return [pscustomobject]@{
            Source        = $File.FullName
            Target        = $target
            Engine        = $usedEngine
            Pages         = $textProfile.Pages
            Images        = $result.Images
            Links         = $result.Links
            Footnotes     = $result.Footnotes
            Anchors       = $result.Anchors
            Tables        = $result.Tables
            Headers       = $result.Headers
            Html          = $result.Html
            TextStats     = $textStats
            TextBoxLoss   = $textBoxLoss
            KeywordCounts = $result.KeywordCounts
        }
    }
    finally {
        Remove-StageFolder -Path $stage
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
$script:LogPath = Join-Path $folders.Logs ('pdfconvert-{0}.log' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))

# Retention: keep the newest 30 run logs. The 'pdfconvert-' prefix deliberately does not match
# the docx script's 'convert-*.log' sweep, so neither script deletes the other's logs; the
# Where-Object re-checks the long name because -Filter also matches 8.3 short names.
Get-ChildItem -LiteralPath $folders.Logs -Filter 'pdfconvert-*.log' -File |
    Where-Object { $_.Name -like 'pdfconvert-*' } |
    Sort-Object Name -Descending |
    Select-Object -Skip 29 |
    Remove-Item -Force -WhatIf:$false -Confirm:$false -ErrorAction SilentlyContinue

Write-Host ''
Write-Host ('Convert-PdfToMd v{0}' -f $script:ScriptVersion) -ForegroundColor Cyan
Write-Host ('-' * 60) -ForegroundColor DarkGray
Write-Log -Level INFO -NoConsole ('Version   : {0}' -f $script:ScriptVersion)
Write-Log -Level INFO -NoConsole ('PowerShell: {0}' -f $PSVersionTable.PSVersion)
Write-Log -Level INFO -NoConsole ("Workspace : {0}" -f $folders.Root)

try {
    $script:PandocPath = Get-PandocPath
    $script:PdfToTextExe = Get-PdfToTextPath -Preferred $PdfToTextPath
    Write-Log -Level INFO -NoConsole ('pandoc    : {0}' -f $script:PandocPath)
    Write-Log -Level INFO -NoConsole ('pdftotext : {0}' -f $script:PdfToTextExe)

    $stale = Clear-StaleStage
    if ($stale -gt 0) {
        Write-Log -Level INFO -NoConsole ('Removed {0} stale staging folder(s) from %TEMP%.' -f $stale)
    }

    # Office automation is unsupported in a non-interactive session and hangs there, which would
    # burn the full timeout on every single file.
    if ($Engine -ne 'Text' -and -not [Environment]::UserInteractive) {
        Write-Log -Level WARN 'No interactive session: Word automation is unsupported there. Using the text engine.'
        $Engine = 'Text'
    }

    if ($Engine -ne 'Text') {
        $running = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue)
        if ($running.Count -gt 0 -and -not $AllowRunningWord) {
            Write-Log -Level WARN ('Word is already running ({0} process(es)). The Word engine is disabled to protect your open documents - close Word and run again for better output.' -f $running.Count)
            Write-OrphanWordHint
        }
    }

    if (-not $ReplacementsPath) { $ReplacementsPath = Join-Path $folders.Root 'replacements.json' }
    $rules = @(Import-ReplacementRule -Path $ReplacementsPath)
    Write-Host ('Engine    : {0}' -f $Engine) -ForegroundColor Gray
    Write-Host ('Rules     : {0} keyword replacement(s)' -f $rules.Count) -ForegroundColor Gray
    Write-Log -Level INFO -NoConsole ('Engine    : {0} (text layout {1})' -f $Engine, $TextLayout)
    Write-Log -Level INFO -NoConsole ('Rules     : {0} from {1}' -f $rules.Count, $ReplacementsPath)

    $documents = @(Get-InputDocuments -InputFolder $folders.Input)

    if ($documents.Count -eq 0) {
        if ($NoPrompt) {
            Write-Log -Level WARN ('No .pdf files found in "{0}". Add files and run again.' -f $folders.Input)
            exit 0
        }
        $copied = @(Request-InputFiles -InputFolder $folders.Input)
        if ($copied.Count -eq 0) {
            Write-Log -Level WARN ('Nothing selected. Put your .pdf files in "{0}" and run again.' -f $folders.Input)
            exit 0
        }
        $documents = @(Get-InputDocuments -InputFolder $folders.Input)
    }

    Write-Host ('Documents : {0}' -f $documents.Count) -ForegroundColor Gray
    Write-Host ('-' * 60) -ForegroundColor DarkGray

    $ok = 0; $failed = 0
    $totals = [ordered]@{ Images = 0; Links = 0; Footnotes = 0; Keywords = 0; Anchors = 0; Tables = 0; Html = 0; Pages = 0; Word = 0; Text = 0 }

    foreach ($doc in $documents) {
        try {
            $result = Convert-OnePdf -File $doc -OutputFolder $folders.Output -Rules $rules

            $keywordHits = 0
            foreach ($v in $result.KeywordCounts.Values) { $keywordHits += $v }
            $totals.Images    += $result.Images
            $totals.Links     += $result.Links
            $totals.Footnotes += $result.Footnotes
            $totals.Keywords  += $keywordHits
            $totals.Anchors   += $result.Anchors
            $totals.Tables    += $result.Tables
            $totals.Html      += $result.Html
            $totals.Pages     += $result.Pages
            if ($result.Engine -eq 'word') { $totals.Word++ } else { $totals.Text++ }
            $ok++

            $detail = '{0}; {1} page(s), links {2}, images {3}, footnotes {4}, tables {5}, keywords {6}' -f
                $result.Engine, $result.Pages, $result.Links, $result.Images, $result.Footnotes, $result.Tables, $keywordHits
            Write-Host ('  OK    {0}' -f $doc.Name) -ForegroundColor Green -NoNewline
            Write-Host ('  ->  output\{0}  ({1})' -f (Split-Path $result.Target -Leaf), $detail) -ForegroundColor DarkGray

            Write-Log -Level OK -NoConsole ('{0} -> {1} | {2}' -f $doc.Name, (Split-Path $result.Target -Leaf), $detail)
            if ($result.TextStats) {
                Write-Log -Level INFO -NoConsole ('    text engine: {0} header/footer line(s), {1} page number(s), {2} hyphen join(s), {3} list marker(s), {4} heading(s), {5} multi-column page(s)' -f
                    $result.TextStats.HeadersRemoved, $result.TextStats.NumbersRemoved, $result.TextStats.HyphensJoined,
                    $result.TextStats.ListsConverted, $result.TextStats.HeadingsFound, $result.TextStats.MultiColumn)
            }
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
    Write-Host ('Engines   : {0} via Word, {1} via text extraction; {2} page(s) total' -f
        $totals.Word, $totals.Text, $totals.Pages) -ForegroundColor Gray
    Write-Host ('Removed {0} link URL(s), {1} image(s), {2} footnote(s), {3} anchor(s)' -f
        $totals.Links, $totals.Images, $totals.Footnotes, $totals.Anchors) -ForegroundColor Gray
    Write-Host ('Converted {0} HTML table(s), unwrapped {1} HTML tag(s); replaced {2} keyword occurrence(s)' -f
        $totals.Tables, $totals.Html, $totals.Keywords) -ForegroundColor Gray
    Write-Host ('Output    : {0}' -f $folders.Output) -ForegroundColor Gray
    Write-Host ('Log       : {0}' -f $script:LogPath) -ForegroundColor DarkGray
    Write-Host ''

    Write-Log -Level INFO -NoConsole ('Done. ok={0} failed={1} word={2} text={3} pages={4} links={5} images={6} footnotes={7} anchors={8} tables={9} html={10} keywords={11}' -f
        $ok, $failed, $totals.Word, $totals.Text, $totals.Pages, $totals.Links, $totals.Images,
        $totals.Footnotes, $totals.Anchors, $totals.Tables, $totals.Html, $totals.Keywords)

    exit $(if ($failed) { 1 } else { 0 })
}
catch {
    Write-Host ''
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Log -Level ERROR -NoConsole ('Fatal: {0}' -f $_.Exception.ToString())
    exit 1
}

#endregion
