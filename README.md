# docx-pdf-to-md

Two PowerShell scripts that batch-convert Word documents and PDFs into clean Markdown.
Hyperlink URLs are stripped, images are removed, tables are normalised, and a configurable set
of keywords is replaced — the same cleanup pipeline for both, so a mixed pile of source
documents comes out looking consistent.

| Script | Converts | Version |
|---|---|---|
| [`Convert-DocxToMd.ps1`](Convert-DocxToMd.ps1) | `.docx` → `.md` | 1.4.0 |
| [`Convert-PdfToMd.ps1`](Convert-PdfToMd.ps1) | `.pdf` → `.md` | 1.0.0 |

Both share one workspace, one `replacements.json` and one `output\` folder. Each script owns its
folder structure and creates it on first run:

```
docx-pdf-to-md\
  Convert-DocxToMd.ps1
  Convert-PdfToMd.ps1
  replacements.json     keyword mapping, used by both
  input\                put your .docx and .pdf files here
  output\               .md files land here
  logs\                 convert-*.log (docx) and pdfconvert-*.log (pdf)
```

## Requirements

| | `Convert-DocxToMd.ps1` | `Convert-PdfToMd.ps1` |
|---|---|---|
| PowerShell 7+ | required | required |
| [pandoc](https://pandoc.org) on `PATH` | required | required |
| Microsoft Word | not used | recommended — gives far better output, see [Engines](#engines) |
| Xpdf `pdftotext` on `PATH` | not used | required (it is the fallback engine and the scanned-PDF check) |

```powershell
winget install --id JohnMacFarlane.Pandoc
```

### About `pdftotext`

`pdftotext.exe` ships with **Git for Windows** (`C:\Program Files\Git\mingw64\bin`), so it is
usually already present and nothing needs installing.

Two things are worth knowing, because both cause confusing failures:

- **`pdftotext.exe` is not `xpdf.exe`.** Xpdf is published as two separate downloads.
  *XpdfReader* is the graphical PDF viewer (`xpdf.exe`) and contains no command line tools;
  it cannot write text to a file and this script cannot use it. What you need is the
  **Xpdf command line tools** package, which contains `pdftotext.exe` in its `bin64` folder.
  Both are at [xpdfreader.com/download.html](https://www.xpdfreader.com/download.html).
- **It must be the Xpdf build, not poppler's.** Poppler ships a `pdftotext` with the same name
  but different options and different exit codes. The script checks the vendor and refuses to run
  rather than silently producing wrong output.

If `pdftotext.exe` is not on `PATH`, the script also looks in the usual install folders. Failing
that, point at it directly:

```powershell
.\Convert-PdfToMd.ps1 -PdfToTextPath "C:\Program Files\xpdf-tools-win64\bin64\pdftotext.exe"
```

## Quick start

Drop your files into `input\` and run the matching script:

```powershell
.\Convert-DocxToMd.ps1
```

```powershell
.\Convert-PdfToMd.ps1
```

If `input\` is empty a dialog opens so you can pick files. **Selected files are copied into
`input\`, never moved** — your originals stay where they are. Conversion starts as soon as the
copy finishes.

Output files are never overwritten. A second run on `report.pdf` produces `report-1.md`, then
`report-2.md`. Because both scripts write into the same `output\`, a `report.docx` and a
`report.pdf` also collide this way: whichever runs first gets `report.md` and the other gets
`report-1.md`. The per-file log line records which source produced which file.

Exit code is `1` if any document failed, otherwise `0`.

### Common variations

Preview a run without writing anything (this never starts Word):

```powershell
.\Convert-PdfToMd.ps1 -WhatIf
```

Skip Word entirely — faster, no dependency on Word being closed, but no tables or bold:

```powershell
.\Convert-PdfToMd.ps1 -Engine Text
```

Insist on Word, so a document fails loudly instead of quietly falling back to weaker output:

```powershell
.\Convert-PdfToMd.ps1 -Engine Word -Verbose
```

Run unattended, with no dialog when `input\` is empty:

```powershell
.\Convert-PdfToMd.ps1 -NoPrompt
```

Convert a folder elsewhere, using a private copy of the keyword list:

```powershell
.\Convert-PdfToMd.ps1 -Root D:\client-work -ReplacementsPath D:\private\replacements.json
```

Give a stubborn, table-heavy PDF a better chance on the text engine:

```powershell
.\Convert-PdfToMd.ps1 -Engine Text -TextLayout Table -FenceLayoutBlocks
```

### Before converting PDFs: close Word

`Convert-PdfToMd.ps1` uses Word to reflow PDFs, and Word is a single-instance program — automating
it while you have documents open would attach to *your* session. To protect your unsaved work the
script **refuses to use Word whenever any Word process is running**, and falls back to plain-text
extraction instead. You will see:

```
Word is already running (1 process(es)). The Word engine is disabled to protect your
open documents - close Word and run again for better output.
```

Close Word and run again to get the better conversion.

## Engines

pandoc cannot read PDF, so `Convert-PdfToMd.ps1` needs a way to get the content out first. It has
two, and `-Engine Auto` (the default) tries them in order.

| | **Word** (preferred) | **Text** (fallback) |
|---|---|---|
| How | Word reflows the PDF into a real `.docx`, then pandoc converts it | Xpdf `pdftotext -layout` extracts plain text, which is then tidied into Markdown |
| Headings | recovered | only from numbering like `2.3 Scope` |
| Tables | real Markdown pipe tables | kept as aligned text, not tables |
| Bold / italic | preserved | lost |
| Page headers, footers, page numbers | dropped automatically | detected and removed heuristically |
| Needs | Word installed and closed | nothing beyond `pdftotext` |
| Speed | a few seconds per document | near-instant |

Force one with `-Engine Word` or `-Engine Text`. `-Engine Word` fails a document rather than
falling back, which is what you want when you are checking whether Word is doing its job.

The Word engine runs in a separate process with a hard timeout, so a PDF that hangs Word fails on
its own and the rest of the batch continues. It is also skipped automatically in non-interactive
sessions (scheduled tasks), where Office automation is unsupported and would simply hang.

### Limitations of PDF conversion

- **Scanned or image-only PDFs cannot be converted.** There is no OCR on this machine. They are
  detected before any work is done and reported as a failure rather than producing an empty file:

  ```
  FAIL  scan-2019.pdf  no extractable text (8 page(s), 0 non-whitespace characters). This looks
                       like a scanned or image-only PDF, and no OCR engine is available.
  ```

  Tune the threshold with `-MinCharsPerPage`, or set it to `0` to disable the check.

- **Multi-column pages defeat the Text engine.** `pdftotext -layout` interleaves both columns onto
  the same line and nothing downstream can unpick that. The script detects it and says so, naming
  the page and recommending `-Engine Word`, which handles columns properly.

- **Encrypted PDFs are rejected**, with the reason reported per file.

## Parameters

### Shared by both scripts

| Parameter | Effect |
|---|---|
| `-Root <path>` | Use a different workspace folder (default: the script's folder) |
| `-ReplacementsPath <file>` | Use a different mapping file (default: `replacements.json`) |
| `-NoPrompt` | Never show dialogs — for scheduled or unattended runs |
| `-KeepFootnotes` | Keep footnotes and endnotes instead of removing them |
| `-StripBareUrls` | Also remove bare URLs typed as plain text (off by default) |
| `-PandocTimeoutSec <n>` | Max seconds per document inside pandoc (default 120) |
| `-Version` | Print the script version and exit |
| `-WhatIf` | Show what would be written without writing anything (never starts Word) |

### `Convert-PdfToMd.ps1` only

| Parameter | Effect |
|---|---|
| `-Engine Auto\|Word\|Text` | Which extraction engine to use (default `Auto`) |
| `-WordTimeoutSec <n>` | Max seconds per document inside Word (default 300) |
| `-TextTimeoutSec <n>` | Max seconds per document inside pdftotext (default 120) |
| `-TextLayout Layout\|Table\|Simple\|Raw` | pdftotext extraction mode (default `Layout`; try `Table` for table-heavy documents) |
| `-PdfToTextPath <file>` | Full path to `pdftotext.exe`, for when it is not on `PATH`. Must be the command line tool, not `xpdf.exe` |
| `-MinCharsPerPage <n>` | Scanned-PDF threshold (default 50; `0` disables the check) |
| `-AllowRunningWord` | **Dangerous.** Automate Word even when it is open — risks your unsaved documents |

Text-engine tuning (ignored when Word does the conversion):

| Parameter | Effect |
|---|---|
| `-KeepRunningHeaders` | Keep page headers and footers that repeat across pages |
| `-KeepPageNumbers` | Keep lines that contain nothing but a page number |
| `-KeepLineHyphens` | Keep words hyphenated across a line break instead of rejoining them |
| `-ReflowParagraphs` | Join wrapped lines into single paragraphs. Off by default — Markdown already renders consecutive lines as one paragraph, so this can only do harm |
| `-DetectCapsHeadings` | Treat short ALL-CAPS lines as headings. Off by default — high false-positive rate on acronym-heavy documents |
| `-FenceLayoutBlocks` | Wrap column-aligned blocks (tables, ASCII art) in code fences so their alignment survives |
| `-TextStrictPipeline` | Run the full docx cleanup chain on text output too (see [Why the text engine skips two stages](#why-the-text-engine-skips-two-stages)) |

## What gets cleaned

| | Behaviour |
|---|---|
| Hyperlinks | Link text is kept, the URL is dropped: `[Contoso docs](https://…)` → `Contoso docs` |
| Bare URLs | Left alone unless `-StripBareUrls` is used |
| Images | Removed entirely; no media files are ever written to disk |
| Tracked changes | Accepted (insertions kept, deletions dropped) |
| Comments | Removed |
| Footnotes / endnotes | Removed, unless `-KeepFootnotes` |
| Document metadata | Never emitted (no YAML front matter) |
| Bookmark anchors | `<span id="_Toc232673667" class="anchor"></span>`, `<a id="…"></a>` and `[]{#_Toc… .anchor}` all removed |
| Inline HTML | Unwrapped — the words stay, the markup goes. See below |
| Tables | Always Markdown pipe tables — see below |
| Whitespace | Trailing spaces trimmed, runs of blank lines collapsed, non-breaking spaces normalised |
| Ligatures (PDF) | `ﬁ ﬂ ﬀ ﬃ ﬄ` expanded to plain letters, so keyword replacement and search cannot miss a word |

## Inline HTML

Word has character formatting Markdown cannot express, so pandoc emits raw HTML for it. All of it
is unwrapped — the text survives, the tags do not:

| Word formatting | Pandoc emits | Result |
|---|---|---|
| Highlight | `<span class="mark">text</span>` | `text` |
| Underline | `<u>text</u>` | `text` |
| Small caps | `<span class="smallcaps">text</span>` | `text` |
| Superscript / subscript | `<sup>2</sup>`, `<sub>2</sub>` | `2` |
| Custom character style | `<span class="MyStyle">text</span>` | `text` |

Formatting Markdown *can* express is kept: `<strong>`/`<b>` become `**bold**` and `<em>`/`<i>`
become `*italic*`. Pandoc's attribute syntax for the same thing (`[text]{.mark}`) is handled too.

**Fenced code blocks and inline code spans are never touched**, so a document containing HTML
samples keeps them intact.

### Why the text engine skips two stages

Those stages exist to clean up *markup pandoc generated*. Plain text extracted from a PDF has no
generated markup, so anything they match is literal document text. Two would actively cause damage
and are skipped:

- **Inline HTML removal** — a PDF that *documents* HTML would have its `<span>` and `<div>`
  examples eaten, and plain text has no code fences to protect them.
- **`[text](url)` link stripping** — legal and academic PDFs are full of `see clause [3](a)` and
  `Smith [2019](p. 4)`, which would silently collapse to `3` and `2019`.

`-TextStrictPipeline` runs the full chain anyway if you want identical processing regardless.

## Tables

Pandoc renders a Word table as a pipe table only when GFM can express it. Anything with merged
cells or multi-paragraph cells falls back to raw `<table style="width:100%;">` HTML, which is why
tables used to come out in two different shapes. Both scripts normalise them into pipe tables:

- Raw `<table>` blocks are parsed and rewritten as pipe tables. Merged cells are flattened — the
  content stays in the first cell of the span and the covered cells are left blank.
- Cell content is converted to inline Markdown: `<strong>`/`<em>`/`<code>` become Markdown,
  `<br>` and multiple `<p>` become spaces, `<a href>` keeps its text and loses its URL, HTML
  entities are decoded, and a literal `|` is escaped.
- A `<caption>` is emitted as a bold line above the table.
- Word tables rarely mark a header row, so pandoc emits an empty one. When the header row is
  entirely blank, the first body row is promoted into it.

Nested tables (a table inside a table cell) are not handled and will come through as raw HTML.

PDFs only get real tables through the Word engine. On the Text engine, column-aligned blocks are
left exactly as they are — never mangled, but never turned into pipe tables either. Use
`-FenceLayoutBlocks` to keep their alignment visible, or `-TextLayout Table`.

## Keyword replacement

Edit `replacements.json`. Only `find` and `replace` are required:

```json
[
  { "find": "Contoso", "replace": "Fabrikam" },
  { "find": "Project Bluebird", "replace": "Project Falcon", "wholeWord": false },
  { "find": "\\bv\\d+\\.\\d+\\b", "replace": "vNEXT", "regex": true, "preserveCase": false }
]
```

| Flag | Default | Meaning |
|---|---|---|
| `wholeWord` | `true` | Replacing `Sun` won't touch `Sunday`. Set to `false` for substring matching |
| `caseInsensitive` | `true` | Match regardless of capitalisation |
| `preserveCase` | `false` | Shape the replacement to match the casing that was found (see below) |
| `regex` | `false` | Treat `find` as a regular expression (`wholeWord` is then ignored) |

### Casing

**The replacement is written exactly as you spell it in `replacements.json`, always.** Matching
still ignores case, so every spelling in the document is found, but the output is verbatim:

| Rule | Found in document | Written to Markdown |
|---|---|---|
| `Contoso` → `ConsultantCompanyXYZ` | `Contoso` | `ConsultantCompanyXYZ` |
| | `contoso` | `ConsultantCompanyXYZ` |
| | `CONTOSO` | `ConsultantCompanyXYZ` |
| | `CoNtOsO` | `ConsultantCompanyXYZ` |

If you want a particular rule to mirror the casing it found instead, opt that entry in:

```json
{ "find": "contoso", "replace": "fabrikam", "preserveCase": true }
```

Then `CONTOSO` → `FABRIKAM`, `Contoso` → `Fabrikam`, `contoso` → `fabrikam`. Even with
`preserveCase` on, a replacement carrying deliberate internal capitals (`ConsultantCompanyXYZ`,
`Company ABC`, `iPhone`) is never flattened by title-casing.

Rules are applied in file order, and replacement output is not re-scanned by later rules.

**Bare URLs are not keyword-replaced.** Since they are kept verbatim, replacing inside one would
only corrupt it — `https://www.contoso.com` would become `https://www.Company ABC.com`. The
consequence is that a name you are scrubbing can still survive inside a URL. If that matters, run
with `-StripBareUrls` to remove those URLs entirely.

## Logging

Every run prints a per-file summary and writes a timestamped log:

- `logs\convert-yyyyMMdd-HHmmss.log` — `Convert-DocxToMd.ps1`
- `logs\pdfconvert-yyyyMMdd-HHmmss.log` — `Convert-PdfToMd.ps1`

The prefixes are deliberately different so that each script's retention sweep (newest 30 runs)
only ever touches its own logs.

Each log records the script version, the resolved tool paths, each source and target file, which
engine converted it, per-keyword hit counts, how many links, images, footnotes and anchors were
removed, how many HTML tables were converted, and any errors.

## Security notes

Audited 2026-07-21 (`Convert-DocxToMd.ps1`) and 2026-07-25 (`Convert-PdfToMd.ps1`); manual review
plus PSScriptAnalyzer. What the scripts protect against:

- External tools are invoked with an argument array — no shell string interpolation, so hostile
  filenames cannot inject options or commands. pandoc additionally gets a `--` terminator.
- Every regex (including `regex: true` patterns from `replacements.json`) runs with a 5-second
  match timeout, and every external tool runs under its own timeout — a malicious or corrupt
  document fails alone; the batch continues.
- Literal `find` terms are regex-escaped; output encoding is pinned to UTF-8 without BOM.
- **PDF only:** each document is staged into a private `%TEMP%` folder first, copying only the
  file's default data stream. That strips the Mark-of-the-Web that would otherwise open it in
  Protected View, and means Word never learns the original path — it cannot lock your file, add it
  to the recent-documents list, or leave a `~$` file beside it. Staging folders are removed after
  every document, and any older than 24 hours are swept at startup.
- **PDF only:** Word automation runs in a child process launched with a plain script path (not an
  encoded command line, which trips endpoint protection). All values reach it through the
  environment block, never the command line. Word is started with macros disabled, alerts
  suppressed, and document repair off, and its settings are never written to your profile.
- **PDF only:** the script refuses to drive a Word instance it did not start, and only ever ends a
  Word process it created — verified by process id, name and start time, so a recycled id cannot
  cause it to kill something unrelated.

**Known residual risks (reviewed and accepted by design):**

- **Log files record the `find` terms** from `replacements.json` (e.g. `keyword "X" replaced 7x`)
  — the names being scrubbed persist in `logs\` (newest 30 runs kept per script). Treat `logs\` as
  being as confidential as `replacements.json`.
- **Output filenames are not keyword-replaced** — content is scrubbed, the filename is copied
  as-is from the source document.
- **`input\` keeps the original, un-scrubbed documents** it copies in; clean it manually when
  the originals should not linger.
- `replacements.json` inherently contains the real terms; keep it (and this whole folder) out of
  any repository or shared location, or point at a private copy with `-ReplacementsPath`.
  Save it as UTF-8 — it is read as UTF-8.
- pandoc and pdftotext are resolved from `PATH`; the resolved paths are written to every log.
- Protected View cannot be switched off through Word's object model. Stripping Mark-of-the-Web
  handles the common case; if your Trust Center marks additional locations unsafe, the Word
  timeout is the remaining backstop.

## Version history

### Convert-PdfToMd.ps1

#### 1.0.0

- Initial release. Word-reflow and `pdftotext` engines feeding the Markdown cleanup pipeline of
  `Convert-DocxToMd` 1.4.0, so both file types produce consistent output.
- Refuses to automate Word while you have it open, and never ends a Word process it did not start.
- Scanned and encrypted PDFs are detected up front and fail with an explanation instead of writing
  an empty file; a final non-empty check covers both engines.
- Text engine removes running headers, footers and page numbers, rejoins hyphenated line breaks
  and detects numbered headings, while leaving column-aligned blocks untouched.
- Output writes are atomic, so the shared `output\` folder is safe when both scripts run at once.
- `pdftotext.exe` is found on `PATH`, in the usual Xpdf install folders, or via `-PdfToTextPath`.
  Pointing at `xpdf.exe` (the XpdfReader GUI) reports that specific mistake instead of a generic
  error, and a poppler build is rejected rather than used.

### Convert-DocxToMd.ps1

#### 1.4.0

- Hardening: 5-second match timeout on every regex (including `replacements.json` patterns) and
  a configurable pandoc timeout (`-PandocTimeoutSec`, default 120 s) — one bad document fails
  alone instead of hanging the batch.
- Code-fence tracking now pairs opening and closing markers (` ``` ` vs `~~~`) correctly.
- `logs\` keeps only the newest 30 runs.
- Performance: compiled regexes reused per table cell; lines without markup skip the HTML sweep.
- Security review documented (see Security notes); no behavior changes to conversion output.

#### 1.3.0

- Leftover inline HTML is unwrapped: `<span class="mark">`, `<u>`, `<sub>`, `<sup>`, small caps
  and custom character styles keep their words but lose the markup.
- `<strong>`/`<em>` are converted to Markdown rather than discarded.
- Fenced code blocks and inline code spans are protected from the sweep.
- Keyword replacement no longer runs inside bare URLs, which used to corrupt them.

#### 1.2.0

- Replacements are written exactly as spelled in `replacements.json`. Mirroring the casing found
  in the document is now opt-in per entry via `"preserveCase": true`.

#### 1.1.0

- Raw `<table>` HTML is converted to Markdown pipe tables, so every table has the same shape.
- Empty table headers are filled from the first body row.
- Word bookmark anchor markup (`<span id="_Toc…" class="anchor"></span>` and friends) is removed.
- Replacements with deliberate internal capitals are no longer flattened by title-casing.
- Added `-Version`; the version is shown in the banner and written to the log.

#### 1.0.0

- Initial release: pandoc conversion, URL stripping, image removal, case-preserving keyword
  replacement, mandatory `input\`/`output\` folders, file-picker copy-in, numbered output files.
