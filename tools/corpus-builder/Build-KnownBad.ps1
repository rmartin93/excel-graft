<#
.SYNOPSIS
    Builds deliberately corrupt .xlsx fixtures in corpus/known-bad/ by
    hand-editing XML parts of clean corpus templates and re-zipping.

.DESCRIPTION
    Each fixture copies every zip entry of a clean template (built by
    Build-Corpus.ps1) in the original order -- so [Content_Types].xml stays
    the first entry -- replacing only the one part being corrupted. The edit
    is a literal string/regex replacement; the script throws if the expected
    text is not found, so a template change can't silently produce a
    "corrupt" fixture that is actually clean.

    Pure ASCII file (Windows PowerShell 5.1).

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File tools/corpus-builder/Build-KnownBad.ps1
#>
[CmdletBinding()]
param([string]$OutDir)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$templates = [System.IO.Path]::GetFullPath((Join-Path $scriptDir "..\..\corpus\templates"))
if ([string]::IsNullOrEmpty($OutDir)) {
	$OutDir = [System.IO.Path]::GetFullPath((Join-Path $scriptDir "..\..\corpus\known-bad"))
}
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$utf8 = New-Object System.Text.UTF8Encoding($false)

function Edit-Once([string]$Text, [string]$Pattern, [string]$Replacement) {
	$re = New-Object System.Text.RegularExpressions.Regex($Pattern)
	if (-not $re.IsMatch($Text)) { throw "Pattern not found: $Pattern" }
	return $re.Replace($Text, $Replacement, 1)
}

function New-Fixture([string]$Name, [string]$BaseTemplate, [hashtable]$Edits) {
	$src = Join-Path $templates "$BaseTemplate\template.xlsx"
	$dst = Join-Path $OutDir "$Name.xlsx"
	if (Test-Path $dst) { Remove-Item $dst -Force }
	$in = [System.IO.Compression.ZipFile]::OpenRead($src)
	$fs = [System.IO.File]::Open($dst, [System.IO.FileMode]::CreateNew)
	$out = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
	$applied = @{}
	try {
		foreach ($e in $in.Entries) {
			$ms = New-Object System.IO.MemoryStream
			$s = $e.Open(); $s.CopyTo($ms); $s.Dispose()
			$bytes = $ms.ToArray()
			if ($Edits.ContainsKey($e.FullName)) {
				$text = $utf8.GetString($bytes)
				$new = & $Edits[$e.FullName] $text
				if ($new -ceq $text) { throw "Edit to $($e.FullName) changed nothing" }
				$bytes = $utf8.GetBytes($new)
				$applied[$e.FullName] = $true
			}
			$ne = $out.CreateEntry($e.FullName, [System.IO.Compression.CompressionLevel]::Optimal)
			$ns = $ne.Open(); $ns.Write($bytes, 0, $bytes.Length); $ns.Dispose()
		}
	}
	finally {
		$out.Dispose(); $fs.Dispose(); $in.Dispose()
	}
	foreach ($k in $Edits.Keys) { if (-not $applied[$k]) { throw "$Name : part $k not found in $BaseTemplate" } }
	Write-Host "  $Name.xlsx  (from $BaseTemplate)"
}

Write-Host "Writing fixtures to $OutDir"

# Control: an unmodified template pushed through the same re-zip path. Must
# NOT be flagged -- proves the zip writer itself is not the corruption.
New-Fixture "control-rezipped-clean" "flat-table-totals" @{}

# (a1) Table autoFilter covers only the header row (the ExcelJS bug).
New-Fixture "a1-autofilter-header-only" "flat-table-totals" @{
	"xl/tables/table1.xml" = { param($t) Edit-Once $t '<autoFilter ref="A5:F10"' '<autoFilter ref="A5:F5"' }
}
# (a2) Table autoFilter narrower than the table (3 of 6 columns).
New-Fixture "a2-autofilter-narrow" "flat-table-totals" @{
	"xl/tables/table1.xml" = { param($t) Edit-Once $t '<autoFilter ref="A5:F10"' '<autoFilter ref="A5:C10"' }
}
# (b) Two <c> elements with the same r in one row.
New-Fixture "b-duplicate-cell-ref" "minimal-single-row" @{
	"xl/worksheets/sheet1.xml" = { param($t) Edit-Once $t '(<c r="C2" t="s"><v>5</v></c>)' '$1<c r="C2"><v>42</v></c>' }
}
# (c) Rows out of order: row 3 written before row 2.
New-Fixture "c-rows-out-of-order" "flat-table-totals" @{
	"xl/worksheets/sheet1.xml" = { param($t) Edit-Once $t '(<row r="2" .*?</row>)(<row r="3" .*?</row>)' '$2$1' }
}
# (d) Inline string longer than Excel's 32767-character cell limit.
New-Fixture "d-inline-string-too-long" "minimal-single-row" @{
	"xl/worksheets/sheet1.xml" = { param($t) Edit-Once $t '<c r="A2" t="s"><v>3</v></c>' ('<c r="A2" t="inlineStr"><is><t>' + ('x' * 40000) + '</t></is></c>') }
}
# (e) Data-validation sqref past the last row (1048576).
New-Fixture "e-sqref-past-max-row" "flat-table-totals" @{
	"xl/worksheets/sheet1.xml" = { param($t) Edit-Once $t 'sqref="C6:C10"' 'sqref="C6:C1048577"' }
}
# (e2) Conditional-formatting sqref past the last row.
New-Fixture "e2-cf-sqref-past-max-row" "flat-table-totals" @{
	"xl/worksheets/sheet1.xml" = { param($t) Edit-Once $t '<conditionalFormatting sqref="D6:D10">' '<conditionalFormatting sqref="D6:D1048577">' }
}
# (f) Table column name no longer matches the header cell text.
New-Fixture "f-table-header-mismatch" "minimal-single-row" @{
	"xl/tables/table1.xml" = { param($t) Edit-Once $t 'name="Name"/>' 'name="FullName"/>' }
}
# (g) A mergeCell overlapping another mergeCell.
New-Fixture "g-overlapping-merges" "stacked-regions" @{
	"xl/worksheets/sheet1.xml" = { param($t) Edit-Once $t '<mergeCells count="1"><mergeCell ref="A21:C24"/>' '<mergeCells count="2"><mergeCell ref="A21:C24"/><mergeCell ref="B23:D26"/>' }
}
# (h) Table ref shorter than its data (table ends one row early; the totals
#     row cell formulas now sit outside the table).
New-Fixture "h-table-ref-too-short" "flat-table-totals" @{
	"xl/tables/table1.xml" = { param($t) Edit-Once $t 'ref="A5:F11" totalsRowCount="1"' 'ref="A5:F10" totalsRowCount="1"' }
}
# (i) Malformed XML: sheet part is not well-formed.
New-Fixture "i-malformed-sheet-xml" "minimal-single-row" @{
	"xl/worksheets/sheet1.xml" = { param($t) Edit-Once $t '</sheetData>' '' }
}
# (j) Shared-formula child with no master (master cell's <f> removed) --
#     what a naive "delete the sample rows" does to indirect-rates column D.
New-Fixture "j-orphan-shared-formula" "indirect-rates" @{
	"xl/worksheets/sheet3.xml" = { param($t) Edit-Once $t '<f t="shared" ref="D5:D17" si="0">C5/\$C\$17</f>' '' }
}

Write-Host "Done."
