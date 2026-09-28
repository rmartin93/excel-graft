<#
.SYNOPSIS
    Opens .xlsx files in real Excel via COM automation: verifies each opens
    and re-saves without throwing, and prints a structural fingerprint
    (sheets, defined names, tables + their ranges/row counts/totals, used
    range) for each.

.DESCRIPTION
    What this DOES catch, reliably (verified live against a real corrupted
    file on 2026-09-28): hard corruption that makes Excel refuse to open or
    save a file, and structural regressions — a table losing its totals
    row, a table range not covering the rows it should, a defined name
    disappearing, a sheet going missing.

    What this does NOT catch: subtle corruption in extLst / x14 extensions
    (newer conditional formatting, slicers, threaded comments, etc). Tested
    live against corpus/templates/exceljs-fork-broken/Test-Output.xlsx,
    which is known to trigger Excel's repair prompt when opened normally:
    with DisplayAlerts=$false, Excel opens it with no thrown exception, no
    change to Workbook.Name, and identical ListObjects/Names/UsedRange
    shape versus the uncorrupted source template. The corruption is real
    (the user saw the repair prompt interactively) but invisible to both
    COM exceptions and the pre-2010 object-model surface this script can
    cheaply inspect. Automating detection of the repair dialog itself
    (DisplayAlerts=$true + watching for the dialog window) was also tried
    live and did not surface a dialog either — Excel appears to resolve
    this particular repair silently even when alerts are enabled for
    automation.

    Net: treat this script as a fast smoke test for hard failures and
    structural regressions, not as a substitute for the manual step in
    PROJECT.md — open a suspect file in Excel interactively, and if the
    repair prompt appears, use the notification bar's "repair log" link to
    save Repairs.xml into corpus/repair-logs/. That log is what tells you
    exactly which XML part was dropped, which neither this script nor the
    Open XML SDK validator (tools/validator/, not yet built) can currently
    give you for this class of corruption.

.PARAMETER Path
    A file or directory of .xlsx files to check. Defaults to corpus/templates.

.EXAMPLE
    ./Test-ExcelFiles.ps1
    ./Test-ExcelFiles.ps1 -Path ../../corpus/templates/exceljs-fork-broken/Test-Output.xlsx
#>

[CmdletBinding()]
param(
	[string]$Path
)

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrEmpty($Path)) {
	$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
	$Path = Join-Path $scriptDir "..\..\corpus\templates"
}

function Get-XlsxFiles([string]$TargetPath) {
	if (Test-Path $TargetPath -PathType Leaf) {
		return @(Get-Item $TargetPath)
	}
	return Get-ChildItem -Path $TargetPath -Filter "*.xlsx" -Recurse -File
}

function Get-Fingerprint($Workbook) {
	$sheets = @()
	foreach ($ws in $Workbook.Worksheets) {
		$tables = @()
		foreach ($lo in $ws.ListObjects) {
			$tables += [ordered]@{
				Name       = $lo.Name
				Range      = $lo.Range.Address()
				DataRows   = $lo.ListRows.Count
				ShowTotals = [bool]$lo.ShowTotals
			}
		}
		$sheets += [ordered]@{
			Name      = $ws.Name
			UsedRange = $ws.UsedRange.Address()
			Charts    = $ws.ChartObjects().Count
			Tables    = $tables
		}
	}
	return [ordered]@{
		Sheets = $sheets
		Names  = @($Workbook.Names | ForEach-Object { $_.Name })
	}
}

function Test-OneFile([System.IO.FileInfo]$File, $Excel) {
	$result = [PSCustomObject]@{
		File        = $File.FullName
		Opened      = $false
		SavedOk     = $false
		Error       = $null
		Fingerprint = $null
	}

	try {
		$wb = $Excel.Workbooks.Open($File.FullName)
		$result.Opened = $true
		$result.Fingerprint = Get-Fingerprint $wb

		$tmpOut = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), [System.IO.Path]::GetRandomFileName() + ".xlsx")
		$wb.SaveAs($tmpOut, 51) # xlOpenXMLWorkbook
		$wb.Close($false)
		Remove-Item $tmpOut -Force -ErrorAction SilentlyContinue
		$result.SavedOk = $true
	}
	catch {
		$result.Error = $_.Exception.Message
	}

	return $result
}

$files = Get-XlsxFiles -TargetPath $Path
if ($files.Count -eq 0) {
	Write-Warning "No .xlsx files found under $Path"
	exit 1
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
	Write-Host "Checking $($files.Count) file(s) against real Excel via COM..."
	$results = @()
	foreach ($f in $files) {
		Write-Host "  $($f.Name)..." -NoNewline
		$r = Test-OneFile -File $f -Excel $excel
		$results += $r
		if (-not $r.Opened) {
			Write-Host " FAILED TO OPEN ($($r.Error))" -ForegroundColor Red
		}
		elseif (-not $r.SavedOk) {
			Write-Host " OPENED but re-save failed ($($r.Error))" -ForegroundColor Red
		}
		else {
			$tableCount = ($r.Fingerprint.Sheets | ForEach-Object { $_.Tables.Count } | Measure-Object -Sum).Sum
			Write-Host " OK ($($r.Fingerprint.Sheets.Count) sheet(s), $tableCount table(s))" -ForegroundColor Green
		}
	}
}
finally {
	$excel.Quit()
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
	[GC]::Collect()
	[GC]::WaitForPendingFinalizers()
}

Write-Host "`n--- Fingerprints ---"
foreach ($r in $results) {
	Write-Host "`n$($r.File):"
	if ($r.Fingerprint) {
		$r.Fingerprint | ConvertTo-Json -Depth 6 | Write-Host
	}
}

$failures = $results | Where-Object { -not $_.Opened -or -not $_.SavedOk }
if ($failures.Count -gt 0) {
	Write-Host "`n$($failures.Count) of $($results.Count) file(s) failed to open/save cleanly." -ForegroundColor Red
	exit 1
}

Write-Host "`nAll $($results.Count) file(s) opened and re-saved without a COM exception." -ForegroundColor Green
Write-Host "Reminder: this does not prove no repair prompt occurred (see script header). For subtle corruption, check interactively and capture the repair log." -ForegroundColor Yellow
exit 0
