<#
.SYNOPSIS
    Opens .xlsx files in real Excel via COM automation: fails any file Excel
    would show the "We found a problem with some content ... Do you want us
    to try to recover" repair prompt for, verifies clean files re-save, and
    prints a structural fingerprint (sheets, defined names, tables + their
    ranges/row counts/totals, used range) for each.

.DESCRIPTION
    Repair detection (-DetectRepairs, on by default). Established
    empirically on 2026-09-28 against the fixtures in corpus/known-bad/ (see
    its README for the per-fixture evidence):

    1. PRIMARY SIGNAL - Workbooks.Open throws. With DisplayAlerts=$false and
       the default CorruptLoad (xlNormalLoad), Excel does NOT repair silently:
       the repair prompt is auto-answered "No" and Workbooks.Open throws
       "Unable to get the Open property of the Workbooks class". For every
       fixture, "Open threw" matched exactly whether an interactive user
       gets the repair prompt (verified with UI Automation on a visible
       instance). So a file that opens here would not have prompted a user
       of this Excel build.

    2. DIAGNOSIS - repair log. When the normal open fails, the file is
       opened again with CorruptLoad=xlRepairFile. That repairs silently and
       Excel writes a recovery log (%TEMP%\errorNNNNNN_01.xml) naming the
       part and record class it repaired or removed, e.g. "Removed Records:
       Merge cells from /xl/worksheets/sheet1.xml part". The log text is
       printed with the failure. (Some corruptions - malformed XML, a sqref
       past row 1048576 - fail in repair mode too; they are reported as
       "Excel could not open it even in repair mode".)

    3. BELT AND BRACES - for files that open normally, a second repair-mode
       open is done and its recovery log inspected. In repair mode Excel
       always writes a log; a log with no <removedRecords>/<repairedRecords>/
       <removedParts>/<repairedParts> entries means nothing was repaired. A
       log WITH entries fails the file. This never fired on the known-bad
       set (Excel had nothing to say about any file that opened normally),
       but it guards against a silent-repair class we have not seen yet.

    What this still does NOT catch: corruption that this Excel build
    tolerates silently but another (older or stricter) build prompts for.
    corpus/templates/exceljs-fork-broken/Test-Output.xlsx is the known
    example: it prompted on another machine, but on this build (16.0 build
    20326) it opens normally and a repair-mode open logs nothing. So "passes
    here" is not proof of validity; keep the structural checks in the unit
    tests for those known cases (see corpus/README.md).

.PARAMETER Path
    A file or directory of .xlsx files to check. Defaults to corpus/templates.

.PARAMETER DetectRepairs
    On by default. Pass -DetectRepairs:$false (in-process) or
    -SkipRepairDetection (works with powershell -File / npm run) to skip the
    repair-mode passes. A file whose normal open throws still fails, just
    without the repair-log diagnosis.

.EXAMPLE
    ./Test-ExcelFiles.ps1
    ./Test-ExcelFiles.ps1 -Path ../../corpus/known-bad    # every file should FAIL
    ./Test-ExcelFiles.ps1 -Path ../../corpus/templates/exceljs-fork-broken/Test-Output.xlsx
#>

[CmdletBinding()]
param(
	[string]$Path,
	[switch]$DetectRepairs = $true,
	[switch]$SkipRepairDetection
)
# "-DetectRepairs:$false" does not survive "powershell -File" argument
# parsing, so -SkipRepairDetection is the -File-friendly way to turn it off.
if ($SkipRepairDetection) { $DetectRepairs = $false }

$ErrorActionPreference = "Stop"

if ([string]::IsNullOrEmpty($Path)) {
	$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
	$Path = Join-Path $scriptDir "..\..\corpus\templates"
}

Add-Type @"
using System; using System.Runtime.InteropServices;
public static class ExcelHarnessWin32 {
	[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
"@

$xlRepairFile = 1
$repairLogDir = [System.IO.Path]::GetTempPath()

function Get-XlsxFiles([string]$TargetPath) {
	if (Test-Path $TargetPath -PathType Leaf) {
		return @(Get-Item $TargetPath)
	}
	return @(Get-ChildItem -Path $TargetPath -Filter "*.xlsx" -Recurse -File | Where-Object { -not $_.Name.StartsWith("~$") })
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

function Get-RepairLogNames {
	return @(Get-ChildItem -Path $repairLogDir -Filter "error*.xml" -File -ErrorAction SilentlyContinue | ForEach-Object { $_.Name })
}

# Finds the recovery log(s) Excel wrote for $FilePath since $Before, returns
# the list of repaired/removed entries (empty list = log found, nothing
# repaired; $null = no log found), and deletes the logs it consumed.
function Read-RepairLog([string]$FilePath, [string[]]$Before) {
	$deadline = (Get-Date).AddSeconds(3)
	do {
		$new = @(Get-ChildItem -Path $repairLogDir -Filter "error*.xml" -File -ErrorAction SilentlyContinue | Where-Object { $Before -notcontains $_.Name })
		$found = $null
		foreach ($f in $new) {
			try { $x = [xml](Get-Content -LiteralPath $f.FullName -Raw) } catch { continue }
			if ($null -eq $x.recoveryLog) { continue }
			if (-not ([string]$x.recoveryLog.summary).Contains($FilePath)) { continue }
			if ($null -eq $found) { $found = @() }
			foreach ($n in $x.recoveryLog.ChildNodes) {
				if (@("logFileName", "summary", "additionalInfo") -contains $n.LocalName) { continue }
				foreach ($entry in $n.ChildNodes) { $found += ($entry.InnerText -replace "\s+", " ").Trim() }
			}
			Remove-Item -LiteralPath $f.FullName -Force -ErrorAction SilentlyContinue
		}
		if ($null -ne $found) { return , $found }
		Start-Sleep -Milliseconds 200
	} while ((Get-Date) -lt $deadline)
	return $null
}

# Opens in xlRepairFile mode (silent repair) and returns the repair entries.
function Invoke-RepairProbe([string]$FilePath, $Excel) {
	$before = Get-RepairLogNames
	$probe = [ordered]@{ Opened = $false; Entries = $null; Error = $null }
	try {
		# The full 15-argument form is required: passing [Type]::Missing for
		# the middle optional arguments makes Open fail on this build.
		$wb = $Excel.Workbooks.Open($FilePath, 0, $true, 5, "", "", $true, 2, ",", $false, $false, 0, $false, $false, $xlRepairFile)
		$probe.Opened = $true
		$wb.Close($false)
	}
	catch {
		$probe.Error = $_.Exception.Message
	}
	$probe.Entries = Read-RepairLog -FilePath $FilePath -Before $before
	return $probe
}

function Test-OneFile([System.IO.FileInfo]$File, $Excel) {
	$result = [PSCustomObject]@{
		File           = $File.FullName
		Opened         = $false
		SavedOk        = $false
		RepairDetected = $false
		RepairDetail   = @()
		Error          = $null
		Fingerprint    = $null
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

	if (-not $result.Opened) {
		# Signal 1: with DisplayAlerts off, a failed normal open is exactly the
		# case where an interactive user gets the repair prompt.
		$result.RepairDetected = $true
		if ($DetectRepairs) {
			$probe = Invoke-RepairProbe -FilePath $File.FullName -Excel $Excel
			if (-not $probe.Opened) {
				$result.RepairDetail = @("Excel could not open it even in repair mode ($($probe.Error))")
			}
			elseif ($null -eq $probe.Entries) {
				$result.RepairDetail = @("repair-mode open succeeded but no recovery log was found")
			}
			else {
				$result.RepairDetail = @($probe.Entries)
			}
		}
	}
	elseif ($DetectRepairs) {
		# Signal 3: a repair-mode open that repairs anything means Excel found
		# something wrong even though the normal open was silent.
		$probe = Invoke-RepairProbe -FilePath $File.FullName -Excel $Excel
		if ($probe.Opened -and $probe.Entries -and $probe.Entries.Count -gt 0) {
			$result.RepairDetected = $true
			$result.RepairDetail = @($probe.Entries)
		}
	}

	return $result
}

$files = Get-XlsxFiles -TargetPath $Path
if ($files.Count -eq 0) {
	Write-Warning "No .xlsx files found under $Path"
	exit 1
}

$excel = New-Object -ComObject Excel.Application
$excelPid = 0
try { [void][ExcelHarnessWin32]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelPid) } catch { }
$excel.Visible = $false
$excel.DisplayAlerts = $false

$results = @()
try {
	Write-Host "Checking $($files.Count) file(s) against real Excel $($excel.Version) (build $($excel.Build)) via COM (repair detection: $([bool]$DetectRepairs))..."
	foreach ($f in $files) {
		Write-Host "  $($f.Directory.Name)/$($f.Name)..." -NoNewline
		$r = Test-OneFile -File $f -Excel $excel
		$results += $r
		if ($r.RepairDetected) {
			$how = if ($r.Opened) { "silent repair found by repair-mode open" } else { "normal open refused - Excel would show the repair prompt" }
			Write-Host " FAILED: REPAIR ($how)" -ForegroundColor Red
			foreach ($d in $r.RepairDetail) { Write-Host "      $d" -ForegroundColor Red }
		}
		elseif (-not $r.Opened) {
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
	try { foreach ($wb in @($excel.Workbooks)) { $wb.Close($false) } } catch { }
	try { $excel.Quit() } catch { }
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
	Remove-Variable excel -ErrorAction SilentlyContinue
	[GC]::Collect()
	[GC]::WaitForPendingFinalizers()
	# Never leave an Excel process behind (a stuck instance breaks the next run).
	if ($excelPid -gt 0) {
		$proc = Get-Process -Id $excelPid -ErrorAction SilentlyContinue
		if ($proc -and -not $proc.WaitForExit(15000)) {
			Write-Warning "Excel (pid $excelPid) did not exit after Quit; killing it."
			Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue
		}
	}
}

Write-Host "`n--- Fingerprints ---"
foreach ($r in $results) {
	Write-Host "`n$($r.File):"
	if ($r.Fingerprint) {
		$r.Fingerprint | ConvertTo-Json -Depth 6 | Write-Host
	}
}

$failures = @($results | Where-Object { $_.RepairDetected -or -not $_.Opened -or -not $_.SavedOk })
if ($failures.Count -gt 0) {
	Write-Host "`n$($failures.Count) of $($results.Count) file(s) FAILED (repair prompt or open/save failure):" -ForegroundColor Red
	foreach ($r in $failures) { Write-Host "  $($r.File)" -ForegroundColor Red }
	exit 1
}

Write-Host "`nAll $($results.Count) file(s) opened without a repair prompt and re-saved." -ForegroundColor Green
Write-Host "Reminder: this Excel build tolerates some corruption that stricter builds prompt for (see script header and corpus/README.md)." -ForegroundColor Yellow
exit 0
