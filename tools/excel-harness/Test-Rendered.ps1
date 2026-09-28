<#
.SYNOPSIS
    Ground-truth check of rendered workbooks in real Excel.

.DESCRIPTION
    `npm test` writes rendered outputs plus *.manifest.json files into
    corpus/generated/. This script opens each output in Excel via COM,
    forces a full recalculation, and checks:
      - every expected cell value (regenerated subtotals/totals must equal
        the sum of the data that was written),
      - every table's data row count,
      - no formula evaluates to #REF! or #NAME? (a broken reference is how
        a wrong row mapping shows up once Excel recalculates).
    It also counts other formula errors (#DIV/0!, #VALUE!, ...) as warnings,
    since hostile test data can cause those legitimately.

    Results go to corpus/generated/excel-results.json. Exit code 1 on any
    failure.

.EXAMPLE
    npm run harness:rendered
#>
[CmdletBinding()]
param(
	[string]$Dir
)

$ErrorActionPreference = "Stop"
if ([string]::IsNullOrEmpty($Dir)) {
	$Dir = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "..\..\corpus\generated"
}
$Dir = (Resolve-Path $Dir).Path

$manifests = Get-ChildItem -Path $Dir -Filter "*.manifest.json" -File
if ($manifests.Count -eq 0) {
	Write-Warning "No manifests in $Dir - run npm test first."
	exit 1
}

$entries = @()
foreach ($m in $manifests) {
	$json = Get-Content $m.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
	foreach ($e in $json) { $entries += $e }
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.ScreenUpdating = $false
$results = @()
$failed = 0

function Test-Close([double]$a, [double]$b) {
	return [math]::Abs($a - $b) -le (1e-6 * [math]::Max(1, [math]::Abs($b)))
}

try {
	Write-Host "Checking $($entries.Count) rendered file(s) in Excel $($excel.Version)..."
	foreach ($e in $entries) {
		$path = Join-Path $Dir $e.file
		$problems = New-Object System.Collections.Generic.List[string]
		$warnings = New-Object System.Collections.Generic.List[string]
		$probeResults = @()
		$started = Get-Date
		Write-Host ("  {0,-48}" -f $e.file) -NoNewline
		$wb = $null
		try {
			$wb = $excel.Workbooks.Open($path, 0, $true)
			$excel.CalculateFull()

			foreach ($c in $e.cells) {
				$ws = $wb.Worksheets.Item($c.sheet)
				$v = $ws.Range($c.cell).Value2
				if ($c.value -is [string]) {
					if ("$v" -ne $c.value) { $problems.Add("$($c.sheet)!$($c.cell) = '$v', expected '$($c.value)' ($($c.note))") }
				}
				elseif ($null -eq $v -or -not ($v -is [double]) -or -not (Test-Close $v $c.value)) {
					$problems.Add("$($c.sheet)!$($c.cell) = $v, expected $($c.value) ($($c.note))")
				}
			}

			foreach ($t in $e.tables) {
				$found = $null
				foreach ($ws in $wb.Worksheets) {
					foreach ($lo in $ws.ListObjects) { if ($lo.Name -eq $t.name) { $found = $lo } }
				}
				if ($null -eq $found) { $problems.Add("table $($t.name) not found") }
				elseif ($found.ListRows.Count -ne $t.dataRows) { $problems.Add("table $($t.name) has $($found.ListRows.Count) data rows, expected $($t.dataRows)") }
			}

			foreach ($p in @($e.probes)) {
				if ($null -eq $p) { continue }
				$ws = $wb.Worksheets.Item($p.sheet)
				$lastCol = [math]::Min($ws.UsedRange.Column + $ws.UsedRange.Columns.Count - 1, 40)
				$cells = @()
				for ($c = 1; $c -le $lastCol; $c++) {
					$cell = $ws.Cells.Item([int]$p.row, $c)
					$text = "$($cell.Text)".Trim()
					if ($text -ne "") { $cells += [PSCustomObject]@{ address = $cell.Address($false, $false); text = $text } }
				}
				$probeResults += [PSCustomObject]@{ label = $p.label; row = $p.row; cells = $cells }
			}

			foreach ($ws in $wb.Worksheets) {
				$errCells = $null
				try { $errCells = $ws.UsedRange.SpecialCells(-4123, 16) } catch { $errCells = $null } # xlCellTypeFormulas, xlErrors
				if ($null -eq $errCells) { continue }
				$broken = 0
				$other = 0
				$example = ""
				foreach ($cell in $errCells.Cells) {
					$text = $cell.Text
					if ($text -eq "#REF!" -or $text -eq "#NAME?") {
						$broken++
						if ($example -eq "") { $example = "$($cell.Address($false, $false)) $($cell.Formula)" }
					}
					else { $other++ }
					if ($broken + $other -gt 5000) { break }
				}
				if ($broken -gt 0) {
					$msg = "$($ws.Name): $broken formula(s) show #REF!/#NAME? (e.g. $example)"
					if ($e.allowBrokenRefs) { $warnings.Add("$msg - expected, the render warned about it") } else { $problems.Add($msg) }
				}
				if ($other -gt 0) { $warnings.Add("$($ws.Name): $other formula(s) show other errors") }
			}
		}
		catch {
			# With DisplayAlerts off, Excel answers its own repair prompt with "No" and Open throws:
			# this is how a repair prompt shows up under automation (see corpus/known-bad/README.md).
			$problems.Add("Excel refused to open it normally - a repair prompt: $($_.Exception.Message)")
		}
		finally {
			if ($null -ne $wb) { $wb.Close($false) | Out-Null }
		}
		$secs = [math]::Round(((Get-Date) - $started).TotalSeconds, 1)
		if ($problems.Count -eq 0) {
			$suffix = if ($warnings.Count -gt 0) { " (warnings: $($warnings -join '; '))" } else { "" }
			Write-Host "OK  $($e.cells.Count) value(s) checked, ${secs}s$suffix" -ForegroundColor Green
		}
		else {
			$failed++
			Write-Host "FAILED" -ForegroundColor Red
			foreach ($p in $problems) { Write-Host "      - $p" -ForegroundColor Red }
		}
		$results += [PSCustomObject]@{ file = $e.file; ok = ($problems.Count -eq 0); problems = @($problems); warnings = @($warnings); probes = @($probeResults); seconds = $secs }
	}
}
finally {
	$excel.Quit()
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
	[GC]::Collect()
	[GC]::WaitForPendingFinalizers()
}

$results | ConvertTo-Json -Depth 5 | Set-Content -Path (Join-Path $Dir "excel-results.json") -Encoding UTF8
if ($failed -gt 0) {
	Write-Host "`n$failed of $($results.Count) file(s) failed in Excel." -ForegroundColor Red
	exit 1
}
Write-Host "`nAll $($results.Count) rendered file(s) computed correctly in Excel." -ForegroundColor Green
exit 0
