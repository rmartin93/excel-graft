<#
.SYNOPSIS
    Builds two grouped-layout corpus templates in real Excel:
      grouped-footers  - department groups with a subtotal row BELOW each group,
                         a blank spacer row after each group, banded detail rows,
                         a running balance column, and a grand total.
      nested-groups    - Region > Category > line items, with subtotal header rows
                         at both levels and a grand total over the regions.
#>
$ErrorActionPreference = "Stop"
$root = Resolve-Path (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "..\..")
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

function Save-Book($wb, $name) {
	$dir = Join-Path $root "corpus\templates\$name"
	New-Item -ItemType Directory -Force $dir | Out-Null
	$path = Join-Path $dir "template.xlsx"
	if (Test-Path $path) { Remove-Item $path -Force }
	$wb.SaveAs($path, 51)
	$wb.Close($false)
	Write-Host "saved $path"
}

try {
	# ---------------------------------------------------------------- grouped-footers
	$wb = $excel.Workbooks.Add()
	$ws = $wb.Worksheets.Item(1)
	$ws.Name = "Payroll"
	$ws.Range("A1").Value2 = "Payroll by Department"
	$ws.Range("A1").Font.Size = 16
	$ws.Range("A1").Font.Bold = $true
	$ws.Range("A2").Value2 = "Period: {{period}}"
	$hdr = @("Employee", "Title", "Gross Pay", "Running Total")
	for ($c = 0; $c -lt 4; $c++) { $ws.Cells.Item(4, $c + 1).Value2 = $hdr[$c] }
	$ws.Range("A4:D4").Font.Bold = $true
	$ws.Range("A4:D4").Interior.Color = 0x5A3A1F   # dark blue (BGR)
	$ws.Range("A4:D4").Font.Color = 0xFFFFFF

	$groups = @(
		@{ Name = "Engineering"; Rows = @(@("Avery Kim", "Engineer III", 9200), @("Sam Patel", "Engineer II", 7800), @("Jordan Silva", "Architect", 11050)) },
		@{ Name = "Finance"; Rows = @(@("Riley Novak", "Controller", 9900), @("Casey Cohen", "Analyst", 6100)) }
	)
	$r = 5
	$footers = @()
	foreach ($g in $groups) {
		$first = $r
		$i = 0
		foreach ($row in $g.Rows) {
			$ws.Cells.Item($r, 1).Value2 = [string]$row[0]
			$ws.Cells.Item($r, 2).Value2 = [string]$row[1]
			$ws.Cells.Item($r, 3).Value2 = [double]$row[2]
			if ($r -eq $first) { $ws.Cells.Item($r, 4).Formula = "=C$r" } else { $ws.Cells.Item($r, 4).Formula = "=D$($r - 1)+C$r" }
			if ($i % 2 -eq 1) { $ws.Range("A${r}:D${r}").Interior.Color = 0xF7EBDD }  # banded light blue
			$i++
			$r++
		}
		$last = $r - 1
		$ws.Cells.Item($r, 1).Value2 = "$($g.Name) subtotal"
		$ws.Cells.Item($r, 3).Formula = "=SUM(C${first}:C${last})"
		$ws.Range("A${r}:D${r}").Font.Bold = $true
		$ws.Range("A${r}:D${r}").Borders.Item(8).LineStyle = 1   # top border
		$footers += "C$r"
		$r++
		$r++   # blank spacer row
	}
	$ws.Cells.Item($r, 1).Value2 = "Total payroll"
	$ws.Cells.Item($r, 3).Formula = "=SUM($($footers -join ','))"
	$ws.Range("A${r}:D${r}").Font.Bold = $true
	$ws.Range("A${r}:D${r}").Interior.Color = 0x5A3A1F
	$ws.Range("A${r}:D${r}").Font.Color = 0xFFFFFF
	$ws.Range("C5:D$r").NumberFormat = '$#,##0.00'
	$wb.Names.Add("Payroll", "=Payroll!`$A`$5:`$D`$$r") | Out-Null
	$ws.Cells.Item($r + 2, 1).Value2 = "Average per department"
	$ws.Cells.Item($r + 2, 3).Formula = "=C$r/$($groups.Count)"
	$ws.Columns.Item(1).ColumnWidth = 24
	$ws.Columns.Item(2).ColumnWidth = 16
	$ws.Columns.Item(3).ColumnWidth = 14
	$ws.Columns.Item(4).ColumnWidth = 15
	Save-Book $wb "grouped-footers"

	# ---------------------------------------------------------------- nested-groups
	$wb = $excel.Workbooks.Add()
	$ws = $wb.Worksheets.Item(1)
	$ws.Name = "Budget"
	$ws.Range("A1").Value2 = "Budget vs Actual"
	$ws.Range("A1").Font.Size = 16
	$ws.Range("A1").Font.Bold = $true
	$hdr = @("Line", "Budget", "Actual", "Variance")
	for ($c = 0; $c -lt 4; $c++) { $ws.Cells.Item(3, $c + 1).Value2 = $hdr[$c] }
	$ws.Range("A3:D3").Font.Bold = $true
	$ws.Range("A3:D3").Borders.Item(9).LineStyle = 1

	$regions = @(
		@{ Name = "North"; Cats = @(@{ Name = "Travel"; Lines = @(@("Airfare", 5000, 4200), @("Hotels", 3000, 3350)) }, @{ Name = "Software"; Lines = @(, @("Licenses", 8000, 7600)) }) },
		@{ Name = "South"; Cats = @(@{ Name = "Travel"; Lines = @(@("Airfare", 4000, 4100), @("Mileage", 1200, 900)) }, @{ Name = "Training"; Lines = @(, @("Courses", 2500, 2500)) }) }
	)
	$r = 4
	$regionRows = @()
	foreach ($reg in $regions) {
		$regRow = $r
		$regionRows += "B$regRow"
		$ws.Cells.Item($r, 1).Value2 = $reg.Name
		$ws.Range("A${r}:D${r}").Interior.Color = 0x9C6A2E
		$ws.Range("A${r}:D${r}").Font.Color = 0xFFFFFF
		$ws.Range("A${r}:D${r}").Font.Bold = $true
		$r++
		$catRows = @()
		foreach ($cat in $reg.Cats) {
			$catRow = $r
			$catRows += $catRow
			$ws.Cells.Item($r, 1).Value2 = "  $($cat.Name)"
			$ws.Range("A${r}:D${r}").Interior.Color = 0xF2E3D5
			$ws.Range("A${r}:D${r}").Font.Bold = $true
			$r++
			foreach ($line in $cat.Lines) {
				$ws.Cells.Item($r, 1).Value2 = "    $([string]$line[0])"
				$ws.Cells.Item($r, 2).Value2 = [double]$line[1]
				$ws.Cells.Item($r, 3).Value2 = [double]$line[2]
				$ws.Cells.Item($r, 4).Formula = "=C$r-B$r"
				$r++
			}
			$ws.Cells.Item($catRow, 2).Formula = "=SUM(B$($catRow + 1):B$($r - 1))"
			$ws.Cells.Item($catRow, 3).Formula = "=SUM(C$($catRow + 1):C$($r - 1))"
			$ws.Cells.Item($catRow, 4).Formula = "=C$catRow-B$catRow"
		}
		$ws.Cells.Item($regRow, 2).Formula = "=SUM($(($catRows | ForEach-Object { "B$_" }) -join ','))"
		$ws.Cells.Item($regRow, 3).Formula = "=SUM($(($catRows | ForEach-Object { "C$_" }) -join ','))"
		$ws.Cells.Item($regRow, 4).Formula = "=C$regRow-B$regRow"
	}
	$ws.Cells.Item($r, 1).Value2 = "Grand total"
	$ws.Cells.Item($r, 2).Formula = "=SUM($($regionRows -join ','))"
	$ws.Cells.Item($r, 3).Formula = "=SUM($(($regionRows | ForEach-Object { $_ -replace 'B', 'C' }) -join ','))"
	$ws.Cells.Item($r, 4).Formula = "=C$r-B$r"
	$ws.Range("A${r}:D${r}").Font.Bold = $true
	$ws.Range("A${r}:D${r}").Borders.Item(8).LineStyle = -4119   # double top border
	$ws.Range("B4:D$r").NumberFormat = '#,##0;[Red](#,##0)'
	$tbl = $ws.ListObjects.Add(1, $ws.Range("A3:D$r"), $null, 1)
	$tbl.Name = "Budget"
	$tbl.TableStyle = "TableStyleLight15"
	$fc = $ws.Range("D4:D$r").FormatConditions.Add(1, 6, "=0")   # xlCellValue, xlLess
	$fc.Font.Color = 0x0000C0
	$ws.Columns.Item(1).ColumnWidth = 22
	for ($c = 2; $c -le 4; $c++) { $ws.Columns.Item($c).ColumnWidth = 12 }
	# A chart over the grand total row's neighbours
	$ch = $ws.Shapes.AddChart2(201, 51, 360, 30, 360, 220).Chart
	$ch.SetSourceData($ws.Range("A4:C4,A10:C10"))
	$ch.HasTitle = $true
	$ch.ChartTitle.Text = "Regions: budget vs actual"
	Save-Book $wb "nested-groups"
}
finally {
	$excel.Quit()
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
	[GC]::Collect()
	[GC]::WaitForPendingFinalizers()
}
