<#
.SYNOPSIS
    Builds corpus/templates/issue-1-disclosure/template.xlsx: a minimized
    stand-in for the workbook in GitHub issue #1.

    Sheet "Disclosure Table": a hierarchy shown only by styling -
    level-1 rows (yellow, label in A), level-2 rows (blue, label in B),
    detail rows - with NO subtotal formulas, no Table and no defined names,
    followed by blank rows and fixed notes + a picture below the data.
    Sheet "2026 CC": a flat block whose formulas refer back to sheet 1.
#>
$ErrorActionPreference = "Stop"
$root = Resolve-Path (Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "..\..")
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
	$wb = $excel.Workbooks.Add()
	while ($wb.Worksheets.Count -lt 2) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
	$ws = $wb.Worksheets.Item(1)
	$ws.Name = "Disclosure Table"
	$ws.Range("A1").Value2 = "FY2026 Disclosure Statement"
	$ws.Range("A1").Font.Size = 16
	$ws.Range("A1").Font.Bold = $true
	$ws.Range("A2").Value2 = "Cost Accounting Standards Board Disclosure Statement - Part VIII"
	$ws.Range("A4").Value2 = "Prepared from the cost accounting system"
	$hdr = @("Pool Reference", "Disclosure Reference", "Segment", "Base Code", "Disclosure Section", "D/S Reference", "Claim Pool Schedule", "Claim Base Schedule", "I", "SKF", "Cost Out", "In", "Out")
	for ($c = 0; $c -lt $hdr.Count; $c++) { $ws.Cells.Item(6, $c + 1).Value2 = $hdr[$c] }
	$ws.Range("A6:M6").Font.Bold = $true
	$ws.Range("A6:M6").Interior.Color = 0x404040
	$ws.Range("A6:M6").Font.Color = 0xFFFFFF
	$ws.Range("A6:M6").WrapText = $true

	$l1 = @(
		@{ Name = "Fringe Pool"; L2 = @(@{ Name = "Payroll Taxes"; D = 3 }, @{ Name = "Group Insurance"; D = 2 }) },
		@{ Name = "Overhead Pool"; L2 = @(@{ Name = "Facilities"; D = 2 }, @{ Name = "Indirect Labor"; D = 3 }) }
	)
	$r = 7
	$n = 0
	foreach ($a in $l1) {
		$ws.Cells.Item($r, 1).Value2 = [string]$a.Name
		$ws.Range("A${r}:M${r}").Interior.Color = 0x66FFFF   # yellow (BGR)
		$ws.Range("A${r}:M${r}").Font.Bold = $true
		$r++
		foreach ($b in $a.L2) {
			$ws.Cells.Item($r, 2).Value2 = [string]$b.Name
			$ws.Range("A${r}:M${r}").Interior.Color = 0xF7D9BD   # light blue
			$ws.Range("A${r}:M${r}").Font.Italic = $true
			$r++
			for ($k = 0; $k -lt $b.D; $k++) {
				$n++
				$ws.Cells.Item($r, 3).Value2 = "Seg " + (1 + $n % 2)
				$ws.Cells.Item($r, 4).Value2 = "B" + (100 + $n)
				$ws.Cells.Item($r, 5).Value2 = "8." + $n
				$ws.Cells.Item($r, 6).Value2 = "DS-" + $n
				$ws.Cells.Item($r, 7).Value2 = "Sch H"
				$ws.Cells.Item($r, 8).Value2 = "Sch I"
				$ws.Cells.Item($r, 11).Value2 = [double](1000 * $n + 0.5)
				$r++
			}
		}
	}
	$lastData = $r - 1
	$ws.Range("K7:K$lastData").NumberFormat = '#,##0.00'
	$ws.Range("A7:M$lastData").Borders.LineStyle = 1
	# Blank capacity rows, then fixed content below the data
	$notes = $lastData + 4
	$ws.Cells.Item($notes, 1).Value2 = "Notes:"
	$ws.Cells.Item($notes, 1).Font.Bold = $true
	$ws.Cells.Item($notes + 1, 1).Value2 = "1. Amounts are unaudited."
	$ws.Cells.Item($notes + 2, 1).Value2 = "2. See Part VIII instructions."
	$png = Join-Path $env:TEMP "issue1-logo.png"
	Add-Type -AssemblyName System.Drawing
	$bmp = New-Object System.Drawing.Bitmap 160, 48
	$g = [System.Drawing.Graphics]::FromImage($bmp)
	$g.Clear([System.Drawing.Color]::FromArgb(16, 124, 65))
	$g.DrawString("SIGNATURE", (New-Object System.Drawing.Font "Arial", 14), [System.Drawing.Brushes]::White, 12, 12)
	$bmp.Save($png)
	$g.Dispose(); $bmp.Dispose()
	$cell = $ws.Cells.Item($notes + 4, 1)
	[void]$ws.Shapes.AddPicture($png, 0, 1, $cell.Left, $cell.Top, 160, 48)
	$ws.Columns.Item(1).ColumnWidth = 18
	$ws.Columns.Item(2).ColumnWidth = 20
	for ($c = 3; $c -le 13; $c++) { $ws.Columns.Item($c).ColumnWidth = 11 }

	$cc = $wb.Worksheets.Item(2)
	$cc.Name = "2026 CC"
	$cc.Range("A1").Value2 = "2026 Cost Centers"
	$cc.Range("A1").Font.Bold = $true
	$h2 = @("Cost Center", "Segment", "Base Code", "Budget", "Actual", "From Disclosure", "Variance", "Owner")
	for ($c = 0; $c -lt $h2.Count; $c++) { $cc.Cells.Item(5, $c + 1).Value2 = $h2[$c] }
	$cc.Range("A5:H5").Font.Bold = $true
	for ($i = 0; $i -lt 4; $i++) {
		$row = 6 + $i
		$cc.Cells.Item($row, 1).Value2 = "CC-" + (200 + $i)
		$cc.Cells.Item($row, 2).Value2 = "Seg " + (1 + $i % 2)
		$cc.Cells.Item($row, 3).Value2 = "B" + (101 + $i)
		$cc.Cells.Item($row, 4).Value2 = [double](5000 + 100 * $i)
		$cc.Cells.Item($row, 5).Value2 = [double](4800 + 150 * $i)
		$cc.Cells.Item($row, 6).Formula = "=SUMIF('Disclosure Table'!`$D:`$D,C$row,'Disclosure Table'!`$K:`$K)"
		$cc.Cells.Item($row, 7).Formula = "=E$row-D$row"
		$cc.Cells.Item($row, 8).Value2 = "Owner " + $i
	}
	$cc.Cells.Item(11, 1).Value2 = "Disclosure total"
	$cc.Cells.Item(11, 6).Formula = "=SUM('Disclosure Table'!K7:K$lastData)"
	$cc.Range("D6:G11").NumberFormat = '#,##0.00'

	$dir = Join-Path $root "corpus\templates\issue-1-disclosure"
	New-Item -ItemType Directory -Force $dir | Out-Null
	$path = Join-Path $dir "template.xlsx"
	if (Test-Path $path) { Remove-Item $path -Force }
	$wb.Worksheets.Item(1).Activate()
	$wb.SaveAs($path, 51)
	$wb.Close($false)
	Write-Host "saved $path (data rows 7-$lastData)"
}
finally {
	$excel.Quit()
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
	[GC]::Collect()
	[GC]::WaitForPendingFinalizers()
}
