<#
.SYNOPSIS
    Builds the excel-graft battle-test template corpus from scratch by
    driving real Excel through COM.

.DESCRIPTION
    Each template is created in a fresh workbook, styled, and saved as
    corpus/templates/<name>/template.xlsx (xlOpenXMLWorkbook). A README.md is
    written next to it describing its regions, scalar targets and the
    features that are hostile to a template engine. If an individual COM
    feature fails on this Excel build, that feature is skipped, the failure
    is recorded in the template's README ("Build log"), and the template is
    still saved.

    Re-runnable: existing template.xlsx / README.md files are overwritten.

    NOTE: this file must stay pure ASCII (Windows PowerShell 5.1 reads BOM-less
    scripts in the ANSI code page).

.PARAMETER Only
    Build only the named template(s), e.g. -Only indirect-rates,date1904

.PARAMETER OutDir
    Root folder for the templates. Defaults to <repo>/corpus/templates.

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File tools/corpus-builder/Build-Corpus.ps1
    powershell -NoProfile -ExecutionPolicy Bypass -File tools/corpus-builder/Build-Corpus.ps1 -Only indirect-rates
#>

[CmdletBinding()]
param(
	[string[]]$Only,
	[string]$OutDir
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Off

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrEmpty($OutDir)) {
	$OutDir = [System.IO.Path]::GetFullPath((Join-Path $scriptDir "..\..\corpus\templates"))
}

$M = [Type]::Missing
$ThemeDir = "C:\Program Files\Microsoft Office\root\Document Themes 16"

# ---- Excel constants ---------------------------------------------------------
$xlCenter = -4108; $xlLeft = -4131; $xlRight = -4152; $xlTop = -4160
$xlSrcRange = 1; $xlYes = 1
$xlTotalsNone = 0; $xlTotalsSum = 1; $xlTotalsAverage = 2; $xlTotalsCount = 3
$xlContinuous = 1; $xlDouble = -4119; $xlThin = 2; $xlMedium = -4138
$xlEdgeTop = 8; $xlEdgeBottom = 9; $xlEdgeLeft = 7; $xlEdgeRight = 10; $xlInsideH = 12; $xlInsideV = 11
$xlThemeColorDark1 = 1; $xlThemeColorLight1 = 2; $xlThemeColorDark2 = 3; $xlThemeColorLight2 = 4
$xlThemeColorAccent1 = 5; $xlThemeColorAccent2 = 6; $xlThemeColorAccent3 = 7; $xlThemeColorAccent4 = 8
$xlThemeColorAccent5 = 9; $xlThemeColorAccent6 = 10
$xlThemeFontMajor = 1

function RGB([int]$r, [int]$g, [int]$b) { return $r + ($g * 256) + ($b * 65536) }

# Accounting / currency formats used across templates.
$FmtAcct = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)'
$FmtAcct0 = '_($* #,##0_);_($* (#,##0);_($* "-"??_);_(@_)'
$FmtCur = '$#,##0.00;[Red]($#,##0.00)'
$FmtInt = '#,##0'
$FmtPct = '0.0%'
$FmtDate = 'mm/dd/yyyy'

# ---- build-log helpers -------------------------------------------------------
$script:BuildLog = New-Object System.Collections.Generic.List[string]

function Invoke-Feature([string]$Name, [scriptblock]$Body) {
	try {
		. $Body
		$script:BuildLog.Add("OK: $Name")
	}
	catch {
		$msg = $_.Exception.Message -replace "\s+", " "
		$script:BuildLog.Add("SKIPPED: $Name -- $msg")
		Write-Warning "  [$Name] skipped: $msg"
	}
}

function Write-Utf8([string]$Path, [string]$Text) {
	$enc = New-Object System.Text.UTF8Encoding($false)
	[System.IO.File]::WriteAllText($Path, ($Text -replace "`r`n", "`n"), $enc)
}

# ---- workbook helpers --------------------------------------------------------
function New-Workbook([string]$Theme) {
	$wb = $excel.Workbooks.Add(-4167) # xlWBATWorksheet: exactly one sheet
	if ($Theme) {
		$thmx = Join-Path $ThemeDir "$Theme.thmx"
		Invoke-Feature "theme '$Theme'" { $wb.ApplyTheme($thmx) }
	}
	return $wb
}

function Add-Sheet($wb, [string]$Name) {
	$ws = $wb.Worksheets.Add($M, $wb.Worksheets.Item($wb.Worksheets.Count))
	$ws.Name = $Name
	return $ws
}

# PowerShell's COM binder caches the argument type of a property-set call
# site, so `$cell.Value2 = $v` fails once a site has seen a string and then
# gets a number. InvokeMember avoids the cached binder.
function Set-ComProp($Obj, [string]$Prop, $Value) {
	[void]$Obj.GetType().InvokeMember($Prop, [System.Reflection.BindingFlags]::SetProperty, $null, $Obj, @($Value))
}

# Write a 2-D block starting at (row, col). Strings starting with "=" are
# formulas; [datetime] values become date serials (1900 or 1904 system).
function Set-Grid($ws, [int]$Row, [int]$Col, $Rows, [switch]$Date1904) {
	for ($i = 0; $i -lt $Rows.Count; $i++) {
		$r = $Rows[$i]
		for ($j = 0; $j -lt $r.Count; $j++) {
			$v = $r[$j]
			if ($null -eq $v) { continue }
			$cell = $ws.Cells.Item($Row + $i, $Col + $j)
			if ($v -is [datetime]) {
				$serial = $v.ToOADate()
				if ($Date1904) { $serial -= 1462 }
				Set-ComProp $cell "Value2" $serial
			}
			elseif ($v -is [string] -and $v.StartsWith("=")) {
				$cell.Formula = $v
			}
			else {
				Set-ComProp $cell "Value2" $v
			}
		}
	}
}

function Set-Widths($ws, [hashtable]$Widths) {
	foreach ($k in $Widths.Keys) { $ws.Columns.Item($k).ColumnWidth = $Widths[$k] }
}

function Set-Title($ws, [string]$Address, [string]$Text, [int]$Size = 18) {
	$r = $ws.Range($Address)
	$r.Value2 = $Text
	$r.Font.ThemeFont = $xlThemeFontMajor
	$r.Font.Size = $Size
	$r.Font.Bold = $true
	$r.Font.ThemeColor = $xlThemeColorAccent1
	$r.Font.TintAndShade = -0.25
}

function Set-Band($rng, [int]$ThemeColor, [double]$Tint, [switch]$Bold) {
	$rng.Interior.ThemeColor = $ThemeColor
	$rng.Interior.TintAndShade = $Tint
	if ($Bold) { $rng.Font.Bold = $true }
}

function Add-Table($ws, [string]$Address, [string]$Name, [string]$Style) {
	$lo = $ws.ListObjects.Add($xlSrcRange, $ws.Range($Address), $M, $xlYes)
	$lo.Name = $Name
	if ($Style) { $lo.TableStyle = $Style }
	return $lo
}

function Get-TotalCell($lo, [string]$Column) {
	return $lo.TotalsRowRange.Cells.Item(1, $lo.ListColumns.Item($Column).Index)
}

function Set-Freeze($ws, [int]$Rows, [int]$Cols) {
	$ws.Activate()
	$w = $excel.ActiveWindow
	$w.FreezePanes = $false
	$w.ScrollRow = 1
	$w.ScrollColumn = 1
	$w.SplitRow = $Rows
	$w.SplitColumn = $Cols
	$w.FreezePanes = $true
}

function Hide-Gridlines($ws) {
	$ws.Activate()
	$excel.ActiveWindow.DisplayGridlines = $false
}

function Set-PageSetup($ws, [string]$PrintArea, [string]$TitleRows, [bool]$Landscape) {
	$excel.PrintCommunication = $false
	try {
		$ps = $ws.PageSetup
		if ($PrintArea) { $ps.PrintArea = $PrintArea }
		if ($TitleRows) { $ps.PrintTitleRows = $TitleRows }
		if ($Landscape) { $ps.Orientation = 2 }
		$ps.Zoom = $false
		$ps.FitToPagesWide = 1
		$ps.FitToPagesTall = $false
		$ps.CenterFooter = "Page &P of &N"
	}
	finally {
		$excel.PrintCommunication = $true
	}
}

function Save-Template($wb, [string]$Name, [string]$Readme) {
	# Tidy: A1 selected on every sheet, first visible sheet active.
	foreach ($ws in $wb.Worksheets) {
		if ($ws.Visible -eq -1) {
			try { $ws.Activate(); $ws.Range("A1").Select() | Out-Null } catch { }
		}
	}
	foreach ($ws in $wb.Worksheets) { if ($ws.Visible -eq -1) { $ws.Activate(); break } }
	$excel.Calculate()

	$dir = Join-Path $OutDir $Name
	New-Item -ItemType Directory -Force -Path $dir | Out-Null
	$path = Join-Path $dir "template.xlsx"
	if (Test-Path $path) { Remove-Item $path -Force }
	$wb.SaveAs($path, 51)
	$wb.Close($false)

	$log = ($script:BuildLog | ForEach-Object { "- $_" }) -join "`n"
	$text = $Readme.TrimEnd() + "`n`n## Build log`n`nGenerated by ``tools/corpus-builder/Build-Corpus.ps1`` on Excel $($excel.Version) (build $($excel.Build)), $(Get-Date -Format 'yyyy-MM-dd').`n`n$log`n"
	Write-Utf8 (Join-Path $dir "README.md") $text
	Write-Host "  saved $path" -ForegroundColor Green
}

function New-LogoPng([string]$Path, [string]$Text, [int]$W = 200, [int]$H = 60) {
	Add-Type -AssemblyName System.Drawing
	$bmp = New-Object System.Drawing.Bitmap($W, $H)
	$g = [System.Drawing.Graphics]::FromImage($bmp)
	try {
		$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
		$g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
		$g.Clear([System.Drawing.Color]::Transparent)
		$navy = [System.Drawing.Color]::FromArgb(255, 31, 56, 100)
		$gold = [System.Drawing.Color]::FromArgb(255, 237, 177, 32)
		$g.FillEllipse((New-Object System.Drawing.SolidBrush($gold)), 4, 6, $H - 12, $H - 12)
		$g.FillEllipse((New-Object System.Drawing.SolidBrush($navy)), 14, 16, $H - 32, $H - 32)
		$font = New-Object System.Drawing.Font("Segoe UI Semibold", 20, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
		$g.DrawString($Text, $font, (New-Object System.Drawing.SolidBrush($navy)), [single]($H - 2), [single](($H - 28) / 2))
		$bmp.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
	}
	finally {
		$g.Dispose(); $bmp.Dispose()
	}
}

# =============================================================================
# 1. flat-table-totals
# =============================================================================
function Build-FlatTableTotals {
	$wb = New-Workbook "Office Theme"
	$ws = $wb.Worksheets.Item(1)
	$ws.Name = "Invoice"

	Set-Title $ws "A1" "INVOICE" 26
	Set-Grid $ws 2 1 @(
		@("Bill To:", "{{customerName}}", $null, $null, "Invoice #", "INV-2025-0412"),
		@("Invoice Date:", "{{invoiceDate}}", $null, $null, "Terms", "Net 30")
	)
	$ws.Range("A2:A3").Font.Bold = $true
	$ws.Range("E2:E3").Font.Bold = $true
	$ws.Range("E2:E3").HorizontalAlignment = $xlRight
	$ws.Range("B3").NumberFormat = 'mmmm d, yyyy'   # placeholder sits in a date-formatted cell
	$ws.Range("B2:B3").Font.Color = (RGB 89 89 89)
	$ws.Range("A4:F4").Borders.Item($xlEdgeBottom).LineStyle = $xlContinuous
	$ws.Range("A4:F4").Borders.Item($xlEdgeBottom).ThemeColor = $xlThemeColorAccent1

	Set-Grid $ws 5 1 @(
		@("SKU", "Description", "Status", "Qty", "UnitPrice", "Amount"),
		@("WDG-1001", "Widget, standard (blue)", "Open", 120, 4.25, $null),
		@("WDG-1002", "Widget, deluxe (brushed steel)", "Paid", 36, 12.5, $null),
		@("BRK-2210", "Mounting bracket, 90 degree", "Open", 240, 1.15, $null),
		@("SRV-0100", "On-site installation (hours)", "Void", 6, 95, $null),
		@("SHP-0001", "Freight & handling", "Paid", 1, 48.9, $null)
	)
	$lo = Add-Table $ws "A5:F10" "Items" "TableStyleMedium2"
	$lo.ListColumns.Item("Amount").DataBodyRange.Formula = "=[@Qty]*[@UnitPrice]"
	$lo.ShowTotals = $true
	$lo.ListColumns.Item("Qty").TotalsCalculation = $xlTotalsSum
	$lo.ListColumns.Item("Amount").TotalsCalculation = $xlTotalsSum
	$lo.ListColumns.Item("Qty").Range.NumberFormat = $FmtInt
	$lo.ListColumns.Item("UnitPrice").Range.NumberFormat = $FmtCur
	$lo.ListColumns.Item("Amount").Range.NumberFormat = $FmtCur
	$lo.ListColumns.Item("Status").DataBodyRange.HorizontalAlignment = $xlCenter

	# Totals row is row 11. Tax + grand total below the table.
	Set-Grid $ws 13 5 @(
		@("Tax (8%)", "=Items[[#Totals],[Amount]]*0.08"),
		@("Grand Total", "=F13+F11")
	)
	$ws.Range("E13:E14").HorizontalAlignment = $xlRight
	$ws.Range("E13:F14").Font.Bold = $true
	$ws.Range("F13:F14").NumberFormat = $FmtCur
	$ws.Range("E14:F14").Borders.Item($xlEdgeTop).LineStyle = $xlContinuous
	$ws.Range("E14:F14").Borders.Item($xlEdgeBottom).LineStyle = $xlDouble
	Set-Band $ws.Range("E14:F14") $xlThemeColorAccent1 0.8
	$ws.Range("A16").Value2 = "Thank you for your business. Please remit payment within 30 days."
	$ws.Range("A16").Font.Italic = $true
	$ws.Range("A16").Font.Color = (RGB 118 118 118)

	Set-Widths $ws @{ "A" = 14; "B" = 34; "C" = 10; "D" = 9; "E" = 14; "F" = 15 }

	Invoke-Feature "data bars on Amount (x14 extLst)" {
		$db = $ws.Range("F6:F10").FormatConditions.AddDatabar()
		$db.BarColor.Color = (RGB 99 142 198)
	}
	Invoke-Feature "formula CF rule Qty>100" {
		$ws.Activate(); $ws.Range("D6").Select() | Out-Null
		$fc = $ws.Range("D6:D10").FormatConditions.Add(2, $M, '=$D6>100')
		$fc.Interior.Color = (RGB 255 235 156)
		$fc.Font.Color = (RGB 156 87 0)
		$fc.Font.Bold = $true
	}
	Invoke-Feature "data validation list on Status" {
		$v = $ws.Range("C6:C10").Validation
		$v.Delete()
		$v.Add(3, 1, 1, "Open,Paid,Void")
		$v.InCellDropdown = $true
	}
	Invoke-Feature "frozen panes below header" { Set-Freeze $ws 5 0 }
	Invoke-Feature "hide gridlines" { Hide-Gridlines $ws }
	Invoke-Feature "print area + print titles" { Set-PageSetup $ws '$A$1:$F$16' '$5:$5' $false }

	Save-Template $wb "flat-table-totals" @'
# flat-table-totals

An invoice: one Excel Table with a calculated column and a totals row, with
formulas below the table that depend on the totals row.

## Regions

| Key | Kind | Sheet | Header | Sample rows | Row roles |
|---|---|---|---|---|---|
| `Items` | Table (`A5:F11`, TableStyleMedium2) | Invoice | row 5 | 6-10 | all detail; row 11 is the table's totals row |

Columns: SKU, Description, Status, Qty, UnitPrice, Amount. `Amount` is a
calculated column `=[@Qty]*[@UnitPrice]`. Totals row: `SUBTOTAL(109,...)` on
Qty and Amount, label "Total" in SKU.

## Scalar targets

- `B2` = `{{customerName}}` (cell text placeholder)
- `B3` = `{{invoiceDate}}` (cell text placeholder, in a cell already
  formatted `mmmm d, yyyy`, so a date written there should be a serial)

## Features

- Theme: Office Theme; TableStyleMedium2; gridlines hidden
- Number formats: `#,##0` on Qty, `$#,##0.00;[Red](...)` on UnitPrice and Amount
- Conditional formatting: data bar on `F6:F10` (Excel writes this with an
  x14 `extLst` twin in the sheet), and an expression rule `=$D6>100` on `D6:D10`
- Data validation: list `Open,Paid,Void` on `C6:C10`
- Frozen panes below row 5 (the header)
- Print area `$A$1:$F$16`, print titles `$5:$5`, fit to 1 page wide
- `F13` Tax: `=Items[[#Totals],[Amount]]*0.08` (structured ref to totals)
- `F14` Grand Total: `=F13+F11` (A1 ref to the totals row cell)

## Hostile for a template engine

- `F14` references the totals row by A1 (`F11`). When Items grows, the totals
  row moves and this must become e.g. `F13+F21`. (v1 constraint says content
  outside a region should use structured refs; this one deliberately does
  not, so it is a Phase 4 fixture.)
- CF sqref and DV sqref cover exactly the sample rows `6:10` and must grow.
- The data-bar rule exists twice: once in `<conditionalFormatting>` and once in
  `<extLst><x14:conditionalFormattings>` linked by `x14:id`. Both sqrefs must
  be updated or Excel will repair.
- Print area and the `_xlnm._FilterDatabase`-style names live in `workbook.xml`.
- Placeholders live in `sharedStrings.xml`; replacing them must not rewrite
  the shared strings table (write inline strings instead).
- The frozen pane (`<pane ySplit="5" topLeftCell="A6">`) sits above the region
  and must not move.
'@
}

# =============================================================================
# 2. named-range-region
# =============================================================================
function Build-NamedRangeRegion {
	$wb = New-Workbook "Facet"
	$ws = $wb.Worksheets.Item(1)
	$ws.Name = "Staff"

	$ws.Range("A1:F1").Merge()
	Set-Title $ws "A1" "Northwind Traders - Staff Roster" 20
	$ws.Range("A1").HorizontalAlignment = $xlRight
	$ws.Range("A1").VerticalAlignment = $xlCenter
	$ws.Range("A1").IndentLevel = 1
	$ws.Rows.Item(1).RowHeight = 48
	$ws.Range("A2").Value2 = "Compensation review, FY2025 (confidential)"
	$ws.Range("A2").Font.Italic = $true
	$ws.Range("A2").Font.Color = (RGB 118 118 118)

	Set-Grid $ws 4 1 @(
		@("Emp ID", "Name", "Department", "Hire Date", "Salary", "Bonus"),
		@("E-1041", "Alice Nakamura", "Engineering", [datetime]"2019-03-11", 128500, 12850),
		@("E-1077", "Brian O'Neill", "Finance", [datetime]"2021-07-19", 96400, 7230),
		@("E-1102", "Carmen Diaz", "Operations", [datetime]"2016-01-04", 88250, 5295),
		@("E-1156", "Dmitri Volkov", "Engineering", [datetime]"2023-10-02", 112000, 8400),
		@("E-1190", "Esther Mensah", "Sales", [datetime]"2020-05-26", 79900, 15980)
	)
	$hdr = $ws.Range("A4:F4")
	Set-Band $hdr $xlThemeColorAccent1 0 -Bold
	$hdr.Font.ThemeColor = $xlThemeColorDark1
	$hdr.HorizontalAlignment = $xlCenter
	$hdr.RowHeight = 22
	$hdr.VerticalAlignment = $xlCenter
	foreach ($r in 5, 7, 9) { Set-Band $ws.Range("A${r}:F${r}") $xlThemeColorAccent1 0.8 }
	$ws.Range("A5:F9").Borders.Item($xlInsideH).LineStyle = $xlContinuous
	$ws.Range("A5:F9").Borders.Item($xlInsideH).Color = (RGB 217 217 217)
	$ws.Range("D5:D9").NumberFormat = $FmtDate
	$ws.Range("D5:D9").HorizontalAlignment = $xlCenter
	$ws.Range("E5:F10").NumberFormat = $FmtAcct0

	Set-Grid $ws 10 1 @(, @("Total", $null, $null, "=COUNTA(A5:A9)&"" staff""", "=SUM(E5:E9)", "=SUM(F5:F9)"))
	$tot = $ws.Range("A10:F10")
	$tot.Font.Bold = $true
	$tot.Borders.Item($xlEdgeTop).LineStyle = $xlContinuous
	$tot.Borders.Item($xlEdgeTop).Weight = $xlMedium
	$tot.Borders.Item($xlEdgeBottom).LineStyle = $xlDouble
	$ws.Range("D10").HorizontalAlignment = $xlCenter

	$wb.Names.Add("Employees", '=Staff!$A$5:$F$9') | Out-Null

	Set-Grid $ws 12 1 @(
		@("Reviewed by HR - see note"),
		@("HR compensation policy")
	)
	Set-Widths $ws @{ "A" = 12; "B" = 22; "C" = 16; "D" = 13; "E" = 14; "F" = 12; "G" = 3 }

	Invoke-Feature "legacy note below region (A12)" {
		$c = $ws.Range("A12").AddComment("HR: figures reconciled to payroll register on 2025-12-31. -- K. Patel")
		$c.Shape.Width = 220; $c.Shape.Height = 60
	}
	Invoke-Feature "hyperlink below region (A13)" {
		$ws.Hyperlinks.Add($ws.Range("A13"), "https://example.com/hr/compensation-policy", $M, "Open the HR compensation policy", "HR compensation policy") | Out-Null
	}
	Invoke-Feature "logo picture above region" {
		$png = Join-Path $env:TEMP "excel-graft-corpus-logo.png"
		New-LogoPng $png "NORTHWIND"
		$pic = $ws.Shapes.AddPicture($png, 0, -1, 4, 6, 120, 36)
		$pic.Name = "Logo"
		$pic.Placement = 2 # xlMove
	}
	Invoke-Feature "shape below region (anchor must shift)" {
		$top = $ws.Range("B15").Top
		$shp = $ws.Shapes.AddShape(5, $ws.Range("B15").Left, $top, 190, 44)
		$shp.Name = "ApprovedStamp"
		$shp.Fill.ForeColor.RGB = (RGB 226 239 218)
		$shp.Line.ForeColor.RGB = (RGB 84 130 53)
		$shp.Line.Weight = 1.5
		$tr = $shp.TextFrame2.TextRange
		$tr.Text = "APPROVED - Compensation Committee"
		$tr.Font.Size = 10
		$tr.Font.Bold = -1
		$tr.Font.Fill.ForeColor.RGB = (RGB 56 87 35)
		$shp.TextFrame2.VerticalAnchor = 3
		$tr.ParagraphFormat.Alignment = 2
	}
	Invoke-Feature "clustered column chart over sample rows" {
		$ws.Range("H20").Select() | Out-Null
		$left = $ws.Range("H4").Left; $top = $ws.Range("H4").Top
		$co = $ws.Shapes.AddChart2(201, 51, $left, $top, 420, 240)
		$co.Name = "SalaryChart"
		$ch = $co.Chart
		while ($ch.SeriesCollection().Count -gt 0) { $ch.SeriesCollection(1).Delete() }
		$s1 = $ch.SeriesCollection().NewSeries()
		$s1.Name = '=Staff!$E$4'
		$s1.Values = '=Staff!$E$5:$E$9'
		$s1.XValues = '=Staff!$B$5:$B$9'
		$s2 = $ch.SeriesCollection().NewSeries()
		$s2.Name = '=Staff!$F$4'
		$s2.Values = '=Staff!$F$5:$F$9'
		$s2.XValues = '=Staff!$B$5:$B$9'
		$ch.HasTitle = $true
		$ch.ChartTitle.Text = "Salary and bonus by employee"
	}
	Invoke-Feature "hide gridlines" { Hide-Gridlines $ws }
	Invoke-Feature "page setup (landscape, fit 1 page wide)" { Set-PageSetup $ws $null $null $true }

	Save-Template $wb "named-range-region" @'
# named-range-region

A region defined by a workbook-scoped defined name instead of a Table, with
A1-formula totals, a chart, pictures/shapes, a note and a hyperlink around it.

## Regions

| Key | Kind | Sheet | Header | Sample rows | Row roles |
|---|---|---|---|---|---|
| `Employees` | defined name `=Staff!$A$5:$F$9` | Staff | row 4 (styled, not a Table) | 5-9 | all detail |

Columns: Emp ID, Name, Department, Hire Date, Salary, Bonus. Rows 5, 7, 9
carry a manual alternating fill (accent1 tint 0.8); rows 6, 8 have none. The
total row 10 is OUTSIDE the name: `D10 =COUNTA(A5:A9)&" staff"`,
`E10 =SUM(E5:E9)`, `F10 =SUM(F5:F9)`.

## Scalar targets

None.

## Features

- Theme: Facet. Merged, centered title `A1:F1` (row height 48)
- Styled header row (accent1 fill, white bold), hairline inner borders
- Date format on Hire Date, accounting (no decimals) on Salary/Bonus
- Picture `Logo` (PNG generated with System.Drawing) anchored in rows 1-2
  ABOVE the region, placement "move but don't size"
- Clustered column chart `SalaryChart` at `H4`, two series whose values and
  categories are `Staff!$E$5:$E$9`, `$F$5:$F$9`, `$B$5:$B$9`
- Legacy note (comment + VML drawing) on `A12`, below the region
- External hyperlink on `A13`, below the region
- Rounded-rectangle shape `ApprovedStamp` at `B15`, below the region

## Hostile for a template engine

- The region is a defined name, not a Table: no `tableN.xml` tells you the
  header; the header is "the row above the name" and is styled manually.
- Alternating fills: the sample rows have two different styles (odd/even).
  The engine must decide whether to cycle the pattern or copy one row.
- Total row sits directly below the region with plain A1 `SUM(E5:E9)`; it
  must move down and its range must grow.
- Chart series ranges (`xl/charts/chart1.xml` `<c:f>`) point at exactly the
  sample rows and must grow; the chart's cached `<c:numCache>`/`<c:strCache>`
  values are stale afterwards.
- The chart anchor starts at row 4 (inside the header row) and ends below
  row 9, so row insertion stretches or moves it depending on `editAs`.
- The logo is above the region and must NOT move; the shape, note (VML
  anchor + `<comment ref>`), and hyperlink (`<hyperlink ref>`) are below and
  MUST move.
- Merged title cell `A1:F1` is above the region (must not shift). The title
  is right-aligned so it does not collide with the logo.
'@
}

# =============================================================================
# 3. multi-sheet-summary
# =============================================================================
function Build-MultiSheetSummary {
	$wb = New-Workbook "Integral"
	$wsSum = $wb.Worksheets.Item(1)
	$wsSum.Name = "Summary"
	$wsData = Add-Sheet $wb "Data"
	$wsPiv = Add-Sheet $wb "Pivot"
	$wsLk = Add-Sheet $wb "Lookups"

	# Lookups (hidden) feeding a DV list.
	Set-Grid $wsLk 1 1 @(@("Regions"), @("North"), @("South"), @("East"), @("West"))
	$wsLk.Range("A1").Font.Bold = $true
	$wb.Names.Add("RegionList", '=Lookups!$A$2:$A$5') | Out-Null

	# Data: Table Sales at A1.
	Set-Grid $wsData 1 1 @(
		@("Region", "Rep", "Month", "Revenue"),
		@("North", "Hannah Brooks", [datetime]"2025-01-01", 48250.00),
		@("South", "Luis Ortega", [datetime]"2025-01-01", 39110.50),
		@("East", "Mei Lin", [datetime]"2025-02-01", 52780.25),
		@("West", "Sam Carter", [datetime]"2025-02-01", 31640.00),
		@("North", "Hannah Brooks", [datetime]"2025-03-01", 50115.75)
	)
	$lo = Add-Table $wsData "A1:D6" "Sales" "TableStyleMedium9"
	$lo.ListColumns.Item("Month").DataBodyRange.NumberFormat = 'mmm yyyy'
	$lo.ListColumns.Item("Revenue").Range.NumberFormat = $FmtAcct
	Set-Widths $wsData @{ "A" = 12; "B" = 18; "C" = 12; "D" = 16 }
	Invoke-Feature "DV list on Sales[Region] from hidden sheet" {
		$v = $lo.ListColumns.Item("Region").DataBodyRange.Validation
		$v.Delete()
		$v.Add(3, 1, 1, "=RegionList")
	}
	Invoke-Feature "frozen header on Data" { Set-Freeze $wsData 1 0 }

	# Summary
	Set-Title $wsSum "A1" "Regional Sales Summary" 20
	$wsSum.Range("A2").Formula = '="Months covered: "&TEXT(MIN(Sales[Month]),"mmm yyyy")&" - "&TEXT(MAX(Sales[Month]),"mmm yyyy")'
	$wsSum.Range("A2").Font.Color = (RGB 118 118 118)
	Set-Grid $wsSum 4 1 @(
		@("Region", "Revenue", "Share"),
		@("North"), @("South"), @("East"), @("West"),
		@("Total", "=SUM(B5:B8)", "=SUM(C5:C8)")
	)
	$wsSum.Range("B5:B8").Formula = '=SUMIFS(Sales[Revenue],Sales[Region],$A5)'
	$wsSum.Range("C5:C8").Formula = '=IF($B$9=0,0,B5/$B$9)'
	Set-Band $wsSum.Range("A4:C4") $xlThemeColorAccent1 0 -Bold
	$wsSum.Range("A4:C4").Font.ThemeColor = $xlThemeColorDark1
	$wsSum.Range("A9:C9").Font.Bold = $true
	$wsSum.Range("A9:C9").Borders.Item($xlEdgeTop).LineStyle = $xlContinuous
	$wsSum.Range("A9:C9").Borders.Item($xlEdgeBottom).LineStyle = $xlDouble
	$wsSum.Range("B5:B9").NumberFormat = $FmtAcct
	$wsSum.Range("C5:C9").NumberFormat = $FmtPct
	Set-Grid $wsSum 11 1 @(
		@("Cross-sheet checks", $null),
		@("Second sample revenue (=Data!D3)", "=Data!D3"),
		@("All sample revenue (=SUM(Data!D2:D6))", "=SUM(Data!D2:D6)"),
		@("Structured total (=SUM(Sales[Revenue]))", "=SUM(Sales[Revenue])"),
		@("Row count (=ROWS(Sales))", "=ROWS(Sales)")
	)
	$wsSum.Range("A11").Font.Bold = $true
	$wsSum.Range("A11:B11").Borders.Item($xlEdgeBottom).LineStyle = $xlContinuous
	$wsSum.Range("B12:B14").NumberFormat = $FmtAcct
	Set-Widths $wsSum @{ "A" = 40; "B" = 16; "C" = 10 }
	Invoke-Feature "hide gridlines (Summary)" { Hide-Gridlines $wsSum }

	Invoke-Feature "PivotTable on Pivot sheet sourced from Sales" {
		$pc = $wb.PivotCaches().Create(1, "Sales", 6)
		$pt = $pc.CreatePivotTable($wsPiv.Range("A3"), "SalesPivot", $true, 6)
		$f = $pt.PivotFields("Region"); $f.Orientation = 1
		$df = $pt.AddDataField($pt.PivotFields("Revenue"), "Total Revenue", -4157)
		$df.NumberFormat = $FmtAcct
		$f2 = $pt.PivotFields("Rep"); $f2.Orientation = 1; $f2.Position = 2
		$pt.RowAxisLayout(1) # tabular
		$pt.TableStyle2 = "PivotStyleMedium9"
		$wsPiv.Range("A1").Value2 = "Revenue by region and rep (refresh after render)"
		$wsPiv.Range("A1").Font.Bold = $true
	}
	Invoke-Feature "slicer on Sales[Region]" {
		$sc = $wb.SlicerCaches.Add2($lo, "Region")
		$sl = $sc.Slicers.Add($wsData, $M, "RegionSlicer", "Region", $wsData.Range("F2").Top, $wsData.Range("F2").Left, 144, 170)
		$sl.Style = "SlicerStyleLight1"
	}
	Invoke-Feature "hide Lookups sheet" { $wsLk.Visible = 0 }

	Save-Template $wb "multi-sheet-summary" @'
# multi-sheet-summary

A data sheet with one Table feeding a summary sheet (structured and A1
cross-sheet refs), a PivotTable, a slicer, and a hidden lookup sheet.

## Regions

| Key | Kind | Sheet | Header | Sample rows | Row roles |
|---|---|---|---|---|---|
| `Sales` | Table (`A1:D6`, TableStyleMedium9, no totals row) | Data | row 1 | 2-6 | all detail |

Columns: Region, Rep, Month (real date serials, format `mmm yyyy`), Revenue.

## Scalar targets

None. (`Summary!A2` is a formula over `Sales[Month]`.)

## Features

- Theme: Integral. Sheets: Summary, Data, Pivot, Lookups (hidden)
- `Summary!B5:B8` `=SUMIFS(Sales[Revenue],Sales[Region],$A5)` (one Range
  assignment, so a shared formula), shares, total row
- `Summary!B12` `=Data!D3` -- A1 ref to ONE sample body cell
- `Summary!B13` `=SUM(Data!D2:D6)` -- A1 ref to the whole sample range
- `Summary!B14` `=SUM(Sales[Revenue])`, `B15` `=ROWS(Sales)` -- structured
- PivotTable `SalesPivot` on sheet Pivot, source `Sales` (pivotCache
  definition + records parts), Region and Rep on rows, sum of Revenue
- Slicer `RegionSlicer` on the Sales table (slicer, slicerCache, x14/x15
  extLst entries in workbook.xml and the sheet)
- Hidden sheet Lookups, defined name `RegionList = Lookups!$A$2:$A$5`, data
  validation list `=RegionList` on `Sales[Region]` body cells
- Frozen header row on Data

## Hostile for a template engine

- `=Data!D3` points at a single sample row. After render it is undefined
  what it should mean (Phase 4 / v1 constraint violation, deliberately).
- `=SUM(Data!D2:D6)` must grow to cover the rendered rows (Phase 4).
- The pivot cache holds a copy of the sample data (`pivotCacheRecords1.xml`)
  and should be flagged `refreshOnLoad="1"`, or the pivot shows sample data.
- The table-based slicer (`tableSlicerCache` in `x15` extLst) references the
  table by id; the table's id/name must not change.
- DV list sqref on `A2:A6` must grow with the table.
- Hidden sheet must stay hidden (`state="hidden"` in workbook.xml).
'@
}

# =============================================================================
# 4. indirect-rates
# =============================================================================
function Build-IndirectRates {
	$wb = New-Workbook "Office 2013 - 2022 Theme"
	$wsCov = $wb.Worksheets.Item(1)
	$wsCov.Name = "Cover"
	$wsDL = Add-Sheet $wb "Direct Labor"
	$wsIP = Add-Sheet $wb "Indirect Pools"
	$wsR = Add-Sheet $wb "Rates"

	# ---- Cover
	$wsCov.Range("A1:D1").Merge()
	Set-Title $wsCov "A1" "Indirect Cost Rates" 22
	$wsCov.Rows.Item(1).RowHeight = 36
	$wsCov.Range("A2:D2").Merge()
	$wsCov.Range("A2").Value2 = "Annual indirect cost pools, allocation bases and rates"
	$wsCov.Range("A2").Font.Italic = $true
	$wsCov.Range("A2").Font.Color = (RGB 89 89 89)
	Set-Grid $wsCov 5 1 @(
		@("Contractor", "Northwind Defense Systems, Inc."),
		@("Fiscal year end", [datetime]"2025-12-31"),
		@("Prepared by", "J. Alvarez, Controller"),
		@("Submission date", "{{submissionDate}}")
	)
	$wsCov.Range("A5:A8").Font.Bold = $true
	$wsCov.Range("B6").NumberFormat = 'mmmm d, yyyy'
	$wsCov.Range("B6").HorizontalAlignment = $xlLeft
	$inputs = $wsCov.Range("B5:B8")
	$inputs.Interior.Color = (RGB 255 242 204)
	$inputs.Borders.LineStyle = $xlContinuous
	$inputs.Borders.Color = (RGB 191 191 191)
	$wsCov.Range("A10").Value2 = "Yellow cells are inputs. All other cells are protected."
	$wsCov.Range("A10").Font.Size = 9
	$wsCov.Range("A10").Font.Color = (RGB 118 118 118)
	$wb.Names.Add("ContractorName", '=Cover!$B$5') | Out-Null
	$wb.Names.Add("FiscalYearEnd", '=Cover!$B$6') | Out-Null
	$wb.Names.Add("PreparedBy", '=Cover!$B$7') | Out-Null
	Set-Widths $wsCov @{ "A" = 20; "B" = 40; "C" = 12; "D" = 12 }
	Invoke-Feature "hide gridlines (Cover)" { Hide-Gridlines $wsCov }

	# ---- Direct Labor
	Set-Title $wsDL "A1" "Schedule: Direct Labor by Employee" 16
	$wsDL.Range("A2").Formula = '=ContractorName&" - FYE "&TEXT(FiscalYearEnd,"mm/dd/yyyy")'
	$wsDL.Range("A2").Font.Color = (RGB 89 89 89)
	Set-Grid $wsDL 4 1 @(
		@("Employee", "LaborCategory", "Project", "Hours", "Rate", "Cost"),
		@("Maria Chen", "Senior Engineer", "W56HZV-24-C-0112", 1880, 72.50, $null),
		@("David Okafor", "Systems Analyst", "W56HZV-24-C-0112", 1920, 58.25, $null),
		@("Priya Raman", "Program Manager", "FA8650-23-D-2040", 1760, 81.00, $null),
		@("Tom Lindqvist", "Test Technician", "FA8650-23-D-2040", 2010, 39.75, $null),
		@("Angela Ruiz", "Software Engineer II", "N00024-25-C-5301", 1840, 64.10, $null),
		@("Kevin Brandt", "Technical Writer", "N00024-25-C-5301", 1500, 44.00, $null)
	)
	$dl = Add-Table $wsDL "A4:F10" "DirectLabor" "TableStyleMedium2"
	$dl.ListColumns.Item("Cost").DataBodyRange.Formula = "=[@Hours]*[@Rate]"
	$dl.ShowTotals = $true
	$dl.ListColumns.Item("Hours").TotalsCalculation = $xlTotalsSum
	$dl.ListColumns.Item("Cost").TotalsCalculation = $xlTotalsSum
	$dl.ListColumns.Item("Hours").Range.NumberFormat = '#,##0.0'
	$dl.ListColumns.Item("Rate").Range.NumberFormat = $FmtAcct
	$dl.ListColumns.Item("Cost").Range.NumberFormat = $FmtAcct
	Set-Widths $wsDL @{ "A" = 18; "B" = 22; "C" = 20; "D" = 10; "E" = 12; "F" = 16 }
	Invoke-Feature "frozen header (Direct Labor)" { Set-Freeze $wsDL 4 1 }
	Invoke-Feature "print setup (Direct Labor)" { Set-PageSetup $wsDL '$A$1:$F$11' '$4:$4' $true }

	# ---- Indirect Pools (defined-name region with groups)
	Set-Title $wsIP "A1" "Schedule: Indirect Expense Pools" 16
	$wsIP.Range("A2").Formula = '=ContractorName'
	$wsIP.Range("A2").Font.Color = (RGB 89 89 89)
	Set-Grid $wsIP 4 1 @(
		@("Account", "Description", "Amount", "Share of Total"),
		@("Fringe", $null, "=SUM(C6:C8)"),
		@(5100, "Employer FICA / Medicare", 50070),
		@(5110, "Group health insurance", 118400),
		@(5120, "401(k) employer match", 39270),
		@("Overhead", $null, "=SUM(C10:C12)"),
		@(6100, "Facilities rent", 214000),
		@(6110, "Utilities", 38650),
		@(6120, "Indirect labor - engineering support", 303900),
		@("G&A", $null, "=SUM(C14:C16)"),
		@(7100, "Executive salaries", 142000),
		@(7110, "Accounting & legal", 36500),
		@(7120, "General liability insurance", 21800),
		@("Total Indirect Pools", $null, "=SUM(C5,C9,C13)")
	)
	# One Range.Formula over 13 rows -> Excel writes a shared formula.
	$wsIP.Range("D5:D17").Formula = '=C5/$C$17'
	$hdrIP = $wsIP.Range("A4:D4")
	Set-Band $hdrIP $xlThemeColorAccent1 0 -Bold
	$hdrIP.Font.ThemeColor = $xlThemeColorDark1
	foreach ($r in 5, 9, 13) {
		Set-Band $wsIP.Range("A${r}:D${r}") $xlThemeColorAccent1 0.6 -Bold
	}
	$gt = $wsIP.Range("A17:D17")
	Set-Band $gt $xlThemeColorAccent1 -0.25 -Bold
	$gt.Font.ThemeColor = $xlThemeColorDark1
	$gt.Borders.Item($xlEdgeTop).LineStyle = $xlContinuous
	$gt.Borders.Item($xlEdgeBottom).LineStyle = $xlDouble
	foreach ($r in 6, 7, 8, 10, 11, 12, 14, 15, 16) { $wsIP.Range("A$r").IndentLevel = 1 }
	$wsIP.Range("A5:A17").HorizontalAlignment = $xlLeft
	$wsIP.Range("C5:C17").NumberFormat = $FmtAcct
	$wsIP.Range("D5:D17").NumberFormat = $FmtPct
	$wb.Names.Add("IndirectPools", "='Indirect Pools'!`$A`$5:`$D`$17") | Out-Null
	Set-Widths $wsIP @{ "A" = 22; "B" = 38; "C" = 16; "D" = 14 }
	Invoke-Feature "hide gridlines (Indirect Pools)" { Hide-Gridlines $wsIP }
	Invoke-Feature "page setup (Indirect Pools)" { Set-PageSetup $wsIP $null $null $false }

	# ---- Rates
	$wb.Names.Add("FringeRate", "='Indirect Pools'!`$C`$5/'Direct Labor'!`$F`$11") | Out-Null
	$wb.Names.Add("OverheadRate", "='Indirect Pools'!`$C`$9/'Direct Labor'!`$F`$11") | Out-Null
	$wb.Names.Add("GARate", "='Indirect Pools'!`$C`$13/('Direct Labor'!`$F`$11+'Indirect Pools'!`$C`$5+'Indirect Pools'!`$C`$9)") | Out-Null

	Set-Title $wsR "A1" "Schedule: Indirect Rate Computation" 16
	$wsR.Range("A2").Formula = '="Fiscal year ended "&TEXT(FiscalYearEnd,"mmmm d, yyyy")'
	$wsR.Range("A2").Font.Color = (RGB 89 89 89)
	Set-Grid $wsR 4 1 @(
		@("Rate", "Pool Costs", "Allocation Base", "FY2025 Rate", "FY2024 Rate", "Change"),
		@("Fringe", "='Indirect Pools'!C5", "='Direct Labor'!F11", "=FringeRate", 0.312),
		@("Overhead", "='Indirect Pools'!C9", "='Direct Labor'!F11", "=OverheadRate", 0.874),
		@("G&A", "='Indirect Pools'!C13", "='Direct Labor'!F11+B5+B6", "=GARate", 0.139)
	)
	$wsR.Range("F5:F7").Formula = '=D5-E5'
	Set-Grid $wsR 9 1 @(, @("Direct labor per table (structured)", $null, "=DirectLabor[[#Totals],[Cost]]"))
	$wsR.Range("A9").Font.Italic = $true
	$hdrR = $wsR.Range("A4:F4")
	Set-Band $hdrR $xlThemeColorAccent1 0 -Bold
	$hdrR.Font.ThemeColor = $xlThemeColorDark1
	$hdrR.WrapText = $true
	$hdrR.HorizontalAlignment = $xlCenter
	$wsR.Rows.Item(4).RowHeight = 30
	$wsR.Range("A5:F7").Borders.Item($xlInsideH).LineStyle = $xlContinuous
	$wsR.Range("A5:F7").Borders.Item($xlInsideH).Color = (RGB 217 217 217)
	$wsR.Range("A7:F7").Borders.Item($xlEdgeBottom).LineStyle = $xlContinuous
	$wsR.Range("B5:C7").NumberFormat = $FmtAcct0
	$wsR.Range("C9").NumberFormat = $FmtAcct0
	$wsR.Range("D5:E7").NumberFormat = '0.00%'
	$wsR.Range("F5:F7").NumberFormat = '+0.00%;-0.00%;0.00%'
	$wsR.Range("E5:E7").Interior.Color = (RGB 255 242 204)
	Set-Widths $wsR @{ "A" = 16; "B" = 15; "C" = 17; "D" = 12; "E" = 12; "F" = 11 }
	Invoke-Feature "icon set CF on Rates!F5:F7" {
		$ic = $wsR.Range("F5:F7").FormatConditions.AddIconSetCondition()
		$ic.IconSet = $wb.IconSets.Item(1) # xl3Arrows
	}
	Invoke-Feature "bar chart on Rates" {
		$wsR.Activate(); $wsR.Range("H30").Select() | Out-Null
		$co = $wsR.Shapes.AddChart2(216, 57, $wsR.Range("A11").Left, $wsR.Range("A11").Top, 440, 220) # xlBarClustered
		$co.Name = "RatesChart"
		$ch = $co.Chart
		while ($ch.SeriesCollection().Count -gt 0) { $ch.SeriesCollection(1).Delete() }
		foreach ($col in "D", "E") {
			$s = $ch.SeriesCollection().NewSeries()
			$s.Name = "=Rates!`$$col`$4"
			$s.Values = "=Rates!`$$col`$5:`$$col`$7"
			$s.XValues = '=Rates!$A$5:$A$7'
		}
		$ch.HasTitle = $true
		$ch.ChartTitle.Text = "Indirect rates, current vs prior year"
	}
	Invoke-Feature "hide gridlines (Rates)" { Hide-Gridlines $wsR }

	Invoke-Feature "protect Cover (no password), inputs unlocked" {
		$wsCov.Range("B5:B8").Locked = $false
		$wsCov.Protect()
	}

	Save-Template $wb "indirect-rates" @'
# indirect-rates

A realistic indirect-rate workbook: scalar cover sheet, a
direct-labor Table, a grouped indirect-pool region defined by name (not a
Table), and a rate computation sheet that reaches into both via A1 refs and
named formulas.

## Regions

| Key | Kind | Sheet | Header | Sample rows | Row roles |
|---|---|---|---|---|---|
| `DirectLabor` | Table (`A4:F11`, TableStyleMedium2, totals row 11) | Direct Labor | row 4 | 5-10 | all detail |
| `IndirectPools` | defined name `='Indirect Pools'!$A$5:$D$17` | Indirect Pools | row 4 | 5-17 | see below |

`DirectLabor` columns: Employee, LaborCategory, Project, Hours, Rate, Cost
(calculated column `=[@Hours]*[@Rate]`). Totals: sum of Hours and Cost.

`IndirectPools` columns: Account, Description, Amount, Share of Total.

| Row | Role | Account | Amount formula |
|---|---|---|---|
| 5 | group header (accent1 tint 0.6, bold) | Fringe | `=SUM(C6:C8)` |
| 6-8 | detail (indented) | 5100, 5110, 5120 | values |
| 9 | group header | Overhead | `=SUM(C10:C12)` |
| 10-12 | detail | 6100, 6110, 6120 | values |
| 13 | group header | G&A | `=SUM(C14:C16)` |
| 14-16 | detail | 7100, 7110, 7120 | values |
| 17 | grand total (accent1 shade -0.25, white bold, double bottom border) | Total Indirect Pools | `=SUM(C5,C9,C13)` |

Column D on every row 5-17 is `=C5/$C$17` ... `=C17/$C$17`, written by ONE
`Range.Formula` assignment, so Excel stores it as a shared formula
(`<f t="shared" ref="D5:D17" si="0">` on D5, bare `<f t="shared" si="0"/>` on
D6:D17).

Unlike `transactions`, the group header is ABOVE its details (a subtotal
row that sums the rows below it), and the grand total sums the group rows.

## Scalar targets

- Defined names: `ContractorName` (`Cover!$B$5`), `FiscalYearEnd`
  (`Cover!$B$6`, date-formatted), `PreparedBy` (`Cover!$B$7`)
- Placeholder: `Cover!B8` = `{{submissionDate}}`
- Other sheets use the names in formulas (`=ContractorName&...`,
  `TEXT(FiscalYearEnd,...)`).

## Features

- Theme: Office 2013-2022. Merged title on Cover
- Cover sheet protected (no password); input cells `B5:B8` unlocked and
  yellow-filled. The scalar targets sit on a protected sheet.
- Named formulas (not cell refs): `FringeRate =
  'Indirect Pools'!$C$5/'Direct Labor'!$F$11`, `OverheadRate`, `GARate`
- Rates sheet: `B5 ='Indirect Pools'!C5`, `C5 ='Direct Labor'!F11`,
  `D5 =FringeRate`, `F5:F7 =D5-E5` (shared formula), icon set (3 arrows) on
  `F5:F7`, clustered bar chart `RatesChart` over `Rates!$D$5:$E$7`
- `Rates!C9` `=DirectLabor[[#Totals],[Cost]]` (structured, for contrast)
- Frozen panes, print area and print titles on Direct Labor (landscape)

## Hostile for a template engine

- The grouped region is a defined name, and its group rows are recognised
  only by their formulas and fills. Group subtotals, the grand total and the
  shared formula in D all need regenerating for any number of groups.
- The shared formula `D5:D17` has its master in the first sample row. If the
  engine deletes row 5 as a "sample" and keeps a D6 that says
  `<f t="shared" si="0"/>`, the file is corrupt. Shared formulas must be
  expanded or re-mastered.
- Cross-sheet A1 refs into both regions: `'Indirect Pools'!C5/C9/C13` (group
  rows, whose positions change) and `'Direct Labor'!F11` (the totals row,
  which moves as the table grows). Same refs are also inside defined-name
  formulas in `workbook.xml` (`FringeRate` etc.), which must be shifted too.
- Rates `C7 ='Direct Labor'!F11+B5+B6` mixes cross-sheet and local refs.
- Protected Cover sheet: the engine writes into unlocked cells of a
  `<sheetProtection>` sheet; nothing about protection should change.
- Sheet names with spaces require quoting in every formula the engine emits.
'@
}

# =============================================================================
# 5. wide-and-styled
# =============================================================================
function Build-WideAndStyled {
	$wb = New-Workbook "Retrospect"
	$ws = $wb.Worksheets.Item(1)
	$ws.Name = "GL"

	$months = "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
	$headers = @("Account No", "Account Name", "Department", "Cost Center", "Owner", "Region", "Currency") + $months +
		@("Q1", "Q2", "Q3", "Q4", "FY Total", "Budget", "Variance", "Variance %", "Trend", "Status", "Last Reviewed", "Notes")
	# 7 + 12 + 12 = 31 columns: A..AE
	Set-Title $ws "A1" "General Ledger - Operating Expense Detail, FY2025" 18
	$ws.Range("A2").Value2 = "All amounts in USD. Months are grouped (click + above the columns to expand)."
	$ws.Range("A2").Font.Color = (RGB 118 118 118)
	Set-Grid $ws 4 1 @(, $headers)

	$accts = @(
		@(61000, "Salaries & wages - admin", "Administration", "CC-100", "R. Singh", "US-East", "USD", 18000),
		@(61500, "Software subscriptions", "IT", "CC-210", "P. Novak", "US-West", "USD", 4200),
		@(62000, "Travel & entertainment", "Sales", "CC-330", "J. Fischer", "EMEA", "USD", 6500),
		@(63000, "Facilities & utilities", "Operations", "CC-400", "L. Moreau", "US-East", "USD", 9800),
		@(64000, "Professional services", "Finance", "CC-120", "A. Haddad", "US-Central", "USD", 7300)
	)
	$rnd = New-Object System.Random 42
	$row = 5
	foreach ($a in $accts) {
		$vals = @($a[0], $a[1], $a[2], $a[3], $a[4], $a[5], $a[6])
		foreach ($m in 1..12) { $vals += [math]::Round($a[7] * (0.8 + $rnd.NextDouble() * 0.45), 2) }
		$budget = [math]::Round($a[7] * 12 * 1.02 / 100) * 100
		$vals += @($null, $null, $null, $null, $null, $budget, $null, $null, $null,
			@("Open", "Reviewed", "Reviewed", "Flagged", "Open")[$row - 5],
			[datetime]"2025-12-15",
			@("Includes Q3 merit increases.", "Annual renewal of ERP and CRM licences moved from March to January; see PO 88213.", "", "Utility rebate received in August.", "Audit fees, legal retainer, and a one-off valuation engagement for the acquisition due diligence.")[$row - 5])
		Set-Grid $ws $row 1 @(, $vals)
		$row++
	}
	$lo = Add-Table $ws "A4:AE9" "Ledger" "TableStyleMedium7"
	$lo.ListColumns.Item("Q1").DataBodyRange.Formula = "=SUM(Ledger[@[Jan]:[Mar]])"
	$lo.ListColumns.Item("Q2").DataBodyRange.Formula = "=SUM(Ledger[@[Apr]:[Jun]])"
	$lo.ListColumns.Item("Q3").DataBodyRange.Formula = "=SUM(Ledger[@[Jul]:[Sep]])"
	$lo.ListColumns.Item("Q4").DataBodyRange.Formula = "=SUM(Ledger[@[Oct]:[Dec]])"
	$lo.ListColumns.Item("FY Total").DataBodyRange.Formula = "=SUM(Ledger[@[Q1]:[Q4]])"
	$lo.ListColumns.Item("Variance").DataBodyRange.Formula = "=[@[FY Total]]-[@Budget]"
	$lo.ListColumns.Item("Variance %").DataBodyRange.Formula = "=IF([@Budget]=0,0,[@Variance]/[@Budget])"
	$lo.ShowTotals = $true
	foreach ($c in $months + @("Q1", "Q2", "Q3", "Q4", "FY Total", "Budget", "Variance")) {
		$lo.ListColumns.Item($c).TotalsCalculation = $xlTotalsSum
	}
	$lo.ListColumns.Item("Account No").DataBodyRange.NumberFormat = '0'
	$ws.Range("H5:AA10").NumberFormat = $FmtAcct0
	$lo.ListColumns.Item("Variance %").Range.NumberFormat = '0.0%;[Red]-0.0%'
	$lo.ListColumns.Item("Last Reviewed").DataBodyRange.NumberFormat = 'dd-mmm-yy'

	$hdr = $lo.HeaderRowRange
	$hdr.WrapText = $true
	$hdr.VerticalAlignment = $xlCenter
	$hdr.HorizontalAlignment = $xlCenter
	$ws.Rows.Item(4).RowHeight = 33
	$ws.Range("A5:AE9").RowHeight = 30
	$ws.Range("A5:AE9").VerticalAlignment = $xlTop
	$lo.ListColumns.Item("Notes").DataBodyRange.WrapText = $true
	$lo.ListColumns.Item("Account Name").DataBodyRange.WrapText = $true
	$b = $lo.Range.Borders
	$b.LineStyle = $xlContinuous
	$b.Weight = $xlThin
	$b.Color = (RGB 191 191 191)
	$ws.Range("T4:X10").Interior.ThemeColor = $xlThemeColorAccent2
	$ws.Range("T4:X10").Interior.TintAndShade = 0.8
	$ws.Range("T4:X4").Font.ThemeColor = $xlThemeColorLight1
	$ws.Range("X5:X10").Font.Bold = $true

	$w = @{ "A" = 9; "B" = 22; "C" = 14; "D" = 10; "E" = 11; "F" = 10; "G" = 8; "S" = 3 }
	foreach ($col in "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S") { $w[$col] = 9 }
	foreach ($col in "T", "U", "V", "W") { $w[$col] = 10 }
	$w["X"] = 11; $w["Y"] = 11; $w["Z"] = 11; $w["AA"] = 9; $w["AB"] = 14; $w["AC"] = 10; $w["AD"] = 11; $w["AE"] = 40
	Set-Widths $ws $w

	Invoke-Feature "hidden columns (Cost Center D, Currency G)" {
		$ws.Columns.Item("D").Hidden = $true
		$ws.Columns.Item("G").Hidden = $true
	}
	Invoke-Feature "column outline grouping on months H:S" {
		$ws.Range("H:S").Columns.Group() | Out-Null
	}
	Invoke-Feature "row outline grouping on sample rows 5:9" {
		$ws.Range("5:9").Rows.Group() | Out-Null
	}
	Invoke-Feature "sparkline per sample row in Trend column" {
		$ws.Range("AB5:AB9").SparklineGroups.Add(1, "GL!H5:S9") | Out-Null
		$sg = $ws.Range("AB5").SparklineGroups.Item(1)
		$sg.SeriesColor.ThemeColor = $xlThemeColorAccent1
		$sg.Points.Highpoint.Visible = $true
		$sg.Points.Highpoint.Color.Color = (RGB 0 176 80)
		$sg.Points.Lowpoint.Visible = $true
		$sg.Points.Lowpoint.Color.Color = (RGB 192 0 0)
	}
	Invoke-Feature "frozen panes (header + 2 columns)" { Set-Freeze $ws 4 2 }
	Invoke-Feature "print setup (landscape, titles)" { Set-PageSetup $ws $null '$4:$4' $true }

	Save-Template $wb "wide-and-styled" @'
# wide-and-styled

A 31-column ledger Table with heavy styling: hidden columns, row and column
outlines, custom heights, wrapping, borders and a sparkline per row.

## Regions

| Key | Kind | Sheet | Header | Sample rows | Row roles |
|---|---|---|---|---|---|
| `Ledger` | Table (`A4:AE10`, TableStyleMedium7, totals row 10) | GL | row 4 | 5-9 | all detail |

Columns (A..AE): Account No, Account Name, Department, Cost Center
(hidden), Owner, Region, Currency (hidden), Jan..Dec (H:S), Q1..Q4 (T:W,
calculated `=SUM(Ledger[@[Jan]:[Mar]])` etc.), FY Total (X), Budget (Y),
Variance (Z, calculated), Variance % (AA, calculated), Trend (AB, sparkline
only), Status, Last Reviewed (date), Notes (long wrapped text). Totals row
sums every month, quarter, FY Total, Budget, Variance.

## Scalar targets

None.

## Features

- Theme: Retrospect. TableStyleMedium7 plus manual accent2 band over the
  quarter/total block, thin grey borders everywhere
- Header row height 33 with wrap, sample rows 30 each (`customHeight`),
  wrapped Notes and Account Name
- Hidden columns D and G
- Column outline: H:S grouped (`outlineLevel="1"` on `<col>`)
- Row outline: rows 5-9 grouped (`outlineLevel="1"` on `<row>`), summary
  row below (the totals row)
- Sparkline group in `AB5:AB9` over `GL!H5:S9` (one sparkline per row,
  stored in the sheet `extLst` as `x14:sparklineGroups`)
- Frozen panes at C5, landscape print with titles `$4:$4`

## Hostile for a template engine

- `<cols>` has many `<col>` spans with hidden/outline attributes; row
  insertion must not touch them but they must survive byte-identical.
- Each inserted row must copy `ht`, `customHeight` and `outlineLevel` from
  the sample row, or the outline breaks.
- The sparkline group has one `<x14:sparkline>` per row with its own
  `<xm:f>GL!H5:S5</xm:f>` / `<xm:sqref>AB5</xm:sqref>`. Growing the table
  means adding sparklines, not just shifting refs -- and they are in an
  extLst the engine otherwise treats as opaque.
- Structured refs with spaces and special chars (`[@[FY Total]]`,
  `[Variance %]`) inside calculated columns.
'@
}

# =============================================================================
# 6. date1904
# =============================================================================
function Build-Date1904 {
	$wb = New-Workbook "Organic"
	$wb.Date1904 = $true
	$ws = $wb.Worksheets.Item(1)
	$ws.Name = "Shipments"

	Set-Title $ws "A1" "Shipment Lead Times" 18
	Set-Grid $ws 2 1 @(, @("As of", [datetime]"2025-06-30")) -Date1904
	$ws.Range("A2").Font.Bold = $true
	$ws.Range("B2").NumberFormat = 'd mmmm yyyy'
	$ws.Range("B2").HorizontalAlignment = $xlLeft
	$wb.Names.Add("AsOfDate", '=Shipments!$B$2') | Out-Null

	Set-Grid $ws 4 1 @(
		@("Order ID", "Customer", "Order Date", "Ship Date", "Days To Ship", "Age (days)"),
		@("SO-50211", "Bluebird Cafe", [datetime]"2025-05-02", [datetime]"2025-05-06"),
		@("SO-50219", "Harbor Supply Co.", [datetime]"2025-05-09", [datetime]"2025-05-20"),
		@("SO-50230", "Peak Outfitters", [datetime]"2025-05-28", [datetime]"2025-05-29"),
		@("SO-50244", "Lumen Labs", [datetime]"2025-06-12", [datetime]"2025-06-19")
	) -Date1904
	$lo = Add-Table $ws "A4:F8" "Shipments" "TableStyleMedium4"
	$lo.ListColumns.Item("Days To Ship").DataBodyRange.Formula = "=[@[Ship Date]]-[@[Order Date]]"
	$lo.ListColumns.Item("Age (days)").DataBodyRange.Formula = "=AsOfDate-[@[Order Date]]"
	$lo.ListColumns.Item("Order Date").DataBodyRange.NumberFormat = 'yyyy-mm-dd'
	$lo.ListColumns.Item("Ship Date").DataBodyRange.NumberFormat = 'yyyy-mm-dd'
	$lo.ListColumns.Item("Days To Ship").DataBodyRange.NumberFormat = '0'
	$lo.ListColumns.Item("Age (days)").DataBodyRange.NumberFormat = '0'
	$lo.ShowTotals = $true
	$lo.ListColumns.Item("Days To Ship").TotalsCalculation = $xlTotalsAverage
	(Get-TotalCell $lo "Days To Ship").NumberFormat = '0.0'
	$lo.ListColumns.Item("Order Date").TotalsCalculation = 5 # min
	(Get-TotalCell $lo "Order Date").NumberFormat = 'yyyy-mm-dd'
	Set-Widths $ws @{ "A" = 12; "B" = 20; "C" = 13; "D" = 13; "E" = 13; "F" = 12 }
	Invoke-Feature "hide gridlines" { Hide-Gridlines $ws }

	Save-Template $wb "date1904" @'
# date1904

A workbook on the 1904 date system (`<workbookPr date1904="1"/>`).

## Regions

| Key | Kind | Sheet | Header | Sample rows | Row roles |
|---|---|---|---|---|---|
| `Shipments` | Table (`A4:F9`, TableStyleMedium4, totals row 9) | Shipments | row 4 | 5-8 | all detail |

Columns: Order ID, Customer, Order Date, Ship Date, Days To Ship
(`=[@[Ship Date]]-[@[Order Date]]`), Age (days) (`=AsOfDate-[@[Order Date]]`).
Totals: MIN of Order Date (date formatted), AVERAGE of Days To Ship.

## Scalar targets

- `AsOfDate` (`Shipments!$B$2`), date-formatted `d mmmm yyyy`

## Features

- Theme: Organic. `Workbook.Date1904 = True`
- Date serials were written as `OADate - 1462`, so e.g. 2025-05-02 is
  stored as `44317`, not `45779`.

## Hostile for a template engine

- Every date the engine writes (table cells and the `AsOfDate` scalar) must
  use the 1904 epoch (1904-01-01 = 0), or everything is 4 years and 1 day
  off. Detect via `workbookPr/@date1904`.
- No 1900 leap-year bug applies in this system.
- Totals row applies a date format to a MIN over a date column.
'@
}

# =============================================================================
# 7. stacked-regions
# =============================================================================
function Build-StackedRegions {
	$wb = New-Workbook "Slice"
	$ws = $wb.Worksheets.Item(1)
	$ws.Name = "P&L"

	Set-Title $ws "A1" "Profit & Loss - Q4 FY2025" 18
	Set-Grid $ws 3 1 @(
		@("Category", "Description", "Amount"),
		@("Product", "Hardware units", 125000),
		@("Services", "Consulting engagements", 48500),
		@("Subscriptions", "SaaS recurring revenue", 36200),
		@("Other", "Interest income", 1150)
	)
	$rev = Add-Table $ws "A3:C7" "Revenue" "TableStyleMedium7"
	$rev.ShowTotals = $true
	$rev.ListColumns.Item("Amount").TotalsCalculation = $xlTotalsSum
	(Get-TotalCell $rev "Category").Value2 = "Total Revenue"
	# Revenue totals row = 8. Gap rows 9-10.
	Set-Grid $ws 11 1 @(
		@("Category", "Description", "Amount"),
		@("Payroll", "Salaries, wages & benefits", 92000),
		@("Occupancy", "Office rent", 14500),
		@("Technology", "Software & cloud hosting", 6800),
		@("Marketing", "Campaigns & events", 11250),
		@("Travel", "Customer visits", 4300)
	)
	$exp = Add-Table $ws "A11:C16" "Expenses" "TableStyleMedium3"
	$exp.ShowTotals = $true
	$exp.ListColumns.Item("Amount").TotalsCalculation = $xlTotalsSum
	(Get-TotalCell $exp "Category").Value2 = "Total Expenses"
	# Expenses totals row = 17.
	$ws.Range("C4:C8").NumberFormat = $FmtAcct
	$ws.Range("C12:C17").NumberFormat = $FmtAcct

	Set-Grid $ws 19 1 @(, @("Net Income", $null, "=C8-C17"))
	$ni = $ws.Range("A19:C19")
	Set-Band $ni $xlThemeColorAccent1 0 -Bold
	$ni.Font.ThemeColor = $xlThemeColorDark1
	$ni.Font.Size = 12
	$ws.Range("C19").NumberFormat = $FmtAcct
	$ws.Range("C19").Borders.Item($xlEdgeBottom).LineStyle = $xlDouble

	$ws.Range("A21:C24").Merge()
	$note = $ws.Range("A21")
	$note.Value2 = "Notes: Revenue is recognised under ASC 606. Subscription revenue is recognised ratably over the contract term. Interest income excludes unrealised gains on marketable securities. Figures are unaudited and subject to year-end adjustments."
	$note.WrapText = $true
	$note.VerticalAlignment = $xlTop
	$ws.Range("A21:C24").Interior.Color = (RGB 242 242 242)
	$ws.Range("A21:C24").Borders.LineStyle = $xlContinuous
	$ws.Range("A21:C24").Borders.Color = (RGB 191 191 191)
	$note.Font.Size = 9
	Set-Widths $ws @{ "A" = 18; "B" = 32; "C" = 16 }
	Invoke-Feature "hide gridlines" { Hide-Gridlines $ws }
	Invoke-Feature "print area covering everything" { Set-PageSetup $ws '$A$1:$C$24' $null $false }

	Save-Template $wb "stacked-regions" @'
# stacked-regions

Two Tables stacked vertically on one sheet: growing the first must push the
second (and everything below) down.

## Regions

| Key | Kind | Sheet | Header | Sample rows | Row roles |
|---|---|---|---|---|---|
| `Revenue` | Table (`A3:C8`, TableStyleMedium7, totals row 8) | P&L | row 3 | 4-7 | all detail |
| `Expenses` | Table (`A11:C17`, TableStyleMedium3, totals row 17) | P&L | row 11 | 12-16 | all detail |

Both tables have columns Category, Description, Amount. Totals rows sum
Amount; the Category totals cell holds a literal label ("Total Revenue",
"Total Expenses") instead of Excel's default "Total".

## Scalar targets

None.

## Features

- Theme: Slice. Sheet name `P&L` (needs quoting and XML-escaping: `P&amp;L`)
- `C19` Net Income `=C8-C17` -- A1 refs to BOTH totals rows
- Merged note block `A21:C24` below both tables (grey fill, border, wrap)
- Print area `'P&L'!$A$1:$C$24`

## Hostile for a template engine

- Rendering `Revenue` with N rows shifts the entire `Expenses` table
  (`tableN.xml ref`, `autoFilter ref`), `C19`, the merge `A21:C24` and the
  print area. Rendering both requires applying shifts in the right order.
- `=C8-C17` references two different totals rows by A1; both move by
  different amounts (Phase 4).
- The 2-row gap between the tables must be preserved.
- Two tables share column names; structured refs must be table-qualified.
'@
}

# =============================================================================
# 8. dynamic-arrays
# =============================================================================
function Build-DynamicArrays {
	$wb = New-Workbook "Ion"
	$wsP = $wb.Worksheets.Item(1)
	$wsP.Name = "Products"
	$wsA = Add-Sheet $wb "Analysis"

	Set-Grid $wsP 1 1 @(
		@("SKU", "Name", "Category", "Price", "Stock"),
		@("SKU-1001", "Trail Runner 2", "Footwear", 129.00, 42),
		@("SKU-1002", "Summit Jacket", "Outerwear", 249.50, 8),
		@("SKU-1003", "Merino Base Layer", "Apparel", 89.95, 65),
		@("SKU-1004", "Alpine Pack 38L", "Gear", 179.00, 12),
		@("SKU-1005", "Rain Shell", "Outerwear", 159.00, 5),
		@("SKU-1006", "Camp Sandal", "Footwear", 59.00, 120)
	)
	$lo = Add-Table $wsP "A1:E7" "Products" "TableStyleMedium6"
	$lo.ListColumns.Item("Price").Range.NumberFormat = $FmtCur
	$lo.ListColumns.Item("Stock").Range.NumberFormat = $FmtInt
	Set-Widths $wsP @{ "A" = 11; "B" = 22; "C" = 13; "D" = 11; "E" = 9 }

	Set-Title $wsA "A1" "Product Analysis (dynamic arrays)" 18
	$labels = @{ "A3" = "Categories (SORT/UNIQUE)"; "C3" = "Low stock < 20 (FILTER)"; "H3" = "Price lookup (XLOOKUP)"; "H6" = "Inventory value (LET)"; "H8" = "Category count (A4#)"; "A12" = "All products by price, desc (SORT over A1 range Products!A2:E7)" }
	foreach ($k in $labels.Keys) { $wsA.Range($k).Value2 = $labels[$k]; $wsA.Range($k).Font.Bold = $true; $wsA.Range($k).Font.ThemeColor = $xlThemeColorAccent1 }
	$wsA.Range("H4").Value2 = "SKU-1003"
	$wsA.Range("H4").Interior.Color = (RGB 255 242 204)
	Invoke-Feature "Formula2 dynamic arrays" {
		$wsA.Range("A4").Formula2 = "=SORT(UNIQUE(Products[Category]))"
		$wsA.Range("C4").Formula2 = '=FILTER(Products[[Name]:[Stock]],Products[Stock]<20,"None")'
		$wsA.Range("I4").Formula2 = '=XLOOKUP(H4,Products[SKU],Products[Price],"Not found")'
		$wsA.Range("I6").Formula2 = "=LET(p,Products[Price],q,Products[Stock],SUMPRODUCT(p,q))"
		$wsA.Range("I8").Formula2 = "=ROWS(A4#)"
		$wsA.Range("A13").Formula2 = "=SORT(Products!A2:E7,4,-1)"
	}
	$wsA.Range("E4:E10").NumberFormat = $FmtCur
	$wsA.Range("I4").NumberFormat = $FmtCur
	$wsA.Range("I6").NumberFormat = $FmtCur
	$wsA.Range("D13:D20").NumberFormat = $FmtCur
	Set-Widths $wsA @{ "A" = 14; "B" = 20; "C" = 20; "D" = 12; "E" = 11; "F" = 8; "G" = 3; "H" = 24; "I" = 14 }
	Invoke-Feature "hide gridlines (Analysis)" { Hide-Gridlines $wsA }

	Save-Template $wb "dynamic-arrays" @'
# dynamic-arrays

A Table feeding dynamic-array formulas (written with `Range.Formula2`) and
their spill ranges on another sheet.

## Regions

| Key | Kind | Sheet | Header | Sample rows | Row roles |
|---|---|---|---|---|---|
| `Products` | Table (`A1:E7`, TableStyleMedium6, no totals row) | Products | row 1 | 2-7 | all detail |

Columns: SKU, Name, Category, Price, Stock.

## Scalar targets

None (`Analysis!H4` is a manual input cell, value `SKU-1003`).

## Features (sheet Analysis)

- `A4 =SORT(UNIQUE(Products[Category]))` -- spills down
- `C4 =FILTER(Products[[Name]:[Stock]],Products[Stock]<20,"None")` -- spills
  4 columns wide, N rows down
- `I4 =XLOOKUP(H4,Products[SKU],Products[Price],"Not found")`
- `I6 =LET(p,Products[Price],q,Products[Stock],SUMPRODUCT(p,q))`
- `I8 =ROWS(A4#)` -- spill-range reference operator
- `A13 =SORT(Products!A2:E7,4,-1)` -- dynamic array over an A1 range that
  is exactly the sample rows

## Hostile for a template engine

- Spilling dynamic-array formulas (`A4`, `C4`, `A13`) are stored as
  `<f t="array" ref="A4:A7">` with `cm="1"` pointing into `xl/metadata.xml` (`XLDAPR` cell metadata), and
  functions are prefixed (`_xlfn._xlws.SORT`, `_xlfn.UNIQUE`,
  `_xlfn._xlws.FILTER`, `_xlfn.XLOOKUP`, `_xlfn.LET`, `_xlpm.p`,
  `_xlfn.ANCHORARRAY` for `A4#`). Losing `metadata.xml` or `cm` turns them
  into legacy CSE formulas showing `@` implicit intersection.
- Spill results are stored as plain cached values in the cells below/right of
  each formula. After a render the spill size changes; stale spilled values
  must be removed or Excel may show `#SPILL!` until recalculated.
- `SORT(Products!A2:E7,...)` must grow with the table (Phase 4 / v1
  constraint violation, deliberately).
'@
}

# =============================================================================
# 9. threaded-comments
# =============================================================================
function Build-ThreadedComments {
	$wb = New-Workbook "Wisp"
	$ws = $wb.Worksheets.Item(1)
	$ws.Name = "Tasks"

	Set-Title $ws "A1" "Month-End Close Checklist" 18
	Set-Grid $ws 3 1 @(
		@("Task", "Owner", "Due", "Status"),
		@("Bank reconciliations", "M. Ito", [datetime]"2025-10-03", "Done"),
		@("Accrued expenses review", "S. Patel", [datetime]"2025-10-04", "In progress"),
		@("Revenue cut-off testing", "D. Okoro", [datetime]"2025-10-06", "Not started"),
		@("Fixed asset roll-forward", "M. Ito", [datetime]"2025-10-07", "Not started")
	)
	$lo = Add-Table $ws "A3:D7" "Tasks" "TableStyleLight9"
	$lo.ListColumns.Item("Due").DataBodyRange.NumberFormat = 'ddd mm/dd'
	Set-Widths $ws @{ "A" = 30; "B" = 14; "C" = 12; "D" = 14 }
	$ws.Range("A9").Value2 = "Sign-off"
	$ws.Range("A9").Font.Bold = $true
	$ws.Range("B9").Value2 = "Controller"
	$ws.Range("A11").Value2 = "Reviewer"

	Invoke-Feature "threaded comment on header cell B3" {
		$t = $ws.Range("B3").AddCommentThreaded("Owner must be a named person, not a team.")
		$t.AddReply("Agreed - updated the template guidance.") | Out-Null
	}
	Invoke-Feature "threaded comment below table (B9)" {
		$ws.Range("B9").AddCommentThreaded("Controller signs off once every task is Done.") | Out-Null
	}
	Invoke-Feature "legacy note inside table body (A5)" {
		$c = $ws.Range("A5").AddComment("Legacy note: includes the new AP subledger this month.")
		$c.Shape.Width = 200; $c.Shape.Height = 50
	}
	Invoke-Feature "legacy note below table (A11)" {
		$ws.Range("A11").AddComment("Legacy note: reviewer initials go in B11.") | Out-Null
	}
	Invoke-Feature "hide gridlines" { Hide-Gridlines $ws }

	Save-Template $wb "threaded-comments" @'
# threaded-comments

A small Table with modern threaded comments and legacy notes, both inside
and below the region.

## Regions

| Key | Kind | Sheet | Header | Sample rows | Row roles |
|---|---|---|---|---|---|
| `Tasks` | Table (`A3:D7`, TableStyleLight9, no totals row) | Tasks | row 3 | 4-7 | all detail |

Columns: Task, Owner, Due (date `ddd mm/dd`), Status.

## Scalar targets

None.

## Features

- Theme: Wisp
- Threaded comment (with one reply) on header cell `B3`
- Threaded comment on `B9`, below the table
- Legacy note on `A5` (inside a sample row) and on `A11` (below the table)

## Hostile for a template engine

- Threaded comments are stored in `xl/threadedComments/threadedCommentN.xml`
  plus `xl/persons/person.xml`, AND Excel writes a legacy fallback comment
  for each thread into `xl/commentsN.xml` (text starting
  `[Threaded comment]`) with a VML shape in `xl/drawings/vmlDrawingN.vml`.
  All three must stay consistent: the `ref` of a thread below the table must
  shift in threadedComments, comments and the VML `<x:Row>` anchor.
- The legacy note on `A5` is attached to a sample row. When sample rows are
  cleared, the note should either be dropped (with its VML shape) or kept
  on the first rendered row -- never left pointing at a stale cell with a
  dangling shape.
- The header-cell thread must not move.
'@
}

# =============================================================================
# 10. minimal-single-row
# =============================================================================
function Build-MinimalSingleRow {
	$wb = New-Workbook $null
	$ws = $wb.Worksheets.Item(1)
	$ws.Name = "Sheet1"
	Set-Grid $ws 1 1 @(
		@("Name", "Email", "Phone"),
		@("Ada Lovelace", "ada@example.com", "+1 555 0100")
	)
	$lo = Add-Table $ws "A1:C2" "Contacts" "TableStyleMedium2"
	Set-Widths $ws @{ "A" = 18; "B" = 24; "C" = 14 }

	Save-Template $wb "minimal-single-row" @'
# minimal-single-row

The smallest possible region: a Table at A1 with exactly one sample row, no
totals row, nothing else in the workbook.

## Regions

| Key | Kind | Sheet | Header | Sample rows | Row roles |
|---|---|---|---|---|---|
| `Contacts` | Table (`A1:C2`, TableStyleMedium2, no totals row) | Sheet1 | row 1 | 2 | detail |

Columns: Name, Email, Phone.

## Scalar targets

None.

## Hostile for a template engine

- Header is row 1: there is no row above the table to anchor anything to.
- Only one sample row: rendering zero rows must still leave a valid table
  (Excel requires at least one body row -- the engine must keep an empty
  row, not produce `ref="A1:C1"`).
- Nothing below the table, so the `<dimension>` is the only thing that
  grows besides the table ref and autoFilter.
- Default theme (no theme applied), default fonts.
'@
}

# =============================================================================
# main
# =============================================================================
$builders = [ordered]@{
	"flat-table-totals"  = "Build-FlatTableTotals"
	"named-range-region" = "Build-NamedRangeRegion"
	"multi-sheet-summary" = "Build-MultiSheetSummary"
	"indirect-rates"     = "Build-IndirectRates"
	"wide-and-styled"    = "Build-WideAndStyled"
	"date1904"           = "Build-Date1904"
	"stacked-regions"    = "Build-StackedRegions"
	"dynamic-arrays"     = "Build-DynamicArrays"
	"threaded-comments"  = "Build-ThreadedComments"
	"minimal-single-row" = "Build-MinimalSingleRow"
}
# powershell -File passes "a,b" as one string; accept both forms.
if ($Only) { $Only = @($Only | ForEach-Object { $_ -split "," } | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
if ($Only) {
	foreach ($o in $Only) { if (-not $builders.Contains($o)) { throw "Unknown template '$o'. Known: $($builders.Keys -join ', ')" } }
}

$excel = New-Object -ComObject Excel.Application
$failed = @()
try {
	$excel.Visible = $false
	$excel.DisplayAlerts = $false
	$excel.ScreenUpdating = $false
	Write-Host "Excel $($excel.Version) build $($excel.Build); output -> $OutDir"
	foreach ($name in $builders.Keys) {
		if ($Only -and ($Only -notcontains $name)) { continue }
		Write-Host "Building $name..."
		$script:BuildLog = New-Object System.Collections.Generic.List[string]
		try {
			& $builders[$name]
		}
		catch {
			$failed += $name
			Write-Host "  FAILED: $($_.Exception.Message) at $($_.InvocationInfo.PositionMessage)" -ForegroundColor Red
			foreach ($wb in @($excel.Workbooks)) { try { $wb.Close($false) } catch { } }
		}
	}
}
finally {
	try { foreach ($wb in @($excel.Workbooks)) { $wb.Close($false) } } catch { }
	$excel.ScreenUpdating = $true
	$excel.Quit()
	[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
	Remove-Variable excel -ErrorAction SilentlyContinue
	[GC]::Collect()
	[GC]::WaitForPendingFinalizers()
}

if ($failed.Count -gt 0) {
	Write-Host "Failed templates: $($failed -join ', ')" -ForegroundColor Red
	exit 1
}
Write-Host "Done." -ForegroundColor Green
exit 0
