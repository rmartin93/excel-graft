# known-bad

Deliberately corrupt workbooks. They exist to prove that
`tools/excel-harness/Test-ExcelFiles.ps1` can tell when Excel wants to
repair a file. They are NOT under `corpus/templates/`, so the default
harness run (`npm run harness:excel`) does not pick them up. Run them on
purpose; every fixture except the ones marked "not flagged" should FAIL:

```
powershell -NoProfile -ExecutionPolicy Bypass -File tools/excel-harness/Test-ExcelFiles.ps1 -Path corpus/known-bad
```

Rebuild them (after `Build-Corpus.ps1`, since they are cut from its output):

```
powershell -NoProfile -ExecutionPolicy Bypass -File tools/corpus-builder/Build-KnownBad.ps1
```

`Build-KnownBad.ps1` copies every zip entry of a clean template in the
original order (`[Content_Types].xml` stays first) with
System.IO.Compression and changes only one part, by a literal/regex edit that
throws if the text it expects is missing.

## Findings (Excel 16.0 build 20326, Windows 11, 2026-09-28)

**Excel does not repair silently under automation.** The old harness header
said it did. That was only tested against `exceljs-fork-broken/Test-Output.xlsx`,
which this Excel build does not repair at all. With `DisplayAlerts = $false`
and the default `CorruptLoad` (xlNormalLoad), Excel answers the
"We found a problem with some content..." prompt with **No**, so
`Workbooks.Open` throws `Unable to get the Open property of the Workbooks class`.

We checked this against the real UI. Each fixture was opened the way a user
opens it (`excel.exe /x <file>`, visible), and UI Automation read the
NUIDialog text. **"COM Open throws" matched "a user gets the repair prompt"
for all 13 fixtures, with no false positives and no false negatives.**

| Signal | Reliable? | Notes |
|---|---|---|
| `Workbooks.Open` throws (DisplayAlerts off, normal load) | **Yes: primary signal** | Matched the interactive prompt for every fixture. This is what the harness fails on. |
| `Open(..., CorruptLoad = 1 /*xlRepairFile*/)` + recovery log `%TEMP%\errorNNNNNN_01.xml` | **Yes: diagnosis** | Repairs silently and writes a `<recoveryLog>` naming the part and record class. In repair mode Excel writes a log for EVERY file, clean ones included. A clean file's log has only `<summary>`/`<additionalInfo>`, while a repaired one has `<removedRecords>`, `<repairedRecords>`, `<removedParts>` or `<repairedParts>`. The harness uses this to explain failures, and runs it as a second check on files that open normally (it has never fired on those so far). |
| Window caption contains "Repaired" | Only interactively | Interactive Excel shows `name.xlsx  -  Repaired - Excel` after you click Yes. Under COM with `xlRepairFile` the `Window.Caption` is unchanged, so it can't be used. |
| Table/cell counts after repair-mode open | Weak | `h-table-ref-too-short` shows the repaired table ref (`A5:F10`). Most others look identical to the original. |
| `CorruptLoad = 2` (xlExtractData) | No | Opens everything, including clean files, as values only, and drops every table. |
| Visible + DisplayAlerts on + watching for the dialog | Works, but fragile | UI Automation finds the NUIDialog as a descendant of `XLMAIN`, not as a top-level window. EnumWindows-style top-level scans miss it. It needs a visible desktop session, and killing Excel with the dialog up can leave recovery state behind. The harness doesn't use it. |

COM quirk: the 15-argument `Workbooks.Open` fails whenever the middle
optional arguments are `[Type]::Missing`. Pass all of them explicitly:
`Open(path, 0, $true, 5, "", "", $true, 2, ",", $false, $false, 0, $false, $false, 1)`.

After a failed or repair-mode open, `Excel.Quit()` sometimes leaves
`EXCEL.EXE /automation -Embedding` running. The harness now records the
PID (from `Application.Hwnd`) and kills the process if it hasn't exited 15 s
after Quit. A stale automation instance left over from an earlier session
caused one false failure run during this investigation. If results look
odd, check `Get-Process EXCEL` first.

## Per-fixture results

"Prompt" is what a user sees opening the file interactively. "COM normal"
is `Workbooks.Open` with alerts off. "Repair log" is the recovery-log entry
from a repair-mode open (or the interactive "Yes").

| Fixture | Base template | Corruption | Prompt | COM normal | COM repair mode | Repair log | Harness |
|---|---|---|---|---|---|---|---|
| `control-rezipped-clean` | flat-table-totals | none (re-zipped only) | no | opens | opens | (nothing) | OK |
| `a1-autofilter-header-only` | flat-table-totals | table `autoFilter ref="A5:F5"` (header only) vs table `A5:F11`. This is the ExcelJS bug. | **no** | opens | opens | (nothing) | OK (**not flagged**) |
| `a2-autofilter-narrow` | flat-table-totals | `autoFilter ref="A5:C10"`, 3 of 6 columns | **no** | opens | opens | (nothing) | OK (**not flagged**) |
| `b-duplicate-cell-ref` | minimal-single-row | second `<c r="C2">` in row 2 | yes | throws | opens | Removed Records: Cell information from /xl/worksheets/sheet1.xml part | FAILED |
| `c-rows-out-of-order` | flat-table-totals | `<row r="3">` written before `<row r="2">` | yes | throws | opens | Removed Records: Cell information from /xl/worksheets/sheet1.xml part | FAILED |
| `d-inline-string-too-long` | minimal-single-row | `t="inlineStr"` of 40,000 chars (limit 32,767) | yes | throws | opens | Repaired Records: String properties from /xl/worksheets/sheet1.xml part | FAILED |
| `e-sqref-past-max-row` | flat-table-totals | dataValidation `sqref="C6:C1048577"` | yes | throws | **throws** | interactive Yes: "Repaired Part: /xl/worksheets/sheet1.xml part with XML error. Load error. Line 2, column 0." The whole sheet is replaced. | FAILED ("could not open even in repair mode") |
| `e2-cf-sqref-past-max-row` | flat-table-totals | conditionalFormatting `sqref="D6:D1048577"` | **no** | opens | opens | (nothing) | OK (**not flagged**) |
| `f-table-header-mismatch` | minimal-single-row | tableColumn `name="FullName"`, header cell text is "Name" | yes | throws | opens | Repaired Records: Table from /xl/tables/table1.xml part (Table) | FAILED |
| `g-overlapping-merges` | stacked-regions | added `<mergeCell ref="B23:D26"/>` overlapping `A21:C24` | yes | throws | opens | Removed Records: Merge cells from /xl/worksheets/sheet1.xml part | FAILED |
| `h-table-ref-too-short` | flat-table-totals | table `ref="A5:F10"` with `totalsRowCount="1"` (should be `A5:F11`) | yes | throws | opens | Repaired Records: Table from /xl/tables/table1.xml part (Table) | FAILED |
| `i-malformed-sheet-xml` | minimal-single-row | `</sheetData>` removed (not well-formed) | yes | throws | **throws** | interactive Yes: "Replaced Part: /xl/worksheets/sheet1.xml part with XML error. The name in the end tag of the element must match the element type in the start tag. Line 2, column 1394." | FAILED ("could not open even in repair mode") |
| `j-orphan-shared-formula` | incurred-cost | removed the master `<f t="shared" ref="D5:D17" si="0">`, leaving children `<f t="shared" si="0"/>` | yes | throws | opens | Removed Records: Shared formula from /xl/worksheets/sheet3.xml part; Removed Records: Formula from /xl/calcChain.xml part | FAILED |
| `templates/exceljs-fork-broken/Test-Output.xlsx` (not in this folder) | n/a | real ExcelJS output | **no** (on this build; it prompted on another machine) | opens | opens | (nothing) | OK (**not flagged**) |

## What this means for the engine

- The harness catches every corruption class that makes THIS Excel build
  prompt, and names the broken part. It can't see the classes this build
  tolerates: autoFilter/table-ref mismatches (a1, a2) and an out-of-range CF
  sqref (e2). Those need structural assertions in the unit tests (see
  `corpus/README.md` for the ExcelJS fixture).
- An out-of-range DV sqref (e) destroys the whole sheet on repair. Any
  sqref the engine grows must be clamped to row 1,048,576.
- A shared-formula master left behind in a deleted sample row (j) is a
  real repair trigger. The engine must expand or re-master shared formulas
  when it removes sample rows (see `templates/incurred-cost`).
