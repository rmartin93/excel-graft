# Corpus

Real `.xlsx` files used as regression fixtures. Every bug gets a corpus
template first — see CLAUDE.md rule 6.

## templates/

One subfolder per template, one feature (or one incident) each.

- `exceljs-fork-broken/` — the first fixture, from the production incident
  that started this project. `Test-Template.xlsx` is the source template
  (made in Excel, not corrupted). `Test-Output.xlsx` is what
  `@rmartin93/exceljs-fork` rendered from it — the file that showed a
  repair prompt when opened in Excel on another machine. On this machine's
  Excel (2026-09-28) it opens with **no** repair prompt, so Excel builds
  differ in strictness and "opens clean here" is not proof of a valid
  file. No repair log is available. A part-by-part diff against the
  template shows what ExcelJS broke:
  - **Table `autoFilter` ref is wrong.** Table `ref="A1:C5"` with
    `totalsRowCount="1"` needs `autoFilter ref="A1:C4"` (header + data rows,
    excluding totals). ExcelJS wrote `A1:C1`. This is the most likely
    repair trigger.
  - **Style lost.** `<dxfs>` was emptied (`count="0"`) and the Name
    column's `dataDxfId="0"` was dropped, so the yellow fill is gone.
  - **Metadata stripped.** Every `xr:uid`/`xr3:uid`, the `xr*` namespaces,
    the workbook `extLst` (`calcFeatures`, `chartTrackingRefBase`) and
    `absPath` were removed; `fileVersion` was downgraded.
  - **Relationships renumbered.** Worksheet rId changed from `rId1` to
    `rId4`.
  - **Cached formula value dropped** from the totals cell, with no
    `fullCalcOnLoad`.

  The render tests for this fixture should include a structural check for
  each of these, since Excel on this machine won't flag them.

- `transactions/Transactions-Template.xlsx` — added 2026-09-28 to test that
  highlighting and formulas survive a render, and to load large data
  volumes. One Table `Table1` (`A4:C11`, columns Merchant / Date / Amount)
  holding highlighted category rows (fill `theme 4` tint) with
  `SUM(C6:C7)`-style subtotals, detail rows (date and currency formats),
  and a highlighted grand-total row `SUM(C5,C8)`. The subtotals use plain
  A1 references, not structured references, so growing the rows under them
  needs either Phase 4 reference shifting or regenerated formulas.
  Placeholders not added yet.

- Generated templates (2026-09-28), built from scratch in Excel via COM by
  `tools/corpus-builder/Build-Corpus.ps1`. Each folder holds
  `template.xlsx` and a `README.md` listing its regions (name, sheet,
  sample rows, row roles), scalar targets, features and engine-hostile
  details:
  - `flat-table-totals/`: invoice Table with a calculated column, totals row,
    x14 data bars, formula CF, DV list, frozen panes, print area/titles,
    `{{placeholders}}`, and an A1 ref to the totals row below the table.
  - `named-range-region/`: defined-name region (not a Table), merged title,
    alternating fills, A1 total row, chart over the sample rows, a picture
    above and a shape/note/hyperlink below the region.
  - `multi-sheet-summary/`: Table `Sales` feeding SUMIFS plus A1 cross-sheet
    refs, a PivotTable, a slicer, and a hidden lookup sheet driving a DV list.
  - `indirect-rates/`: indirect cost pools and rates. Scalar named cells on a
    protected cover sheet, Table `DirectLabor`, grouped defined-name region
    `IndirectPools` (pool header rows with subtotals, then a grand total),
    rates via cross-sheet A1 refs and named formulas, an icon set, a chart,
    and shared formulas.
  - `wide-and-styled/`: 31-column Table with hidden columns, row/column
    outlines, custom heights, wrap, borders, and a sparkline per row.
  - `date1904/`: 1904 date system.
  - `stacked-regions/`: two Tables stacked on one sheet, a net-income cell
    referencing both totals rows, a merged block, and a print area.
  - `dynamic-arrays/`: FILTER/SORT/UNIQUE/XLOOKUP/LET over a Table, with spills.
  - `threaded-comments/`: threaded comments and legacy notes, in and around
    a Table.
  - `minimal-single-row/`: a Table at A1 with one sample row and nothing else.
- Built by `tools/corpus-builder/Build-GroupedExtras.ps1`:
  - `grouped-footers/`: named-range region with a subtotal row *below* each
    group ("Engineering subtotal"), a blank spacer row after each group,
    banded detail rows, a running balance whose first row differs, a grand
    total over the footers, and a cell below the region that uses it.
  - `nested-groups/`: a Table with Region > Category > line items, subtotal
    header rows at both levels, a grand total, variance formulas, CF and a chart.

Still needed (see PROJECT.md phase 0 exit criteria): more real templates
from production.

## known-bad/

Deliberately corrupt fixtures, each with one hand-made corruption, and the
empirical record of how Excel reacts to them. This is how we know the
harness's repair detection works. It is not part of the default harness
run. See `known-bad/README.md`.

## repair-logs/

Repair logs saved from Excel's "View the repair log" flow after Excel
fixes a corrupted file. These are the first regression tests for any new
corruption class found.
