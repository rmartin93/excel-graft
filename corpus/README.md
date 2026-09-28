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

Still needed (see PROJECT.md phase 0 exit criteria): 20-40 templates
covering tables-with-totals, conditional formatting, data validation,
merged cells, images, charts, print areas, frozen panes, multiple sheets,
and the 1904 date system, plus more real templates from production.

## repair-logs/

Repair logs saved from Excel's "View the repair log" flow after Excel
fixes a corrupted file. These are the first regression tests for any new
corruption class found.
