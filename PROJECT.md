# PROJECT.md

## Why this exists

`excel-graft` replaces an Express API's use of a forked ExcelJS
(`@rmartin93/exceljs-fork`, see `../../LocalSites/exceljs-fork`) to fill
Excel templates with data from a database and serve the result.

ExcelJS parses the whole workbook into its own object model and re-serializes
a new file from that model. Anything the model doesn't understand — newer
`extLst` extensions, x14 conditional formatting, slicers, threaded comments —
gets dropped or mangled, and Excel shows a repair prompt on open. That
happened in production; `corpus/templates/exceljs-fork-broken/` holds the
actual template and the actual corrupted output from that incident, which is
this project's first regression fixture.

The fix is a different architecture, not a better model: treat `.xlsx` as a
zip of XML parts, copy every part byte-for-byte, and patch only the XML
that has to change. See [CLAUDE.md](CLAUDE.md) for the invariants this
implies and must never be violated.

## Template model

**The template decides the shape, never the engine.** A template is a
workbook a business user designed in Excel, often full of sample data:
highlighted group rows, subtotals, grand totals, number formats. The
engine's only job is to clear the sample data and insert current data
in the same shape. It infers the structure from the sample rows (which
rows are group headers, details, subtotals and totals; which styles go
with each) and regenerates the formulas so they cover the inserted rows.
It never asks the caller to choose a layout. The first real example is
`corpus/templates/transactions/`.

The table below is the original placeholder-based sketch. Placeholders
stay useful for single cells like a report date, but a template must not
need them to be filled.

| Feature | Syntax / mechanism | Notes |
|---|---|---|
| Scalar values | `{{customer.name}}` in a cell | Keeps the cell's style (`s` attribute), writes a typed value |
| Repeating rows | An Excel Table named in the template (e.g. `tblItems`) | Primary mechanism. Structured refs (`tblItems[Amount]`) keep working as the table grows |
| Repeating rows (fallback) | Marker row: `{{#each items}}` … `{{item.qty}}` | For layouts that can't use a Table |
| Formats | Taken from the template row | The designer controls styling in Excel, not this tool |

## What must be patched when rows are inserted

| Part | What to update |
|---|---|
| `sheetN.xml` | Row `r` values, cell refs, `<dimension>`, `<mergeCells>`, `<conditionalFormatting sqref>`, `<dataValidations>`, `<hyperlinks>`, `<autoFilter>` |
| `tables/tableN.xml` | `ref` and the `autoFilter` ref |
| `workbook.xml` | Defined names (print areas, named ranges); `<calcPr fullCalcOnLoad="1"/>` |
| `calcChain.xml` | Delete it (plus its `.rels` entry and content-type override) |
| Formula cells | Strip cached `<v>` so Excel recalculates on open |
| `drawings/` | Shift anchors for images/charts below inserted rows |
| Charts | Update series ranges over the grown region (or require charts to reference Tables) |

## Stack

| Concern | Choice |
|---|---|
| Language | TypeScript `strict`, `noUncheckedIndexedAccess` |
| Zip | `fflate` |
| XML | `saxes` tokenizer for sheet XML; raw string passthrough for everything else |
| Build | `tsup`, dual ESM/CJS, Node >= 18 |
| Tests | Vitest, `fast-check` (property tests), Stryker (mutation testing, later) |

## API sketch (target shape, not all implemented yet)

```ts
const tpl = await Template.load(buffer);
const out = await tpl.render({
  customer: { name: "Acme" },
  tblItems: rows, // Table name -> array
}, { onMissing: "error" | "blank" });
out.pipe(res);
```

`tpl.inspect()` lists the placeholders and tables a template contains, for
validating templates at startup.

## Battle-testing strategy

1. **Template corpus** — `corpus/templates/`, 20-40 real templates in
   current Excel, one feature each (tables+totals, CF, DV, merged cells,
   images, charts, print areas, frozen panes, multi-sheet, 1904 date
   system), plus real templates from production.
2. **Passthrough invariant** — every untouched part byte-identical,
   asserted on every render test.
3. **Structural validation in CI** — Open XML SDK validator
   (`tools/validator/`, not yet built — needs the .NET SDK, not currently
   installed on this machine) + LibreOffice headless conversion as a
   second check.
4. **Real Excel smoke test** — `tools/excel-harness/`, PowerShell + Excel
   COM, opens every corpus output with alerts captured, fails on any
   repair prompt, re-saves to confirm Excel accepts it. Ground truth; run
   before every release.
5. **Property tests with hostile data** — unicode/emoji, XML-invalid
   control chars, strings > 32,767 chars, `NaN`/`Infinity`,
   `null`/`undefined`, 1900 leap-bug boundary dates, huge/tiny numbers,
   leading `=` (must not become a formula).
6. **Limits and scale** — 1,048,576-row overflow, 100k-row renders within
   memory/time budget, concurrent renders in Express.
7. **Mutation testing** — Stryker on ref-shifting and escaping modules.

## Phases

| Phase | Scope | Exit criteria | Status |
|---|---|---|---|
| 0 | Repo, CI, validator tool, Excel COM harness, first corpus templates | Harness catches the ExcelJS-corrupted file | **in progress** |
| 1 | Zip passthrough and scalar `{{placeholders}}` with typed values | Zero repairs across the corpus; passthrough invariant holds | not started |
| 2 | Table-based repeats, including all dependent-part updates | Tables with totals, CF, DV, charts all open clean | not started |
| 3 | Marker-row repeats, multiple sheets, streaming output | 100k rows within budget | not started |
| 4 | A1 formula reference shifting (tokenizer, shared formulas) | Property tests show formula refs stay correct after shifts | not started |
| 5 | Docs, `inspect()`, 1.0 release, `test.html` demo harness against a local Express API | Tested in a real Express API for a few weeks | not started |

## Prior art

`xlsx-template` (npm) uses the same patch-don't-rebuild idea in older
JavaScript. Worth reading before Phase 2's Table logic.
