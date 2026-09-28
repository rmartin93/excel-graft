# excel-graft

Fill real Excel workbooks with fresh data — without breaking them.

Business users design a workbook in Excel: highlighted category rows,
subtotals, a grand total, conditional formatting, charts, a rates sheet that
points at the totals. They fill it with **sample data** so it looks right.
excel-graft uses that workbook as the template: it clears the sample rows,
inserts your data **in the same shape**, rewrites every formula that pointed
at the old rows, and copies every other part of the file through
byte-for-byte.

It was built to replace an ExcelJS-based export that produced files Excel
offered to "repair". ExcelJS rebuilds the whole workbook from its own object
model and silently drops what it doesn't understand. excel-graft never
rebuilds anything.

```ts
import { Template } from "excel-graft";

const tpl = Template.loadSync(fs.readFileSync("transactions.xlsx")); // once, at startup

app.get("/api/transactions.xlsx", async (req, res) => {
	const xlsx = await tpl.render({
		Table1: [
			{ Merchant: "Groceries", rows: [{ Merchant: "HEB", Date: new Date("2026-01-03"), Amount: 12.5 }] },
			{ Merchant: "Fuel", rows: [{ Merchant: "Shell", Date: new Date("2026-01-06"), Amount: 40 }] },
		],
	});
	res.attachment("transactions.xlsx").send(Buffer.from(xlsx));
});
```

```
npm install excel-graft
```

Node 18+, ESM and CommonJS, TypeScript types included. Two small
dependencies (`fflate`, `saxes`).

## How a template is read

You don't annotate anything. The engine reads the workbook the way a person
would.

**Where the data goes** — any of:

| In the workbook | Data key | Notes |
|---|---|---|
| An Excel Table (Insert → Table) | the table name, e.g. `Table1` | Best choice. Body rows between the header and the totals row are the sample rows. Field names are the column headers. |
| A named range over the sample rows (select them, type a name in the Name Box) | the name, e.g. `Employees` | For layouts that can't be a Table (merged cells, several header rows). Field names come from the row just above the range. |
| A named single cell | the name, e.g. `ContractorName` | Gets one value, keeps its formatting. |
| `{{customer.name}}` typed in a cell | `customer` (`{ name }`) | Whole-cell placeholders get a typed value (numbers stay numbers); `Report for {{month}}` interpolates. |

Regions you don't pass data for are left exactly as they are.

**No Tables or names in the workbook?** Declare regions in code instead —
nothing is written into the workbook (don't inject `<definedNames>` by hand:
in the wrong place Excel deletes them and shows the repair prompt):

```ts
const tpl = Template.loadSync(bytes, {
	regions: {
		DisclosureData: "'Disclosure Table'!A7:M207",
		CostCenterData: "'2026 CC'!A6:H66",
	},
});
```

Blank rows at the ends of a declared range are spare capacity: they're
cleared, so the range can be generous. From the CLI:
`npx excel-graft inspect book.xlsx --region "DisclosureData='Disclosure Table'!A7:M207"`.

**What the sample rows mean** — decided by their formulas:

- A row whose formula totals *other sample rows* (`=SUM(C6:C7)`) is a
  **group row**. If it sits above the rows it totals it's a group header; below
  them, a group footer. Its label cells (e.g. "Groceries") become group
  fields.
- A row that totals the group rows (`=SUM(C5,C8)`), or everything, is a
  **grand total**: kept once, with its formula rewritten to cover all the
  groups you pass.
- Every other row is a **detail row**. Formulas in detail rows behave like
  Excel's fill-down: `=C6*$B$2` becomes `=C7*$B$2`, `=D5+C6` stays a running
  balance, `=C6/C$5` keeps pointing at its own group's header.
- **No subtotal formulas?** The styling is read instead: the most common row
  style is the detail row, and rows styled differently that fill fewer cells
  (a yellow pool row with just a label, a blue sub-pool row) are group levels,
  outermost first. Rows below the last detail that look different (notes, a
  total line) are kept once. `inspect()` reports which was used in
  `structure` (`"formulas"`, `"styles"`, `"layout"` or `"flat"`).
- **Still not what you meant?** Spell the rows out with a layout:
  `{ range, layout: { levels: [{ header: 7 }, { header: 8 }], detail: 9, fixedRows: [26, 27] } }`
  (sheet row numbers; levels outermost first, each with a `header` row, a
  `footer` row, or both).
- Blank rows between groups are spacers and repeat with each group.
- Alternating fills (banded rows) are repeated in the same pattern. If the
  first detail row has a different formula (an opening balance), it's used
  for the first row only.

So the `transactions.xlsx` layout — a highlighted category row with
`=SUM(...)` over its transactions, repeated, then a Total row — becomes:

```ts
type Data = {
	Table1: { Merchant: string; rows: { Merchant: string; Date: Date; Amount: number }[] }[];
};
```

Run `npx excel-graft inspect template.xlsx` to see exactly what was inferred,
or open the demo (below) and look at the card for your workbook.

**Tips for template designers**

- Put at least **two sample groups** in a grouped layout. With one group and
  a total below it, the engine can't tell "a group and its subtotal" from
  "a list and its grand total" (it picks the latter).
- Formatting comes from the sample rows: style the first detail row, the
  group row and the total row the way you want every one of them to look.
- Anything outside the data rows — titles, notes, charts, a rates sheet —
  stays put or moves down with the rows below the data, formulas included.

## Formulas and references

Every formula in the workbook that points into or below a filled region is
rewritten: other cells on the sheet, other sheets, defined names, conditional
formatting and data validation (including Excel 2010+ extensions),
sparklines, chart series, pivot-table sources, print areas, merged cells,
hyperlinks, comments, images and charts anchored below the data.

References from *outside* a region to one specific sample row follow that row
by meaning, not by position:

- A rates sheet cell `='Indirect Pools'!C9` pointing at the **Overhead** pool's
  subtotal follows the group labeled "Overhead", wherever it lands.
- If the data has no such group, the reference becomes `#REF!` (and
  `onWarning` says why) — never a different pool's numbers.
- A range over the sample rows (`=SUM(Data!D2:D6)`) grows to cover all rows.

Excel recalculates everything when the file opens (`fullCalcOnLoad`, and the
stale calc chain is removed), so cached values never lie.

## Type safety

Generate a type from the template so a renamed column or a new grouping
level is a compile error, not a blank column in production:

```
npx excel-graft types reports/incurred-cost.xlsx --name IncurredCostData --out src/incurred-cost.xlsx.ts
```

```ts
import type { IncurredCostData } from "./incurred-cost.xlsx";
const tpl = Template.loadSync<IncurredCostData>(bytes);
await tpl.render({ IndirectPools: [{ Account: "Fringe", rows: [...] }] }); // checked
```

Field types are inferred from the sample values and their number formats
(`Date` for date-formatted cells, `number`, `string`, `boolean`). At runtime,
a data key that matches nothing in the template throws with the list of
valid keys; `strict: true` also rejects unknown record fields.

## API

```ts
Template.load(bytes) / Template.loadSync(bytes)   // parse + analyze once; reuse for every request
tpl.inspect(): TemplateSchema                     // regions, inferred shapes, field types, scalars
tpl.render(data, options?): Promise<Uint8Array>   // also renderSync
verify(xlsxBytes): VerifyIssue[]                  // structural checks behind Excel's repair prompt
generateTypes(schema, typeName?): string          // TypeScript source
```

`RenderOptions`:

| Option | Default | |
|---|---|---|
| `onMissing` | `"blank"` | `"error"` throws when a record lacks a field or a placeholder has no value |
| `strict` | `false` | Throw on record properties that match no column (extra query columns are ignored by default) |
| `onWarning` | — | Notes about things that were dropped or became `#REF!` |
| `onReport` | — | Where each region landed: output rows, row kinds, group spans |

Cell values: `string`, `number`, `bigint`, `boolean`, `Date`, `null`/`undefined`
(empty cell, formatting kept) or `{ formula: "SUM(A1:A3)" }`. Strings are
always text — `"=1+1"` is never evaluated. Dates are converted from their UTC
fields (what the `mssql` driver returns by default); the 1904 date system is
honored. Control characters, text over Excel's 32,767-character limit, `NaN`
and `Infinity` are all made safe rather than producing a corrupt file.

Errors: `RenderDataError` (bad data: wrong key, wrong shape, row limit) and
`TemplateStructureError` (a region whose sample rows can't be interpreted —
those are listed in `inspect().skipped` instead of failing the load).

## How it's tested

The claim is "no repair prompts", so the test suite is built to catch them:

1. **A corpus of real Excel workbooks** (`corpus/templates/`) made in Excel —
   tables with totals, grouped regions, named ranges with merged titles,
   charts, images, pivot tables, slicers, sparklines, x14 conditional
   formatting, data validation, dynamic arrays, threaded comments, protected
   sheets, the 1904 date system, and a DCAA-style incurred cost submission.
   Every one is rendered with 0, 1, 7 and 300 rows and with hostile values
   (emoji, control characters, 40,000-character strings, `NaN`, bigints,
   1900 leap-bug dates, strings that look like formulas).
2. **Byte-identical passthrough**: every part a render didn't need to touch
   must be identical to the template's — asserted on every render.
3. **`verify()`**: structural checks for the known repair triggers (table
   filter ranges, row/cell order, types, limits, overlapping merges,
   relationships, content types, stale calc chains, element order in
   `workbook.xml` and worksheets, invalid or cell-like names, formula syntax
   and argument counts). It flags the exact bug in the original ExcelJS
   output, and every hand-corrupted file in `corpus/known-bad/`.
4. **Real Excel**: `npm run harness:rendered` opens every rendered file in
   Excel via COM, recalculates, and checks that every regenerated subtotal
   and total equals the sum of the data that was written, that table row
   counts match, and that no formula shows `#REF!` or `#NAME?`.
5. **Property tests** on the formula tokenizer and reference rewriting, and
   on text escaping.
6. **Scale**: 100,000 rows in 1,000 groups render in about a second.

## Try it: the test bench

```
npm run demo
```

Opens a local page with export buttons for every corpus template (from 10
to 100,000 rows), a **Verify** button, a **Check in Excel** button (Windows
+ Excel), a JSON editor for custom data, the generated TypeScript types, and
a drop zone to **test your own workbook** with generated data.

## Known limits

- Two filled regions that share rows (side-by-side tables) can't both grow;
  this throws instead of guessing.
- Merged cells that span several sample rows, and comments or hyperlinks on
  sample rows, are removed (with a warning) — they belong to the sample data.
- A grand total listing thousands of groups can exceed Excel's 8,192-character
  formula limit; that throws. Total with `SUBTOTAL` over one range instead.
- `{{placeholders}}` in text boxes, headers/footers and sheet names aren't
  filled.
- 3-D references (`=SUM(Jan:Feb!B7)`) across filled sheets can't follow
  each sheet's rows; they're left as-is with an `onWarning` message.
- A sample cell inside a non-aggregating function (`=ROUND(Pools!C9,2)`)
  means that one group; inside `SUM`/`AVERAGE`/`COUNT`/… listed together with
  its sibling sample groups (`=SUM(C5,C9,C13)`) it means all of them.
- Dynamic-array formulas whose spill runs into other content will show
  `#SPILL!` when the data grows — the same thing Excel would do.

## Reporting a problem

If Excel shows a repair prompt, or a number is wrong, please open an issue with:

1. `npx excel-graft verify output.xlsx`
2. `npx excel-graft inspect template.xlsx`
3. The template (with sample data scrubbed) if you can share it, and Excel's
   repair log if it offered one (in the "repaired" notification: *View*/*log*).

Every bug gets a corpus workbook that reproduces it before it's fixed.

## Development

```
npm install
npm test                 # unit, property, corpus and passthrough tests
npm run typecheck
npm run build
npm run harness:rendered # Windows + Excel: check the rendered outputs in real Excel
npm run demo
```

See [CLAUDE.md](CLAUDE.md) for the invariants the code must never violate
and [PROJECT.md](PROJECT.md) for the design.

MIT © Ryan Martin
