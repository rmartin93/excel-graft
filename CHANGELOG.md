# Changelog

## 0.2.0

Fixes [#1](https://github.com/rmartin93/excel-graft/issues/1).

- **Declared regions**: `Template.load(bytes, { regions: { Key: "'Sheet'!A7:M207" } })`
  (and `--region` on the CLI) for workbooks without Tables or named ranges.
  Nothing is written into the workbook.
- **Explicit layouts**: `{ range, layout: { levels, detail, fixedRows } }` when the
  row roles can't be inferred.
- **Style-driven groups**: hierarchies shown only by row styling (label rows
  styled apart from the details) are inferred as nested groups; distinct rows
  below the data (notes, total lines) are kept once. `inspect()` reports
  `structure`.
- Blank rows at the edges of a region are cleared instead of repeated; pictures
  anchored in them move with the content above.
- **verify()** now also checks element order in `workbook.xml` and worksheets
  (a misplaced `<definedNames>` makes Excel delete the names and repair),
  defined/table name validity (including names that look like cell
  references), names scoped to missing sheets or pointing at missing sheets,
  and formula syntax (unbalanced brackets/quotes, >255 arguments, >64 levels,
  argument counts of common functions).
- Test bench: a **Regions** button to declare regions for any workbook.

## 0.1.0

First release.
