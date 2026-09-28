# CLAUDE.md

Instructions for agents (Claude Code or otherwise) working in this repo.
See [PROJECT.md](PROJECT.md) for why this project exists.

## The invariants (never violate these)

1. **Never rebuild the workbook.** `.xlsx` is a zip of XML parts. Every part
   the engine did not need to touch must survive byte-identical in the
   output. Do not run anything through a full parse-model-serialize cycle
   for parts we don't understand.
2. **The passthrough invariant is a test, not a hope.** Every test that
   renders a template must assert that untouched parts (everything except
   the sheet(s) actually rendered and their direct dependents) are
   byte-identical between input and output zip. See `test/passthrough.test.ts`.
3. **Write new strings as `t="inlineStr"`.** Never add entries to
   `sharedStrings.xml`. This avoids ever needing to rewrite the shared
   strings table.
4. **Delete `calcChain.xml`** (plus its `.rels` entry and its
   `[Content_Types].xml` override) whenever any formula cell could have
   changed. A stale calc chain is one of the most common causes of the
   Excel repair prompt.
5. **Strip cached `<v>` values on formula cells** that sit inside or below
   a rendered/grown region, and set `<calcPr fullCalcOnLoad="1"/>` in
   `workbook.xml`, so Excel recalculates on open instead of trusting stale
   cached results.
6. **Every bug gets a corpus template first.** If you find a corruption or
   a repair prompt, the fix does not start with the code — it starts with
   a new template under `corpus/templates/` (or a repair log under
   `corpus/repair-logs/`) that reproduces it, committed before the fix.
7. **Row insertion touches more than `sheetN.xml`.** Check the full list in
   PROJECT.md's "What must be patched" table before considering a
   row-insertion change complete: dimension, mergeCells, conditional
   formatting sqref, data validations, hyperlinks, autoFilter, table refs,
   defined names, drawing anchors, chart series ranges.
8. **All A1 reference rewriting goes through `src/formula.ts` (tokenizer)
   and `src/layout.ts` (`WorkbookMapper`).** No ad hoc regex shifting of
   formulas anywhere else. Any change to either file needs property tests in
   `test/formula.test.ts` and a green `npm run harness:rendered` — the
   Excel harness is what caught the double-counted grand total.
9. **A reference must never silently point at the wrong data.** When a
   reference to a specific sample row has no counterpart in the rendered data
   (e.g. a rates sheet pointing at a pool the data doesn't have), it becomes
   `#REF!` with an `onWarning` message — never a neighbouring row.

## Working method

- **Test-first, one phase at a time.** Follow the phase table in
  PROJECT.md. Do not start Phase N+1 work while Phase N's exit criteria are
  unmet.
- **A change is not done** until:
  - `npm test` passes,
  - the Open XML SDK validator passes on every corpus output (once
    `tools/validator` exists — see its README for current status),
  - `npm run harness:rendered` passes on every rendered output (it opens
    them in Excel via COM, recalculates, and checks computed totals, table
    row counts and broken references — it is the ground truth, prefer it
    over your own judgment about whether a file is "probably fine"). Note
    that some Excel builds silently accept files others repair, so
    `verify()` must pass too.
- **Review test changes yourself.** Agents are good at the reference-shifting
  grind but will happily loosen or delete a failing test to make it pass.
  Never let an agent "fix" a test without a human reading the diff — if
  you're an agent reading this, flag test changes explicitly in your
  summary rather than folding them silently into a larger diff.
- **XML handling:** use `saxes` (a tokenizer) for the parts we actually
  parse, never a full-DOM parser that normalizes namespaces, reorders
  attributes, or touches `mc:Ignorable`. For every other part, treat it as
  an opaque byte blob and copy it through.
- **No dependencies beyond the stack in PROJECT.md** without discussing it
  first — the entire point of this project is to stay small and legible
  compared to ExcelJS.
