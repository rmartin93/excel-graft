# corpus-builder

These scripts generate the battle-test corpus by driving real Excel over COM.

- `Build-Corpus.ps1` creates each template in `corpus/templates/<name>/`
  from scratch (`template.xlsx` plus a `README.md` describing its regions,
  scalar targets, features and hostile points). Each README ends with a
  build log listing every optional feature as `OK:` or
  `SKIPPED: <reason>`.
- `Build-KnownBad.ps1` cuts the deliberately corrupt fixtures in
  `corpus/known-bad/` from those templates (see that folder's README).

## Built with

Excel **16.0, build 20326** (Microsoft 365, 64-bit) on Windows 11 Home
10.0.26200, 2026-09-28. All 10 templates built with **no skipped
features**: data bars, slicer, PivotTable, sparklines, `Formula2` dynamic
arrays and `AddCommentThreaded` all worked on this build.

## Run

```
# all templates (overwrites existing template.xlsx / README.md)
powershell -NoProfile -ExecutionPolicy Bypass -File tools/corpus-builder/Build-Corpus.ps1

# a subset
powershell -NoProfile -ExecutionPolicy Bypass -File tools/corpus-builder/Build-Corpus.ps1 -Only incurred-cost,date1904

# known-bad fixtures (run after Build-Corpus)
powershell -NoProfile -ExecutionPolicy Bypass -File tools/corpus-builder/Build-KnownBad.ps1

# verify
npm run harness:excel
powershell -NoProfile -ExecutionPolicy Bypass -File tools/excel-harness/Test-ExcelFiles.ps1 -Path corpus/known-bad
```

The whole corpus takes about 30 s to build. Excel runs invisible with
alerts off, and it is always quit and released in `finally`.

## Notes

- Rebuilding changes the bytes of every template, because Excel writes
  fresh `xr:uid` GUIDs, timestamps and so on. Templates are committed
  fixtures, so only rebuild when you mean to, and re-run any tests that
  pin their bytes.
- Keep both scripts pure ASCII. Windows PowerShell 5.1 reads BOM-less
  scripts in the ANSI code page.
- PowerShell's COM binder caches argument types per call site. After a
  site has set `Range.Value2` to a string, setting it to a number throws
  "Unable to cast Int32 to String". `Set-ComProp` works around this with
  `InvokeMember`.
- `ListColumn` has no `TotalsRowRange`. Use `Get-TotalCell` instead, which
  indexes into `ListObject.TotalsRowRange`.
- `FormatConditions.Add(xlExpression, ...)` resolves relative references
  against the active cell, so the builder selects the range's top-left cell
  first.
