# Open XML SDK validator (not built yet)

## Status

Blocked on the .NET SDK — not currently installed on this machine
(`dotnet --version` fails with "No .NET SDKs were found").

## What it needs to be

A small .NET console tool, checked into this folder, that:

1. Takes a path to an `.xlsx` (or a directory of them) as an argument.
2. Loads it with the [Open XML SDK](https://github.com/dotnet/Open-XML-SDK)
   (`DocumentFormat.OpenXml`).
3. Runs `OpenXmlValidator` against every part for the current Office
   version (`FileFormatVersions.Microsoft365`).
4. Prints every validation error with its part URI and exits non-zero if
   there are any.

This is the structural-validity check in CI — it catches schema violations
(malformed `sqref`, bad relationship IDs, orphaned parts) that our own code
doesn't otherwise check for. It will NOT catch the class of corruption
found in `corpus/templates/exceljs-fork-broken/Test-Output.xlsx` on
2026-09-28 (subtle `extLst`/x14-level loss that's schema-valid but wrong)
— see `tools/excel-harness/Test-ExcelFiles.ps1` for what that harness does
and does not catch, and PROJECT.md's battle-testing strategy for how the
layers fit together.

## To build it (next session, once the .NET SDK is installed)

```
winget install Microsoft.DotNet.SDK.8
cd tools/validator
dotnet new console -n XlsxValidator
dotnet add XlsxValidator package DocumentFormat.OpenXml
```

Then implement the loop described above in `Program.cs`, and wire
`npm run validate` in the root `package.json` to shell out to
`dotnet run --project tools/validator/XlsxValidator`.
