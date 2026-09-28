# excel-graft

A lossless Excel template engine. Treats `.xlsx` as a zip of XML parts,
copies every part byte-for-byte, and patches only the XML it has to touch.
Built to replace an ExcelJS-based render pipeline that produced files Excel
flagged for repair on open — see [PROJECT.md](PROJECT.md) for the full
story and design, [CLAUDE.md](CLAUDE.md) for the invariants this codebase
must never violate.

Status: **Phase 0** (repo scaffolding, corpus, test harnesses). Nothing
renders yet — `Template.render()` currently throws. See the phase table in
PROJECT.md.

## Quick start

```
npm install
npm run build
npm test
npm run typecheck
npm run harness:excel   # Windows + Excel installed only; opens the corpus in real Excel via COM
```

## Layout

- `src/` — the engine
- `test/` — Vitest, including the passthrough-invariant test (CLAUDE.md rule 2)
- `corpus/` — real template fixtures and Excel repair logs (CLAUDE.md rule 6)
- `tools/excel-harness/` — PowerShell + Excel COM smoke test (ground truth; see its script header for what it does and does not catch)
- `tools/validator/` — Open XML SDK structural validator (not built yet — needs the .NET SDK)
- `express-demo/` — Phase 5 demo API + `test.html` (not started)

## Prior art

`xlsx-template` (npm) uses the same patch-don't-rebuild idea in older
JavaScript.
