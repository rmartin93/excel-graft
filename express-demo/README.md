# excel-graft test bench

A small Express API plus `public/test.html`, shaped like a real export
button: templates are loaded once at startup, each request renders and
streams the bytes back.

```
npm run demo        # from the repo root: builds the library, installs, starts, opens the browser
```

Then, for each workbook in `corpus/templates/` (and any you drop onto the page):

- **Export**: renders generated data (0 to 100,000 rows per region) and downloads the `.xlsx`.
- **Verify**: renders, then runs `verify()`, the structural checks behind Excel's repair prompt.
- **Check in Excel**: renders, then opens the file in real Excel (Windows + Excel only),
  recalculates, and reports broken references, table row counts and the computed totals.
- **Custom data**: edit the JSON passed to `render()` (ISO date strings become Dates).
- **Types**: the TypeScript type generated from the template.

Endpoints (all under `/api`): `GET templates`, `GET export/:id?rows=N`,
`POST export/:id` (JSON body), `GET verify/:id?rows=N`, `GET excel-check/:id?rows=N`,
`GET sample/:id`, `GET template/:id`, `POST upload?name=file.xlsx` (raw bytes).

The server binds to `127.0.0.1:3939` (`PORT` to change) and keeps uploads in memory only.
