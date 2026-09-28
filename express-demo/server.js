// excel-graft demo API: the same shape as a real Express + React export
// button — load templates once at startup, render per request, stream bytes.
import { execFile, exec } from "node:child_process";
import { randomUUID } from "node:crypto";
import { mkdirSync, readdirSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import express from "express";
import { generateTypes, RenderDataError, Template, verify } from "excel-graft";
import { fakeData } from "./fake-data.js";

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.join(here, "..");
const corpusDir = path.join(root, "corpus", "templates");
const workDir = path.join(tmpdir(), "excel-graft-demo");
mkdirSync(workDir, { recursive: true });
const PORT = Number(process.env.PORT ?? 3939);

/** id -> { id, title, file, bytes, template, schema, source } */
const templates = new Map();

function titleFor(dir, file) {
	const base = dir.replace(/[-_]+/g, " ").replace(/\b\w/g, (c) => c.toUpperCase());
	return /template\.xlsx$/i.test(file) ? base : `${base} — ${file.replace(/\.xlsx$/i, "")}`;
}

function register(id, title, bytes, source, file, regions = undefined) {
	const template = Template.loadSync(bytes, regions ? { regions } : {});
	const schema = template.inspect();
	const entry = { id, title, file, bytes, template, schema, source, readme: templates.get(id)?.readme, regions };
	templates.set(id, entry);
	return entry;
}

function loadCorpus() {
	for (const dir of readdirSync(corpusDir, { withFileTypes: true })) {
		if (!dir.isDirectory()) continue;
		const folder = path.join(corpusDir, dir.name);
		for (const f of readdirSync(folder)) {
			if (!f.toLowerCase().endsWith(".xlsx") || /output/i.test(f) || f.startsWith("~$")) continue;
			const id = `${dir.name}--${f.replace(/\.xlsx$/i, "")}`.toLowerCase().replace(/[^a-z0-9-]+/g, "-");
			try {
				const entry = register(id, titleFor(dir.name, f), readFileSync(path.join(folder, f)), "corpus", `${dir.name}/${f}`);
				try {
					entry.readme = readFileSync(path.join(folder, "README.md"), "utf8");
				} catch {}
			} catch (err) {
				console.warn(`skipping ${dir.name}/${f}: ${err.message}`);
			}
		}
	}
}
loadCorpus();

const app = express();
app.use(express.json({ limit: "100mb" }));

app.get("/", (_req, res) => res.sendFile(path.join(here, "public", "test.html")));

function summary(t) {
	return {
		id: t.id,
		title: t.title,
		file: t.file,
		source: t.source,
		schema: t.schema,
		readme: t.readme,
		types: generateTypes(t.schema, `${t.id.replace(/(^|-)(\w)/g, (_m, _d, c) => c.toUpperCase()).replace(/[^A-Za-z0-9]/g, "")}Data`),
		templateIssues: verify(t.bytes),
		regions: t.regions ?? {},
		bytes: t.bytes.length,
	};
}

app.get("/api/templates", (_req, res) => {
	res.json([...templates.values()].map(summary));
});

function getTemplate(req, res) {
	const t = templates.get(req.params.id);
	if (!t) res.status(404).json({ error: `no template ${req.params.id}` });
	return t;
}

function renderFor(t, data) {
	const warnings = [];
	let report;
	const started = performance.now();
	const bytes = t.template.renderSync(data, { onWarning: (w) => warnings.push(w), onReport: (r) => (report = r) });
	const ms = performance.now() - started;
	return { bytes, ms, warnings, report };
}

function sendXlsx(res, t, result, suffix) {
	const name = `${path.basename(t.file, ".xlsx")}${suffix ? `-${suffix}` : ""}.xlsx`;
	const rows = result.report?.regions.reduce((n, r) => n + r.rows.length, 0) ?? 0;
	res.set({
		"Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
		"Content-Disposition": `attachment; filename="${name}"`,
		"X-Render-Ms": result.ms.toFixed(1),
		"X-Rows": String(rows),
		"X-Warnings": encodeURIComponent(JSON.stringify(result.warnings)),
		"Access-Control-Expose-Headers": "X-Render-Ms, X-Rows, X-Warnings, Content-Disposition",
	});
	res.send(Buffer.from(result.bytes.buffer, result.bytes.byteOffset, result.bytes.byteLength));
}

function reviveDates(value) {
	if (typeof value === "string" && /^\d{4}-\d{2}-\d{2}(T[\d:.]+Z?)?$/.test(value)) return new Date(value.length === 10 ? `${value}T00:00:00Z` : value);
	if (Array.isArray(value)) return value.map(reviveDates);
	if (value && typeof value === "object" && !("formula" in value)) return Object.fromEntries(Object.entries(value).map(([k, v]) => [k, reviveDates(v)]));
	return value;
}

function handleError(res, err) {
	const status = err instanceof RenderDataError ? 400 : 500;
	res.status(status).json({ error: err.message, kind: err.name });
}

// The export button: GET /api/export/:id?rows=1000
app.get("/api/export/:id", (req, res) => {
	const t = getTemplate(req, res);
	if (!t) return;
	const rows = Math.min(Math.max(Number(req.query.rows ?? 25), 0), 500_000);
	try {
		sendXlsx(res, t, renderFor(t, fakeData(t.schema, rows, Number(req.query.seed ?? 42))), `${rows}-rows`);
	} catch (err) {
		handleError(res, err);
	}
});

// Custom data from the page's JSON editor (ISO date strings become Dates).
app.post("/api/export/:id", (req, res) => {
	const t = getTemplate(req, res);
	if (!t) return;
	try {
		sendXlsx(res, t, renderFor(t, reviveDates(req.body)), "custom");
	} catch (err) {
		handleError(res, err);
	}
});

app.get("/api/sample/:id", (req, res) => {
	const t = getTemplate(req, res);
	if (t) res.json(fakeData(t.schema, Number(req.query.rows ?? 4)));
});

app.get("/api/template/:id", (req, res) => {
	const t = getTemplate(req, res);
	if (!t) return;
	res.set({ "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Content-Disposition": `attachment; filename="${path.basename(t.file)}"` });
	res.send(Buffer.from(t.bytes));
});

// Structural verification of a render (the checks behind Excel's repair prompt).
app.get("/api/verify/:id", (req, res) => {
	const t = getTemplate(req, res);
	if (!t) return;
	try {
		const rows = Math.min(Number(req.query.rows ?? 25), 500_000);
		const result = renderFor(t, fakeData(t.schema, rows));
		const started = performance.now();
		const issues = verify(result.bytes);
		res.json({ issues, renderMs: result.ms, verifyMs: performance.now() - started, bytes: result.bytes.length, warnings: result.warnings, regions: result.report?.regions.map((r) => ({ key: r.key, sheet: r.sheet, firstRow: r.firstRow, lastRow: r.lastRow })) ?? [] });
	} catch (err) {
		handleError(res, err);
	}
});

// Ground truth: open the render in real Excel (Windows + Excel only).
app.get("/api/excel-check/:id", (req, res) => {
	const t = getTemplate(req, res);
	if (!t) return;
	if (process.platform !== "win32") return res.status(501).json({ error: "Excel checks need Windows with Excel installed." });
	let result;
	try {
		result = renderFor(t, fakeData(t.schema, Math.min(Number(req.query.rows ?? 25), 200_000)));
	} catch (err) {
		return handleError(res, err);
	}
	const run = randomUUID();
	const dir = path.join(workDir, run);
	mkdirSync(dir, { recursive: true });
	const file = `${t.id}.xlsx`;
	writeFileSync(path.join(dir, file), result.bytes);
	const tables = t.schema.regions
		.filter((r) => r.kind === "table")
		.map((r) => ({ name: r.key, dataRows: result.report?.regions.find((x) => x.key === r.key)?.rows.length }))
		.filter((x) => x.dataRows !== undefined);
	const probes = (result.report?.regions ?? []).flatMap((r) =>
		r.rows.filter((x) => x.kind === "fixed").map((x) => ({ sheet: r.sheet, row: x.row, label: `${r.key} total row` })),
	);
	writeFileSync(path.join(dir, "demo.manifest.json"), JSON.stringify([{ file, cells: [], tables, probes }]));
	const script = path.join(root, "tools", "excel-harness", "Test-Rendered.ps1");
	execFile("powershell", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script, "-Dir", dir], { timeout: 180_000 }, (err, stdout) => {
		let results = [];
		try {
			results = JSON.parse(readFileSync(path.join(dir, "excel-results.json"), "utf8").replace(/^﻿/, ""));
		} catch {}
		const r = Array.isArray(results) ? results[0] : results;
		res.json({ ok: !err && r?.ok === true, result: r, log: stdout, renderMs: result.ms });
	});
});

// Declare regions in code (no workbook edits): body { regions: { Key: "'Sheet'!A7:M207" | { range, layout } } }.
app.post("/api/regions/:id", (req, res) => {
	const t = getTemplate(req, res);
	if (!t) return;
	try {
		const regions = req.body?.regions && Object.keys(req.body.regions).length ? req.body.regions : undefined;
		res.json(summary(register(t.id, t.title, t.bytes, t.source, t.file, regions)));
	} catch (err) {
		res.status(400).json({ error: err.message });
	}
});

// Try your own workbook: POST the raw .xlsx bytes.
app.post("/api/upload", express.raw({ type: "*/*", limit: "100mb" }), (req, res) => {
	try {
		const name = String(req.query.name ?? "upload.xlsx").replace(/[^\w .()-]/g, "_");
		const id = `upload-${randomUUID().slice(0, 8)}`;
		const entry = register(id, name.replace(/\.xlsx$/i, ""), new Uint8Array(req.body), "upload", name);
		res.json(summary(entry));
	} catch (err) {
		res.status(400).json({ error: `Couldn't read that workbook: ${err.message}` });
	}
});

app.listen(PORT, "127.0.0.1", () => {
	const url = `http://localhost:${PORT}`;
	console.log(`excel-graft demo: ${url}  (${templates.size} templates loaded)`);
	if (process.argv.includes("--open")) {
		const cmd = process.platform === "win32" ? `start "" "${url}"` : process.platform === "darwin" ? `open "${url}"` : `xdg-open "${url}"`;
		exec(cmd);
	}
});
