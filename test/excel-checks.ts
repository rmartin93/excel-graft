import { writeFileSync } from "node:fs";
import path from "node:path";
import { colToIndex, indexToCol } from "../src/a1.js";
import type { RenderReport, TemplateSchema } from "../src/index.js";
import { readParts } from "../src/zip.js";
import { parseXml, findAll, textOf, decodeXml } from "../src/xml.js";
import { GENERATED, writeGenerated } from "./helpers.js";

/**
 * Expectations the Excel harness (tools/excel-harness/Test-Rendered.ps1)
 * checks after opening a rendered file in real Excel and recalculating:
 * computed cell values, table row counts, and no broken references.
 */
export interface ExcelCheck {
	file: string;
	cells: { sheet: string; cell: string; value: number | string; note: string }[];
	tables: { name: string; dataRows: number }[];
}

export class ExcelManifest {
	private readonly entries: ExcelCheck[] = [];

	constructor(private readonly name: string) {}

	add(file: string, bytes: Uint8Array, check: Omit<ExcelCheck, "file">): void {
		writeGenerated(file, bytes);
		this.entries.push({ file, ...check });
	}

	save(): void {
		writeFileSync(path.join(GENERATED, `${this.name}.manifest.json`), JSON.stringify(this.entries, null, 1));
	}
}

const decoder = new TextDecoder();

/**
 * Derives the values Excel must compute for every regenerated SUM total:
 * a group row's SUM over column X must equal the sum of X over the
 * group's detail records, a fixed total row's over all of them, and a
 * table totals row with totalsRowFunction="sum" likewise.
 */
export function sumChecks(schema: TemplateSchema, report: RenderReport, output: Uint8Array): ExcelCheck["cells"] {
	const parts = readParts(output);
	const checks: ExcelCheck["cells"] = [];
	const sheetParts = sheetPartsByName(parts);
	for (const region of report.regions) {
		const rs = schema.regions.find((r) => r.key === region.key);
		const sheetPart = sheetParts.get(region.sheet);
		if (!rs || !sheetPart) continue;
		const xml = decoder.decode(parts.get(sheetPart) as Uint8Array);
		const formulas = cellFormulas(xml);
		const leafRows = region.rows.filter((r) => r.kind === "leaf");
		const sumOver = (field: string, first: number, last: number) => {
			let total = 0;
			let any = false;
			for (const r of leafRows) {
				if (r.row < first || r.row > last) continue;
				const v = r.record?.[field];
				if (typeof v === "number" && Number.isFinite(v)) {
					total += v;
					any = true;
				} else if (typeof v === "bigint") {
					total += Number(v);
					any = true;
				} else if (v !== undefined && v !== null) return undefined; // non-numeric: skip the check
			}
			return any || leafRows.length > 0 ? total : undefined;
		};
		for (const r of region.rows) {
			if (r.kind === "leaf" || r.kind === "spacer") continue;
			for (const [col, field] of Object.entries(rs.columns)) {
				const f = formulas.get(`${col}${r.row}`);
				if (!f || !/^SUM\([^()]*\)$/i.test(f.replace(/\$/g, ""))) continue;
				if (!leafRows.some((l) => l.record && field in l.record)) continue;
				const expected = sumOver(field, r.kind === "fixed" ? region.firstRow : r.groupFirst, r.kind === "fixed" ? region.lastRow : r.groupLast);
				if (expected === undefined) continue;
				checks.push({ sheet: region.sheet, cell: `${col}${r.row}`, value: round(expected), note: `${r.kind} ${f}` });
			}
		}
		// Table totals row.
		if (rs.kind === "table") {
			const totalsRow = region.lastRow + 1;
			for (const [col, field] of Object.entries(rs.columns)) {
				const f = formulas.get(`${col}${totalsRow}`);
				if (!f || !/^SUBTOTAL\((109|9),/i.test(f)) continue;
				if (!leafRows.some((l) => l.record && field in l.record)) continue;
				const expected = sumOver(field, region.firstRow, region.lastRow);
				if (expected !== undefined) checks.push({ sheet: region.sheet, cell: `${col}${totalsRow}`, value: round(expected), note: `totals ${f}` });
			}
		}
	}
	return checks;
}

function round(n: number): number {
	return Math.round(n * 1e6) / 1e6;
}

function sheetPartsByName(parts: Map<string, Uint8Array>): Map<string, string> {
	const wb = decoder.decode(parts.get("xl/workbook.xml") as Uint8Array);
	const rels = decoder.decode(parts.get("xl/_rels/workbook.xml.rels") as Uint8Array);
	const targets = new Map<string, string>();
	for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
		const id = /Id="([^"]+)"/.exec(m[0])?.[1];
		const target = /Target="([^"]+)"/.exec(m[0])?.[1];
		if (id && target) targets.set(id, target.replace(/^\/?(xl\/)?/, "xl/"));
	}
	const out = new Map<string, string>();
	for (const m of wb.matchAll(/<sheet\b[^>]*>/g)) {
		const name = /name="([^"]+)"/.exec(m[0])?.[1];
		const id = /r:id="([^"]+)"/.exec(m[0])?.[1];
		if (name && id && targets.has(id)) out.set(decodeXml(name), targets.get(id) as string);
	}
	return out;
}

function cellFormulas(xml: string): Map<string, string> {
	const root = parseXml(xml);
	const out = new Map<string, string>();
	for (const c of findAll(root, "c")) {
		const f = c.children.find((x) => x.name === "f");
		if (f && c.attrs.r) out.set(c.attrs.r, textOf(xml, f));
	}
	return out;
}

export { colToIndex, indexToCol };
