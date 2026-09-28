import { readFileSync } from "node:fs";
import { afterAll, describe, expect, it } from "vitest";
import { fakeData } from "../express-demo/fake-data.js";
import { type RenderReport, type ShapeSchema, Template, type TemplateSchema, verify } from "../src/index.js";
import { ExcelManifest, sumChecks } from "./excel-checks.js";
import { corpusTemplates, expectPassthrough, HOSTILE_DATES, HOSTILE_NUMBERS, HOSTILE_STRINGS, renderablePart } from "./helpers.js";

/**
 * The battle test: every workbook under corpus/templates/ is rendered at
 * several sizes and with hostile values, and every output must pass
 * verify(), keep untouched parts byte-identical, and (via the Excel
 * harness) compute its regenerated totals correctly in real Excel.
 * Drop a new .xlsx into corpus/templates/<name>/ and it's covered.
 */
const manifest = new ExcelManifest("corpus");
afterAll(() => manifest.save());

const SIZES = [0, 1, 7, 300];

function hostileShape(shape: ShapeSchema, n: number, seed: number): unknown[] {
	const pickAny = (i: number) => {
		const pools = [HOSTILE_STRINGS, HOSTILE_NUMBERS, HOSTILE_DATES, [null, undefined, true, false]] as unknown[][];
		const pool = pools[(i + seed) % pools.length] as unknown[];
		return pool[(i * 7 + seed) % pool.length];
	};
	return Array.from({ length: n }, (_, i) => {
		const rec: Record<string, unknown> = {};
		shape.fields.forEach((f, j) => (rec[f.name] = pickAny(i * 31 + j)));
		if (shape.kind === "groups") rec[shape.childKey] = hostileShape(shape.child, (i % 3) + 1, seed + i);
		return rec;
	});
}

function hostileData(schema: TemplateSchema): Record<string, unknown> {
	const data: Record<string, unknown> = {};
	schema.regions.forEach((r, i) => (data[r.key] = hostileShape(r.shape, 12, i)));
	for (const s of schema.scalars) {
		if (s.key.includes(".")) continue;
		data[s.key] = HOSTILE_STRINGS[s.key.length % HOSTILE_STRINGS.length];
	}
	return data;
}

function tableRows(report: RenderReport, schema: TemplateSchema) {
	return report.regions
		.filter((r) => schema.regions.find((x) => x.key === r.key)?.kind === "table")
		.map((r) => ({ name: r.key, dataRows: r.rows.length }));
}

const templates = corpusTemplates();

describe.each(templates)("$name", ({ name, file }) => {
	const input = new Uint8Array(readFileSync(file));
	const slug = name.replace(/\.xlsx$/i, "").replace(/[^A-Za-z0-9-]+/g, "-");

	it("the template itself is structurally clean and round-trips byte-for-byte", async () => {
		expect(verify(input)).toEqual([]);
		const tpl = await Template.load(input);
		expectPassthrough(input, await tpl.toBuffer(), () => false);
		expectPassthrough(input, await tpl.render({}), () => false);
	});

	it.each(SIZES)("renders %i rows per region cleanly", async (rows) => {
		const tpl = await Template.load(input);
		const schema = tpl.inspect();
		const data = fakeData(schema, rows, rows + 1);
		let report: RenderReport | undefined;
		const warnings: string[] = [];
		const out = await tpl.render(data, { onReport: (r) => (report = r), onWarning: (w) => warnings.push(w) });
		expect(verify(out)).toEqual([]);
		expectPassthrough(input, out, renderablePart);
		if (!report) throw new Error("no report");
		const cells = sumChecks(schema, report, out);
		manifest.add(`corpus/${slug}-${rows}.xlsx`, out, { cells, tables: tableRows(report, schema), allowBrokenRefs: warnings.some((w) => w.includes("#REF!")) });
	});

	it("renders hostile values cleanly", async () => {
		const tpl = await Template.load(input);
		const schema = tpl.inspect();
		let report: RenderReport | undefined;
		const warnings: string[] = [];
		const out = await tpl.render(hostileData(schema), { onReport: (r) => (report = r), onWarning: (w) => warnings.push(w) });
		expect(verify(out)).toEqual([]);
		expectPassthrough(input, out, renderablePart);
		manifest.add(`corpus/${slug}-hostile.xlsx`, out, { cells: [], tables: tableRows(report as RenderReport, schema), allowBrokenRefs: warnings.some((w) => w.includes("#REF!")) });
	});
});
