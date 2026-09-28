import { afterAll, describe, expect, it } from "vitest";
import { type RenderReport, Template, verify } from "../src/index.js";
import { badName, formulaProblems } from "../src/verify-rules.js";
import { readParts, writeParts } from "../src/zip.js";
import { ExcelManifest } from "./excel-checks.js";
import { corpusFile, expectPassthrough, partText, renderablePart } from "./helpers.js";

/**
 * GitHub issue #1: a disclosure-statement workbook whose hierarchy (yellow
 * level-1 rows, blue level-2 rows, details) is shown only by styling, with
 * no Table or names; and an export Excel repaired that verify() passed.
 */
const TEMPLATE = corpusFile("issue-1-disclosure", "template.xlsx");
const REGIONS = { DisclosureData: "'Disclosure Table'!A7:M207", CostCenterData: "'2026 CC'!A6:H66" };
const manifest = new ExcelManifest("issue-1");
afterAll(() => manifest.save());

const detail = (base: string, cost: number) => ({ Segment: "Seg 1", "Base Code": base, "Disclosure Section": 8.1, "Cost Out": cost });
const DATA = {
	DisclosureData: [
		{ "Pool Reference": "Fringe", rows: [{ "Disclosure Reference": "Payroll Taxes", rows: [detail("B101", 100), detail("B102", 200)] }] },
		{
			"Pool Reference": "G&A",
			rows: [
				{ "Disclosure Reference": "Executive", rows: [detail("B101", 10), detail("B103", 20), detail("B104", 30)] },
				{ "Disclosure Reference": "Legal", rows: [detail("B102", 5)] },
			],
		},
	],
	CostCenterData: [
		{ "Cost Center": "CC-1", Segment: "Seg 1", "Base Code": "B101", Budget: 100, Actual: 120, Owner: "A" },
		{ "Cost Center": "CC-2", Segment: "Seg 1", "Base Code": "B102", Budget: 200, Actual: 150, Owner: "B" },
	],
};
// DisclosureData output: Fringe 7, Payroll Taxes 8, details 9-10; G&A 11, Executive 12, 13-15, Legal 16, 17; gap 18-20; notes 21-23.

describe("issue #1: style-grouped named range", () => {
	it("a workbook with no Tables or names has no regions until they're declared", () => {
		expect(Template.loadSync(TEMPLATE).inspect().regions).toEqual([]);
		const schema = Template.loadSync(TEMPLATE, { regions: REGIONS }).inspect();
		const d = schema.regions.find((r) => r.key === "DisclosureData")!;
		expect(d).toMatchObject({ kind: "name", structure: "styles", sampleRange: "A7:M207" });
		expect(d.fixedRows).toEqual(["A23:M23", "A24:M24", "A25:M25", "A26:M26", "A27:M27", "A28:M28"]);
		expect(d.shape).toMatchObject({
			kind: "groups",
			fields: [{ name: "Pool Reference", samples: ["Fringe", "Overhead"], labelFormat: "{value} Pool" }],
			child: { kind: "groups", fields: [{ name: "Disclosure Reference", samples: ["Payroll Taxes", "Group Insurance", "Facilities", "Indirect Labor"] }], child: { kind: "rows" } },
		});
		const cc = schema.regions.find((r) => r.key === "CostCenterData")!;
		expect(cc).toMatchObject({ structure: "styles", fixedRows: ["A10:H10", "A11:H11"], shape: { kind: "rows" } });
	});

	it("renders every level with its own row style, keeps the notes and picture below, and updates sheet 2", () => {
		let report: RenderReport | undefined;
		const out = Template.loadSync(TEMPLATE, { regions: REGIONS }).renderSync(DATA, { onReport: (r) => (report = r) });
		expect(verify(out)).toEqual([]);
		expectPassthrough(TEMPLATE, out, renderablePart);
		const rows = report!.regions.find((r) => r.key === "DisclosureData")!.rows;
		expect(rows.map((r) => `${r.row}:${r.kind}`).join(" ")).toBe(
			"7:header 8:header 9:leaf 10:leaf 11:header 12:header 13:leaf 14:leaf 15:leaf 16:header 17:leaf 18:fixed 19:fixed 20:fixed 21:fixed 22:fixed 23:fixed",
		);
		// Each output row copies the template row of its own level (not always row 7).
		expect(new Set(rows.filter((r) => r.kind === "header" && r.groupLast - r.groupFirst > 3).map((r) => r.templateRow))).toEqual(new Set([7]));
		expect(rows.find((r) => r.row === 8)?.templateRow).toBe(8);
		expect(rows.find((r) => r.row === 9)?.templateRow).toBe(9);
		const sheet = partText(out, "xl/worksheets/sheet1.xml");
		const style = (cell: string) => new RegExp(`<c r="${cell}" s="(\\d+)"`).exec(sheet)?.[1];
		expect(style("A7")).toBe(style("A11")); // level 1 (yellow)
		expect(style("A8")).toBe(style("A16")); // level 2 (blue)
		expect(style("A9")).toBe(style("A17")); // details
		expect(new Set([style("A7"), style("A8"), style("A9")]).size).toBe(3);
		expect(sheet).toContain("<t>Fringe Pool</t>");
		expect(sheet).toContain("<t>G&amp;A Pool</t>");
		expect(sheet).toMatch(/<c r="A21"[^>]*t="s"/); // "Notes:" kept once, right after the gap
		// The signature picture sat in the cleared capacity rows; it stays 4 rows below the notes heading.
		const drawing = [...readParts(out).keys()].find((k) => /drawings\/drawing\d+\.xml$/.test(k))!;
		expect(partText(out, drawing)).toContain("<xdr:row>24</xdr:row>");
		const cc = partText(out, "xl/worksheets/sheet2.xml");
		expect(cc).toContain("<f>SUMIF('Disclosure Table'!$D:$D,C6,'Disclosure Table'!$K:$K)</f>");
		expect(cc).toContain("<f>SUM('Disclosure Table'!K7:K17)</f>");
		manifest.add("issue-1/styles.xlsx", out, {
			cells: [
				{ sheet: "2026 CC", cell: "F6", value: 110, note: "SUMIF B101 over the disclosure data" },
				{ sheet: "2026 CC", cell: "F7", value: 205, note: "SUMIF B102" },
				{ sheet: "2026 CC", cell: "F9", value: 365, note: "disclosure total" },
				{ sheet: "Disclosure Table", cell: "A7", value: "Fringe Pool", note: "level-1 label" },
			],
			tables: [],
		});
	});

	it("a declared layout gives the same result without relying on styles", () => {
		const tpl = Template.loadSync(TEMPLATE, {
			regions: {
				...REGIONS,
				DisclosureData: { range: REGIONS.DisclosureData, layout: { levels: [{ header: 7 }, { header: 8 }], detail: 9, fixedRows: [26, 27, 28] } },
			},
		});
		const d = tpl.inspect().regions.find((r) => r.key === "DisclosureData")!;
		expect(d).toMatchObject({ structure: "layout", fixedRows: ["A26:M26", "A27:M27", "A28:M28"] });
		expect(d.shape).toMatchObject({ kind: "groups", child: { kind: "groups", child: { kind: "rows" } } });
		const out = tpl.renderSync(DATA);
		expect(verify(out)).toEqual([]);
		expect(partText(out, "xl/worksheets/sheet1.xml")).toContain("<t>Executive</t>");
		expect(() => Template.loadSync(TEMPLATE, { regions: { X: { range: REGIONS.DisclosureData, layout: { levels: [{ header: 9 }], detail: 8 } } } }).renderSync({ X: [] })).toThrow();
	});

	it("banded rows are still banding, not groups", () => {
		const r = Template.loadSync(corpusFile("named-range-region", "template.xlsx")).inspect().regions[0]!;
		expect(r.structure).toBe("flat");
		expect(r.shape.kind).toBe("rows");
	});
});

describe("issue #1: verify() catches what Excel repairs", () => {
	const withNames = (where: "end" | "after-sheets") => {
		const parts = readParts(TEMPLATE);
		const names = `<definedNames><definedName name="DisclosureData">'Disclosure Table'!$A$7:$M$207</definedName></definedNames>`;
		let wb = new TextDecoder().decode(parts.get("xl/workbook.xml"));
		wb = where === "end" ? wb.replace("</workbook>", `${names}</workbook>`) : wb.replace("</sheets>", `</sheets>${names}`);
		parts.set("xl/workbook.xml", new TextEncoder().encode(wb));
		return writeParts(parts);
	};

	it("flags <definedNames> placed after <calcPr> (Excel removes the name and repairs; confirmed in Excel 16)", () => {
		expect(verify(withNames("end")).map((i) => i.message)).toEqual([expect.stringMatching(/<definedNames> must come before/)]);
		expect(verify(withNames("after-sheets"))).toEqual([]);
	});

	it("flags names that look like cell references and invalid formulas", () => {
		expect(badName("CC2026", "defined")).toMatch(/cell reference/);
		expect(badName("T1", "table")).toMatch(/cell reference/);
		expect(badName("R1C1", "defined")).toMatch(/cell reference/);
		expect(badName("Disclosure Data", "defined")).toMatch(/characters/);
		expect(badName("DisclosureData", "defined")).toBeUndefined();
		expect(badName("_xlnm.Print_Area", "defined")).toBeUndefined();
		expect(formulaProblems("ROUND(C3,C5,C8,2)")).toEqual([expect.stringMatching(/4 argument\(s\) to ROUND, which takes 2/)]);
		expect(formulaProblems("SUM(A1")).toEqual([expect.stringMatching(/unclosed \(/)]);
		expect(formulaProblems('"abc')).toEqual([expect.stringMatching(/unclosed string/)]);
		expect(formulaProblems("IF(A1,,B1)")).toEqual([]);
		expect(formulaProblems("TODAY()+ROW()")).toEqual([]);
		expect(formulaProblems("SUM((C5,C8))")).toEqual([]);
	});
});
