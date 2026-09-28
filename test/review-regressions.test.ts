import { afterAll, describe, expect, it } from "vitest";
import { Template, verify } from "../src/index.js";
import { mkxlsx } from "./build-xlsx.js";
import { ExcelManifest } from "./excel-checks.js";
import { corpusFile, partText } from "./helpers.js";

/**
 * Regressions from the adversarial review (2026-09-28). Each case is the
 * reviewer's minimal reproduction; the Excel harness re-checks the values
 * it proved wrong.
 */
const manifest = new ExcelManifest("review");
afterAll(() => manifest.save());

function render(tpl: Uint8Array, data: Record<string, unknown>) {
	const warnings: string[] = [];
	const out = Template.loadSync(tpl).renderSync(data, { onWarning: (w) => warnings.push(w) });
	return { out, warnings, sheet: (n = 1) => partText(out, `xl/worksheets/sheet${n}.xml`) };
}

const f = (xml: string, cell: string) => new RegExp(`<c r="${cell}"[^>]*><f[^>]*>([^<]*)</f>`).exec(xml)?.[1]?.replace(/&gt;/g, ">").replace(/&amp;/g, "&");

// Pools: Fringe (3; 4-5), Overhead (6; 7-8), Total 9 — rendered as G&A, Fringe, Overhead.
const poolRows = (extra: Record<number, Record<string, string>>) => ({
	2: { A: "Pool", B: "Account", C: "Amount" },
	3: { A: "Fringe", C: "=SUM(C4:C5)" },
	4: { B: "f1", C: 10 },
	5: { B: "f2", C: 20 },
	6: { A: "Overhead", C: "=SUM(C7:C8)" },
	7: { B: "o1", C: 100 },
	8: { B: "o2", C: 200 },
	9: { A: "Total", C: "=SUM(C3,C6)" },
	...extra,
});
const pools = {
	Pools: [
		{ Pool: "G&A", rows: [{ Account: "g", Amount: 5 }] },
		{ Pool: "Fringe", rows: [{ Account: "x", Amount: 1 }, { Account: "y", Amount: 2 }] },
		{ Pool: "Overhead", rows: [{ Account: "x", Amount: 1000 }, { Account: "y", Amount: 2000 }] },
	],
};
// Output: G&A 3 (4), Fringe 5 (6-7), Overhead 8 (9-10), Total 11; rows below shift by 2.

describe("review regressions", () => {
	it("#1 whole-column references keep their columns when a shared formula is rewritten", () => {
		const sh = (ref: string | null, text?: string) => ({ raw: text ? `<f t="shared" ref="${ref}" si="0">${text}</f>` : `<f t="shared" si="0"/>` });
		const tpl = mkxlsx([
			{
				name: "Sheet1",
				rows: { 2: { A: "Cat", B: "Amount" }, 3: { A: "x", B: 1 }, 4: { A: "y", B: 2 }, 5: { A: "x", B: 3 }, 8: { D: "x", E: sh("E8:E9", "SUMIF(A:A,D8,B:B)") }, 9: { D: "y", E: sh(null) } },
				tables: [{ name: "Sales", ref: "A2:B5", cols: ["Cat", "Amount"] }],
			},
		]);
		const { out, sheet } = render(tpl, { Sales: [{ Cat: "x", Amount: 10 }, { Cat: "y", Amount: 20 }, { Cat: "y", Amount: 30 }, { Cat: "x", Amount: 40 }] });
		expect(verify(out)).toEqual([]);
		expect(f(sheet(), "E9")).toBe("SUMIF(A:A,D9,B:B)");
		expect(f(sheet(), "E10")).toBe("SUMIF(A:A,D10,B:B)");
		manifest.add("review/1-shared-cols.xlsx", out, { cells: [{ sheet: "Sheet1", cell: "E9", value: 50, note: "SUMIF x" }, { sheet: "Sheet1", cell: "E10", value: 50, note: "SUMIF y" }], tables: [] });
	});

	it("#2 a sample group row inside a non-aggregating function follows its group instead of expanding", () => {
		const tpl = mkxlsx(
			[
				{ name: "Pools", rows: poolRows({ 11: { A: "round", B: "=ROUND(C6,2)" }, 12: { A: "if", B: "=IF(C6>100,C6,0)" }, 13: { A: "max", B: "=MAX(C6,0)" }, 14: { A: "sum", B: "=SUM(C3,C6)" } }) },
				{ name: "Rates", rows: { 1: { A: "=ROUND(Pools!C6,2)", B: "=IFERROR(Pools!C6,0)" } } },
			],
			{ names: [{ name: "Pools", ref: "Pools!$A$3:$C$9" }] },
		);
		const { out, sheet } = render(tpl, pools);
		expect(verify(out)).toEqual([]);
		expect(f(sheet(), "B13")).toBe("ROUND(C8,2)");
		expect(f(sheet(), "B14")).toBe("IF(C8>100,C8,0)");
		expect(f(sheet(), "B15")).toBe("MAX(C8,0)");
		expect(f(sheet(), "B16")).toBe("SUM(C3,C5,C8)"); // SUM still means "every group"
		expect(f(sheet(2), "A1")).toBe("ROUND(Pools!C8,2)");
		manifest.add("review/2-args.xlsx", out, {
			cells: [
				{ sheet: "Pools", cell: "B13", value: 3000, note: "ROUND(Overhead)" },
				{ sheet: "Pools", cell: "B15", value: 3000, note: "MAX(Overhead,0)" },
				{ sheet: "Pools", cell: "B16", value: 3008, note: "SUM of all pools" },
				{ sheet: "Rates", cell: "A1", value: 3000, note: "cross-sheet ROUND(Overhead)" },
			],
			tables: [],
		});
	});

	it("#3 a merge on every sample row is written once per output row", () => {
		const tpl = mkxlsx(
			[
				{
					name: "Sheet1",
					rows: { 2: { A: "Item", B: "Description", D: "Amount" }, 3: { A: "a", B: "desc a", D: 10 }, 4: { A: "b", B: "desc b", D: 20 }, 5: { A: "c", B: "desc c", D: 30 }, 6: { A: "Total", D: "=SUM(D3:D5)" } },
					extra: `<mergeCells count="3"><mergeCell ref="B3:C3"/><mergeCell ref="B4:C4"/><mergeCell ref="B5:C5"/></mergeCells>`,
				},
			],
			{ names: [{ name: "Items", ref: "Sheet1!$A$3:$D$5" }] },
		);
		const { out, sheet } = render(tpl, { Items: [1, 2, 3, 4].map((n) => ({ Item: `i${n}`, Description: `d${n}`, Amount: n })) });
		expect(verify(out)).toEqual([]);
		expect(sheet()).toContain('<mergeCells count="4"><mergeCell ref="B3:C3"/><mergeCell ref="B4:C4"/><mergeCell ref="B5:C5"/><mergeCell ref="B6:C6"/></mergeCells>');
		manifest.add("review/3-merges.xlsx", out, { cells: [{ sheet: "Sheet1", cell: "D7", value: 10, note: "total" }], tables: [] });
	});

	it("#4 an absolute range over the body, written in a detail row, grows with the data", () => {
		const tpl = mkxlsx([
			{
				name: "Sheet1",
				rows: {
					2: { A: "Item", B: "Amount", C: "Pct", D: "Running" },
					3: { A: "a", B: 10, C: "=B3/SUM($B$3:$B$5)", D: "=SUM($B$3:B3)" },
					4: { A: "b", B: 20, C: "=B4/SUM($B$3:$B$5)", D: "=SUM($B$3:B4)" },
					5: { A: "c", B: 30, C: "=B5/SUM($B$3:$B$5)", D: "=SUM($B$3:B5)" },
				},
				tables: [{ name: "TblA", ref: "A2:D5", cols: ["Item", "Amount", "Pct", "Running"] }],
			},
		]);
		const { out, sheet } = render(tpl, { TblA: [1, 2, 3, 4, 5, 6].map((n) => ({ Item: `i${n}`, Amount: n * 100 })) });
		expect(verify(out)).toEqual([]);
		expect(f(sheet(), "C3")).toBe("B3/SUM($B$3:$B$8)");
		expect(f(sheet(), "C8")).toBe("B8/SUM($B$3:$B$8)");
		expect(f(sheet(), "D8")).toBe("SUM($B$3:B8)");
		manifest.add("review/4-pct.xlsx", out, {
			cells: [
				{ sheet: "Sheet1", cell: "C3", value: Math.round((100 / 2100) * 1e6) / 1e6, note: "percent of total" },
				{ sheet: "Sheet1", cell: "D8", value: 2100, note: "running total" },
			],
			tables: [{ name: "TblA", dataRows: 6 }],
		});
	});

	it("#4b a lookup into another region on the same sheet grows with that region", () => {
		const tpl = mkxlsx([
			{
				name: "Sheet1",
				rows: {
					1: { A: "Code", B: "Rate" },
					2: { A: "X", B: 1 },
					3: { A: "Y", B: 2 },
					5: { A: "Code", B: "Hours", C: "Cost" },
					6: { A: "X", B: 10, C: "=B6*VLOOKUP(A6,$A$2:$B$3,2,FALSE)" },
					7: { A: "Y", B: 20, C: "=B7*VLOOKUP(A7,$A$2:$B$3,2,FALSE)" },
				},
				tables: [
					{ name: "Rates", ref: "A1:B3", cols: ["Code", "Rate"] },
					{ name: "Work", ref: "A5:C7", cols: ["Code", "Hours", "Cost"] },
				],
			},
		]);
		const { out, sheet } = render(tpl, {
			Rates: [{ Code: "X", Rate: 1 }, { Code: "Y", Rate: 2 }, { Code: "Z", Rate: 3 }, { Code: "W", Rate: 4 }],
			Work: [{ Code: "Z", Hours: 10 }, { Code: "W", Hours: 10 }, { Code: "X", Hours: 10 }],
		});
		expect(verify(out)).toEqual([]);
		expect(f(sheet(), "C8")).toBe("B8*VLOOKUP(A8,$A$2:$B$5,2,FALSE)");
		manifest.add("review/4b-lookup.xlsx", out, {
			cells: [
				{ sheet: "Sheet1", cell: "C8", value: 30, note: "Z" },
				{ sheet: "Sheet1", cell: "C9", value: 40, note: "W" },
			],
			tables: [{ name: "Rates", dataRows: 4 }, { name: "Work", dataRows: 3 }],
		});
	});

	it("#5 #6 CF/DV absolute ranges grow; ranges to the last row stay at the last row", () => {
		const extra =
			`<conditionalFormatting sqref="B3:B5"><cfRule type="expression" dxfId="0" priority="1"><formula>$B3&gt;AVERAGE($B$3:$B$5)</formula></cfRule></conditionalFormatting>` +
			`<dataValidations count="1"><dataValidation type="custom" allowBlank="1" showErrorMessage="1" sqref="A3:A5"><formula1>COUNTIF($A$3:$A$5,A3)=1</formula1></dataValidation></dataValidations>`;
		const tpl = mkxlsx(
			[
				{
					name: "Sheet1",
					rows: {
						2: { A: "Item", B: "Amount" },
						3: { A: "a", B: 10 },
						4: { A: "b", B: 20 },
						5: { A: "c", B: 30 },
						7: { A: "Total", B: "=SUM(B3:B5)" },
						8: { A: "count to bottom", B: "=COUNTA($A$3:$A$1048576)" },
						10: { A: "rows to bottom", B: "=ROWS(3:1048576)" },
					},
					extra,
					tables: [{ name: "Items", ref: "A2:B5", cols: ["Item", "Amount"] }],
				},
			],
			{ names: [{ name: "ToBottom", ref: "Sheet1!$B$3:$B$1048576" }] },
		);
		const { out, sheet } = render(tpl, { Items: [1, 2, 3, 4, 5, 6].map((n) => ({ Item: `i${n}`, Amount: n * 10 })) });
		expect(verify(out)).toEqual([]);
		expect(sheet()).toContain("<formula>$B3&gt;AVERAGE($B$3:$B$8)</formula>");
		expect(sheet()).toContain("<formula1>COUNTIF($A$3:$A$8,A3)=1</formula1>");
		expect(f(sheet(), "B11")).toBe("COUNTA($A$3:$A$1048576)");
		expect(f(sheet(), "B13")).toBe("ROWS(3:1048576)");
		expect(partText(out, "xl/workbook.xml")).toContain("Sheet1!$B$3:$B$1048576");
		manifest.add("review/5-6-ranges.xlsx", out, {
			cells: [
				{ sheet: "Sheet1", cell: "B11", value: 9, note: "COUNTA to bottom" },
				{ sheet: "Sheet1", cell: "B13", value: 1048574, note: "ROWS to bottom" },
			],
			tables: [{ name: "Items", dataRows: 6 }],
		});
	});

	it("#7 a pivot table below the data moves with it", () => {
		const TEMPLATE = corpusFile("review-pivot-below", "template.xlsx");
		const out = Template.loadSync(TEMPLATE).renderSync({ Sales: [1, 2, 3, 4, 5, 6, 7, 8].map((n) => ({ Cat: n % 2 ? "x" : "y", Amount: n })) });
		expect(verify(out)).toEqual([]);
		expect(partText(out, "xl/pivotTables/pivotTable1.xml")).toContain('<location ref="A15:B18"');
		manifest.add("review/7-pivot.xlsx", out, { cells: [], tables: [{ name: "Sales", dataRows: 8 }] });
	});

	it("#8 ranges over one sample group's details follow that group", () => {
		const tpl = mkxlsx(
			[{ name: "Pools", rows: poolRows({ 11: { A: "sum", C: "=SUM(C4:C5,C7:C8)" }, 12: { A: "count", C: "=COUNT(C4:C5,C7:C8)" }, 13: { A: "max fringe", C: "=MAX(C4:C5)" } }) }],
			{ names: [{ name: "Pools", ref: "Pools!$A$3:$C$9" }] },
		);
		const { out, sheet } = render(tpl, pools);
		expect(verify(out)).toEqual([]);
		expect(f(sheet(), "C13")).toBe("SUM(C6:C7,C9:C10)");
		expect(f(sheet(), "C15")).toBe("MAX(C6:C7)");
		manifest.add("review/8-detail-areas.xlsx", out, {
			cells: [
				{ sheet: "Pools", cell: "C13", value: 3003, note: "Fringe + Overhead accounts" },
				{ sheet: "Pools", cell: "C14", value: 4, note: "count" },
				{ sheet: "Pools", cell: "C15", value: 2, note: "max Fringe account" },
			],
			tables: [],
		});
	});

	it("#9 references into nested groups match the whole label path", () => {
		const tpl = mkxlsx(
			[
				{
					name: "Sheet1",
					rows: {
						2: { A: "Dept", B: "Item", C: "Amount" },
						3: { A: "Eng", C: "=SUM(C4,C7)" },
						4: { B: "Travel", C: "=SUM(C5:C6)" },
						5: { B: "t1", C: 1 },
						6: { B: "t2", C: 2 },
						7: { B: "Supplies", C: "=SUM(C8:C9)" },
						8: { B: "s1", C: 3 },
						9: { B: "s2", C: 4 },
						10: { A: "Sales", C: "=SUM(C11,C14)" },
						11: { B: "Travel", C: "=SUM(C12:C13)" },
						12: { B: "t1", C: 5 },
						13: { B: "t2", C: 6 },
						14: { B: "Supplies", C: "=SUM(C15:C16)" },
						15: { B: "s1", C: 7 },
						16: { B: "s2", C: 8 },
						17: { A: "Total", C: "=SUM(C3,C10)" },
						19: { A: "Sales travel", C: "=C11" },
					},
				},
			],
			{ names: [{ name: "Spend", ref: "Sheet1!$A$3:$C$17" }] },
		);
		const { out, sheet } = render(tpl, {
			Spend: [
				{ Dept: "Eng", rows: [{ Item: "Travel", rows: [{ Item: "x", Amount: 100 }] }, { Item: "Supplies", rows: [{ Item: "y", Amount: 200 }] }] },
				{ Dept: "Sales", rows: [{ Item: "Travel", rows: [{ Item: "x", Amount: 1000 }] }, { Item: "Supplies", rows: [{ Item: "y", Amount: 2000 }] }] },
			],
		});
		expect(verify(out)).toEqual([]);
		expect(f(sheet(), "C15")).toBe("C9");
		manifest.add("review/9-nested-labels.xlsx", out, { cells: [{ sheet: "Sheet1", cell: "C15", value: 1000, note: "Sales/Travel" }], tables: [] });
	});

	it("#10 a 3-D reference spanning rendered sheets is reported, not silently left stale", () => {
		const month = () => ({ 2: { A: "Item", B: "Amount" }, 3: { A: "a", B: 1 }, 4: { A: "b", B: 2 }, 5: { A: "c", B: 3 }, 7: { A: "Total", B: "=SUM(B3:B5)" } });
		const tpl = mkxlsx([
			{ name: "Jan", rows: month(), tables: [{ name: "JanRows", ref: "A2:B5", cols: ["Item", "Amount"] }] },
			{ name: "Feb", rows: month(), tables: [{ name: "FebRows", ref: "A2:B5", cols: ["Item", "Amount"] }] },
			{ name: "Summary", rows: { 1: { A: "Both", B: "=SUM(Jan:Feb!B7)" } } },
		]);
		const { warnings } = render(tpl, { JanRows: [{ Amount: 1 }], FebRows: [{ Amount: 2 }, { Amount: 3 }] });
		expect(warnings.some((w) => /Jan:Feb!B7/.test(w))).toBe(true);
	});

	it("lesser: date-labelled groups match; hyperlink locations and totalsRowFormula follow; sqrefs aren't duplicated", () => {
		const tpl = mkxlsx(
			[
				{
					name: "S",
					rows: {
						2: { A: "Day", B: "Item", C: "Amount" },
						3: { A: { v: 46023, s: 2 }, C: "=SUM(C4:C5)" },
						4: { B: "a", C: 1 },
						5: { B: "b", C: 2 },
						6: { A: { v: 46024, s: 2 }, C: "=SUM(C7:C8)" },
						7: { B: "c", C: 3 },
						8: { B: "d", C: 4 },
						9: { A: "Total", C: "=SUM(C3,C6)" },
						11: { A: "Jan 2 subtotal", C: "=C6" },
					},
					extra: `<dataValidations count="1"><dataValidation type="whole" sqref="C4:C5 C7:C8"><formula1>0</formula1></dataValidation></dataValidations>`,
				},
			],
			{ names: [{ name: "Days", ref: "S!$A$3:$C$9" }] },
		);
		const { out, sheet, warnings } = render(tpl, {
			Days: [
				{ Day: new Date(Date.UTC(2026, 0, 1)), rows: [{ Item: "x", Amount: 1 }] },
				{ Day: new Date(Date.UTC(2026, 0, 2)), rows: [{ Item: "y", Amount: 2 }] },
			],
		});
		expect(warnings).toEqual([]);
		expect(f(sheet(), "C9")).toBe("C5");
		expect(sheet()).toContain('sqref="C4 C6"');

		const misc = corpusFile("review-misc", "template.xlsx");
		const r2 = Template.loadSync(misc).renderSync({ Sales: [1, 2, 3, 4, 5, 6].map((n) => ({ Cat: `c${n}`, Amount: n })) });
		expect(verify(r2)).toEqual([]);
		expect(partText(r2, "xl/worksheets/sheet1.xml")).toContain('location="Data!B15"');
		expect(partText(r2, "xl/tables/table1.xml")).toContain("<totalsRowFormula>SUM(B2:B7)*2</totalsRowFormula>");
		manifest.add("review/lesser-dates.xlsx", out, { cells: [{ sheet: "S", cell: "C9", value: 2, note: "Jan 2 group" }], tables: [] });
		manifest.add("review/lesser-misc.xlsx", r2, { cells: [{ sheet: "Data", cell: "B15", value: 42, note: "=B8 totals" }], tables: [{ name: "Sales", dataRows: 6 }] });
	});
});
