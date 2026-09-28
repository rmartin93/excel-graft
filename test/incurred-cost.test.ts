import { afterAll, describe, expect, it } from "vitest";
import { type RenderReport, Template, verify } from "../src/index.js";
import { ExcelManifest, sumChecks } from "./excel-checks.js";
import { corpusFile, expectPassthrough, partText, renderablePart } from "./helpers.js";

const TEMPLATE = corpusFile("incurred-cost", "template.xlsx");
const manifest = new ExcelManifest("incurred-cost");
afterAll(() => manifest.save());

const labor = Array.from({ length: 50 }, (_, i) => ({
	Employee: `Employee ${i + 1}`,
	LaborCategory: i % 2 ? "Engineer III" : "Analyst II",
	Project: `P-${100 + (i % 5)}`,
	Hours: 100 + i,
	Rate: 50 + (i % 7),
}));
const laborCost = labor.reduce((s, r) => s + r.Hours * r.Rate, 0);
const acct = (Account: number, Amount: number) => ({ Account, Description: `Account ${Account}`, Amount });

describe("incurred-cost template (cross-sheet rates over grouped pools)", () => {
	it("the template's pool groups carry their sample labels", async () => {
		const tpl = await Template.load(TEMPLATE);
		const pools = tpl.inspect().regions.find((r) => r.key === "IndirectPools");
		expect(pools?.shape).toMatchObject({ kind: "groups", fields: [{ name: "Account", type: "string", samples: ["Fringe", "Overhead", "G&A"] }] });
	});

	it("references to a specific pool follow that pool by label, wherever it lands", async () => {
		const tpl = await Template.load(TEMPLATE);
		const warnings: string[] = [];
		let report: RenderReport | undefined;
		const out = await tpl.render(
			{
				ContractorName: "Meraki Digital, LLC",
				FiscalYearEnd: new Date(Date.UTC(2026, 11, 31)),
				PreparedBy: "Ryan Martin",
				submissionDate: new Date(Date.UTC(2027, 5, 30)),
				DirectLabor: labor,
				// Overhead first, Fringe second, no G&A pool this year.
				IndirectPools: [
					{ Account: "Overhead", rows: [acct(6100, 1000), acct(6110, 250), acct(6120, 125), acct(6130, 5)] },
					{ Account: "Fringe", rows: [acct(5100, 700), acct(5110, 300)] },
				],
			},
			{ onWarning: (w) => warnings.push(w), onReport: (r) => (report = r) },
		);
		expect(verify(out)).toEqual([]);
		expectPassthrough(TEMPLATE, out, renderablePart);

		// Indirect Pools: Overhead header row 5 (details 6-9), Fringe header row 10 (details 11-12), total row 13.
		const pools = partText(out, "xl/worksheets/sheet3.xml");
		expect(pools).toContain("<f>SUM(C6:C9)</f>");
		expect(pools).toContain("<f>SUM(C11:C12)</f>");
		expect(pools).toContain("<f>SUM(C5,C10)</f>");

		const rates = partText(out, "xl/worksheets/sheet4.xml");
		expect(rates).toContain("<f>'Indirect Pools'!C10</f>"); // Fringe (was C5)
		expect(rates).toContain("<f>'Indirect Pools'!C5</f>"); // Overhead (was C9)
		expect(rates).toContain("<f>'Indirect Pools'!#REF!</f>"); // G&A isn't in the data: visible, not silently wrong
		expect(rates).toContain("<f>'Direct Labor'!F55</f>"); // totals row moved from 11 to 55
		expect(warnings.some((w) => /G&A/.test(w))).toBe(true);

		const workbook = partText(out, "xl/workbook.xml");
		expect(workbook).toContain("<definedName name=\"FringeRate\">'Indirect Pools'!$C$10/'Direct Labor'!$F$55</definedName>");
		expect(workbook).toContain("<definedName name=\"OverheadRate\">'Indirect Pools'!$C$5/'Direct Labor'!$F$55</definedName>");
		expect(workbook).toContain("<definedName name=\"IndirectPools\">'Indirect Pools'!$A$5:$D$13</definedName>");
		expect(workbook).toContain("'Direct Labor'!$A$1:$F$55");

		const cells = [
			...sumChecks(tpl.inspect(), report as RenderReport, out),
			{ sheet: "Rates", cell: "B5", value: 1000, note: "Fringe pool" },
			{ sheet: "Rates", cell: "B6", value: 1380, note: "Overhead pool" },
			{ sheet: "Rates", cell: "C5", value: laborCost, note: "direct labor base" },
			{ sheet: "Rates", cell: "D5", value: Math.round((1000 / laborCost) * 1e6) / 1e6, note: "FringeRate" },
			{ sheet: "Rates", cell: "D6", value: Math.round((1380 / laborCost) * 1e6) / 1e6, note: "OverheadRate" },
			{ sheet: "Cover", cell: "B5", value: "Meraki Digital, LLC", note: "named cell" },
		];
		manifest.add("incurred-cost/reordered-pools.xlsx", out, { cells, tables: [{ name: "DirectLabor", dataRows: 50 }], allowBrokenRefs: true });
	});
});
