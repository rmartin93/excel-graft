import { afterAll, describe, expect, it } from "vitest";
import { type RenderReport, Template, verify } from "../src/index.js";
import { ExcelManifest, sumChecks } from "./excel-checks.js";
import { corpusFile, expectPassthrough, partText, renderablePart } from "./helpers.js";

const manifest = new ExcelManifest("grouped-layouts");
afterAll(() => manifest.save());

function render(tpl: Template, data: Record<string, unknown>) {
	let report: RenderReport | undefined;
	const warnings: string[] = [];
	const out = tpl.renderSync(data, { onReport: (r) => (report = r), onWarning: (w) => warnings.push(w) });
	return { out, report: report as RenderReport, warnings };
}

describe("grouped-footers: subtotal below each group, spacer rows, banding, running balance", () => {
	const TEMPLATE = corpusFile("grouped-footers", "template.xlsx");

	it("infers footer groups and keeps the label's fixed text as layout", async () => {
		const tpl = await Template.load(TEMPLATE);
		const region = tpl.inspect().regions[0]!;
		expect(region).toMatchObject({ key: "Payroll", kind: "name", fixedRows: ["A14:D14"] });
		expect(region.shape).toMatchObject({
			kind: "groups",
			fields: [{ name: "Employee", samples: ["Engineering", "Finance"], labelFormat: "{value} subtotal" }],
			child: { kind: "rows", fields: [{ name: "Employee" }, { name: "Title" }, { name: "Gross Pay" }] },
		});
	});

	it("renders footers, spacers, banding and running balances for any number of groups", async () => {
		const tpl = await Template.load(TEMPLATE);
		const person = (n: number, pay: number) => ({ Employee: `Person ${n}`, Title: "Staff", "Gross Pay": pay });
		const { out, report } = render(tpl, {
			period: "September 2026",
			Payroll: [
				{ Employee: "Sales", rows: [person(1, 100), person(2, 200), person(3, 300), person(4, 400)] },
				{ Employee: "Support", rows: [person(5, 50)] },
				{ Employee: "Research", rows: [person(6, 10), person(7, 20)] },
			],
		});
		expect(verify(out)).toEqual([]);
		expectPassthrough(TEMPLATE, out, renderablePart);
		const xml = partText(out, "xl/worksheets/sheet1.xml");
		// Sales: rows 5-8, footer 9, spacer 10; Support: 11, footer 12, spacer 13; Research: 14-15, footer 16, spacer 17; total 18.
		expect(xml).toContain("<t>Sales subtotal</t>");
		expect(xml).toContain("<f>SUM(C5:C8)</f>");
		expect(xml).toContain("<f>SUM(C11:C11)</f>");
		expect(xml).toContain("<f>SUM(C14:C15)</f>");
		expect(xml).toContain("<f>SUM(C9,C12,C16)</f>");
		// Running balance: the first row of each group starts over, the rest add the previous row.
		expect(xml).toMatch(/<c r="D5"[^>]*><f>C5<\/f>/);
		expect(xml).toMatch(/<c r="D6"[^>]*><f>D5\+C6<\/f>/);
		expect(xml).toMatch(/<c r="D14"[^>]*><f>C14<\/f>/);
		expect(xml).toMatch(/<c r="D15"[^>]*><f>D14\+C15<\/f>/);
		// The row below the region followed the grand total down.
		expect(xml).toMatch(/<c r="C20"[^>]*><f>C18\/2<\/f>/);
		const workbook = partText(out, "xl/workbook.xml");
		expect(workbook).toContain("Payroll!$A$5:$D$18");
		const cells = [
			...sumChecks(tpl.inspect(), report, out),
			{ sheet: "Payroll", cell: "D8", value: 1000, note: "running balance, end of Sales" },
			{ sheet: "Payroll", cell: "D15", value: 30, note: "running balance restarts per group" },
			{ sheet: "Payroll", cell: "A2", value: "Period: September 2026", note: "placeholder" },
		];
		expect(cells.find((c) => c.cell === "C18")?.value).toBe(1080);
		manifest.add("grouped/footers.xlsx", out, { cells, tables: [] });
	});

	it("banded detail rows alternate their fills like the sample", async () => {
		const tpl = await Template.load(TEMPLATE);
		const { out } = render(tpl, { Payroll: [{ Employee: "A", rows: [{}, {}, {}, {}] }] });
		const xml = partText(out, "xl/worksheets/sheet1.xml");
		const styleOf = (cell: string) => new RegExp(`<c r="${cell}"(?: s="(\\d+)")?`).exec(xml)?.[1];
		expect(styleOf("A5")).toBe(styleOf("A7"));
		expect(styleOf("A6")).toBe(styleOf("A8"));
		expect(styleOf("A5")).not.toBe(styleOf("A6"));
	});
});

describe("nested-groups: Region > Category > lines, a Table, CF, and a chart", () => {
	const TEMPLATE = corpusFile("nested-groups", "template.xlsx");

	it("infers two levels of groups under a grand total", async () => {
		const tpl = await Template.load(TEMPLATE);
		const region = tpl.inspect().regions[0]!;
		expect(region).toMatchObject({ key: "Budget", kind: "table", fixedRows: ["A16:D16"] });
		expect(region.shape).toMatchObject({
			kind: "groups",
			fields: [{ name: "Line", samples: ["North", "South"] }],
			child: { kind: "groups", fields: [{ name: "Line" }], child: { kind: "rows" } },
		});
	});

	it("regenerates totals at every level", async () => {
		const tpl = await Template.load(TEMPLATE);
		const line = (Line: string, Budget: number, Actual: number) => ({ Line, Budget, Actual });
		const { out, report } = render(tpl, {
			Budget: [
				{ Line: "West", rows: [{ Line: "Travel", rows: [line("Air", 100, 90), line("Rail", 50, 70), line("Car", 20, 20)] }] },
				{
					Line: "East",
					rows: [
						{ Line: "Travel", rows: [line("Air", 300, 310)] },
						{ Line: "Software", rows: [line("SaaS", 1000, 900), line("Laptops", 2000, 2100)] },
						{ Line: "Training", rows: [] },
					],
				},
			],
		});
		expect(verify(out)).toEqual([]);
		expectPassthrough(TEMPLATE, out, renderablePart);
		const xml = partText(out, "xl/worksheets/sheet1.xml");
		// West 4, Travel 5, lines 6-8; East 9, Travel 10, line 11, Software 12, lines 13-14, Training 15, blank 16; total 17.
		expect(xml).toContain("<f>SUM(B6:B8)</f>");
		expect(xml).toContain("<f>SUM(B5)</f>");
		expect(xml).toContain("<f>SUM(B10,B12,B15)</f>");
		expect(xml).toContain("<f>SUM(B4,B9)</f>");
		expect(partText(out, "xl/tables/table1.xml")).toContain('ref="A3:D17"');
		const cells = [
			...sumChecks(tpl.inspect(), report, out),
			{ sheet: "Budget", cell: "B17", value: 3470, note: "grand total budget" },
			{ sheet: "Budget", cell: "C17", value: 3490, note: "grand total actual" },
			{ sheet: "Budget", cell: "D9", value: 10, note: "East variance" },
		];
		manifest.add("grouped/nested.xlsx", out, { cells, tables: [{ name: "Budget", dataRows: 14 }] });
	});
});
