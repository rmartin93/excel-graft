import { readFile } from "node:fs/promises";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { Template } from "../src/index.js";
import { readParts } from "../src/zip.js";
import { verify } from "../src/index.js";
import { expectPassthrough, partText, renderablePart } from "./helpers.js";

const CORPUS_DIR = path.join(__dirname, "..", "corpus", "templates");

async function loadCorpusFile(...segments: string[]): Promise<Uint8Array> {
	const filePath = path.join(CORPUS_DIR, ...segments);
	return new Uint8Array(await readFile(filePath));
}

describe("passthrough invariant", () => {
	it("round-trips every part of a real template byte-for-byte with zero edits applied", async () => {
		const original = await loadCorpusFile("exceljs-fork-broken", "Test-Template.xlsx");
		const originalParts = readParts(original);

		const tpl = await Template.load(original);
		const output = await tpl.toBuffer();
		const outputParts = readParts(output);

		expect(outputParts.size).toBe(originalParts.size);
		for (const [name, originalBytes] of originalParts) {
			const outputBytes = outputParts.get(name);
			expect(outputBytes, `part "${name}" missing from output`).toBeDefined();
			expect(
				Buffer.from(outputBytes as Uint8Array).equals(Buffer.from(originalBytes)),
				`part "${name}" was not byte-identical`,
			).toBe(true);
		}
	});

	it("finds no part names in the output that weren't in the input", async () => {
		const original = await loadCorpusFile("exceljs-fork-broken", "Test-Template.xlsx");
		const originalParts = readParts(original);

		const tpl = await Template.load(original);
		const outputParts = readParts(await tpl.toBuffer());

		const extraNames = [...outputParts.keys()].filter((name) => !originalParts.has(name));
		expect(extraNames).toEqual([]);
	});
});

describe("exceljs-fork incident regression", () => {
	// The data the ExcelJS fork rendered into this template in production.
	const data = {
		Table1: [
			{ ID: "todo1", Name: "Buy groceries", Value: 0 },
			{ ID: "todo2", Name: "Walk the dog", Value: 50 },
			{ ID: "todo3", Name: "Finish project", Value: 100 },
		],
	};

	it("verify() catches the corruption ExcelJS produced", async () => {
		const broken = await loadCorpusFile("exceljs-fork-broken", "Test-Output.xlsx");
		expect(verify(broken)).toEqual([
			{ part: "xl/tables/table1.xml", message: "autoFilter ref A1:C1 should be A1:C4 (table ref minus totals rows)" },
		]);
	});

	it("rendering the same data here avoids every defect ExcelJS introduced", async () => {
		const original = await loadCorpusFile("exceljs-fork-broken", "Test-Template.xlsx");
		const tpl = await Template.load(original);
		expect(tpl.inspect().regions.map((r) => r.key)).toEqual(["Table1"]);
		const out = await tpl.render(data);
		expect(verify(out)).toEqual([]);
		const table = partText(out, "xl/tables/table1.xml");
		expect(table).toContain('ref="A1:C5"');
		expect(table).toContain('<autoFilter ref="A1:C4" xr:uid="{4B8469E7-0906-4C85-A657-C1DE99AD1531}"/>');
		expect(table).toContain('dataDxfId="0"'); // ExcelJS dropped this (and the fill it points at)
		const workbook = partText(out, "xl/workbook.xml");
		expect(workbook).toContain("<extLst>"); // ExcelJS stripped calcFeatures and more
		expect(workbook).toContain('r:id="rId1"'); // ExcelJS renumbered relationships
		const sheet = partText(out, "xl/worksheets/sheet1.xml");
		expect(sheet).toContain('mc:Ignorable="x14ac xr xr2 xr3"');
		expect(sheet).toContain('<c r="C5"><f>SUBTOTAL(109,Table1[Value])</f></c>');
		// styles.xml (the dxf ExcelJS emptied), theme, docProps, sharedStrings: byte-identical.
		expectPassthrough(original, out, renderablePart);
	});
});
