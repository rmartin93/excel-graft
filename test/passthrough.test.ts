import { readFile } from "node:fs/promises";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { Template } from "../src/index.js";
import { readParts } from "../src/zip.js";

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

describe("inspect()", () => {
	it("finds the table in the exceljs-fork-broken corpus template", async () => {
		const original = await loadCorpusFile("exceljs-fork-broken", "Test-Template.xlsx");
		const tpl = await Template.load(original);

		const result = tpl.inspect();
		expect(result.tables.length).toBeGreaterThan(0);
	});
});

describe("render()", () => {
	it("is not implemented yet (Phase 1 has not started)", async () => {
		const original = await loadCorpusFile("exceljs-fork-broken", "Test-Template.xlsx");
		const tpl = await Template.load(original);

		await expect(tpl.render({})).rejects.toThrow(/not implemented/i);
	});
});
