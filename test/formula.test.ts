import fc from "fast-check";
import { describe, expect, it } from "vitest";
import { indexToCol } from "../src/a1.js";
import { rewriteRefs, sheetPrefix, tokenize } from "../src/formula.js";
import { shiftFormula } from "../src/sheet.js";

const refs = (f: string) =>
	tokenize(f)
		.filter((t) => t.kind === "ref")
		.map((t) => f.slice(t.start, t.end));

describe("tokenize finds exactly the A1 references", () => {
	it.each([
		["=SUM(C6:C7)", ["C6:C7"]],
		["SUM(C5,C8)", ["C5", "C8"]],
		["=$A$1+B$2*$C3", ["$A$1", "B$2", "$C3"]],
		["='My Sheet'!A1+Sheet2!B2:C3", ["'My Sheet'!A1", "Sheet2!B2:C3"]],
		["='O''Brien'!$A$1", ["'O''Brien'!$A$1"]],
		["=LOG10(A1)+ATAN2(B1,C1)", ["A1", "B1", "C1"]],
		['="A1 in a string "&A1', ["A1"]],
		["=SUM(Table1[Amount])+Table1[[#Totals],[Amount]]", []],
		["=[@Qty]*[@[Unit Price]]", []],
		["=SUM(5:5)+SUM($A:$C)", ["5:5", "$A:$C"]],
		["=A1#", ["A1#"]],
		["=#REF!+A1", ["A1"]],
		["=Sheet1!#REF!", []],
		["=R1C1+TaxRate", []],
		["=1E+5+2.5E-3+A1", ["A1"]],
		["=[1]Sheet1!A1+A2", ["[1]Sheet1!A1", "A2"]],
		["=_xlfn.XLOOKUP(A2,B:B,C:C)", ["A2", "B:B", "C:C"]],
		["=Sheet1:Sheet3!A1", ["Sheet1:Sheet3!A1"]],
		["=IF(A1>=B1,\"x\",\"y\")", ["A1", "B1"]],
		["={1,2;3,4}", []],
		["=XFD1048576+XFE1", ["XFD1048576"]],
	])("%s", (formula, expected) => {
		expect(refs(formula)).toEqual(expected);
	});

	it("marks foreign (other-workbook and 3-D) references", () => {
		const [a, b] = tokenize("=[1]Sheet1!A1+Sheet1:Sheet3!A1").filter((t) => t.kind === "ref");
		expect(a && "foreign" in a && a.foreign).toBe(true);
		expect(b && "foreign" in b && b.foreign).toBe(true);
	});
});

describe("rewriteRefs", () => {
	it("knows which references are whole function arguments", () => {
		const seen: [string, boolean][] = [];
		const f = "=SUM(C5,C8)+C9*2+SUM((C1,C2))+MAX(C3)";
		rewriteRefs(f, (ref, ctx) => {
			seen.push([f.slice(ref.start, ref.end), ctx.inArgList]);
			return undefined;
		});
		expect(seen).toEqual([
			["C5", true],
			["C8", true],
			["C9", false],
			["C1", false],
			["C2", false],
			["C3", true],
		]);
	});

	it("drops later arguments that expand to the same thing", () => {
		const out = rewriteRefs("SUM(C5,C8)", (ref) => (ref.a.row === 5 || ref.a.row === 8 ? { text: "C5,C9,C13", dedupeKey: "k" } : undefined));
		expect(out).toBe("SUM(C5,C9,C13)");
	});

	it("returns the identical string when the mapper changes nothing", () => {
		fc.assert(
			fc.property(fc.string(), (s) => {
				expect(rewriteRefs(s, () => undefined)).toBe(s);
			}),
		);
	});
});

const cellArb = fc.record({
	col: fc.integer({ min: 1, max: 200 }),
	row: fc.integer({ min: 1, max: 5000 }),
	colAbs: fc.boolean(),
	rowAbs: fc.boolean(),
});
const cellText = (c: { col: number; row: number; colAbs: boolean; rowAbs: boolean }) =>
	`${c.colAbs ? "$" : ""}${indexToCol(c.col)}${c.rowAbs ? "$" : ""}${c.row}`;
const sheetArb = fc.constantFrom("", "Data!", "'My Sheet'!", "'O''Brien'!", "S2!");
const refArb = fc.tuple(sheetArb, cellArb, fc.option(cellArb, { nil: undefined })).map(([s, a, b]) => `${s}${cellText(a)}${b ? `:${cellText(b)}` : ""}`);
const fillerArb = fc.constantFrom("+", "*", ",", " ", "-", "/", "&\"txt A1\"&", "+SUM(Table1[Amount])+", "+LOG10(2)+", "+TaxRate+");
const formulaArb = fc.array(fc.tuple(refArb, fillerArb), { minLength: 1, maxLength: 8 }).map((parts) => `=SUM(${parts.map(([r, f]) => r + f).join("")}1)`);

describe("property: generated formulas", () => {
	it("every generated reference is found, in order, with its exact text", () => {
		fc.assert(
			fc.property(fc.array(fc.tuple(refArb, fillerArb), { minLength: 1, maxLength: 8 }), (parts) => {
				const f = `=SUM(${parts.map(([r, fl]) => r + fl).join("")}1)`;
				expect(refs(f)).toEqual(parts.map(([r]) => r));
			}),
		);
	});

	it("shifting by +d then -d is the identity when nothing falls off the sheet", () => {
		fc.assert(
			fc.property(formulaArb, fc.integer({ min: 0, max: 400 }), fc.integer({ min: 0, max: 20 }), (f, dr, dc) => {
				const there = shiftFormula(f, dr, dc);
				expect(shiftFormula(there, -dr, -dc)).toBe(f);
			}),
		);
	});

	it("shifting never touches absolute parts", () => {
		fc.assert(
			fc.property(cellArb, fc.integer({ min: -50, max: 50 }), (c, d) => {
				const f = `=${cellText({ ...c, colAbs: true, rowAbs: true })}`;
				expect(shiftFormula(f, d, d)).toBe(f);
			}),
		);
	});

	it("references pushed off the sheet become #REF!", () => {
		expect(shiftFormula("=A1+B2", -1, 0)).toBe("=#REF!+B1");
		expect(shiftFormula("=Data!A1", -1, 0)).toBe("=Data!#REF!");
	});
});

describe("sheetPrefix", () => {
	it.each([
		["Data", "Data!"],
		["My Sheet", "'My Sheet'!"],
		["O'Brien", "'O''Brien'!"],
		["A1", "'A1'!"],
		["R1C1", "'R1C1'!"],
		["2024", "'2024'!"],
	])("%s", (name, expected) => expect(sheetPrefix(name)).toBe(expected));
});
