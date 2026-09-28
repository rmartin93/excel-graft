import fc from "fast-check";
import { describe, expect, it } from "vitest";
import { isDateFormat } from "../src/styles.js";
import { dateToSerial, encodeValue, MAX_CELL_TEXT, sanitizeText } from "../src/values.js";
import { parseXml } from "../src/xml.js";

describe("dates", () => {
	it.each([
		[Date.UTC(1900, 0, 1), 1],
		[Date.UTC(1900, 1, 28), 59],
		[Date.UTC(1900, 2, 1), 61],
		[Date.UTC(2026, 0, 1), 46023],
		[Date.UTC(2026, 0, 1, 12), 46023.5],
	])("1900 system: %s → %s", (ms, serial) => expect(dateToSerial(new Date(ms), false)).toBe(serial));

	it("1904 system", () => {
		expect(dateToSerial(new Date(Date.UTC(1904, 0, 1)), true)).toBe(0);
		expect(dateToSerial(new Date(Date.UTC(2026, 0, 1)), true)).toBe(46023 - 1462);
	});

	it("dates Excel can't show are written as text instead", () => {
		expect(encodeValue(new Date(Date.UTC(1850, 0, 1)), false).t).toBe("inlineStr");
		expect(encodeValue(new Date(Number.NaN), false)).toEqual({ t: "e", inner: "<v>#VALUE!</v>" });
	});
});

describe("text sanitizing", () => {
	it("escapes control characters the way Excel reads them back", () => {
		expect(sanitizeText("a\u0001b")).toBe("a_x0001_b");
		expect(sanitizeText("_x0041_")).toBe("_x005F_x0041_");
	});

	it("truncates at Excel's cell limit without splitting a surrogate pair", () => {
		const s = `${"a".repeat(MAX_CELL_TEXT - 1)}😀tail`;
		const out = sanitizeText(s);
		expect(out.length).toBeLessThanOrEqual(MAX_CELL_TEXT);
		expect(out.endsWith("\ud83d")).toBe(false);
	});

	it("property: any string encodes to well-formed XML within the length limit", () => {
		fc.assert(
			fc.property(fc.fullUnicodeString({ maxLength: 200 }), fc.string({ maxLength: 50 }), (a, b) => {
				const enc = encodeValue(a + b, false);
				if (enc.inner === "") return;
				expect(() => parseXml(`<c>${enc.inner}</c>`)).not.toThrow();
			}),
		);
	});

	it("property: text round-trips through sanitize + Excel's decoding", () => {
		const decode = (s: string) => s.replace(/_x([0-9A-F]{4})_/g, (_m, h: string) => String.fromCharCode(parseInt(h, 16)));
		fc.assert(
			fc.property(fc.string({ maxLength: 100 }), (s) => {
				expect(decode(sanitizeText(s))).toBe(s);
			}),
		);
	});
});

describe("numbers", () => {
	it("non-finite numbers become #NUM! instead of invalid XML", () => {
		expect(encodeValue(Number.NaN, false)).toEqual({ t: "e", inner: "<v>#NUM!</v>" });
		expect(encodeValue(Number.POSITIVE_INFINITY, false).t).toBe("e");
	});

	it("-0 is written as 0 and bigints stay exact", () => {
		expect(encodeValue(-0, false).inner).toBe("<v>0</v>");
		expect(encodeValue(42n, false).inner).toBe("<v>42</v>");
		expect(encodeValue(123456789012345678n, false)).toMatchObject({ t: "inlineStr" });
	});

	it("strings that look like formulas stay strings", () => {
		expect(encodeValue("=1+1", false)).toEqual({ t: "inlineStr", inner: "<is><t>=1+1</t></is>" });
	});
});

describe("isDateFormat", () => {
	it.each([
		["m/d/yyyy", true],
		["[$-409]mmmm d, yyyy;@", true],
		["h:mm AM/PM", true],
		["[h]:mm:ss", true],
		['_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)', false],
		["0.00%", false],
		["General", false],
		['"Day "0', false],
		["[Red]#,##0", false],
		["0.00E+00", false],
	])("%s → %s", (code, expected) => expect(isDateFormat(code)).toBe(expected));
});
