import { escapeText } from "./xml.js";

/** A value that can be written into a cell. */
export type CellValue = string | number | boolean | bigint | Date | null | undefined | FormulaValue;

/** Writes a formula instead of a value. `formula` has no leading `=`. */
export interface FormulaValue {
	formula: string;
}

export const MAX_CELL_TEXT = 32_767;

const MS_PER_DAY = 86_400_000;
/** 1899-12-30: the 1900 system's epoch once Excel's fictional 1900-02-29 is accounted for. */
const EPOCH_1900 = Date.UTC(1899, 11, 30);
const EPOCH_1904 = Date.UTC(1904, 0, 1);

/**
 * Converts a Date to an Excel serial number from its UTC fields (the mssql
 * driver's default `useUTC: true` hands DATE/DATETIME columns back as UTC).
 */
export function dateToSerial(d: Date, date1904: boolean): number {
	const ms = d.getTime();
	if (date1904) return (ms - EPOCH_1904) / MS_PER_DAY;
	const serial = (ms - EPOCH_1900) / MS_PER_DAY;
	// Before 1900-03-01 Excel's serials are one lower (it pretends 1900 was a leap year).
	return serial < 61 ? serial - 1 : serial;
}

/**
 * Makes a string safe for an OOXML text node: XML-invalid control
 * characters become Excel's `_xHHHH_` escapes, literal `_xHHHH_`
 * sequences are escaped so Excel doesn't decode them, lone surrogates are
 * replaced, and text over Excel's 32,767-character cell limit is cut.
 */
export function sanitizeText(s: string): string {
	let text = s;
	if (text.length > MAX_CELL_TEXT) {
		let cut = MAX_CELL_TEXT;
		const last = text.charCodeAt(cut - 1);
		if (last >= 0xd800 && last <= 0xdbff) cut--;
		text = text.slice(0, cut);
	}
	text = text.replace(/_(x[0-9A-Fa-f]{4}_)/g, "_x005F_$1");
	// eslint-disable-next-line no-control-regex
	text = text.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, (c) => `_x${c.charCodeAt(0).toString(16).toUpperCase().padStart(4, "0")}_`);
	text = text.replace(/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g, "�");
	return text;
}

export function isFormulaValue(v: unknown): v is FormulaValue {
	return typeof v === "object" && v !== null && !(v instanceof Date) && typeof (v as FormulaValue).formula === "string";
}

export interface EncodedCell {
	/** Value for the `t` attribute, or null for none (numbers). */
	t: string | null;
	/** Inner XML of the `<c>` element. Empty for a blank cell. */
	inner: string;
}

export function encodeValue(value: CellValue, date1904: boolean): EncodedCell {
	if (value === null || value === undefined) return { t: null, inner: "" };
	if (typeof value === "string") {
		if (value === "") return { t: null, inner: "" };
		const text = escapeText(sanitizeText(value));
		const space = /^\s|\s$|\n/.test(value) ? ' xml:space="preserve"' : "";
		return { t: "inlineStr", inner: `<is><t${space}>${text}</t></is>` };
	}
	if (typeof value === "number") {
		if (!Number.isFinite(value)) return { t: "e", inner: "<v>#NUM!</v>" };
		return { t: null, inner: `<v>${numberText(value)}</v>` };
	}
	if (typeof value === "bigint") {
		const n = Number(value);
		if (Number.isSafeInteger(n)) return { t: null, inner: `<v>${n}</v>` };
		return encodeValue(value.toString(), date1904);
	}
	if (typeof value === "boolean") return { t: "b", inner: `<v>${value ? 1 : 0}</v>` };
	if (value instanceof Date) {
		if (Number.isNaN(value.getTime())) return { t: "e", inner: "<v>#VALUE!</v>" };
		const serial = dateToSerial(value, date1904);
		if (serial < 0 || serial > 2_958_465) return encodeValue(value.toISOString(), date1904);
		return { t: null, inner: `<v>${numberText(serial)}</v>` };
	}
	if (isFormulaValue(value)) {
		const f = value.formula.startsWith("=") ? value.formula.slice(1) : value.formula;
		return { t: null, inner: `<f>${escapeText(f)}</f>` };
	}
	throw new TypeError(`Unsupported cell value: ${Object.prototype.toString.call(value)}`);
}

function numberText(n: number): string {
	if (Object.is(n, -0)) return "0";
	// Excel stores 15-17 significant digits; JS's shortest round-trip form is exact.
	return String(n);
}
