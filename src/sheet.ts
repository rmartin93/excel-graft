import { colToIndex, MAX_COL, MAX_ROW } from "./a1.js";
import { rewriteRefs } from "./formula.js";
import type { Package } from "./package.js";
import { child, children, decodeXml, findAll, localName, parseXml, type XmlElement } from "./xml.js";

export interface SheetCell {
	el: XmlElement;
	col: number;
	row: number;
	/** `t` attribute (s, str, inlineStr, b, e, n) or undefined. */
	t: string | undefined;
	/** Formula text (no `=`), with shared formulas expanded for followers. */
	formula: string | undefined;
	formulaEl: XmlElement | undefined;
	/** True when the cell holds a non-formula value. */
	hasValue: boolean;
}

export interface SheetRow {
	el: XmlElement;
	r: number;
	cells: SheetCell[];
}

export interface SheetModel {
	part: string;
	xml: string;
	root: XmlElement;
	worksheet: XmlElement;
	sheetData: XmlElement;
	rows: SheetRow[];
	rowMap: Map<number, SheetRow>;
}

const CELL_REF_RE = /^([A-Za-z]{1,3})(\d+)$/;

export function readSheet(pkg: Package, part: string): SheetModel {
	const xml = pkg.text(part);
	const root = parseXml(xml);
	const worksheet = root.children.find((c) => localName(c.name) === "worksheet");
	if (!worksheet) throw new Error(`${part}: no <worksheet> root`);
	const sheetData = worksheet.children.find((c) => localName(c.name) === "sheetData");
	if (!sheetData) throw new Error(`${part}: no <sheetData>`);

	const rows: SheetRow[] = [];
	const masters = new Map<string, { text: string; row: number; col: number }>();
	let lastRow = 0;
	for (const rowEl of sheetData.children) {
		if (localName(rowEl.name) !== "row") continue;
		const r = rowEl.attrs.r === undefined ? lastRow + 1 : Number(rowEl.attrs.r);
		if (!Number.isInteger(r) || r < 1 || r > MAX_ROW) throw new Error(`${part}: bad row number ${rowEl.attrs.r}`);
		lastRow = r;
		const cells: SheetCell[] = [];
		let lastCol = 0;
		for (const cEl of rowEl.children) {
			if (localName(cEl.name) !== "c") continue;
			let col = lastCol + 1;
			if (cEl.attrs.r !== undefined) {
				const m = CELL_REF_RE.exec(cEl.attrs.r);
				if (!m) throw new Error(`${part}: bad cell reference ${cEl.attrs.r}`);
				col = colToIndex(m[1] as string);
			}
			if (col > MAX_COL) throw new Error(`${part}: column out of range in ${cEl.attrs.r}`);
			lastCol = col;
			const fEl = cEl.children.find((c) => localName(c.name) === "f");
			let formula: string | undefined;
			if (fEl) {
				const text = decodeXml(xml.slice(fEl.openEnd, fEl.closeStart));
				if (fEl.attrs.t === "shared" && fEl.attrs.si !== undefined) {
					if (text.length > 0) {
						masters.set(fEl.attrs.si, { text, row: r, col });
						formula = text;
					} else {
						const master = masters.get(fEl.attrs.si);
						formula = master ? shiftFormula(master.text, r - master.row, col - master.col) : undefined;
					}
				} else {
					formula = text;
				}
			}
			const hasValue = !fEl && cEl.children.some((c) => ["v", "is"].includes(localName(c.name)));
			cells.push({ el: cEl, col, row: r, t: cEl.attrs.t, formula, formulaEl: fEl, hasValue });
		}
		rows.push({ el: rowEl, r, cells });
	}
	const rowMap = new Map(rows.map((row) => [row.r, row]));
	return { part, xml, root, worksheet, sheetData, rows, rowMap };
}

/** Shifts the relative parts of every reference, like Excel's copy/fill. */
export function shiftFormula(formula: string, dRow: number, dCol: number): string {
	if (dRow === 0 && dCol === 0) return formula;
	return rewriteRefs(formula, (ref) => {
		if (ref.foreign) return undefined;
		const shiftCell = (c: { col: number; row: number; colAbs: boolean; rowAbs: boolean }) => ({
			...c,
			row: c.rowAbs || ref.shape === "cols" ? c.row : c.row + dRow,
			col: c.colAbs || ref.shape === "rows" ? c.col : c.col + dCol,
		});
		const a = shiftCell(ref.a);
		const b = ref.b ? shiftCell(ref.b) : undefined;
		return formatRefToken(ref.prefix, ref.shape, a, b, ref.spill);
	});
}

export function formatRefToken(
	prefix: string,
	shape: "cell" | "area" | "rows" | "cols",
	a: { col: number; row: number; colAbs: boolean; rowAbs: boolean },
	b: { col: number; row: number; colAbs: boolean; rowAbs: boolean } | undefined,
	spill = "",
): string {
	const bad = (c: { col: number; row: number }) =>
		(shape !== "cols" && (c.row < 1 || c.row > MAX_ROW)) || (shape !== "rows" && (c.col < 1 || c.col > MAX_COL));
	if (bad(a) || (b && bad(b))) return `${prefix}#REF!`;
	const colText = (c: { col: number; colAbs: boolean }) => `${c.colAbs ? "$" : ""}${colLetters(c.col)}`;
	const rowText = (c: { row: number; rowAbs: boolean }) => `${c.rowAbs ? "$" : ""}${c.row}`;
	if (shape === "rows") return `${prefix}${rowText(a)}:${rowText(b ?? a)}`;
	if (shape === "cols") return `${prefix}${colText(a)}:${colText(b ?? a)}`;
	const first = `${colText(a)}${rowText(a)}`;
	if (shape === "cell" || !b) return `${prefix}${first}${spill}`;
	return `${prefix}${first}:${colText(b)}${rowText(b)}${spill}`;
}

function colLetters(n: number): string {
	let s = "";
	while (n > 0) {
		const rem = (n - 1) % 26;
		s = String.fromCharCode(65 + rem) + s;
		n = Math.floor((n - 1) / 26);
	}
	return s;
}

/** The shared string table as plain text per index (rich-text runs concatenated). */
export function readSharedStrings(pkg: Package, workbookPath: string): string[] {
	const rel = pkg
		.rels(workbookPath)
		.find((r) => r.type === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings");
	if (!rel || !pkg.has(rel.target)) return [];
	const xml = pkg.text(rel.target);
	const root = parseXml(xml);
	const sst = root.children[0];
	if (!sst) return [];
	return children(sst, "si").map((si) => stringItemText(xml, si));
}

function stringItemText(xml: string, si: XmlElement): string {
	// Only <t> directly in <si> or inside <r> runs count; <rPh> phonetic text does not.
	const parts: string[] = [];
	for (const c of si.children) {
		const n = localName(c.name);
		if (n === "t") parts.push(decodeXml(xml.slice(c.openEnd, c.closeStart)));
		else if (n === "r") {
			const t = child(c, "t") ?? c.children.find((x) => localName(x.name) === "t");
			if (t) parts.push(decodeXml(xml.slice(t.openEnd, t.closeStart)));
		}
	}
	return decodeText(parts.join(""));
}

/** Decodes Excel's `_xHHHH_` escapes. */
export function decodeText(s: string): string {
	return s.replace(/_x([0-9A-Fa-f]{4})_/g, (_m, hex: string) => String.fromCharCode(parseInt(hex, 16)));
}

/** The displayed text of a string cell, or undefined for non-string cells. */
export function cellText(model: SheetModel, cell: SheetCell, sst: string[]): string | undefined {
	if (cell.t === "s") {
		const v = cell.el.children.find((c) => localName(c.name) === "v");
		if (!v) return undefined;
		return sst[Number(model.xml.slice(v.openEnd, v.closeStart))];
	}
	if (cell.t === "inlineStr") {
		const is = cell.el.children.find((c) => localName(c.name) === "is");
		return is ? stringItemText(model.xml, is) : undefined;
	}
	return undefined;
}

export { findAll };
