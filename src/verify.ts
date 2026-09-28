import { type Area, cellName, MAX_COL, MAX_ROW, parseArea, parseSqref } from "./a1.js";
import { Package, readWorkbook, REL } from "./package.js";
import { cellText, readSharedStrings, readSheet, type SheetModel } from "./sheet.js";
import { MAX_CELL_TEXT } from "./values.js";
import { findAll, localName, parseXml, textOf } from "./xml.js";
import { readParts } from "./zip.js";

export interface VerifyIssue {
	part: string;
	message: string;
}

const VALID_ERRORS = new Set(["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA", "#SPILL!", "#CALC!", "#FIELD!", "#BLOCKED!", "#CONNECT!", "#BUSY!", "#UNKNOWN!", "#PYTHON!"]);

/**
 * Structural checks for the mistakes that make Excel show its repair
 * prompt. Different Excel builds tolerate different subsets of these, so
 * "it opened fine on my machine" is not proof — this is. Returns an empty
 * array for a clean file.
 */
export function verify(input: Uint8Array): VerifyIssue[] {
	const issues: VerifyIssue[] = [];
	const add = (part: string, message: string) => issues.push({ part, message });
	let pkg: Package;
	try {
		pkg = new Package(readParts(input));
	} catch (err) {
		return [{ part: "(zip)", message: `not a readable zip: ${(err as Error).message}` }];
	}

	// Package: well-formed XML, content types, relationships.
	for (const name of pkg.parts.keys()) {
		if (name.endsWith("/")) continue;
		if (/\.(xml|rels)$/i.test(name)) {
			try {
				parseXml(pkg.text(name));
			} catch (err) {
				add(name, `XML is not well-formed: ${(err as Error).message}`);
			}
		}
	}
	if (!pkg.has("[Content_Types].xml")) {
		add("[Content_Types].xml", "missing");
		return issues;
	}
	const ct = parseXml(pkg.text("[Content_Types].xml"));
	const defaults = new Set(findAll(ct, "Default").map((d) => (d.attrs.Extension ?? "").toLowerCase()));
	const overrides = findAll(ct, "Override").map((o) => (o.attrs.PartName ?? "").replace(/^\//, ""));
	for (const o of overrides) if (!pkg.has(o) && !pkg.has(decodeURIComponent(o))) add("[Content_Types].xml", `Override for missing part /${o}`);
	const overrideSet = new Set(overrides);
	for (const name of pkg.parts.keys()) {
		if (name.endsWith("/") || name === "[Content_Types].xml") continue;
		const ext = name.slice(name.lastIndexOf(".") + 1).toLowerCase();
		if (!overrideSet.has(name) && !defaults.has(ext)) add("[Content_Types].xml", `no content type for ${name}`);
	}
	for (const name of pkg.parts.keys()) {
		if (!name.endsWith(".rels")) continue;
		const source = name.replace(/_rels\/([^/]*)\.rels$/, "$1");
		const ids = new Set<string>();
		for (const rel of pkg.rels(source)) {
			if (ids.has(rel.id)) add(name, `duplicate relationship id ${rel.id}`);
			ids.add(rel.id);
			if (!rel.external && !pkg.has(rel.target) && !pkg.has(decodeURIComponent(rel.target))) {
				add(name, `relationship ${rel.id} points at missing part ${rel.target}`);
			}
		}
	}

	let wb;
	try {
		wb = readWorkbook(pkg);
	} catch (err) {
		add("xl/workbook.xml", (err as Error).message);
		return issues;
	}
	const sst = readSharedStrings(pkg, wb.path);
	for (const [i, s] of sst.entries()) if (s.length > MAX_CELL_TEXT) add("xl/sharedStrings.xml", `string ${i} is ${s.length} characters (limit ${MAX_CELL_TEXT})`);

	const names = new Set<string>();
	for (const s of wb.sheets) {
		const k = s.name.toLowerCase();
		if (names.has(k)) add(wb.path, `duplicate sheet name "${s.name}"`);
		names.add(k);
	}
	const dn = new Set<string>();
	for (const n of wb.names) {
		const k = `${n.name.toLowerCase()}|${n.localSheetId ?? ""}`;
		if (dn.has(k)) add(wb.path, `duplicate defined name "${n.name}"`);
		dn.add(k);
		if (n.formula.length > 8192) add(wb.path, `defined name "${n.name}" formula is too long`);
	}

	const stylesRel = pkg.rels(wb.path).find((r) => r.type.endsWith("/styles"));
	let xfCount = Infinity;
	if (stylesRel && pkg.has(stylesRel.target)) {
		const cellXfs = findAll(parseXml(pkg.text(stylesRel.target)), "cellXfs")[0];
		if (cellXfs) xfCount = cellXfs.children.filter((c) => localName(c.name) === "xf").length;
	}

	const models = new Map<string, SheetModel>();
	for (const sheet of wb.sheets) {
		if (!pkg.has(sheet.part)) continue;
		let model: SheetModel;
		try {
			model = readSheet(pkg, sheet.part);
		} catch (err) {
			add(sheet.part, (err as Error).message);
			continue;
		}
		models.set(sheet.part, model);
		checkSheet(model, sst, xfCount, add);
	}

	// Tables.
	const tableIds = new Set<string>();
	const tableNames = new Set<string>();
	const tableAreas = new Map<string, { name: string; area: Area }[]>();
	for (const table of wb.tables) {
		const xml = pkg.text(table.part);
		const el = findAll(parseXml(xml), "table")[0];
		if (!el) continue;
		const id = el.attrs.id ?? "";
		if (tableIds.has(id)) add(table.part, `duplicate table id ${id}`);
		tableIds.add(id);
		const nm = (el.attrs.displayName ?? el.attrs.name ?? "").toLowerCase();
		if (tableNames.has(nm)) add(table.part, `duplicate table name ${nm}`);
		tableNames.add(nm);
		const area = parseArea(el.attrs.ref ?? "");
		if (!area) {
			add(table.part, `bad table ref "${el.attrs.ref}"`);
			continue;
		}
		const list = tableAreas.get(table.sheet.part) ?? [];
		for (const other of list) {
			if (overlaps(area, other.area)) add(table.part, `table ${nm} overlaps table ${other.name}`);
		}
		list.push({ name: nm, area });
		tableAreas.set(table.sheet.part, list);
		const headerRows = el.attrs.headerRowCount === "0" ? 0 : 1;
		const totals = Number(el.attrs.totalsRowCount ?? 0);
		if (area.r2 - area.r1 + 1 - headerRows - totals < 1 && el.attrs.insertRow !== "1") add(table.part, "table has no data rows");
		const af = el.children.find((c) => localName(c.name) === "autoFilter");
		if (af) {
			const expected = { ...area, r2: area.r2 - totals };
			const got = parseArea(af.attrs.ref ?? "");
			if (!got || got.r1 !== expected.r1 || got.r2 !== expected.r2 || got.c1 !== expected.c1 || got.c2 !== expected.c2) {
				add(table.part, `autoFilter ref ${af.attrs.ref} should be ${cellName(expected.c1, expected.r1)}:${cellName(expected.c2, expected.r2)} (table ref minus totals rows)`);
			}
		}
		const cols = findAll(el, "tableColumn");
		if (cols.length !== area.c2 - area.c1 + 1) add(table.part, `${cols.length} tableColumns for a ${area.c2 - area.c1 + 1}-column ref`);
		const model = models.get(table.sheet.part);
		if (model && headerRows) {
			const row = model.rowMap.get(area.r1);
			cols.forEach((col, i) => {
				const cell = row?.cells.find((c) => c.col === area.c1 + i);
				const text = cell ? cellText(model, cell, sst) ?? (cell.hasValue ? rawValue(model, cell) : "") : "";
				const expected = (col.attrs.name ?? "").replace(/_x([0-9A-Fa-f]{4})_/g, (_m, h: string) => String.fromCharCode(parseInt(h, 16)));
				if (text !== expected) add(table.part, `header cell ${cellName(area.c1 + i, area.r1)} is "${text}" but the column is named "${expected}"`);
			});
		}
		if (model) {
			for (const m of mergeAreas(model)) if (overlaps(m, area)) add(table.sheet.part, `merged cells overlap table ${nm}`);
		}
	}

	// calcChain must only list formula cells.
	const calcRel = pkg.rels(wb.path).find((r) => r.type === REL.calcChain);
	if (calcRel && pkg.has(calcRel.target)) {
		const xml = pkg.text(calcRel.target);
		let sheetIndex = 0;
		const idToSheet = new Map<number, SheetModel>();
		const wbXml = parseXml(pkg.text(wb.path));
		const sheetEls = findAll(wbXml, "sheet");
		for (const s of wb.sheets) {
			const el = sheetEls[s.index];
			const model = models.get(s.part);
			if (el && model) idToSheet.set(Number(el.attrs.sheetId), model);
		}
		for (const c of findAll(parseXml(xml), "c")) {
			if (c.attrs.i !== undefined) sheetIndex = Number(c.attrs.i);
			const model = idToSheet.get(sheetIndex);
			const a = parseArea(c.attrs.r ?? "");
			if (!model || !a) continue;
			const cell = model.rowMap.get(a.r1)?.cells.find((x) => x.col === a.c1);
			if (!cell?.formulaEl) add(calcRel.target, `calcChain lists ${c.attrs.r} on sheet ${sheetIndex}, which has no formula`);
		}
	}

	// Drawing anchors.
	for (const name of pkg.parts.keys()) {
		if (!/^xl\/drawings\/drawing\d+\.xml$/.test(name)) continue;
		const root = parseXml(pkg.text(name));
		for (const anchor of findAll(root, "xdr:twoCellAnchor")) {
			const row = (n: string) => {
				const el = anchor.children.find((c) => localName(c.name) === n)?.children.find((c) => localName(c.name) === "row");
				return el ? Number(textOf(pkg.text(name), el)) : undefined;
			};
			const f = row("from");
			const t = row("to");
			if (f !== undefined && t !== undefined && t < f) add(name, `drawing anchor ends (row ${t}) before it starts (row ${f})`);
			if (f !== undefined && (f < 0 || f >= MAX_ROW)) add(name, `drawing anchor row ${f} out of range`);
		}
	}
	return issues;
}

function rawValue(model: SheetModel, cell: { el: { children: { name: string; openEnd: number; closeStart: number }[] } }): string {
	const v = cell.el.children.find((c) => localName(c.name) === "v");
	return v ? model.xml.slice(v.openEnd, v.closeStart) : "";
}

function overlaps(a: Area, b: Area): boolean {
	return a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2;
}

function mergeAreas(model: SheetModel): Area[] {
	const mc = model.worksheet.children.find((c) => localName(c.name) === "mergeCells");
	if (!mc) return [];
	return mc.children.map((m) => parseArea(m.attrs.ref ?? "")).filter((a): a is Area => a !== undefined);
}

function checkSheet(model: SheetModel, sst: string[], xfCount: number, add: (part: string, message: string) => void): void {
	const part = model.part;
	let lastRow = 0;
	const masters = new Set<string>();
	for (const row of model.rows) {
		if (row.r <= lastRow) add(part, `row ${row.r} is out of order or duplicated (after row ${lastRow})`);
		lastRow = row.r;
		if (row.el.attrs.s !== undefined && Number(row.el.attrs.s) >= xfCount) add(part, `row ${row.r} uses missing style ${row.el.attrs.s}`);
		let lastCol = 0;
		for (const cell of row.cells) {
			const ref = cell.el.attrs.r;
			if (ref !== undefined) {
				const a = parseArea(ref);
				if (!a || a.r1 !== row.r) add(part, `cell ${ref} sits in row ${row.r}`);
			}
			if (cell.col <= lastCol) add(part, `cell ${ref ?? cell.col} in row ${row.r} is out of order or duplicated`);
			lastCol = cell.col;
			if (cell.el.attrs.s !== undefined && Number(cell.el.attrs.s) >= xfCount) add(part, `cell ${ref} uses missing style ${cell.el.attrs.s}`);
			const t = cell.t ?? "n";
			const vEl = cell.el.children.find((c) => localName(c.name) === "v");
			const v = vEl ? model.xml.slice(vEl.openEnd, vEl.closeStart) : undefined;
			if (!["n", "s", "str", "inlineStr", "b", "e", "d"].includes(t)) add(part, `cell ${ref} has unknown type t="${t}"`);
			if (t === "s" && v !== undefined && !(Number(v) >= 0 && Number(v) < sst.length)) add(part, `cell ${ref} points at shared string ${v} of ${sst.length}`);
			if (t === "inlineStr" && !cell.el.children.some((c) => localName(c.name) === "is") && !cell.formulaEl) add(part, `cell ${ref} is inlineStr without <is>`);
			if (t === "b" && v !== undefined && v !== "0" && v !== "1") add(part, `cell ${ref} boolean value "${v}"`);
			if (t === "e" && v !== undefined && !VALID_ERRORS.has(v)) add(part, `cell ${ref} error value "${v}"`);
			if (t === "n" && v !== undefined && (v.trim() === "" || Number.isNaN(Number(v)))) add(part, `cell ${ref} numeric value "${v}"`);
			if (t === "inlineStr" || t === "str") {
				const text = cellText(model, cell, sst) ?? "";
				if (text.length > MAX_CELL_TEXT) add(part, `cell ${ref} text is ${text.length} characters (limit ${MAX_CELL_TEXT})`);
			}
			if (cell.formulaEl) {
				const f = model.xml.slice(cell.formulaEl.openEnd, cell.formulaEl.closeStart);
				if (f.length > 8192) add(part, `cell ${ref} formula is ${f.length} characters (limit 8192)`);
				const ft = cell.formulaEl.attrs.t;
				const si = cell.formulaEl.attrs.si;
				if (ft === "shared" && si !== undefined) {
					if (f.length > 0) masters.add(si);
					else if (!masters.has(si)) add(part, `cell ${ref} follows shared formula ${si}, which has no master before it`);
				}
			}
		}
	}

	const ws = model.worksheet;
	const merges = mergeAreas(model);
	for (let i = 0; i < merges.length; i++) {
		const a = merges[i] as Area;
		if (a.r1 === a.r2 && a.c1 === a.c2) add(part, `merge ${cellName(a.c1, a.r1)} is a single cell`);
		for (let j = i + 1; j < merges.length; j++) if (overlaps(a, merges[j] as Area)) add(part, `merged ranges ${fmt(a)} and ${fmt(merges[j] as Area)} overlap`);
	}
	const mc = ws.children.find((c) => localName(c.name) === "mergeCells");
	if (mc?.attrs.count !== undefined && Number(mc.attrs.count) !== mc.children.length) add(part, `mergeCells count ${mc.attrs.count} but ${mc.children.length} entries`);
	const dv = ws.children.find((c) => localName(c.name) === "dataValidations");
	if (dv?.attrs.count !== undefined && Number(dv.attrs.count) !== dv.children.length) add(part, `dataValidations count ${dv.attrs.count} but ${dv.children.length} entries`);

	const sqrefHolders = [...findAll(ws, "conditionalFormatting"), ...findAll(ws, "dataValidation"), ...findAll(ws, "selection"), ...findAll(ws, "ignoredError"), ...findAll(ws, "protectedRange")];
	for (const h of sqrefHolders) {
		if (h.attrs.sqref === undefined) continue;
		const areas = parseSqref(h.attrs.sqref);
		if (areas.length === 0 && localName(h.name) !== "selection") add(part, `empty sqref on <${h.name}>`);
		for (const a of areas) if (!a || a.r2 > MAX_ROW || a.c2 > MAX_COL) add(part, `bad sqref "${h.attrs.sqref}" on <${h.name}>`);
	}
	for (const sq of findAll(ws, "xm:sqref")) {
		for (const a of parseSqref(textOf(model.xml, sq))) if (!a) add(part, `bad xm:sqref "${textOf(model.xml, sq)}"`);
	}
	for (const h of findAll(ws, "hyperlink")) {
		if (!parseArea(h.attrs.ref ?? "")) add(part, `bad hyperlink ref "${h.attrs.ref}"`);
	}
	const dim = ws.children.find((c) => localName(c.name) === "dimension");
	if (dim?.attrs.ref && !parseArea(dim.attrs.ref)) add(part, `bad dimension "${dim.attrs.ref}"`);
}

function fmt(a: Area): string {
	return `${cellName(a.c1, a.r1)}:${cellName(a.c2, a.r2)}`;
}
