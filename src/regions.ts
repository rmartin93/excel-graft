import { indexToCol, parseArea } from "./a1.js";
import { rewriteRefs, tokenize } from "./formula.js";
import type { DefinedName, SheetInfo, TableInfo, WorkbookInfo } from "./package.js";
import type { Package } from "./package.js";
import type { StyleInfo } from "./styles.js";
import { cellText, decodeText, type SheetCell, type SheetModel } from "./sheet.js";
import { children, findAll, parseXml } from "./xml.js";

/**
 * A role is "what kind of row this is" inside a region: every sample row
 * of the same kind (e.g. every group header) shares one role, and every
 * output row is instantiated from a role's prototype.
 */
export interface Role {
	id: number;
	kind: "leaf" | "header" | "footer" | "spacer" | "fixed";
	/** Every template row with this role, across all sample groups. */
	rows: number[];
}

export type FieldType = "string" | "number" | "date" | "boolean" | "unknown";

export interface FieldDef {
	name: string;
	col: number;
	/** Inferred from the sample cells' values and number formats. */
	type: FieldType;
	/** For group label fields: the sample groups' values, in order (e.g. Fringe, Overhead, G&A). */
	samples?: string[];
	/** Fixed text around a group label that every sample shares ("Engineering subtotal" -> suffix " subtotal"). */
	affix?: { prefix: string; suffix: string };
}

export interface LeafPattern {
	kind: "leaf";
	role: Role;
	/** Prototype rows, used cyclically (length 2 for banded rows). */
	/** Sample row whose formulas and constants every detail row copies. */
	formulaProto: number;
	/** A different formula source for the first detail row (e.g. an opening balance `=C5` before `=D5+C6`). */
	firstFormulaProto: number | undefined;
	/** Sample rows whose styles detail rows cycle through (two for banded rows). */
	styleCycle: number[];
	/** A different style source for the first detail row (e.g. a heavier top border). */
	firstStyle: number | undefined;
	fields: FieldDef[];
}

export interface GroupPattern {
	kind: "group";
	header: { role: Role; proto: number; fields: FieldDef[] } | undefined;
	footer: { role: Role; proto: number; fields: FieldDef[] } | undefined;
	spacers: { role: Role; proto: number }[];
	child: Pattern;
	childKey: string;
}

export type Pattern = LeafPattern | GroupPattern;

export interface Region {
	key: string;
	kind: "table" | "name";
	sheet: SheetInfo;
	table: TableInfo | undefined;
	name: DefinedName | undefined;
	c1: number;
	c2: number;
	headerRow: number | undefined;
	bodyStart: number;
	bodyEnd: number;
	/** For tables: the rows after the body that belong to the table (totals). */
	totalsRows: number;
	columnNames: Map<number, string>;
	leading: Role[];
	trailing: Role[];
	unit: Pattern;
	roleOfRow: Map<number, Role>;
	/** For sample group header/footer rows: their label path, e.g. "Sales › Travel". */
	sampleLabels: Map<number, string>;
	/** Every sample group: its label row (header, else footer) and its rows. */
	sampleBlocks: { row: number; first: number; last: number }[];
}

export interface ScalarTarget {
	key: string;
	kind: "placeholder" | "name";
	sheet: SheetInfo;
	col: number;
	row: number;
}

export class TemplateStructureError extends Error {
	override name = "TemplateStructureError";
}

interface RowInfo {
	r: number;
	empty: boolean;
	covered: Set<number>;
	aggregate: boolean;
	ext: [number, number];
}

export function discoverRegions(
	pkg: Package,
	wb: WorkbookInfo,
	sheets: Map<string, SheetModel>,
	sst: string[],
	styles: StyleInfo,
): { regions: Region[]; skipped: { key: string; reason: string }[] } {
	const regions: Region[] = [];
	const skipped: { key: string; reason: string }[] = [];

	for (const table of wb.tables) {
		const model = sheets.get(table.sheet.part);
		if (!model) continue;
		const xml = pkg.text(table.part);
		const tableEl = children(parseXml(xml), "table")[0];
		if (!tableEl) continue;
		const area = parseArea(tableEl.attrs.ref ?? "");
		if (!area) continue;
		const headerRows = tableEl.attrs.headerRowCount === "0" ? 0 : 1;
		const totalsRows = Number(tableEl.attrs.totalsRowCount ?? 0);
		const columnNames = new Map<number, string>();
		findAll(tableEl, "tableColumn").forEach((col, i) => {
			columnNames.set(area.c1 + i, decodeText(col.attrs.name ?? indexToCol(area.c1 + i)));
		});
		const bodyStart = area.r1 + headerRows;
		const bodyEnd = area.r2 - totalsRows;
		if (bodyEnd < bodyStart) {
			skipped.push({ key: table.displayName, reason: "table has no data rows" });
			continue;
		}
		try {
			regions.push(
				buildRegion({
					key: table.displayName,
					kind: "table",
					sheet: table.sheet,
					table,
					name: undefined,
					c1: area.c1,
					c2: area.c2,
					headerRow: headerRows ? area.r1 : undefined,
					bodyStart,
					bodyEnd,
					totalsRows,
					columnNames,
					model,
					styles,
					sst,
				}),
			);
		} catch (err) {
			if (err instanceof TemplateStructureError) skipped.push({ key: table.displayName, reason: err.message });
			else throw err;
		}
	}

	for (const dn of wb.names) {
		const target = nameTarget(dn, wb);
		if (!target || target.area.r1 === target.area.r2 && target.area.c1 === target.area.c2) continue;
		const { sheet, area } = target;
		const model = sheets.get(sheet.part);
		if (!model) continue;
		const overlapsTable = regions.some(
			(r) => r.sheet === sheet && r.kind === "table" && area.r1 <= r.bodyEnd + r.totalsRows && area.r2 >= (r.headerRow ?? r.bodyStart),
		);
		if (overlapsTable) continue;
		const headerRow = area.r1 > 1 ? area.r1 - 1 : undefined;
		const columnNames = new Map<number, string>();
		const used = new Set<string>();
		for (let c = area.c1; c <= area.c2; c++) {
			const cell = headerRow === undefined ? undefined : model.rowMap.get(headerRow)?.cells.find((x) => x.col === c);
			let label = cell ? (cellText(model, cell, sst) ?? "").trim() : "";
			if (label === "" || used.has(label)) label = label === "" ? indexToCol(c) : `${label} (${indexToCol(c)})`;
			used.add(label);
			columnNames.set(c, label);
		}
		try {
			regions.push(
				buildRegion({
					key: dn.name,
					kind: "name",
					sheet,
					table: undefined,
					name: dn,
					c1: area.c1,
					c2: area.c2,
					headerRow,
					bodyStart: area.r1,
					bodyEnd: area.r2,
					totalsRows: 0,
					columnNames,
					model,
					styles,
					sst,
				}),
			);
		} catch (err) {
			if (err instanceof TemplateStructureError) skipped.push({ key: dn.name, reason: err.message });
			else throw err;
		}
	}
	return { regions, skipped };
}

/** Resolves a defined name to a single rectangular area on one sheet, or undefined. */
export function nameTarget(dn: DefinedName, wb: WorkbookInfo): { sheet: SheetInfo; area: { c1: number; r1: number; c2: number; r2: number } } | undefined {
	if (dn.hidden || dn.name.startsWith("_xlnm.") || dn.name.startsWith("_")) return undefined;
	const tokens = tokenize(dn.formula);
	const significant = tokens.filter((t) => t.kind !== "ws");
	const only = significant[0];
	if (significant.length !== 1 || !only || only.kind !== "ref" || only.foreign || !only.sheet) return undefined;
	if (only.shape !== "cell" && only.shape !== "area") return undefined;
	const sheet = wb.sheets.find((s) => s.name === only.sheet);
	if (!sheet) return undefined;
	const b = only.b ?? only.a;
	return {
		sheet,
		area: {
			c1: Math.min(only.a.col, b.col),
			r1: Math.min(only.a.row, b.row),
			c2: Math.max(only.a.col, b.col),
			r2: Math.max(only.a.row, b.row),
		},
	};
}

interface RegionInput {
	key: string;
	kind: "table" | "name";
	sheet: SheetInfo;
	table: TableInfo | undefined;
	name: DefinedName | undefined;
	c1: number;
	c2: number;
	headerRow: number | undefined;
	bodyStart: number;
	bodyEnd: number;
	totalsRows: number;
	columnNames: Map<number, string>;
	model: SheetModel;
	styles: StyleInfo;
	sst: string[];
}

function buildRegion(input: RegionInput): Region {
	const { model, bodyStart, bodyEnd } = input;
	const info = new Map<number, RowInfo>();
	for (let r = bodyStart; r <= bodyEnd; r++) info.set(r, rowInfo(model, input.sheet.name, r, bodyStart, bodyEnd));
	// A total over group rows reaches everything those groups reach: SUM(C8,C12) over two
	// footer subtotals covers both departments, not just rows 8-12.
	for (let changed = true; changed; ) {
		changed = false;
		for (const row of info.values()) {
			if (!row.aggregate) continue;
			for (const c of row.covered) {
				const other = info.get(c) as RowInfo;
				if (!other.aggregate) continue;
				const lo = Math.min(row.ext[0], other.ext[0]);
				const hi = Math.max(row.ext[1], other.ext[1]);
				if (lo !== row.ext[0] || hi !== row.ext[1]) {
					row.ext = [lo, hi];
					changed = true;
				}
			}
		}
	}

	let roleId = 0;
	const roleOfRow = new Map<number, Role>();
	const newRole = (kind: Role["kind"], rows: number[]): Role => {
		const role = { id: roleId++, kind, rows };
		for (const r of rows) roleOfRow.set(r, role);
		return role;
	};

	// Peel fixed summary rows off the ends: a summary footer is an aggregate
	// covering all content above it with no other aggregates below it.
	let lo = bodyStart;
	let hi = bodyEnd;
	const leading: Role[] = [];
	const trailing: Role[] = [];
	const isContentCovered = (row: RowInfo, from: number, to: number): boolean => {
		for (let r = from; r <= to; r++) {
			const x = info.get(r) as RowInfo;
			if (!x.empty && (r < row.ext[0] || r > row.ext[1])) return false;
		}
		return true;
	};
	for (let changed = true; changed && lo <= hi; ) {
		changed = false;
		for (let f = hi; f > lo; f--) {
			const row = info.get(f) as RowInfo;
			if (!row.aggregate) continue;
			let aggBelow = false;
			for (let r = f + 1; r <= hi; r++) if ((info.get(r) as RowInfo).aggregate) aggBelow = true;
			if (aggBelow) break;
			if (isContentCovered(row, lo, f - 1) && hasContent(info, lo, f - 1)) {
				for (let r = hi; r >= f; r--) trailing.unshift(newRole("fixed", [r]));
				hi = f - 1;
				changed = true;
			}
			break;
		}
		for (let h = lo; h < hi; h++) {
			const row = info.get(h) as RowInfo;
			if (!row.aggregate) continue;
			let aggAbove = false;
			for (let r = lo; r < h; r++) if ((info.get(r) as RowInfo).aggregate) aggAbove = true;
			if (aggAbove) break;
			if (isContentCovered(row, h + 1, hi) && hasContent(info, h + 1, hi)) {
				for (let r = lo; r <= h; r++) leading.push(newRole("fixed", [r]));
				lo = h + 1;
				changed = true;
			}
			break;
		}
	}
	if (lo > hi) throw new TemplateStructureError(`${input.key}: no repeatable sample rows found`);

	const seq = parseSeq(info, lo, hi, input.key);
	const ownLabels = new Map<number, string>();
	const sampleBlocks: BlockNode[] = [];
	const unit = toPattern([seq], input, newRole, ownLabels, sampleBlocks);
	// Label paths: a nested group is identified by its parents' labels too (Sales › Travel, not just Travel).
	const sampleLabels = new Map<number, string>();
	for (const b of sampleBlocks) {
		const own = b.header ?? b.footer;
		if (own === undefined) continue;
		const chain = sampleBlocks
			.filter((a) => a !== b && a.first <= b.first && b.last <= a.last)
			.sort((x, y) => y.last - y.first - (x.last - x.first))
			.map((a) => ownLabels.get((a.header ?? a.footer) as number) ?? "");
		const path = [...chain, ownLabels.get(own) ?? ""].join(" › ");
		if (path.replace(/ › /g, "") === "") continue;
		for (const r of [b.header, b.footer]) if (r !== undefined) sampleLabels.set(r, path);
	}
	return {
		key: input.key,
		kind: input.kind,
		sheet: input.sheet,
		table: input.table,
		name: input.name,
		c1: input.c1,
		c2: input.c2,
		headerRow: input.headerRow,
		bodyStart,
		bodyEnd,
		totalsRows: input.totalsRows,
		columnNames: input.columnNames,
		leading,
		trailing,
		unit,
		roleOfRow,
		sampleLabels,
		sampleBlocks: sampleBlocks.map((b) => ({ row: (b.header ?? b.footer) as number, first: b.first, last: b.last })),
	};
}

function hasContent(info: Map<number, RowInfo>, from: number, to: number): boolean {
	for (let r = from; r <= to; r++) if (!(info.get(r) as RowInfo).empty) return true;
	return false;
}

function rowInfo(model: SheetModel, sheetName: string, r: number, bodyStart: number, bodyEnd: number): RowInfo {
	const row = model.rowMap.get(r);
	const covered = new Set<number>();
	let empty = true;
	for (const cell of row?.cells ?? []) {
		if (cell.hasValue || cell.formula !== undefined) empty = false;
		if (cell.formula === undefined) continue;
		rewriteRefs(cell.formula, (ref, ctx) => {
			if (ref.foreign || (ref.sheet !== undefined && ref.sheet !== sheetName)) return undefined;
			if (ref.shape === "cols") return undefined;
			const b = ref.b ?? ref.a;
			const ra = Math.min(ref.a.row, b.row);
			const rb = Math.max(ref.a.row, b.row);
			const isRange = ref.shape === "area" || ref.shape === "rows";
			if (isRange ? r >= ra && r <= rb : !ctx.inArgList) return undefined;
			for (let x = Math.max(ra, bodyStart); x <= Math.min(rb, bodyEnd); x++) if (x !== r) covered.add(x);
			return undefined;
		});
	}
	const all = [...covered, r];
	return {
		r,
		empty,
		covered,
		aggregate: covered.size > 0,
		ext: [Math.min(...all), Math.max(...all)],
	};
}

type SeqNode = { kind: "leaf"; rows: number[] } | { kind: "groups"; blocks: BlockNode[] };

interface BlockNode {
	first: number;
	last: number;
	header: number | undefined;
	footer: number | undefined;
	spacers: number[];
	child: SeqNode;
}

function parseSeq(info: Map<number, RowInfo>, lo: number, hi: number, key: string): SeqNode {
	const aggs = [...info.values()].filter((x) => x.r >= lo && x.r <= hi && x.aggregate);
	if (aggs.length === 0) {
		const rows: number[] = [];
		for (let r = lo; r <= hi; r++) rows.push(r);
		return { kind: "leaf", rows };
	}
	for (const a of aggs) {
		if (a.ext[0] < lo || a.ext[1] > hi) {
			throw new TemplateStructureError(`${key}: the formula in row ${a.r} covers rows outside its group`);
		}
	}
	// Top-level aggregates: not strictly inside another aggregate's extent.
	const top = aggs.filter(
		(a) => !aggs.some((b) => b !== a && b.ext[0] <= a.ext[0] && b.ext[1] >= a.ext[1] && (b.ext[0] < a.ext[0] || b.ext[1] > a.ext[1])),
	);
	const blocks: [number, number][] = [];
	for (const a of top.sort((x, y) => x.ext[0] - y.ext[0])) {
		const last = blocks[blocks.length - 1];
		if (last && last[0] === a.ext[0] && last[1] === a.ext[1]) continue;
		if (last && a.ext[0] <= last[1]) throw new TemplateStructureError(`${key}: group formulas in rows ${last[0]}-${a.ext[1]} overlap`);
		blocks.push([a.ext[0], a.ext[1]]);
	}
	const nodes: BlockNode[] = [];
	let cursor = lo;
	for (const [s, e] of blocks) {
		for (let r = cursor; r < s; r++) {
			const x = info.get(r) as RowInfo;
			if (!x.empty) {
				throw new TemplateStructureError(
					`${key}: row ${r} is neither inside a group nor blank; can't tell how it repeats alongside the groups`,
				);
			}
			const prev = nodes[nodes.length - 1];
			if (!prev) throw new TemplateStructureError(`${key}: blank row ${r} before the first group`);
			prev.spacers.push(r);
		}
		const header = (info.get(s) as RowInfo).aggregate && (info.get(s) as RowInfo).ext[1] === e ? s : undefined;
		const footer = e !== header && (info.get(e) as RowInfo).aggregate && (info.get(e) as RowInfo).ext[0] === s ? e : undefined;
		if (header === undefined && footer === undefined) {
			throw new TemplateStructureError(`${key}: rows ${s}-${e} form a group but neither the first nor last row totals it`);
		}
		const innerLo = header === undefined ? s : s + 1;
		const innerHi = footer === undefined ? e : e - 1;
		if (innerLo > innerHi) throw new TemplateStructureError(`${key}: group in rows ${s}-${e} has no detail rows`);
		nodes.push({ first: s, last: e, header, footer, spacers: [], child: parseSeq(info, innerLo, innerHi, key) });
		cursor = e + 1;
	}
	for (let r = cursor; r <= hi; r++) {
		if (!(info.get(r) as RowInfo).empty) {
			throw new TemplateStructureError(`${key}: row ${r} is neither inside a group nor blank`);
		}
		nodes[nodes.length - 1]?.spacers.push(r);
	}
	return { kind: "groups", blocks: nodes };
}

function toPattern(
	seqs: SeqNode[],
	input: RegionInput,
	newRole: (kind: Role["kind"], rows: number[]) => Role,
	labels: Map<number, string>,
	blocksOut: BlockNode[],
): Pattern {
	const first = seqs[0] as SeqNode;
	if (seqs.some((s) => s.kind !== first.kind)) {
		throw new TemplateStructureError(`${input.key}: sample groups have different structures (some have sub-groups, some don't)`);
	}
	if (first.kind === "leaf") {
		const allRows = seqs.flatMap((s) => (s.kind === "leaf" ? s.rows : []));
		const role = newRole("leaf", allRows);
		const sample = first.rows.filter((r) => !isEmptyRow(input.model, r));
		const candidates = sample.length > 0 ? sample : first.rows;
		const protos = chooseProtos(input.model, candidates);
		return { kind: "leaf", role, ...protos, fields: leafFields(input, protos.formulaProto, allRows) };
	}
	const blocks = seqs.flatMap((s) => (s.kind === "groups" ? s.blocks : []));
	blocksOut.push(...blocks);
	const b0 = blocks[0] as BlockNode;
	const hasHeader = b0.header !== undefined;
	const hasFooter = b0.footer !== undefined;
	if (blocks.some((b) => (b.header !== undefined) !== hasHeader || (b.footer !== undefined) !== hasFooter)) {
		throw new TemplateStructureError(`${input.key}: sample groups disagree on whether they have a header or footer total row`);
	}
	const header = hasHeader
		? {
				role: newRole("header", blocks.map((b) => b.header as number)),
				proto: b0.header as number,
				fields: constantFields(input, b0.header as number, blocks.map((b) => b.header as number)),
			}
		: undefined;
	const footer = hasFooter
		? {
				role: newRole("footer", blocks.map((b) => b.footer as number)),
				proto: b0.footer as number,
				fields: constantFields(input, b0.footer as number, blocks.map((b) => b.footer as number)),
			}
		: undefined;
	const maxSpacers = Math.max(...blocks.map((b) => b.spacers.length));
	const spacers: { role: Role; proto: number }[] = [];
	for (let i = 0; i < maxSpacers && i < b0.spacers.length; i++) {
		spacers.push({
			role: newRole("spacer", blocks.map((b) => b.spacers[i]).filter((r): r is number => r !== undefined)),
			proto: b0.spacers[i] as number,
		});
	}
	for (const part of [header, footer]) {
		for (const f of part?.fields ?? []) {
			const rows = blocks.map((b) => (part === header ? b.header : b.footer) as number);
			const raw = rows.map((r) => cellDisplay(input, r, f.col)).filter((x) => x !== "");
			if (f.type === "string") {
				const affix = commonAffix(raw);
				if (affix) f.affix = affix;
			}
			f.samples = raw.map((x) => stripAffix(x, f.affix));
		}
	}
	for (const b of blocks) {
		const r = b.header ?? b.footer;
		const part = b.header !== undefined ? header : footer;
		if (r === undefined || !part) continue;
		labels.set(r, rowLabel(input, r, part.fields));
	}
	const child = toPattern(
		blocks.map((b) => b.child),
		input,
		newRole,
		labels,
		blocksOut,
	);
	const fieldNames = new Set([...(header?.fields ?? []), ...(footer?.fields ?? [])].map((f) => f.name));
	return { kind: "group", header, footer, spacers, child, childKey: fieldNames.has("rows") ? "children" : "rows" };
}

function isEmptyRow(model: SheetModel, r: number): boolean {
	const row = model.rowMap.get(r);
	return !row || row.cells.every((c) => !c.hasValue && c.formula === undefined);
}

/** Style + relative-formula signature, for spotting banded rows and a special first row. */
function rowSignature(model: SheetModel, r: number): { style: string; formulas: string } {
	const row = model.rowMap.get(r);
	const style = [`row:${row?.el.attrs.s ?? ""}`, ...(row?.cells ?? []).map((c) => `${c.col}:${c.el.attrs.s ?? ""}`)].join("|");
	const formulas = (row?.cells ?? [])
		.filter((c) => c.formula !== undefined)
		.map((c) => `${c.col}:${relativeForm(c)}`)
		.join("|");
	return { style, formulas };
}

function relativeForm(cell: SheetCell): string {
	return rewriteRefs(cell.formula as string, (ref) => {
		const fmt = (c: { col: number; row: number; colAbs: boolean; rowAbs: boolean }) =>
			`R${c.rowAbs ? c.row : `[${c.row - cell.row}]`}C${c.colAbs ? c.col : `[${c.col - cell.col}]`}`;
		return `${ref.prefix}${fmt(ref.a)}${ref.b ? `:${fmt(ref.b)}` : ""}${ref.spill}`;
	});
}

/**
 * Picks prototypes for detail rows. Formulas and styles are chosen
 * independently, so a sample whose first row opens a running balance AND
 * whose rows are banded keeps both behaviours.
 */
function chooseProtos(
	model: SheetModel,
	rows: number[],
): { formulaProto: number; firstFormulaProto: number | undefined; styleCycle: number[]; firstStyle: number | undefined } {
	const r0 = rows[0] as number;
	const sigs = rows.map((r) => rowSignature(model, r));
	const pick = <K extends "style" | "formulas">(key: K) => sigs.map((s) => s[key]);

	// Formulas: all alike, or a distinct first row followed by rows that agree.
	const f = pick("formulas");
	let formulaProto = r0;
	let firstFormulaProto: number | undefined;
	if (rows.length >= 2 && f.slice(1).every((x) => x === f[1]) && f[0] !== f[1]) {
		formulaProto = rows[1] as number;
		firstFormulaProto = r0;
	}

	// Styles: all alike, banded (A,B,A,B...), or a distinct first row followed by rows that agree.
	const s = pick("style");
	let styleCycle = [r0];
	let firstStyle: number | undefined;
	if (rows.length >= 2 && !s.every((x) => x === s[0])) {
		if (s.every((x, i) => x === s[i % 2])) styleCycle = [r0, rows[1] as number];
		else if (s.slice(1).every((x) => x === s[1])) {
			styleCycle = [rows[1] as number];
			firstStyle = r0;
		}
	}
	return { formulaProto, firstFormulaProto, styleCycle, firstStyle };
}

function leafFields(input: RegionInput, protoRow: number, sampleRows: number[]): FieldDef[] {
	const row = input.model.rowMap.get(protoRow);
	const fields: FieldDef[] = [];
	for (let c = input.c1; c <= input.c2; c++) {
		const cell = row?.cells.find((x) => x.col === c);
		if (cell?.formula !== undefined) continue;
		const name = input.columnNames.get(c);
		if (name !== undefined) fields.push({ name, col: c, type: sampleType(input, c, [protoRow, ...sampleRows]) });
	}
	return fields;
}

function constantFields(input: RegionInput, protoRow: number, sampleRows: number[]): FieldDef[] {
	const row = input.model.rowMap.get(protoRow);
	const fields: FieldDef[] = [];
	for (let c = input.c1; c <= input.c2; c++) {
		const cell = row?.cells.find((x) => x.col === c);
		if (!cell?.hasValue) continue;
		const name = input.columnNames.get(c);
		if (name !== undefined) fields.push({ name, col: c, type: sampleType(input, c, [protoRow, ...sampleRows]) });
	}
	return fields;
}

function cellDisplay(input: RegionInput, r: number, col: number): string {
	const cell = input.model.rowMap.get(r)?.cells.find((x) => x.col === col);
	if (!cell?.hasValue) return "";
	const text = cellText(input.model, cell, input.sst);
	if (text !== undefined) return text.trim();
	const v = cell.el.children.find((c) => c.name === "v");
	return v ? input.model.xml.slice(v.openEnd, v.closeStart).trim() : "";
}

/** A group row's identity: its label fields' text, e.g. "Overhead". */
export function rowLabel(input: { model: SheetModel; sst: string[] } & Pick<RegionInput, "c1">, r: number, fields: FieldDef[]): string {
	return fields
		.map((f) => stripAffix(cellDisplay(input as RegionInput, r, f.col), f.affix))
		.join(" / ");
}

/**
 * Shared leading/trailing text across sample labels, cut at a word boundary:
 * ["Engineering subtotal", "Finance subtotal"] -> suffix " subtotal".
 * Needs two or more distinct samples, and every sample must keep a non-empty middle.
 */
export function commonAffix(samples: string[]): { prefix: string; suffix: string } | undefined {
	const distinct = [...new Set(samples)];
	if (distinct.length < 2) return undefined;
	let prefix = distinct.reduce((p, s) => {
		let i = 0;
		while (i < p.length && i < s.length && p[i] === s[i]) i++;
		return p.slice(0, i);
	});
	let suffix = distinct.reduce((p, s) => {
		let i = 0;
		while (i < p.length && i < s.length && p[p.length - 1 - i] === s[s.length - 1 - i]) i++;
		return p.slice(p.length - i);
	});
	// Only whole words: "Total " / " subtotal", not "Tra" from Travel/Training.
	prefix = prefix.slice(0, prefix.search(/\s\S*$/) + 1 || 0);
	if (!/\s$/.test(prefix)) prefix = "";
	const sm = /^\S*\s/.exec(suffix);
	suffix = sm ? suffix.slice(sm[0].length - 1) : "";
	if (!/^\s/.test(suffix)) suffix = "";
	if (prefix === "" && suffix === "") return undefined;
	if (distinct.some((s) => s.length <= prefix.length + suffix.length)) return undefined;
	return { prefix, suffix };
}

export function stripAffix(text: string, affix: { prefix: string; suffix: string } | undefined): string {
	if (!affix) return text;
	let t = text;
	if (affix.prefix && t.startsWith(affix.prefix)) t = t.slice(affix.prefix.length);
	if (affix.suffix && t.endsWith(affix.suffix)) t = t.slice(0, t.length - affix.suffix.length);
	return t;
}

/** The inverse: "Engineering" -> "Engineering subtotal" (unless the value already has the text). */
export function applyAffix(value: string, affix: { prefix: string; suffix: string }): string {
	return `${value.startsWith(affix.prefix) ? "" : affix.prefix}${value}${value.endsWith(affix.suffix) ? "" : affix.suffix}`;
}

/** The type of the first non-empty sample value in this column. */
function sampleType(input: RegionInput, col: number, rows: number[]): FieldType {
	for (const r of rows) {
		const cell = input.model.rowMap.get(r)?.cells.find((x) => x.col === col);
		if (!cell?.hasValue) continue;
		if (cell.t === "s" || cell.t === "inlineStr" || cell.t === "str") return "string";
		if (cell.t === "b") return "boolean";
		if (cell.t === "e") continue;
		const s = cell.el.attrs.s === undefined ? 0 : Number(cell.el.attrs.s);
		return input.styles.isDate[s] ? "date" : "number";
	}
	return "unknown";
}

const PLACEHOLDER_RE = /\{\{\s*([A-Za-z_$][\w$]*(?:\.[A-Za-z_$][\w$]*|\[\d+\])*)\s*\}\}/g;

export function placeholderPaths(text: string): string[] {
	return [...text.matchAll(PLACEHOLDER_RE)].map((m) => m[1] as string);
}

export { PLACEHOLDER_RE };

export function discoverScalars(
	wb: WorkbookInfo,
	sheets: Map<string, SheetModel>,
	sst: string[],
): ScalarTarget[] {
	const out: ScalarTarget[] = [];
	for (const sheet of wb.sheets) {
		const model = sheets.get(sheet.part);
		if (!model) continue;
		for (const row of model.rows) {
			for (const cell of row.cells) {
				if (cell.t !== "s" && cell.t !== "inlineStr") continue;
				const text = cellText(model, cell, sst);
				if (!text || !text.includes("{{")) continue;
				for (const key of placeholderPaths(text)) out.push({ key, kind: "placeholder", sheet, col: cell.col, row: row.r });
			}
		}
	}
	for (const dn of wb.names) {
		const target = nameTarget(dn, wb);
		if (!target || target.area.r1 !== target.area.r2 || target.area.c1 !== target.area.c2) continue;
		out.push({ key: dn.name, kind: "name", sheet: target.sheet, col: target.area.c1, row: target.area.r1 });
	}
	return out;
}
