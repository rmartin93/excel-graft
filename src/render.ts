import { type Area, cellName, MAX_ROW, parseArea, parseSqref, rowRuns } from "./a1.js";
import { rewriteRefs } from "./formula.js";
import { type OutRow, RenderDataError, type RefOrigin, RegionLayout, SheetPlan, WorkbookMapper } from "./layout.js";
import { forceRecalc, type Package, REL, type SheetInfo, type WorkbookInfo } from "./package.js";
import { placeholderPaths, type Region, type ScalarTarget } from "./regions.js";
import { cellText, type SheetCell, type SheetModel, type SheetRow } from "./sheet.js";
import { type CellValue, encodeValue, type EncodedCell, isFormulaValue } from "./values.js";
import { escapeText, findAll, localName, parseXml, rawInner, rawStartTag, setAttrs, Splicer, textOf, type XmlElement } from "./xml.js";

export interface RenderOptions {
	/** What to do when a record lacks a field or a placeholder has no value. Default "blank". */
	onMissing?: "blank" | "error";
	/** Throw when a record has properties that match no column. Default false (extra columns from a query are ignored). */
	strict?: boolean;
	/** Receives notes about template features that were dropped or approximated. */
	onWarning?: (message: string) => void;
	/** Receives where each region landed in the output (row numbers, row kinds, group spans). */
	onReport?: (report: RenderReport) => void;
}

export interface RenderReport {
	regions: RegionReport[];
}

export interface RegionReport {
	key: string;
	sheet: string;
	firstRow: number;
	lastRow: number;
	rows: ReportRow[];
}

export interface ReportRow {
	row: number;
	templateRow: number;
	kind: "leaf" | "header" | "footer" | "spacer" | "fixed";
	/** Output span of the group this row belongs to (the whole region for fixed rows). */
	groupFirst: number;
	groupLast: number;
	record: Record<string, unknown> | undefined;
}

export interface RenderContext {
	pkg: Package;
	wb: WorkbookInfo;
	sheets: Map<string, SheetModel>;
	sst: string[];
	regions: Region[];
	scalars: ScalarTarget[];
}

const MAX_FORMULA_LENGTH = 8192;

export function render(ctx: RenderContext, data: Record<string, unknown>, options: RenderOptions): void {
	const warn = options.onWarning ?? (() => {});
	const onMissing = options.onMissing ?? "blank";

	// Validate top-level keys so a typo fails loudly instead of rendering nothing.
	const known = new Set<string>([...ctx.regions.map((r) => r.key), ...ctx.scalars.map((s) => s.key.split(/[.[]/)[0] as string)]);
	for (const key of Object.keys(data)) {
		if (!known.has(key)) {
			throw new RenderDataError(
				`"${key}" doesn't match any table, named range, named cell or {{placeholder}} in the template. Known keys: ${[...known].sort().join(", ") || "(none)"}`,
			);
		}
	}

	// Build a plan per sheet for the regions that received data.
	const plans = new Map<string, SheetPlan>();
	const bySheet = new Map<SheetInfo, Region[]>();
	for (const region of ctx.regions) {
		if (!(region.key in data)) continue;
		const list = bySheet.get(region.sheet) ?? [];
		list.push(region);
		bySheet.set(region.sheet, list);
	}
	for (const [sheet, regions] of bySheet) {
		regions.sort((a, b) => a.bodyStart - b.bodyStart);
		for (let i = 1; i < regions.length; i++) {
			const prev = regions[i - 1] as Region;
			const cur = regions[i] as Region;
			if (cur.bodyStart <= prev.bodyEnd + prev.totalsRows) {
				throw new RenderDataError(`"${prev.key}" and "${cur.key}" share rows on sheet "${sheet.name}"; filling both would need to insert rows through each other`);
			}
		}
		const layouts: RegionLayout[] = [];
		let offset = 0;
		for (const region of regions) {
			const items = data[region.key];
			if (!Array.isArray(items)) throw new RenderDataError(`"${region.key}" must be an array of rows`);
			if (options.strict) checkStrict(region, items);
			const layout = new RegionLayout(region, region.bodyStart + offset, items);
			layouts.push(layout);
			offset += layout.delta;
		}
		const last = layouts[layouts.length - 1] as RegionLayout;
		const model = ctx.sheets.get(sheet.part) as SheetModel;
		const lastUsed = model.rows[model.rows.length - 1]?.r ?? 0;
		if (lastUsed + offset > MAX_ROW || last.outEnd > MAX_ROW) {
			throw new RenderDataError(`sheet "${sheet.name}" would pass Excel's ${MAX_ROW.toLocaleString("en-US")}-row limit`);
		}
		plans.set(sheet.name, new SheetPlan(sheet.name, layouts));
	}
	const mapper = new WorkbookMapper(plans, warn);
	if (options.onReport) {
		options.onReport({
			regions: [...plans.values()].flatMap((plan) =>
				plan.layouts.map((l) => ({
					key: l.region.key,
					sheet: plan.sheetName,
					firstRow: l.outStart,
					lastRow: l.outEnd,
					rows: l.rows.map((r) => ({
						row: r.out,
						templateRow: r.proto,
						kind: r.role.kind,
						groupFirst: r.scope.first,
						groupLast: r.scope.last,
						record: r.record,
					})),
				})),
			),
		});
	}

	const scalarValues = new Map<ScalarTarget, unknown>();
	for (const s of ctx.scalars) {
		const value = resolvePath(data, s.key);
		scalarValues.set(s, value);
	}

	const writer = new CellWriter(ctx, options, onMissing);
	for (const sheet of ctx.wb.sheets) {
		const model = ctx.sheets.get(sheet.part);
		if (!model) continue;
		const plan = plans.get(sheet.name);
		const scalars = ctx.scalars.filter((s) => s.sheet === sheet && (s.key.split(/[.[]/)[0] as string) in data);
		const xml = renderSheetXml(ctx, model, sheet, plan, mapper, scalars, scalarValues, writer, warn);
		if (xml !== model.xml) ctx.pkg.setText(sheet.part, xml);
		if (plan) patchSheetParts(ctx, sheet, plan, mapper, warn);
	}

	if (mapper.active) patchWorkbookParts(ctx, mapper);
	// Anything written can feed a formula whose cached value is now stale.
	if (Object.keys(data).length > 0) forceRecalc(ctx.pkg, ctx.wb);
}

function checkStrict(region: Region, items: unknown[]): void {
	const names = new Set(region.columnNames.values());
	const visit = (list: unknown[], path: string, childKey: string | undefined): void => {
		list.forEach((item, i) => {
			if (typeof item !== "object" || item === null) return;
			for (const k of Object.keys(item)) {
				if (k === childKey) continue;
				if (!names.has(k)) throw new RenderDataError(`${path}[${i}].${k} doesn't match any column (strict mode)`);
			}
			const kids = childKey ? (item as Record<string, unknown>)[childKey] : undefined;
			if (Array.isArray(kids)) visit(kids, `${path}[${i}].${childKey}`, nextChildKey(region, path.split(".").length));
		});
	};
	visit(items, region.key, nextChildKey(region, 0));
}

function nextChildKey(region: Region, depth: number): string | undefined {
	let p = region.unit;
	for (let d = 0; d < depth; d++) {
		if (p.kind !== "group") return undefined;
		p = p.child;
	}
	return p.kind === "group" ? p.childKey : undefined;
}

function resolvePath(data: Record<string, unknown>, path: string): unknown {
	let cur: unknown = data;
	for (const seg of path.split(/\.|\[(\d+)\]/).filter((s) => s !== undefined && s !== "")) {
		if (cur === null || cur === undefined || typeof cur !== "object") return undefined;
		cur = (cur as Record<string, unknown>)[seg];
	}
	return cur;
}

/** Builds `<c>` elements, caching the per-prototype tag split so 100k-row renders stay fast. */
class CellWriter {
	private readonly tagCache = new Map<XmlElement, [string, string]>();
	private readonly formulaTagCache = new Map<XmlElement, [string, string]>();

	constructor(
		private readonly ctx: RenderContext,
		private readonly options: RenderOptions,
		readonly onMissing: "blank" | "error",
	) {}

	/** Start tag parts around the `r` value, with `t`, `vm` and (optionally) `cm` removed. */
	private tagParts(el: XmlElement, xml: string, keepCm: boolean): [string, string] {
		const cache = keepCm ? this.formulaTagCache : this.tagCache;
		let parts = cache.get(el);
		if (!parts) {
			let tag = rawStartTag(xml, el);
			if (tag.endsWith("/>")) tag = `${tag.slice(0, -2)}>`;
			tag = setAttrs(tag, { r: "\u0000", t: null, vm: null, ...(keepCm ? {} : { cm: null }) });
			const [before, after] = tag.split("\u0000") as [string, string];
			parts = [before, after.slice(0, -1)];
			cache.set(el, parts);
		}
		return parts;
	}

	cell(proto: SheetCell | undefined, xml: string, ref: string, t: string | null, inner: string, isFormula: boolean): string {
		const tAttr = t ? ` t="${t}"` : "";
		if (!proto) return inner === "" ? "" : `<c r="${ref}"${tAttr}>${inner}</c>`;
		const [before, after] = this.tagParts(proto.el, xml, isFormula);
		return `${before}${ref}${after}${tAttr}>${inner}</c>`;
	}

	value(v: unknown, where: () => string): EncodedCell {
		if (v === undefined) {
			if (this.onMissing === "error") throw new RenderDataError(`missing value for ${where()}`);
			return { t: null, inner: "" };
		}
		if (v !== null && typeof v === "object" && !(v instanceof Date) && !isFormulaValue(v)) {
			throw new RenderDataError(`${where()} is an object; cells take strings, numbers, booleans, dates, null, or { formula }`);
		}
		if (typeof v === "symbol" || typeof v === "function") throw new RenderDataError(`${where()} has unsupported type ${typeof v}`);
		return encodeValue(v as CellValue, this.ctx.wb.date1904);
	}

	formulaInner(fEl: XmlElement, xml: string, formula: string, arrayRef: string | undefined): string {
		if (formula.length > MAX_FORMULA_LENGTH) {
			throw new RenderDataError(
				`a regenerated formula is ${formula.length} characters, over Excel's ${MAX_FORMULA_LENGTH} limit (usually a total over thousands of groups). Consider totalling with SUBTOTAL over one range instead of listing each group.`,
			);
		}
		let tag = rawStartTag(xml, fEl);
		if (tag.endsWith("/>")) tag = `${tag.slice(0, -2)}>`;
		const t = fEl.attrs.t;
		if (t === "shared") tag = setAttrs(tag, { t: null, si: null, ref: null });
		else if (t === "array" && arrayRef !== undefined) tag = setAttrs(tag, { ref: arrayRef });
		return `${tag}${escapeText(formula)}</${fEl.name}>`;
	}
}

function renderSheetXml(
	ctx: RenderContext,
	model: SheetModel,
	sheet: SheetInfo,
	plan: SheetPlan | undefined,
	mapper: WorkbookMapper,
	scalars: ScalarTarget[],
	scalarValues: Map<ScalarTarget, unknown>,
	writer: CellWriter,
	warn: (m: string) => void,
): string {
	const xml = model.xml;
	const outside: RefOrigin = { kind: "outside", sheet: sheet.name };
	const mapFormula = (f: string, origin: RefOrigin) =>
		mapper.active ? rewriteRefs(f, (ref, c) => mapper.map(ref, origin, c.inArgList)) : f;

	const scalarAt = new Map<string, ScalarTarget[]>();
	for (const s of scalars) {
		const k = `${s.row}:${s.col}`;
		scalarAt.set(k, [...(scalarAt.get(k) ?? []), s]);
	}

	// Pass 1: decide which static rows need rewriting, and which shared formulas must be unshared.
	const staticRows = model.rows.filter((row) => !plan?.layoutFor(row.r));
	const newFormula = new Map<SheetCell, string>();
	const rewrite = new Set<SheetRow>();
	const unshare = new Set<string>();
	for (const row of staticRows) {
		const newR = plan ? plan.mapRow(row.r) : row.r;
		let touched = newR !== row.r;
		for (const cell of row.cells) {
			if (scalarAt.has(`${row.r}:${cell.col}`)) touched = true;
			if (cell.formula === undefined) continue;
			const f = mapFormula(cell.formula, outside);
			if (f !== cell.formula) {
				newFormula.set(cell, f);
				touched = true;
			}
		}
		if (touched) rewrite.add(row);
	}
	const sharedMembers = new Map<string, SheetRow[]>();
	for (const row of model.rows) {
		for (const cell of row.cells) {
			const si = cell.formulaEl?.attrs.t === "shared" ? cell.formulaEl.attrs.si : undefined;
			if (si === undefined) continue;
			sharedMembers.set(si, [...(sharedMembers.get(si) ?? []), row]);
			if (rewrite.has(row) || plan?.layoutFor(row.r)) unshare.add(si);
		}
	}
	for (const si of unshare) for (const row of sharedMembers.get(si) ?? []) if (!plan?.layoutFor(row.r)) rewrite.add(row);

	// Pass 2: emit sheetData.
	const out: string[] = [];
	const emitStatic = (row: SheetRow) => {
		if (!rewrite.has(row)) {
			out.push(xml.slice(row.el.start, row.el.end));
			return;
		}
		const newR = plan ? plan.mapRow(row.r) : row.r;
		out.push(openRow(xml, row.el, newR));
		for (const c of row.el.children) {
			if (localName(c.name) !== "c") {
				out.push(xml.slice(c.start, c.end));
				continue;
			}
			const cell = row.cells.find((x) => x.el === c) as SheetCell;
			out.push(staticCell(cell, newR));
		}
		out.push(`</${row.el.name}>`);
	};
	const staticCell = (cell: SheetCell, newR: number): string => {
		const ref = cellName(cell.col, newR);
		const targets = scalarAt.get(`${cell.row}:${cell.col}`);
		if (targets) return scalarCell(cell, ref, targets);
		const shared = cell.formulaEl?.attrs.t === "shared" && unshare.has(cell.formulaEl.attrs.si ?? "");
		const f = newFormula.get(cell);
		if (cell.formula !== undefined && cell.formulaEl && (f !== undefined || shared || newR !== cell.row)) {
			const formula = f ?? cell.formula;
			const arrayRef = arrayRefFor(cell.formulaEl, (a) => mapArea(mapper, plan, sheet.name, a));
			return writer.cell(cell, xml, ref, null, writer.formulaInner(cell.formulaEl, xml, formula, arrayRef) + extLstOf(xml, cell), true);
		}
		// Value cell that only moved: keep everything but the reference.
		const tag = setAttrs(rawStartTag(xml, cell.el), { r: ref });
		return cell.el.selfClosing ? tag : `${tag}${rawInner(xml, cell.el)}</${cell.el.name}>`;
	};
	const scalarCell = (cell: SheetCell, ref: string, targets: ScalarTarget[]): string => {
		const nameTarget = targets.find((t) => t.kind === "name");
		if (nameTarget) {
			const enc = writer.value(scalarValues.get(nameTarget), () => `"${nameTarget.key}"`);
			return writer.cell(cell, xml, ref, enc.t, enc.inner, false);
		}
		const text = cellText(model, cell, ctx.sst) ?? "";
		const whole = /^\s*\{\{\s*([^}]+?)\s*\}\}\s*$/.exec(text);
		if (whole && placeholderPaths(text).length === 1) {
			const target = targets[0] as ScalarTarget;
			const enc = writer.value(scalarValues.get(target), () => `{{${target.key}}}`);
			return writer.cell(cell, xml, ref, enc.t, enc.inner, false);
		}
		let missing: string | undefined;
		const replaced = text.replace(/\{\{\s*([^}]+?)\s*\}\}/g, (m, path: string) => {
			const target = targets.find((t) => t.key === path);
			if (!target) return m;
			const v = scalarValues.get(target);
			if (v === undefined) missing ??= path;
			return formatInline(v);
		});
		if (missing !== undefined && writer.onMissing === "error") throw new RenderDataError(`missing value for {{${missing}}}`);
		const enc = encodeValue(replaced, ctx.wb.date1904);
		return writer.cell(cell, xml, ref, enc.t, enc.inner, false);
	};

	if (!plan) {
		// Unrendered sheet: splice only the rows that changed, keep every other byte.
		const splicer = new Splicer(xml);
		for (const row of rewrite) {
			const parts: string[] = [];
			const saved = out.length;
			emitStatic(row);
			parts.push(...out.splice(saved));
			splicer.replaceElement(row.el, parts.join(""));
		}
		patchCrossSheetFormulas(xml, model, sheet, mapper, splicer);
		return splicer.apply();
	}

	const layouts = plan?.layouts ?? [];
	let li = 0;
	let emittedLayout = -1;
	for (const row of model.rows) {
		while (li < layouts.length && row.r > (layouts[li] as RegionLayout).region.bodyEnd) {
			if (emittedLayout < li) emitLayout(layouts[li] as RegionLayout);
			emittedLayout = li;
			li++;
		}
		const lay = layouts[li];
		if (lay && lay.inBody(row.r)) {
			if (emittedLayout < li) {
				emitLayout(lay);
				emittedLayout = li;
			}
			continue;
		}
		emitStatic(row);
	}
	for (; li < layouts.length; li++) if (emittedLayout < li) emitLayout(layouts[li] as RegionLayout);

	function emitLayout(layout: RegionLayout): void {
		for (const r of layout.rows) out.push(emitInstanceRow(layout, r));
	}

	function emitInstanceRow(layout: RegionLayout, r: OutRow): string {
		const proto = model.rowMap.get(r.proto);
		const origin: RefOrigin = { kind: "instance", sheet: sheet.name, layout, row: r };
		const parts: string[] = [proto ? openRow(xml, proto.el, r.out) : `<row r="${r.out}">`];
		const fieldByCol = new Map(r.fields.map((f) => [f.col, f]));
		const cols = new Set<number>([...(proto?.cells.map((c) => c.col) ?? []), ...fieldByCol.keys()]);
		for (const col of [...cols].sort((a, b) => a - b)) {
			const cell = proto?.cells.find((c) => c.col === col);
			const ref = cellName(col, r.out);
			const field = fieldByCol.get(col);
			if (cell?.formula !== undefined && cell.formulaEl && !field) {
				const f = mapFormula(cell.formula, origin);
				const arrayRef = arrayRefFor(cell.formulaEl, (a) => ({ ...a, r1: a.r1 + (r.out - r.proto), r2: a.r2 + (r.out - r.proto) }));
				parts.push(writer.cell(cell, xml, ref, null, writer.formulaInner(cell.formulaEl, xml, f, arrayRef) + extLstOf(xml, cell), true));
				continue;
			}
			if (field) {
				const value = r.record ? r.record[field.name] : null;
				const enc = writer.value(value, () => `${layout.region.key} row ${r.out} column "${field.name}"`);
				const c = writer.cell(cell, xml, ref, enc.t, enc.inner, isFormulaValue(value));
				if (c !== "") parts.push(c);
				continue;
			}
			if (cell) {
				// A constant that isn't data (a label in a total row, or a cell outside the columns): copy it.
				const tag = setAttrs(rawStartTag(xml, cell.el), { r: ref });
				parts.push(cell.el.selfClosing ? tag : `${tag}${rawInner(xml, cell.el)}</${cell.el.name}>`);
			}
		}
		parts.push(`</${proto?.el.name ?? "row"}>`);
		return parts.join("");
	}

	const splicer = new Splicer(xml);
	const sd = model.sheetData;
	if (sd.selfClosing) splicer.replaceElement(sd, `<${sd.name}>${out.join("")}</${sd.name}>`);
	else splicer.replace(sd.openEnd, sd.closeStart, out.join(""));
	patchWorksheetElements(xml, model, sheet, plan, mapper, splicer, warn);
	return splicer.apply();
}

function formatInline(v: unknown): string {
	if (v === null || v === undefined) return "";
	if (v instanceof Date) return Number.isNaN(v.getTime()) ? "" : v.toISOString().slice(0, v.getUTCHours() || v.getUTCMinutes() ? 16 : 10).replace("T", " ");
	return String(v);
}

const rowTagCache = new WeakMap<XmlElement, [string, string]>();

function openRow(xml: string, rowEl: XmlElement, r: number): string {
	let parts = rowTagCache.get(rowEl);
	if (!parts) {
		let tag = setAttrs(rawStartTag(xml, rowEl), { r: "\u0000" });
		if (tag.endsWith("/>")) tag = `${tag.slice(0, -2)}>`;
		parts = tag.split("\u0000") as [string, string];
		rowTagCache.set(rowEl, parts);
	}
	return `${parts[0]}${r}${parts[1]}`;
}

function extLstOf(xml: string, cell: SheetCell): string {
	const ext = cell.el.children.find((c) => localName(c.name) === "extLst");
	return ext ? xml.slice(ext.start, ext.end) : "";
}

function arrayRefFor(fEl: XmlElement, map: (a: Area) => Area): string | undefined {
	if (fEl.attrs.t !== "array" || !fEl.attrs.ref) return undefined;
	const a = parseArea(fEl.attrs.ref);
	if (!a) return undefined;
	const m = map(a);
	return m.r1 === m.r2 && m.c1 === m.c2 ? cellName(m.c1, m.r1) : `${cellName(m.c1, m.r1)}:${cellName(m.c2, m.r2)}`;
}

/** Maps an area as seen from outside the data. */
function mapArea(mapper: WorkbookMapper, plan: SheetPlan | undefined, _sheet: string, a: Area): Area {
	if (!plan) return a;
	return { c1: a.c1, c2: a.c2, r1: mapper.rangeStart(plan, a.r1, a.r2, undefined), r2: mapper.rangeEnd(plan, a.r1, a.r2, undefined) };
}

/** Maps an sqref area into output areas: outside segments shift, body segments project onto every instance. */
function mapSqrefArea(plan: SheetPlan, a: Area): Area[] {
	const rows: [number, number][] = [];
	let cursor = a.r1;
	for (const lay of plan.layouts) {
		const bs = lay.region.bodyStart;
		const be = lay.region.bodyEnd;
		if (be < cursor || bs > a.r2) continue;
		if (cursor < bs) rows.push([plan.mapRow(cursor), plan.mapRow(bs - 1)]);
		const projected = lay.project(lay.bodyRange(Math.max(cursor, bs), Math.min(a.r2, be)));
		rows.push(...rowRuns(projected));
		cursor = be + 1;
	}
	if (cursor <= a.r2) rows.push([plan.mapRow(cursor), plan.mapRow(a.r2)]);
	const merged: [number, number][] = [];
	for (const [s, e] of rows.sort((x, y) => x[0] - y[0])) {
		const last = merged[merged.length - 1];
		if (last && s <= last[1] + 1) last[1] = Math.max(last[1], e);
		else merged.push([s, Math.min(e, MAX_ROW)]);
	}
	return merged.filter(([s]) => s <= MAX_ROW).map(([s, e]) => ({ c1: a.c1, c2: a.c2, r1: s, r2: e }));
}

function formatAreaList(areas: Area[]): string {
	return areas.map((a) => (a.r1 === a.r2 && a.c1 === a.c2 ? cellName(a.c1, a.r1) : `${cellName(a.c1, a.r1)}:${cellName(a.c2, a.r2)}`)).join(" ");
}

function mapSqref(plan: SheetPlan, text: string): { text: string; firstFrom: number | undefined; firstTo: number | undefined } {
	const areas = parseSqref(text);
	const out: Area[] = [];
	let firstFrom: number | undefined;
	let firstTo: number | undefined;
	for (const a of areas) {
		if (!a) continue;
		const mapped = mapSqrefArea(plan, a);
		if (firstFrom === undefined && mapped[0]) {
			firstFrom = a.r1;
			firstTo = mapped[0].r1;
		}
		out.push(...mapped);
	}
	return { text: formatAreaList(out), firstFrom, firstTo };
}

/** Everything in the worksheet part outside sheetData that holds row references. */
function patchWorksheetElements(
	xml: string,
	model: SheetModel,
	sheet: SheetInfo,
	plan: SheetPlan,
	mapper: WorkbookMapper,
	splicer: Splicer,
	warn: (m: string) => void,
): void {
	const ws = model.worksheet;
	const outside: RefOrigin = { kind: "outside", sheet: sheet.name };
	const mapFormula = (f: string, origin: RefOrigin) => rewriteRefs(f, (ref, c) => mapper.map(ref, origin, c.inArgList));
	const editAttrs = (el: XmlElement, updates: Record<string, string | null>) =>
		splicer.replace(el.start, el.openEnd, setAttrs(rawStartTag(xml, el), updates));
	const anchoredFor = (sqref: { firstFrom: number | undefined; firstTo: number | undefined }): RefOrigin =>
		sqref.firstFrom === undefined
			? outside
			: { kind: "anchored", sheet: sheet.name, from: sqref.firstFrom, to: sqref.firstTo as number, layout: plan.layoutFor(sqref.firstFrom) };
	const mapCellRef = (ref: string): string | undefined => {
		const a = parseArea(ref);
		if (!a) return undefined;
		return cellName(a.c1, mapper.pointRow(plan, a.r1, undefined));
	};

	for (const el of ws.children) {
		const name = localName(el.name);
		if (name === "sheetData") continue;
		if (name === "dimension" && el.attrs.ref) {
			const a = parseArea(el.attrs.ref);
			if (a) {
				const m = mapArea(mapper, plan, sheet.name, a);
				const r2 = Math.min(Math.max(m.r2, ...plan.layouts.map((l) => l.outEnd)), MAX_ROW);
				editAttrs(el, { ref: formatAreaList([{ ...m, r2 }]).replace(/^([A-Z]+\d+)$/, "$1:$1") });
			}
		} else if (name === "sheetViews" || name === "customSheetViews") {
			for (const pane of findAll(el, "pane")) if (pane.attrs.topLeftCell) {
				const m = mapCellRef(pane.attrs.topLeftCell);
				if (m) editAttrs(pane, { topLeftCell: m });
			}
			for (const sel of findAll(el, "selection")) {
				const updates: Record<string, string> = {};
				if (sel.attrs.activeCell) updates.activeCell = mapCellRef(sel.attrs.activeCell) ?? sel.attrs.activeCell;
				if (sel.attrs.sqref) updates.sqref = mapSqref(plan, sel.attrs.sqref).text || "A1";
				editAttrs(sel, updates);
			}
		} else if (name === "mergeCells") {
			const kept: string[] = [];
			for (const mc of el.children) {
				const a = parseArea(mc.attrs.ref ?? "");
				if (!a) continue;
				const lay = plan.layouts.find((l) => a.r1 <= l.region.bodyEnd && a.r2 >= l.region.bodyStart);
				if (!lay) {
					kept.push(formatAreaList(mapSqrefArea(plan, a)));
				} else if (a.r1 === a.r2) {
					for (const r of lay.project([a.r1])) kept.push(formatAreaList([{ ...a, r1: r, r2: r }]).replace(/^([A-Z]+\d+)$/, "$1:$1"));
				} else {
					warn(`merged cells ${mc.attrs.ref} span several sample rows of "${lay.region.key}" and were removed`);
				}
			}
			if (kept.length === 0) splicer.replaceElement(el, "");
			else {
				const tag = setAttrs(rawStartTag(xml, el), { count: String(kept.length) });
				splicer.replaceElement(el, `${tag.endsWith("/>") ? tag.slice(0, -2) + ">" : tag}${kept.map((r) => `<mergeCell ref="${r}"/>`).join("")}</${el.name}>`);
			}
		} else if (name === "conditionalFormatting" || name === "dataValidations" || name === "ignoredErrors" || name === "protectedRanges") {
			const holders = name === "conditionalFormatting" ? [el] : el.children;
			let removed = 0;
			for (const h of holders) {
				if (!h.attrs.sqref) continue;
				const sq = mapSqref(plan, h.attrs.sqref);
				if (sq.text === "") {
					splicer.replaceElement(h, "");
					removed++;
					continue;
				}
				const origin = anchoredFor(sq);
				const formulaEls = findAll(h, "formula").concat(findAll(h, "formula1"), findAll(h, "formula2"));
				if (formulaEls.length === 0) {
					editAttrs(h, { sqref: sq.text });
					continue;
				}
				// Rebuild the holder so the attribute edit and the formula edits don't overlap.
				const inner = rebuildInner(xml, h, formulaEls, (f) => mapFormula(f, origin));
				const tag = setAttrs(rawStartTag(xml, h), { sqref: sq.text });
				splicer.replaceElement(h, h.selfClosing ? tag : `${tag}${inner}</${h.name}>`);
			}
			if (name === "dataValidations" && removed > 0) {
				const remaining = el.children.length - removed;
				if (remaining === 0) splicer.replaceElement(el, "");
				else splicer.replace(el.start, el.openEnd, setAttrs(rawStartTag(xml, el), { count: String(remaining) }));
			}
		} else if (name === "hyperlinks") {
			let keptCount = 0;
			for (const h of el.children) {
				const a = parseArea(h.attrs.ref ?? "");
				if (!a) {
					keptCount++;
					continue;
				}
				if (plan.layoutFor(a.r1)) {
					splicer.replaceElement(h, "");
					continue;
				}
				keptCount++;
				editAttrs(h, { ref: formatAreaList(mapSqrefArea(plan, a)) });
			}
			if (keptCount === 0) splicer.replaceElement(el, "");
		} else if (name === "autoFilter") {
			patchAreaAttrs(xml, el, plan, mapper, splicer);
		} else if (name === "rowBreaks") {
			let count = 0;
			let manual = 0;
			const brks: string[] = [];
			for (const brk of el.children) {
				const id = Number(brk.attrs.id ?? "0");
				if (plan.layoutFor(id) && plan.layoutFor(id + 1)) continue;
				const mapped = plan.layoutFor(id) ? (plan.layoutFor(id) as RegionLayout).outEnd : plan.mapRow(id);
				brks.push(setAttrs(rawOuterSelfClosing(xml, brk), { id: String(mapped) }));
				count++;
				if (brk.attrs.man === "1") manual++;
			}
			if (count === 0) splicer.replaceElement(el, "");
			else {
				const tag = setAttrs(rawStartTag(xml, el), { count: String(count), manualBreakCount: String(manual) });
				splicer.replaceElement(el, `${tag}${brks.join("")}</${el.name}>`);
			}
		} else if (name === "cellWatches") {
			for (const cw of el.children) if (cw.attrs.r) editAttrs(cw, { r: mapCellRef(cw.attrs.r) ?? cw.attrs.r });
		} else if (name === "extLst") {
			patchWorksheetExtensions(xml, el, sheet, plan, mapper, splicer, warn);
		} else if (name === "controls" || name === "oleObjects" || name === "AlternateContent") {
			patchAnchorRows(xml, el, plan, splicer);
		}
	}
}

function rawOuterSelfClosing(xml: string, el: XmlElement): string {
	return xml.slice(el.start, el.end);
}

/** Rebuilds an element's inner XML with the text of `targets` rewritten. */
function rebuildInner(xml: string, el: XmlElement, targets: XmlElement[], fn: (text: string) => string): string {
	const sorted = [...targets].sort((a, b) => a.start - b.start);
	const parts: string[] = [];
	let pos = el.openEnd;
	for (const t of sorted) {
		parts.push(xml.slice(pos, t.openEnd), escapeText(fn(textOf(xml, t))));
		pos = t.closeStart;
	}
	parts.push(xml.slice(pos, el.closeStart));
	return parts.join("");
}

function patchAreaAttrs(xml: string, el: XmlElement, plan: SheetPlan, mapper: WorkbookMapper, splicer: Splicer): void {
	const targets = [el, ...findAll(el, "sortState"), ...findAll(el, "sortCondition")].filter((x) => x.attrs.ref);
	for (const t of targets) {
		const a = parseArea(t.attrs.ref as string);
		if (!a) continue;
		const m = mapArea(mapper, plan, plan.sheetName, a);
		splicer.replace(t.start, t.openEnd, setAttrs(rawStartTag(xml, t), { ref: formatAreaList([m]).replace(/^([A-Z]+\d+)$/, "$1:$1") }));
	}
}

/** Shifts `xdr:row`-style anchor rows (0-based) found under `el`. */
function patchAnchorRows(xml: string, el: XmlElement, plan: SheetPlan, splicer: Splicer): void {
	const anchors = findAll(el, "from").concat(findAll(el, "xdr:from"));
	for (const from of anchors) {
		const parent = from.parent;
		if (!parent) continue;
		const to = parent.children.find((c) => localName(c.name) === "to");
		// Form controls / OLE objects only move with cells when moveWithCells="1".
		shiftAnchor(xml, from, to, plan, splicer, parent.attrs.moveWithCells === "1" ? undefined : "absolute");
	}
}

function anchorRowEl(el: XmlElement | undefined): XmlElement | undefined {
	return el?.children.find((c) => localName(c.name) === "row");
}

function shiftAnchor(xml: string, from: XmlElement, to: XmlElement | undefined, plan: SheetPlan, splicer: Splicer, editAs: string | undefined): void {
	if (editAs === "absolute") return;
	const fromRowEl = anchorRowEl(from);
	if (!fromRowEl) return;
	const f0 = Number(textOf(xml, fromRowEl));
	const toRowEl = anchorRowEl(to);
	const t0 = toRowEl ? Number(textOf(xml, toRowEl)) : undefined;
	const [f1, t1] = mapAnchor(plan, f0 + 1, t0 === undefined ? undefined : t0 + 1, editAs === "oneCell");
	if (f1 !== f0 + 1) splicer.replace(fromRowEl.openEnd, fromRowEl.closeStart, String(f1 - 1));
	if (toRowEl && t1 !== undefined && t1 !== (t0 as number) + 1) splicer.replace(toRowEl.openEnd, toRowEl.closeStart, String(t1 - 1));
}

/** Maps a (1-based) anchor from/to pair. Objects over the data keep their size; objects spanning it stretch. */
export function mapAnchor(plan: SheetPlan, from: number, to: number | undefined, keepSize: boolean): [number, number | undefined] {
	const place = (t: number): number => {
		const lay = plan.layoutFor(t);
		if (lay) return Math.min(t + (lay.outStart - lay.region.bodyStart), lay.outEnd);
		return plan.mapRow(t);
	};
	const f1 = place(from);
	if (to === undefined) return [f1, undefined];
	if (keepSize) return [f1, f1 + (to - from)];
	return [f1, Math.max(f1, place(to))];
}

function patchWorksheetExtensions(
	xml: string,
	extLst: XmlElement,
	sheet: SheetInfo,
	plan: SheetPlan,
	mapper: WorkbookMapper,
	splicer: Splicer,
	warn: (m: string) => void,
): void {
	const outside: RefOrigin = { kind: "outside", sheet: sheet.name };
	const mapFormula = (f: string, origin: RefOrigin) => rewriteRefs(f, (ref, c) => mapper.map(ref, origin, c.inArgList));

	// Sparklines: one per cell; body sparklines are replicated per output row from the first sample row's sparkline.
	const sparklineGroups = findAll(extLst, "x14:sparklineGroup");
	const handled = new Set<XmlElement>();
	for (const group of sparklineGroups) {
		const container = group.children.find((c) => localName(c.name) === "sparklines");
		if (!container) continue;
		const out: string[] = [];
		const seenRoles = new Set<string>();
		for (const sp of container.children) {
			handled.add(sp);
			const fEl = sp.children.find((c) => localName(c.name) === "f");
			const sqEl = sp.children.find((c) => localName(c.name) === "sqref");
			if (!sqEl) continue;
			const cell = parseArea(textOf(xml, sqEl));
			if (!cell) continue;
			const f = fEl ? textOf(xml, fEl) : "";
			const lay = plan.layoutFor(cell.r1);
			if (lay) {
				const role = lay.region.roleOfRow.get(cell.r1);
				const key = `${role?.id}:${cell.c1}`;
				if (seenRoles.has(key)) continue;
				seenRoles.add(key);
				for (const o of lay.project([cell.r1])) {
					const row = lay.rowAt(o) as OutRow;
					const shifted = rewriteRefs(f, (ref, c) => mapper.map(ref, { kind: "instance", sheet: sheet.name, layout: lay, row: { ...row, proto: cell.r1 } }, c.inArgList));
					out.push(sparklineXml(sp.name, fEl?.name, sqEl.name, shifted, cellName(cell.c1, o)));
				}
			} else {
				out.push(sparklineXml(sp.name, fEl?.name, sqEl.name, mapFormula(f, outside), cellName(cell.c1, plan.mapRow(cell.r1))));
			}
		}
		if (out.length === 0) {
			splicer.replaceElement(group, "");
			warn(`a sparkline group on "${sheet.name}" ended up empty and was removed`);
		} else {
			splicer.replace(container.openEnd, container.closeStart, out.join(""));
		}
	}

	// Anything else with an xm:sqref (x14 conditional formatting, data validation, ignored errors...).
	for (const sq of findAll(extLst, "xm:sqref")) {
		const holder = sq.parent;
		if (!holder || handled.has(holder)) continue;
		const mapped = mapSqref(plan, textOf(xml, sq));
		if (mapped.text === "") {
			splicer.replaceElement(holder, "");
			continue;
		}
		const origin: RefOrigin =
			mapped.firstFrom === undefined
				? outside
				: { kind: "anchored", sheet: sheet.name, from: mapped.firstFrom, to: mapped.firstTo as number, layout: plan.layoutFor(mapped.firstFrom) };
		for (const f of findAll(holder, "xm:f")) {
			const text = textOf(xml, f);
			const nf = mapFormula(text, origin);
			if (nf !== text) splicer.replace(f.openEnd, f.closeStart, escapeText(nf));
		}
		splicer.replace(sq.openEnd, sq.closeStart, mapped.text);
	}
}

function sparklineXml(name: string, fName: string | undefined, sqName: string, f: string, cell: string): string {
	return `<${name}>${fName ? `<${fName}>${escapeText(f)}</${fName}>` : ""}<${sqName}>${cell}</${sqName}></${name}>`;
}

/** For sheets that weren't rendered: only formulas pointing at rendered sheets change. */
function patchCrossSheetFormulas(xml: string, model: SheetModel, sheet: SheetInfo, mapper: WorkbookMapper, splicer: Splicer): void {
	if (!mapper.active) return;
	const outside: RefOrigin = { kind: "outside", sheet: sheet.name };
	for (const el of findAll(model.worksheet, "formula").concat(findAll(model.worksheet, "formula1"), findAll(model.worksheet, "formula2"), findAll(model.worksheet, "xm:f"))) {
		const text = textOf(xml, el);
		const nf = rewriteRefs(text, (ref, c) => mapper.map(ref, outside, c.inArgList));
		if (nf !== text) splicer.replace(el.openEnd, el.closeStart, escapeText(nf));
	}
}

/** Drawings, comments, tables and threaded comments that belong to a rendered sheet. */
function patchSheetParts(ctx: RenderContext, sheet: SheetInfo, plan: SheetPlan, mapper: WorkbookMapper, warn: (m: string) => void): void {
	const { pkg } = ctx;
	for (const rel of pkg.rels(sheet.part)) {
		if (rel.external || !pkg.has(rel.target)) continue;
		if (rel.type === REL.table) {
			const xml = pkg.text(rel.target);
			const root = parseXml(xml);
			const table = root.children.find((c) => localName(c.name) === "table");
			if (!table) continue;
			const splicer = new Splicer(xml);
			const a = parseArea(table.attrs.ref ?? "");
			if (a) {
				const m = mapArea(mapper, plan, sheet.name, a);
				splicer.replace(table.start, table.openEnd, setAttrs(rawStartTag(xml, table), { ref: formatAreaList([m]).replace(/^([A-Z]+\d+)$/, "$1:$1") }));
			}
			for (const af of table.children.filter((c) => localName(c.name) === "autoFilter")) patchAreaAttrs(xml, af, plan, mapper, splicer);
			const out = splicer.apply();
			if (out !== xml) pkg.setText(rel.target, out);
		} else if (rel.type === REL.drawing) {
			const xml = pkg.text(rel.target);
			const root = parseXml(xml);
			const splicer = new Splicer(xml);
			for (const anchor of findAll(root, "xdr:twoCellAnchor").concat(findAll(root, "xdr:oneCellAnchor"))) {
				const from = anchor.children.find((c) => localName(c.name) === "from");
				const to = anchor.children.find((c) => localName(c.name) === "to");
				if (!from) continue;
				const editAs = localName(anchor.name) === "oneCellAnchor" ? "oneCell" : anchor.attrs.editAs;
				shiftAnchor(xml, from, to, plan, splicer, editAs);
			}
			const out = splicer.apply();
			if (out !== xml) pkg.setText(rel.target, out);
		} else if (rel.type === REL.comments) {
			patchComments(ctx, rel.target, plan, mapper, warn, "comment");
		} else if (rel.type === REL.threadedComment) {
			patchComments(ctx, rel.target, plan, mapper, warn, "threadedComment");
		} else if (rel.type === REL.vmlDrawing) {
			patchVml(ctx, rel.target, plan);
		}
	}
}

function patchComments(ctx: RenderContext, part: string, plan: SheetPlan, mapper: WorkbookMapper, warn: (m: string) => void, tag: string): void {
	const xml = ctx.pkg.text(part);
	const root = parseXml(xml);
	const splicer = new Splicer(xml);
	for (const c of findAll(root, tag)) {
		const a = parseArea(c.attrs.ref ?? "");
		if (!a) continue;
		if (plan.layoutFor(a.r1)) {
			splicer.replaceElement(c, "");
			warn(`a comment on ${c.attrs.ref} sat on a sample data row and was removed`);
			continue;
		}
		const r = mapper.pointRow(plan, a.r1, undefined);
		if (r !== a.r1) splicer.replace(c.start, c.openEnd, setAttrs(rawStartTag(xml, c), { ref: cellName(a.c1, r) }));
	}
	const out = splicer.apply();
	if (out !== xml) ctx.pkg.setText(part, out);
}

/**
 * Legacy VML (comment shapes) is often not well-formed XML, so it is
 * edited with targeted patterns rather than parsed — the one exception to
 * the tokenizer rule.
 */
function patchVml(ctx: RenderContext, part: string, plan: SheetPlan): void {
	const xml = ctx.pkg.text(part);
	const out = xml.replace(/<v:shape\b[\s\S]*?<\/v:shape>/g, (shape) => {
		const rowM = /<x:Row>(\d+)<\/x:Row>/.exec(shape);
		if (!rowM || !/ObjectType="Note"/.test(shape)) return shape;
		const t = Number(rowM[1]) + 1;
		if (plan.layoutFor(t)) return "";
		const nr = plan.mapRow(t);
		const d = nr - t;
		if (d === 0) return shape;
		return shape
			.replace(/<x:Row>\d+<\/x:Row>/, `<x:Row>${nr - 1}</x:Row>`)
			.replace(/<x:Anchor>([^<]*)<\/x:Anchor>/, (_m, list: string) => {
				const v = list.split(",").map((s) => s.trim());
				if (v.length === 8) {
					v[2] = String(Number(v[2]) + d);
					v[6] = String(Number(v[6]) + d);
				}
				return `<x:Anchor>${v.join(", ")}</x:Anchor>`;
			});
	});
	if (out !== xml) ctx.pkg.setText(part, out);
}

/** Defined names, charts and pivot caches can point at any sheet. */
function patchWorkbookParts(ctx: RenderContext, mapper: WorkbookMapper): void {
	const { pkg, wb } = ctx;
	const outside: RefOrigin = { kind: "outside", sheet: undefined };
	const mapFormula = (f: string, sheet?: string) =>
		rewriteRefs(f, (ref, c) => mapper.map(ref, sheet === undefined ? outside : { kind: "outside", sheet }, c.inArgList));

	{
		const xml = pkg.text(wb.path);
		const root = parseXml(xml);
		const splicer = new Splicer(xml);
		for (const dn of findAll(root, "definedName")) {
			const text = textOf(xml, dn);
			const local = dn.attrs.localSheetId === undefined ? undefined : wb.sheets.find((s) => s.index === Number(dn.attrs.localSheetId))?.name;
			const nf = mapFormula(text, local);
			if (nf !== text) splicer.replace(dn.openEnd, dn.closeStart, escapeText(nf));
		}
		if (splicer.changed) pkg.setText(wb.path, splicer.apply());
	}

	for (const name of [...pkg.parts.keys()]) {
		if (/^xl\/charts\/chart(Ex)?\d*\.xml$/.test(name)) {
			const xml = pkg.text(name);
			const root = parseXml(xml);
			const splicer = new Splicer(xml);
			for (const f of findAll(root, "c:f").concat(findAll(root, "cx:f"))) {
				const text = textOf(xml, f);
				const nf = mapFormula(text);
				if (nf !== text) splicer.replace(f.openEnd, f.closeStart, escapeText(nf));
			}
			if (splicer.changed) pkg.setText(name, splicer.apply());
		} else if (/^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/.test(name)) {
			const xml = pkg.text(name);
			const root = parseXml(xml);
			const def = root.children.find((c) => localName(c.name) === "pivotCacheDefinition");
			const src = findAll(root, "worksheetSource")[0];
			if (!def || !src) continue;
			const splicer = new Splicer(xml);
			let touched = false;
			const plan = mapper.plan(src.attrs.sheet);
			if (plan && src.attrs.ref) {
				const a = parseArea(src.attrs.ref);
				if (a) {
					const m = mapArea(mapper, plan, plan.sheetName, a);
					splicer.replace(src.start, src.openEnd, setAttrs(rawStartTag(xml, src), { ref: formatAreaList([m]) }));
					touched = true;
				}
			}
			if (src.attrs.name && ctx.regions.some((r) => r.key === src.attrs.name && mapper.plan(r.sheet.name))) touched = true;
			if (touched) {
				splicer.replace(def.start, def.openEnd, setAttrs(rawStartTag(xml, def), { refreshOnLoad: "1" }));
				pkg.setText(name, splicer.apply());
			}
		}
	}
}

export { mapSqrefArea };
