import { MAX_ROW, rowRuns } from "./a1.js";
import type { RefContext, RefReplacement, RefToken } from "./formula.js";
import { dateToSerial } from "./values.js";
import type { FieldDef, GroupPattern, Pattern, Region, Role } from "./regions.js";
import { formatRefToken } from "./sheet.js";

/** An instance of a group (or the whole region) in the output: a contiguous span of rows. */
export interface Scope {
	first: number;
	last: number;
	parent: Scope | undefined;
	pattern: GroupPattern | undefined;
	headerOut: number | undefined;
	footerOut: number | undefined;
}

export interface OutRow {
	out: number;
	/** Template row this row's formulas and constants come from. */
	proto: number;
	/** Template row this row's styles come from (differs for banded rows). */
	styleProto: number;
	role: Role;
	/** Values keyed by field name, or undefined for rows that only carry formulas/constants. */
	record: Record<string, unknown> | undefined;
	fields: FieldDef[];
	/** The innermost group instance containing this row (for headers/footers: their own group). */
	scope: Scope;
	/** True for the placeholder row written when a data array is empty. */
	blank: boolean;
}

export class RenderDataError extends Error {
	override name = "RenderDataError";
}

export class RegionLayout {
	readonly rows: OutRow[] = [];
	readonly root: Scope;
	private readonly roleAt: Int32Array;
	private readonly roleRowsCache = new Map<number, number[]>();

	constructor(
		readonly region: Region,
		readonly outStart: number,
		data: unknown[],
	) {
		this.root = { first: outStart, last: outStart, parent: undefined, pattern: undefined, headerOut: undefined, footerOut: undefined };
		let cursor = outStart;
		const push = (row: Omit<OutRow, "out">): OutRow => {
			if (cursor > MAX_ROW) {
				throw new RenderDataError(`${region.key}: output would pass Excel's ${MAX_ROW.toLocaleString("en-US")}-row limit`);
			}
			const r = { ...row, out: cursor++ };
			this.rows.push(r);
			return r;
		};
		for (const role of region.leading) push({ proto: role.rows[0] as number, styleProto: role.rows[0] as number, role, record: undefined, fields: [], scope: this.root, blank: false });
		const walk = (pattern: Pattern, items: unknown[], scope: Scope, path: string): void => {
			const list = items.length === 0 ? [undefined] : items;
			if (pattern.kind === "leaf") {
				list.forEach((item, i) => {
					const proto = i === 0 && pattern.firstFormulaProto !== undefined ? pattern.firstFormulaProto : pattern.formulaProto;
					const styleProto =
						i === 0 && pattern.firstStyle !== undefined
							? pattern.firstStyle
							: (pattern.styleCycle[(pattern.firstStyle !== undefined ? i - 1 : i) % pattern.styleCycle.length] as number);
					push({
						proto,
						styleProto,
						role: pattern.role,
						record: asRecord(item, `${path}[${i}]`),
						fields: pattern.fields,
						scope,
						blank: item === undefined,
					});
				});
				return;
			}
			list.forEach((item, i) => {
				const record = asRecord(item, `${path}[${i}]`);
				const g: Scope = { first: cursor, last: cursor, parent: scope, pattern, headerOut: undefined, footerOut: undefined };
				if (pattern.header) {
					g.headerOut = push({ proto: pattern.header.proto, styleProto: pattern.header.proto, role: pattern.header.role, record, fields: pattern.header.fields, scope: g, blank: item === undefined }).out;
				}
				const kids = record?.[pattern.childKey];
				if (kids !== undefined && !Array.isArray(kids)) {
					throw new RenderDataError(`${path}[${i}].${pattern.childKey} must be an array`);
				}
				walk(pattern.child, (kids as unknown[] | undefined) ?? [], g, `${path}[${i}].${pattern.childKey}`);
				if (pattern.footer) {
					g.footerOut = push({ proto: pattern.footer.proto, styleProto: pattern.footer.proto, role: pattern.footer.role, record, fields: pattern.footer.fields, scope: g, blank: item === undefined }).out;
				}
				for (const sp of pattern.spacers) push({ proto: sp.proto, styleProto: sp.proto, role: sp.role, record: undefined, fields: [], scope: g, blank: false });
				g.last = cursor - 1;
			});
		};
		walk(region.unit, data, this.root, region.key);
		for (const role of region.trailing) push({ proto: role.rows[0] as number, styleProto: role.rows[0] as number, role, record: undefined, fields: [], scope: this.root, blank: false });
		this.root.last = cursor - 1;
		this.roleAt = new Int32Array(this.rows.length);
		this.rows.forEach((r, i) => (this.roleAt[i] = r.role.id));
	}

	get outEnd(): number {
		return this.root.last;
	}

	get templateLength(): number {
		return this.region.bodyEnd - this.region.bodyStart + 1;
	}

	get delta(): number {
		return this.outEnd - this.outStart + 1 - this.templateLength;
	}

	inBody(templateRow: number): boolean {
		return templateRow >= this.region.bodyStart && templateRow <= this.region.bodyEnd;
	}

	rowAt(out: number): OutRow | undefined {
		return this.rows[out - this.outStart];
	}

	/** Output rows instantiated from the role of `templateRow`, within `scope` (sorted). */
	project(templateRows: Iterable<number>, scope: Scope = this.root): number[] {
		const roles = new Set<number>();
		for (const t of templateRows) {
			const role = this.region.roleOfRow.get(t);
			if (role) roles.add(role.id);
		}
		if (scope === this.root) {
			const out: number[] = [];
			for (const id of roles) out.push(...this.roleRows(id));
			return roles.size > 1 ? out.sort((a, b) => a - b) : out;
		}
		const out: number[] = [];
		for (let o = scope.first; o <= scope.last; o++) {
			if (roles.has(this.roleAt[o - this.outStart] as number)) out.push(o);
		}
		return out;
	}

	private roleRows(id: number): number[] {
		let cached = this.roleRowsCache.get(id);
		if (!cached) {
			cached = [];
			for (let i = 0; i < this.roleAt.length; i++) if (this.roleAt[i] === id) cached.push(this.outStart + i);
			this.roleRowsCache.set(id, cached);
		}
		return cached;
	}

	/**
	 * Where a position inside the body lands, for things anchored to rows
	 * (pictures, shapes): a row with a role goes to its first output row; a
	 * cleared row (blank capacity) follows the nearest row above it that has one.
	 */
	anchorRow(t: number): number {
		if (this.region.roleOfRow.has(t)) return this.project([t])[0] ?? this.outStart;
		for (let r = t - 1; r >= this.region.bodyStart; r--) {
			if (!this.region.roleOfRow.has(r)) continue;
			const rows = this.project([r]);
			return (rows[rows.length - 1] ?? this.outStart) + (t - r);
		}
		return this.outStart;
	}

	/** Rows of the body in [from, to] (template), clipped. */
	bodyRange(from: number, to: number): number[] {
		const rows: number[] = [];
		for (let t = Math.max(from, this.region.bodyStart); t <= Math.min(to, this.region.bodyEnd); t++) rows.push(t);
		return rows;
	}
}

function asRecord(item: unknown, path: string): Record<string, unknown> | undefined {
	if (item === undefined) return undefined;
	if (typeof item !== "object" || item === null || Array.isArray(item) || item instanceof Date) {
		throw new RenderDataError(`${path} must be an object with one property per column`);
	}
	return item as Record<string, unknown>;
}

/** All rendered regions on one sheet, plus the row mapping for everything outside them. */
export class SheetPlan {
	constructor(
		readonly sheetName: string,
		readonly layouts: RegionLayout[],
	) {}

	/** Maps a template row that is not inside any rendered body. */
	mapRow(t: number): number {
		// A range that runs to the bottom of the sheet keeps running to the bottom (Excel does the same on insert).
		if (t >= MAX_ROW) return MAX_ROW;
		let d = 0;
		for (const l of this.layouts) if (t > l.region.bodyEnd) d += l.delta;
		return Math.min(t + d, MAX_ROW);
	}

	layoutFor(t: number): RegionLayout | undefined {
		return this.layouts.find((l) => l.inBody(t));
	}
}

/** Where a formula (or other reference holder) lives, which decides how its references follow the data. */
export type RefOrigin =
	| { kind: "outside"; sheet: string | undefined }
	| { kind: "instance"; sheet: string; layout: RegionLayout; row: OutRow }
	/** CF/DV style formula: relative refs are anchored at a cell that moved from `from` to `to`. */
	| { kind: "anchored"; sheet: string; from: number; to: number; layout: RegionLayout | undefined };

export class WorkbookMapper {
	private readonly warned = new Set<string>();

	constructor(
		readonly plans: Map<string, SheetPlan>,
		private readonly warn: (message: string) => void = () => {},
		private readonly options: { date1904?: boolean; sheetOrder?: string[] } = {},
	) {}

	private warnOnce(message: string): void {
		if (this.warned.has(message)) return;
		this.warned.add(message);
		this.warn(message);
	}

	/**
	 * A single sample cell referenced from outside its group: the output row
	 * that plays the same part. Group rows are matched by label (a rate sheet's
	 * reference to the "Overhead" pool row follows the Overhead pool wherever
	 * it lands); other rows by position. No match means #REF!, never a
	 * different row's numbers.
	 */
	correspondingRow(lay: RegionLayout, t: number): number {
		const region = lay.region;
		const role = region.roleOfRow.get(t);
		const candidates = lay.project([t]);
		if (!role || candidates.length === 0) return lay.outStart;
		const label = region.sampleLabels.get(t);
		if (label !== undefined && (role.kind === "header" || role.kind === "footer")) {
			for (const o of candidates) {
				const r = lay.rowAt(o);
				if (r && this.labelPath(lay, r).toLowerCase() === label.toLowerCase()) return o;
			}
			this.warnOnce(`"${region.key}" has no group labeled "${label}" in the data, so references to that sample group (row ${t}) became #REF!`);
			return 0;
		}
		const k = role.rows.indexOf(t);
		const hit = candidates[k];
		if (hit !== undefined) return hit;
		this.warnOnce(`"${region.key}": a reference to sample row ${t} (row ${k + 1} of its kind) has no counterpart in the data, so it became #REF!`);
		return 0;
	}

	private formulaPoint(plan: SheetPlan, t: number): number {
		const lay = plan.layoutFor(t);
		return lay ? this.correspondingRow(lay, t) : plan.mapRow(t);
	}

	get active(): boolean {
		return this.plans.size > 0;
	}

	plan(sheet: string | undefined): SheetPlan | undefined {
		return sheet === undefined ? undefined : this.plans.get(sheet);
	}

	/** "Sales › Travel": a group row's label plus its parents', matching Region.sampleLabels. */
	private labelPath(lay: RegionLayout, r: OutRow): string {
		const parts: string[] = [];
		for (let s: Scope | undefined = r.scope; s?.pattern; s = s.parent) {
			const at = s.headerOut ?? s.footerOut;
			const row = at === undefined ? undefined : lay.rowAt(at);
			parts.unshift(row ? outputLabel(row, this.options.date1904 ?? false) : "");
		}
		return parts.join(" › ");
	}

	/**
	 * A range that lies inside one sample group (MAX(C4:C5) over Fringe's
	 * accounts) means that group: find the output group that corresponds to
	 * it. Returns null when the data has no such group (the range becomes #REF!).
	 */
	private sampleGroupScope(lay: RegionLayout, ra: number, rb: number): Scope | null | undefined {
		let best: { row: number; first: number; last: number } | undefined;
		for (const b of lay.region.sampleBlocks) {
			if (b.first <= ra && rb <= b.last && (!best || b.last - b.first < best.last - best.first)) best = b;
		}
		if (!best) return undefined;
		const o = this.correspondingRow(lay, best.row);
		if (o === 0) return null;
		return lay.rowAt(o)?.scope;
	}

	/** 3-D references (Jan:Feb!B7) can't follow per-sheet row moves; say so instead of leaving them silently stale. */
	private check3d(ref: RefToken): void {
		if (!ref.sheet || !ref.sheet.includes(":") || ref.sheet.startsWith("[")) return;
		const [first, last] = ref.sheet.split(":") as [string, string];
		const order = this.options.sheetOrder ?? [];
		const i = order.indexOf(first);
		const j = order.indexOf(last);
		const span = i >= 0 && j >= 0 ? order.slice(Math.min(i, j), Math.max(i, j) + 1) : [first, last];
		const b = ref.b ?? ref.a;
		const moved = span.some((name) => {
			const p = this.plans.get(name);
			return p !== undefined && (p.layoutFor(ref.a.row) !== undefined || p.layoutFor(b.row) !== undefined || p.mapRow(ref.a.row) !== ref.a.row || p.mapRow(b.row) !== b.row);
		});
		if (moved) {
			const text = formatRefToken(ref.prefix, ref.shape, ref.a, ref.b, ref.spill);
			this.warnOnce(`the 3-D reference ${text} spans sheets whose rows moved, so it was left unchanged and may be wrong; reference each sheet (or a Table) instead`);
		}
	}

	/** Returns a replacement for `ref`, or undefined to leave it alone. */
	map(ref: RefToken, origin: RefOrigin, ctx: Pick<RefContext, "inArgList" | "inUnion" | "func"> & Partial<Pick<RefContext, "siblings">>): RefReplacement | string | undefined {
		const { inArgList, inUnion } = ctx;
		if (ref.shape === "cols") return undefined;
		if (ref.foreign) {
			this.check3d(ref);
			return undefined;
		}
		const targetSheet = ref.sheet ?? origin.sheet;
		const plan = this.plan(targetSheet);
		const sameSheet = ref.sheet === undefined || ref.sheet === origin.sheet;
		const originLayout = origin.kind === "instance" && sameSheet ? origin.layout : undefined;
		if (!plan && !originLayout && origin.kind !== "anchored") return undefined;

		const b = ref.b ?? ref.a;
		const ra = Math.min(ref.a.row, b.row);
		const rb = Math.max(ref.a.row, b.row);
		const aIsTop = ref.a.row <= b.row;

		// A fully absolute range ($B$3:$B$5) names a block of rows, not a fill-down offset: it grows with the data.
		const absRange = ref.b !== undefined && (ref.shape === "area" || ref.shape === "rows") && ref.a.rowAbs && ref.b.rowAbs;

		if (origin.kind === "instance" && sameSheet && origin.row.role.kind === "leaf" && !absRange) {
			// Fill-down semantics for detail rows.
			const mapEnd = (c: RefToken["a"]) => ({ ...c, row: this.leafRow(origin, c.row, c.rowAbs, plan) });
			return formatRefToken(ref.prefix, ref.shape, mapEnd(ref.a), ref.b ? mapEnd(ref.b) : undefined, ref.spill);
		}

		if (origin.kind === "anchored" && sameSheet && !absRange) {
			const mapEnd = (c: RefToken["a"]) => {
				const lay = origin.layout;
				if (!c.rowAbs && lay && lay.inBody(origin.from) && lay.inBody(c.row)) return { ...c, row: origin.to + (c.row - origin.from) };
				return { ...c, row: this.pointRow(plan, c.row, undefined) };
			};
			if (!plan) return undefined;
			return formatRefToken(ref.prefix, ref.shape, mapEnd(ref.a), ref.b ? mapEnd(ref.b) : undefined, ref.spill);
		}
		if (!plan) return undefined;

		const scope = origin.kind === "instance" && sameSheet ? origin.row.scope : undefined;
		if (ref.shape === "cell") {
			const t = ref.a.row;
			const lay = plan.layoutFor(t);
			// A sample cell listed in an aggregate stands for every row of its kind when the list says so:
			// a group/total row totalling its children, or two or more sample rows of one kind (SUM(C3,C6)).
			// MAX(C6,0) from outside names one specific group.
			const role = lay?.region.roleOfRow.get(t);
			const listsKind =
				(origin.kind === "instance" && sameSheet && origin.row.role.kind !== "leaf") ||
				(ctx.siblings ?? [ref]).filter((s) => s.shape === "cell" && (s.sheet ?? origin.sheet) === targetSheet && lay?.region.roleOfRow.get(s.a.row) === role).length >= 2;
			if (lay && inArgList && (inUnion || AGGREGATES.has(ctx.func ?? "")) && listsKind) {
				// Always dedupe: two sample groups listed side by side must not count one output group twice.
				const inScope = scope && this.scopeInLayout(scope, lay) ? lay.project([t], scope) : [];
				const rows = inScope.length > 0 ? inScope : lay.project([t]);
				if (rows.length > 0) {
					return { text: expandList(ref, rows, inUnion), dedupeKey: `${ref.sheet ?? ""}|${lay.region.roleOfRow.get(t)?.id}|${ref.a.col}|${ref.a.colAbs}` };
				}
			}
			const scoped = lay !== undefined && scope !== undefined && scope !== lay.root && this.scopeInLayout(scope, lay);
			const row = scoped ? this.pointRow(plan, t, scope) : this.formulaPoint(plan, t);
			return formatRefToken(ref.prefix, "cell", { ...ref.a, row }, undefined, ref.spill);
		}
		let rangeScope = scope;
		const lay = plan.layoutFor(ra);
		if (lay && lay === plan.layoutFor(rb) && !(scope && scope !== lay.root && this.scopeInLayout(scope, lay))) {
			const g = this.sampleGroupScope(lay, ra, rb);
			if (g === null) return `${ref.prefix}#REF!`;
			if (g) rangeScope = g;
		}
		const top = this.rangeStart(plan, ra, rb, rangeScope);
		const bottom = this.rangeEnd(plan, ra, rb, rangeScope);
		const a = { ...ref.a, row: aIsTop ? top : bottom };
		const bb = { ...b, row: aIsTop ? bottom : top };
		return formatRefToken(ref.prefix, ref.shape, a, ref.b ? bb : undefined, ref.spill);
	}

	/** A single template row as seen from outside the data (first matching output row). */
	pointRow(plan: SheetPlan | undefined, t: number, scope: Scope | undefined): number {
		if (!plan) return t;
		const lay = plan.layoutFor(t);
		if (!lay) return plan.mapRow(t);
		const inScope = scope && scope !== lay.root && this.scopeInLayout(scope, lay) ? lay.project([t], scope) : [];
		const rows = inScope.length > 0 ? inScope : lay.project([t]);
		return rows[0] ?? lay.outStart;
	}

	rangeStart(plan: SheetPlan, ra: number, rb: number, scope: Scope | undefined): number {
		const lay = plan.layoutFor(ra);
		if (!lay) return plan.mapRow(ra);
		const rows = this.projectRange(lay, lay.bodyRange(ra, rb), scope);
		return rows.length > 0 ? Math.min(...rows.slice(0, 1)) : lay.outStart;
	}

	rangeEnd(plan: SheetPlan, ra: number, rb: number, scope: Scope | undefined): number {
		const lay = plan.layoutFor(rb);
		if (!lay) return plan.mapRow(rb);
		const rows = this.projectRange(lay, lay.bodyRange(ra, rb), scope);
		return rows.length > 0 ? (rows[rows.length - 1] as number) : lay.outEnd;
	}

	private projectRange(lay: RegionLayout, templateRows: number[], scope: Scope | undefined): number[] {
		if (scope && scope !== lay.root && this.scopeInLayout(scope, lay)) {
			const rows = lay.project(templateRows, scope);
			if (rows.length > 0) return rows;
		}
		return lay.project(templateRows);
	}

	private scopeInLayout(scope: Scope, lay: RegionLayout): boolean {
		return scope.first >= lay.outStart && scope.last <= lay.outEnd;
	}

	/** One endpoint of a reference written in a detail row, following Excel's fill-down rules. */
	private leafRow(origin: Extract<RefOrigin, { kind: "instance" }>, t: number, abs: boolean, plan: SheetPlan | undefined): number {
		const { layout, row } = origin;
		const region = layout.region;
		if (layout.inBody(t)) {
			const role = region.roleOfRow.get(t);
			if (role === row.role) {
				if (!abs) return row.out + (t - row.proto);
				return this.firstInScope(layout, t, row.scope);
			}
			if (role && (role.kind === "header" || role.kind === "footer" || role.kind === "spacer")) {
				for (let s: Scope | undefined = row.scope; s; s = s.parent) {
					if (s.pattern?.header?.role === role && s.headerOut !== undefined) return s.headerOut;
					if (s.pattern?.footer?.role === role && s.footerOut !== undefined) return s.footerOut;
				}
			}
			return layout.project([t])[0] ?? layout.outStart;
		}
		if (!abs) return row.out + (t - row.proto);
		return plan ? this.formulaPoint(plan, t) : t;
	}

	private firstInScope(layout: RegionLayout, t: number, scope: Scope): number {
		const rows = layout.project([t], scope);
		return rows[0] ?? layout.project([t])[0] ?? layout.outStart;
	}
}

/** `C5,C9,C13` (runs collapse to ranges); wrapped as a union when it would exceed Excel's 255 arguments. */
function expandList(ref: RefToken, rows: number[], inUnion: boolean): string {
	const parts = rowRuns(rows).map(([s, e]) =>
		s === e
			? formatRefToken(ref.prefix, "cell", { ...ref.a, row: s }, undefined)
			: formatRefToken(ref.prefix, "area", { ...ref.a, row: s }, { ...ref.a, row: e }),
	);
	const joined = parts.join(",");
	return parts.length > 200 && !inUnion ? `(${joined})` : joined;
}

/**
 * Functions whose arguments are all "more of the same", so a sample cell
 * listed as an argument means every row of its kind. Anywhere else
 * (ROUND(C6,2), IF(...)) a sample cell means one specific row.
 */
const AGGREGATES = new Set([
	"SUM", "SUMSQ", "PRODUCT", "AVERAGE", "AVERAGEA", "COUNT", "COUNTA", "MAX", "MAXA", "MIN", "MINA", "MEDIAN",
	"MODE", "MODE.SNGL", "MODE.MULT", "STDEV", "STDEV.S", "STDEV.P", "STDEVA", "STDEVP", "STDEVPA", "VAR", "VAR.S",
	"VAR.P", "VARA", "VARP", "VARPA", "GEOMEAN", "HARMEAN", "AVEDEV", "DEVSQ", "SUBTOTAL", "AGGREGATE", "CONCAT",
	"CONCATENATE", "TEXTJOIN", "VSTACK", "HSTACK", "CHOOSE",
]);

function display(v: unknown, date1904: boolean): string {
	if (v === null || v === undefined) return "";
	// Template-side labels of date cells are their serial numbers.
	if (v instanceof Date) return Number.isNaN(v.getTime()) ? "" : String(dateToSerial(v, date1904));
	return String(v).trim();
}

/** Same shape as the template side's rowLabel(): the label field texts. */
function outputLabel(r: OutRow, date1904: boolean): string {
	return r.fields.map((f) => display(r.record?.[f.name], date1904)).join(" / ");
}
