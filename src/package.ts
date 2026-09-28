import type { Parts } from "./zip.js";
import { children, findAll, parseXml, rawStartTag, setAttrs, Splicer, textOf, type XmlElement } from "./xml.js";

const decoder = new TextDecoder("utf-8");
const encoder = new TextEncoder();

export const REL = {
	worksheet: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
	table: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table",
	calcChain: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain",
	drawing: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
	chart: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
	comments: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
	vmlDrawing: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
	threadedComment: "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment",
	pivotCacheDefinition: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition",
	officeDocument: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
} as const;

export interface Relationship {
	id: string;
	type: string;
	/** Resolved absolute part name without leading slash (e.g. `xl/tables/table1.xml`), or the raw target for external links. */
	target: string;
	external: boolean;
}

/**
 * Mutable view over the package parts. Text is decoded lazily and parts
 * are only re-encoded when `setText` is called, so untouched parts keep
 * their exact bytes.
 */
export class Package {
	private readonly textCache = new Map<string, string>();

	constructor(readonly parts: Parts) {}

	has(name: string): boolean {
		return this.parts.has(name);
	}

	text(name: string): string {
		let t = this.textCache.get(name);
		if (t === undefined) {
			const data = this.parts.get(name);
			if (!data) throw new Error(`missing part ${name}`);
			t = decoder.decode(data);
			this.textCache.set(name, t);
		}
		return t;
	}

	setText(name: string, text: string): void {
		if (this.textCache.get(name) === text && this.parts.has(name)) return;
		this.textCache.set(name, text);
		this.parts.set(name, encoder.encode(text));
	}

	delete(name: string): void {
		this.parts.delete(name);
		this.textCache.delete(name);
	}

	relsPath(partName: string): string {
		const slash = partName.lastIndexOf("/");
		return `${partName.slice(0, slash + 1)}_rels/${partName.slice(slash + 1)}.rels`;
	}

	rels(partName: string): Relationship[] {
		const path = this.relsPath(partName);
		if (!this.has(path)) return [];
		const xml = this.text(path);
		const root = parseXml(xml);
		const base = partName.slice(0, partName.lastIndexOf("/") + 1);
		return findAll(root, "Relationship").map((el) => {
			const external = el.attrs.TargetMode === "External";
			const raw = el.attrs.Target ?? "";
			return {
				id: el.attrs.Id ?? "",
				type: el.attrs.Type ?? "",
				target: external ? raw : resolvePath(base, raw),
				external,
			};
		});
	}

	/** Removes relationships matching `predicate` from `partName`'s rels. */
	removeRels(partName: string, predicate: (r: Relationship) => boolean): void {
		const path = this.relsPath(partName);
		if (!this.has(path)) return;
		const rels = this.rels(partName);
		const xml = this.text(path);
		const root = parseXml(xml);
		const splicer = new Splicer(xml);
		const els = findAll(root, "Relationship");
		els.forEach((el, i) => {
			const rel = rels[i];
			if (rel && predicate(rel)) splicer.replaceElement(el, "");
		});
		if (splicer.changed) this.setText(path, splicer.apply());
	}

	/** Removes a content-type Override for `partName` (e.g. after deleting calcChain). */
	removeOverride(partName: string): void {
		const path = "[Content_Types].xml";
		if (!this.has(path)) return;
		const xml = this.text(path);
		const root = parseXml(xml);
		const splicer = new Splicer(xml);
		for (const el of findAll(root, "Override")) {
			if (el.attrs.PartName === `/${partName}`) splicer.replaceElement(el, "");
		}
		if (splicer.changed) this.setText(path, splicer.apply());
	}

	mainWorkbookPath(): string {
		const rel = this.rels("").find((r) => r.type === REL.officeDocument);
		return rel?.target ?? "xl/workbook.xml";
	}
}

/** Resolves a relationship target against the source part's directory. */
export function resolvePath(baseDir: string, target: string): string {
	const segments = (target.startsWith("/") ? target.slice(1) : baseDir + target).split("/");
	const out: string[] = [];
	for (const seg of segments) {
		if (seg === "..") out.pop();
		else if (seg !== "." && seg !== "") out.push(seg);
	}
	return out.join("/");
}

export interface SheetInfo {
	name: string;
	/** 0-based position in `<sheets>`, which is what `localSheetId` refers to. */
	index: number;
	part: string;
	state: string;
}

export interface DefinedName {
	name: string;
	localSheetId: number | undefined;
	hidden: boolean;
	formula: string;
	el: XmlElement;
}

export interface TableInfo {
	part: string;
	name: string;
	displayName: string;
	sheet: SheetInfo;
}

export interface WorkbookInfo {
	path: string;
	sheets: SheetInfo[];
	names: DefinedName[];
	tables: TableInfo[];
	date1904: boolean;
}

export function readWorkbook(pkg: Package): WorkbookInfo {
	const path = pkg.mainWorkbookPath();
	const xml = pkg.text(path);
	const root = parseXml(xml);
	const rels = pkg.rels(path);
	const sheets: SheetInfo[] = [];
	const sheetEls = findAll(root, "sheet");
	sheetEls.forEach((el, index) => {
		const rid = el.attrs["r:id"] ?? Object.entries(el.attrs).find(([k]) => k.endsWith(":id"))?.[1];
		const rel = rels.find((r) => r.id === rid);
		if (!rel || rel.type !== REL.worksheet) return; // chartsheets, dialogsheets
		sheets.push({ name: el.attrs.name ?? "", index, part: rel.target, state: el.attrs.state ?? "visible" });
	});

	const names: DefinedName[] = findAll(root, "definedName").map((el) => ({
		name: el.attrs.name ?? "",
		localSheetId: el.attrs.localSheetId === undefined ? undefined : Number(el.attrs.localSheetId),
		hidden: el.attrs.hidden === "1" || el.attrs.hidden === "true",
		formula: textOf(xml, el),
		el,
	}));

	const pr = findAll(root, "workbookPr")[0];
	const date1904 = pr?.attrs.date1904 === "1" || pr?.attrs.date1904 === "true";

	const tables: TableInfo[] = [];
	for (const sheet of sheets) {
		for (const rel of pkg.rels(sheet.part)) {
			if (rel.type !== REL.table || !pkg.has(rel.target)) continue;
			const t = parseXml(pkg.text(rel.target));
			const tableEl = children(t, "table")[0];
			if (!tableEl) continue;
			tables.push({
				part: rel.target,
				name: tableEl.attrs.name ?? "",
				displayName: tableEl.attrs.displayName ?? tableEl.attrs.name ?? "",
				sheet,
			});
		}
	}

	return { path, sheets, names, tables, date1904 };
}

/**
 * Deletes calcChain (part, relationship and content-type override) and
 * forces a full recalculation on open. See CLAUDE.md invariants 4 and 5.
 */
export function forceRecalc(pkg: Package, wb: WorkbookInfo): void {
	const rels = pkg.rels(wb.path);
	for (const rel of rels) {
		if (rel.type === REL.calcChain) {
			pkg.delete(rel.target);
			pkg.removeOverride(rel.target);
		}
	}
	pkg.removeRels(wb.path, (r) => r.type === REL.calcChain);

	const xml = pkg.text(wb.path);
	const root = parseXml(xml);
	const splicer = new Splicer(xml);
	const calcPr = findAll(root, "calcPr")[0];
	if (calcPr) {
		splicer.replace(calcPr.start, calcPr.openEnd, setAttrs(rawStartTag(xml, calcPr), { fullCalcOnLoad: "1" }));
	} else {
		const workbook = root.children.find((c) => c.name === "workbook");
		if (!workbook) throw new Error("workbook.xml has no <workbook> root");
		// calcPr follows sheets/functionGroups/externalReferences/definedNames.
		const before = ["definedNames", "externalReferences", "functionGroups", "sheets"];
		const anchor = before.map((n) => workbook.children.find((c) => c.name === n)).find((c) => c !== undefined);
		if (!anchor) throw new Error("workbook.xml has no <sheets>");
		splicer.insert(anchor.end, '<calcPr fullCalcOnLoad="1"/>');
	}
	pkg.setText(wb.path, splicer.apply());
}
