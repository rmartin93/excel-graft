import { cellName, indexToCol } from "./a1.js";
import { Package, readWorkbook, type WorkbookInfo } from "./package.js";
import { discoverRegions, discoverScalars, type FieldType, type Pattern, type Region, type ScalarTarget } from "./regions.js";
import { render, type RenderOptions } from "./render.js";
import { readSharedStrings, readSheet, type SheetModel } from "./sheet.js";
import { readStyles } from "./styles.js";
import { readParts, writeParts, type Parts } from "./zip.js";

export type { RenderOptions } from "./render.js";

/** The data shape one region accepts, as inferred from its sample rows. */
export interface FieldSchema {
	name: string;
	/** Inferred from the sample value and its number format; "unknown" when the sample cell was empty. */
	type: FieldType;
}

export type ShapeSchema =
	| { kind: "rows"; fields: FieldSchema[] }
	| { kind: "groups"; fields: FieldSchema[]; childKey: string; child: ShapeSchema };

export interface RegionSchema {
	key: string;
	kind: "table" | "name";
	sheet: string;
	/** The sample rows that get replaced, e.g. `A5:C10`. */
	sampleRange: string;
	/** Rows inside the region kept as-is (grand totals and the like), e.g. `["A11:C11"]`. */
	fixedRows: string[];
	/** Column letter → field name. */
	columns: Record<string, string>;
	shape: ShapeSchema;
}

export interface ScalarSchema {
	key: string;
	kind: "placeholder" | "name";
	sheet: string;
	cell: string;
}

export interface TemplateSchema {
	regions: RegionSchema[];
	scalars: ScalarSchema[];
	/** Tables/names that were found but can't be filled, with the reason. */
	skipped: { key: string; reason: string }[];
}

/** @deprecated use TemplateSchema */
export type InspectResult = TemplateSchema;

interface Analysis {
	wb: WorkbookInfo;
	sheets: Map<string, SheetModel>;
	sst: string[];
	regions: Region[];
	scalars: ScalarTarget[];
	skipped: { key: string; reason: string }[];
}

/**
 * A loaded template. Parts are kept as exact bytes; `render()` patches
 * only what the data changes and copies everything else through. A single
 * Template can be rendered many times, concurrently.
 */
export class Template<TData extends object = Record<string, unknown>> {
	private analysis: Analysis | undefined;

	private constructor(private readonly parts: Parts) {}

	static async load<TData extends object = Record<string, unknown>>(input: Uint8Array | ArrayBuffer): Promise<Template<TData>> {
		return Template.loadSync<TData>(input);
	}

	static loadSync<TData extends object = Record<string, unknown>>(input: Uint8Array | ArrayBuffer): Template<TData> {
		const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
		const tpl = new Template<TData>(readParts(bytes));
		tpl.analyze();
		return tpl;
	}

	/** Repacks the template with no edits applied (tests the passthrough invariant). */
	async toBuffer(): Promise<Uint8Array> {
		return writeParts(this.parts);
	}

	private analyze(): Analysis {
		if (this.analysis) return this.analysis;
		const pkg = new Package(this.parts);
		const wb = readWorkbook(pkg);
		const sheets = new Map<string, SheetModel>();
		for (const s of wb.sheets) if (pkg.has(s.part)) sheets.set(s.part, readSheet(pkg, s.part));
		const sst = readSharedStrings(pkg, wb.path);
		const { regions, skipped } = discoverRegions(pkg, wb, sheets, sst, readStyles(pkg, wb.path));
		const scalars = discoverScalars(wb, sheets, sst).filter(
			(s) => !regions.some((r) => r.sheet === s.sheet && s.row >= r.bodyStart && s.row <= r.bodyEnd && s.col >= r.c1 && s.col <= r.c2),
		);
		this.analysis = { wb, sheets, sst, regions, scalars, skipped };
		return this.analysis;
	}

	/** Describes what the template accepts: its regions (and their inferred shapes) and scalar targets. */
	inspect(): TemplateSchema {
		const a = this.analyze();
		const seen = new Set<string>();
		const scalars: ScalarSchema[] = [];
		for (const s of a.scalars) {
			const k = `${s.key}|${s.kind}`;
			if (seen.has(k)) continue;
			seen.add(k);
			scalars.push({ key: s.key, kind: s.kind, sheet: s.sheet.name, cell: cellName(s.col, s.row) });
		}
		return {
			regions: a.regions.map((r) => ({
				key: r.key,
				kind: r.kind,
				sheet: r.sheet.name,
				sampleRange: `${indexToCol(r.c1)}${r.bodyStart}:${indexToCol(r.c2)}${r.bodyEnd}`,
				fixedRows: [...r.leading, ...r.trailing].map((role) => {
					const row = role.rows[0] as number;
					return `${indexToCol(r.c1)}${row}:${indexToCol(r.c2)}${row}`;
				}),
				columns: Object.fromEntries([...r.columnNames].map(([c, n]) => [indexToCol(c), n])),
				shape: shapeOf(r.unit),
			})),
			scalars,
			skipped: a.skipped,
		};
	}

	/** Renders the template with `data` and returns the .xlsx bytes. */
	async render(data: TData, options: RenderOptions = {}): Promise<Uint8Array> {
		return this.renderSync(data, options);
	}

	renderSync(data: TData, options: RenderOptions = {}): Uint8Array {
		const a = this.analyze();
		const pkg = new Package(new Map(this.parts));
		render({ pkg, wb: a.wb, sheets: a.sheets, sst: a.sst, regions: a.regions, scalars: a.scalars }, data as Record<string, unknown>, options);
		return writeParts(pkg.parts);
	}
}

function shapeOf(p: Pattern): ShapeSchema {
	if (p.kind === "leaf") return { kind: "rows", fields: p.fields.map((f) => ({ name: f.name, type: f.type })) };
	const byName = new Map([...(p.header?.fields ?? []), ...(p.footer?.fields ?? [])].map((f) => [f.name, { name: f.name, type: f.type }]));
	const fields = [...byName.values()];
	return { kind: "groups", fields, childKey: p.childKey, child: shapeOf(p.child) };
}
