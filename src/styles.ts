import type { Package } from "./package.js";
import { decodeXml, findAll, localName, parseXml } from "./xml.js";

/** Built-in number formats that display dates/times. */
const BUILTIN_DATE_IDS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58]);

export interface StyleInfo {
	/** For each cellXfs index: true when its number format displays a date or time. */
	isDate: boolean[];
}

export function readStyles(pkg: Package, workbookPath: string): StyleInfo {
	const rel = pkg.rels(workbookPath).find((r) => r.type.endsWith("/styles"));
	if (!rel || !pkg.has(rel.target)) return { isDate: [] };
	const xml = pkg.text(rel.target);
	const root = parseXml(xml);
	const custom = new Map<number, string>();
	for (const nf of findAll(root, "numFmt")) custom.set(Number(nf.attrs.numFmtId), decodeXml(nf.attrs.formatCode ?? ""));
	const cellXfs = findAll(root, "cellXfs")[0];
	const isDate = (cellXfs?.children ?? [])
		.filter((c) => localName(c.name) === "xf")
		.map((xf) => {
			const id = Number(xf.attrs.numFmtId ?? 0);
			const code = custom.get(id);
			return code === undefined ? BUILTIN_DATE_IDS.has(id) : isDateFormat(code);
		});
	return { isDate };
}

/** True when a format code displays a date or time (ignoring quoted text, escapes, colors and elapsed-time brackets). */
export function isDateFormat(code: string): boolean {
	const section = code.split(";")[0] ?? "";
	const stripped = section
		.replace(/"[^"]*"/g, "")
		.replace(/\\./g, "")
		.replace(/_.|\*./g, "")
		.replace(/\[(?:h+|m+|s+)\]/gi, "h")
		.replace(/\[[^\]]*\]/g, "");
	return /[dmyhs]/i.test(stripped) && !/^[#0.,%\sE+-]*$/i.test(stripped);
}
