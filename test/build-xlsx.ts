// Builds small .xlsx templates from a compact description, for regression tests
// that need one precise feature (ported from the adversarial review's repro kit).
/* eslint-disable */
// @ts-nocheck
import { strToU8, zipSync } from "fflate";

export type CellSpec = string | number | null | { v?: unknown; s?: number; raw?: string; t?: string };
export interface SheetSpec {
	name: string;
	rows: Record<number, Record<string, CellSpec>>;
	tables?: { name: string; ref: string; cols: (string | { name: string; attrs?: string; inner?: string })[]; totals?: boolean; headerRowCount?: number; afRef?: string }[];
	extra?: string;
	pre?: string;
	post?: string;
	beforeTables?: string;
	rels?: string[];
	parts?: Record<string, string>;
	ct?: string;
}
export interface NameSpec {
	name: string;
	ref: string;
	local?: number;
}

const esc = (s: unknown) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
function colIdx(l: string): number {
	let n = 0;
	for (const c of l) n = n * 26 + c.charCodeAt(0) - 64;
	return n;
}

export function mkxlsx(sheets: SheetSpec[], { names = [] as NameSpec[], date1904 = false, extraParts = {} as Record<string, string>, extraCt = "", extraWbRels = "" } = {}): Uint8Array {
	const files: Record<string, string> = {};
	let ct = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>`;
	let wbRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`;
	let sheetsXml = "";
	let tableN = 0;
	const sst: string[] = [];
	sheets.forEach((sh, i) => {
		const n = i + 1;
		ct += `<Override PartName="/xl/worksheets/sheet${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
		wbRels += `<Relationship Id="rId${n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${n}.xml"/>`;
		sheetsXml += `<sheet name="${esc(sh.name)}" sheetId="${n}" r:id="rId${n}"/>`;
		const rowNums = Object.keys(sh.rows).map(Number).sort((a, b) => a - b);
		let sd = "";
		for (const r of rowNums) {
			const cells = sh.rows[r];
			const cols = Object.keys(cells).sort((a, b) => colIdx(a) - colIdx(b));
			sd += `<row r="${r}">`;
			for (const c of cols) {
				let v = cells[c];
				let s = "";
				if (v && typeof v === "object" && !Array.isArray(v)) {
					s = v.s !== undefined ? ` s="${v.s}"` : "";
					if (v.raw !== undefined) { sd += `<c r="${c}${r}"${s}${v.t ? ` t="${v.t}"` : ""}>${v.raw}</c>`; continue; }
					v = v.v;
				}
				if (typeof v === "number") sd += `<c r="${c}${r}"${s}><v>${v}</v></c>`;
				else if (typeof v === "string" && v.startsWith("=")) sd += `<c r="${c}${r}"${s}><f>${esc(v.slice(1))}</f></c>`;
				else if (v === null || v === undefined) sd += `<c r="${c}${r}"${s}/>`;
				else { let k = sst.indexOf(String(v)); if (k < 0) { sst.push(String(v)); k = sst.length - 1; } sd += `<c r="${c}${r}"${s} t="s"><v>${k}</v></c>`; }
			}
			sd += `</row>`;
		}
		let rels = "";
		let tparts = "";
		(sh.tables || []).forEach((t, k) => {
			tableN++;
			const tn = tableN;
			ct += `<Override PartName="/xl/tables/table${tn}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>`;
			rels += `<Relationship Id="rIdT${k}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table${tn}.xml"/>`;
			tparts += `<tablePart r:id="rIdT${k}"/>`;
			const hdr = t.headerRowCount === 0 ? ` headerRowCount="0"` : "";
			const tot = t.totals ? ` totalsRowCount="1"` : "";
			const afRef = t.afRef || t.ref;
			files[`xl/tables/table${tn}.xml`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="${tn}" name="${t.name}" displayName="${t.name}" ref="${t.ref}"${hdr}${tot}>${t.headerRowCount === 0 ? "" : `<autoFilter ref="${afRef}"/>`}<tableColumns count="${t.cols.length}">${t.cols.map((c, j) => `<tableColumn id="${j + 1}" name="${esc(typeof c === "string" ? c : c.name)}"${typeof c === "object" && c.attrs ? " " + c.attrs : ""}>${typeof c === "object" && c.inner ? c.inner : ""}</tableColumn>`).join("")}</tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>`;
		});
		(sh.rels || []).forEach((rel) => { rels += rel; });
		if (rels) files[`xl/worksheets/_rels/sheet${n}.xml.rels`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`;
		files[`xl/worksheets/sheet${n}.xml`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" mc:Ignorable="x14ac xr">${sh.pre || ""}<sheetData>${sd}</sheetData>${sh.extra || ""}${sh.beforeTables || ""}${tparts ? `<tableParts count="${(sh.tables || []).length}">${tparts}</tableParts>` : ""}${sh.post || ""}</worksheet>`;
		for (const [k, v] of Object.entries(sh.parts || {})) files[k] = v;
		ct += sh.ct || "";
	});
	wbRels += extraWbRels + `</Relationships>`;
	ct += extraCt + `</Types>`;
	const dn = names.length ? `<definedNames>${names.map((d) => `<definedName name="${d.name}"${d.local !== undefined ? ` localSheetId="${d.local}"` : ""}>${esc(d.ref)}</definedName>`).join("")}</definedNames>` : "";
	files["[Content_Types].xml"] = ct;
	files["_rels/.rels"] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
	files["xl/workbook.xml"] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${date1904 ? '<workbookPr date1904="1"/>' : "<workbookPr/>"}<sheets>${sheetsXml}</sheets>${dn}<calcPr calcId="191029"/></workbook>`;
	files["xl/_rels/workbook.xml.rels"] = wbRels;
	files["xl/styles.xml"] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs><dxfs count="1"><dxf><font><b/></font></dxf></dxfs></styleSheet>`;
	files["xl/sharedStrings.xml"] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sst.length}" uniqueCount="${sst.length}">${sst.map((t) => `<si><t xml:space="preserve">${esc(t)}</t></si>`).join("")}</sst>`;
	files["[Content_Types].xml"] = files["[Content_Types].xml"].replace("</Types>", `<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>`);
	files["xl/_rels/workbook.xml.rels"] = files["xl/_rels/workbook.xml.rels"].replace("</Relationships>", `<Relationship Id="rIdSST" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>`);
	for (const [k, v] of Object.entries(extraParts)) files[k] = v;
	const z: Record<string, Uint8Array> = {};
	for (const [k, v] of Object.entries(files)) z[k] = strToU8(v);
	return zipSync(z);
}

