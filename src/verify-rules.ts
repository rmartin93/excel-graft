import { parseCell } from "./a1.js";
import { tokenize } from "./formula.js";
import { localName, type XmlElement } from "./xml.js";

/**
 * Schema-order and syntax rules behind Excel's repair prompt that aren't
 * about cell data: element order in workbook.xml and worksheets, defined
 * and table names, and formula syntax.
 */

/** CT_Workbook child order (main namespace). */
export const WORKBOOK_ORDER = [
	"fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews", "sheets", "functionGroups",
	"externalReferences", "definedNames", "calcPr", "oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr",
	"smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst",
];

/** CT_Worksheet child order (main namespace). conditionalFormatting may repeat. */
export const WORKSHEET_ORDER = [
	"sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetCalcPr", "sheetProtection",
	"protectedRanges", "scenarios", "autoFilter", "sortState", "dataConsolidate", "customSheetViews", "mergeCells",
	"phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup",
	"headerFooter", "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing",
	"legacyDrawing", "legacyDrawingHF", "drawingHF", "picture", "oleObjects", "controls", "webPublishItems", "tableParts",
	"extLst",
];

const REPEATABLE = new Set(["conditionalFormatting"]);

/** Checks that `root`'s unprefixed children follow `order`, each at most once. Prefixed (mc:, x14ac:...) children are skipped. */
export function checkChildOrder(root: XmlElement, order: string[], what: string): string[] {
	const problems: string[] = [];
	const seen = new Set<string>();
	let last = -1;
	let lastName = "";
	for (const c of root.children) {
		if (c.name.includes(":")) continue;
		const name = localName(c.name);
		const idx = order.indexOf(name);
		if (idx === -1) continue;
		if (seen.has(name) && !REPEATABLE.has(name)) problems.push(`<${name}> appears more than once in ${what}`);
		if (idx < last) problems.push(`<${name}> must come before <${lastName}> in ${what} (Excel repairs out-of-order elements)`);
		else {
			last = idx;
			lastName = name;
		}
		seen.add(name);
	}
	return problems;
}

/** Why `name` isn't a valid defined/table name, or undefined when it is. */
export function badName(name: string, kind: "defined" | "table"): string | undefined {
	if (name.length === 0) return "is empty";
	if (name.length > 255) return "is longer than 255 characters";
	const bare = kind === "defined" ? name.replace(/^_xlnm\./, "") : name;
	if (!/^[A-Za-z_\\¡-￿][A-Za-z0-9_.\\?¡-￿]*$/.test(bare)) return "has characters Excel doesn't allow in names";
	if (parseCell(bare) || /^R(\d+)?C(\d+)?$/i.test(bare) || /^[RC]$/i.test(bare)) return "looks like a cell reference";
	return undefined;
}

/** Fixed argument counts for common functions: [min, max]. */
const ARITY: Record<string, [number, number]> = {
	ABS: [1, 1], INT: [1, 1], SQRT: [1, 1], LEN: [1, 1], UPPER: [1, 1], LOWER: [1, 1], TRIM: [1, 1], NOT: [1, 1],
	ISBLANK: [1, 1], ISNUMBER: [1, 1], ISTEXT: [1, 1], ISERROR: [1, 1], ISNA: [1, 1], YEAR: [1, 1], MONTH: [1, 1],
	DAY: [1, 1], ROUND: [2, 2], ROUNDUP: [2, 2], ROUNDDOWN: [2, 2], MOD: [2, 2], POWER: [2, 2], IFERROR: [2, 2],
	IFNA: [2, 2], TEXT: [2, 2], LEFT: [1, 2], RIGHT: [1, 2], MID: [3, 3], DATE: [3, 3], TIME: [3, 3], IF: [1, 3],
	VLOOKUP: [3, 4], HLOOKUP: [3, 4], MATCH: [2, 3], INDEX: [2, 4], SUMIF: [2, 3], COUNTIF: [2, 2], AVERAGEIF: [2, 3],
	EOMONTH: [2, 2], EDATE: [2, 2], FIND: [2, 3], SEARCH: [2, 3], SUBSTITUTE: [3, 4], REPT: [2, 2], VALUE: [1, 1],
	XLOOKUP: [3, 6], FILTER: [2, 3], UNIQUE: [1, 3], SORT: [1, 4], ROWS: [1, 1], COLUMNS: [1, 1], ROW: [0, 1],
	COLUMN: [0, 1], TODAY: [0, 0], NOW: [0, 0], PI: [0, 0], NA: [0, 0], TRUE: [0, 0], FALSE: [0, 0], OFFSET: [3, 5],
	INDIRECT: [1, 2], CHOOSE: [2, 255], SUBTOTAL: [2, 255], SUMPRODUCT: [1, 255],
};

/** Syntax problems Excel won't load: unbalanced brackets/quotes, too many arguments, too deep, wrong argument counts. */
export function formulaProblems(formula: string): string[] {
	const problems: string[] = [];
	const tokens = tokenize(formula);
	const last = tokens[tokens.length - 1];
	if (last && last.kind === "str" && !/"$/.test(formula.slice(last.start, last.end))) problems.push("has an unclosed string");
	if (last && last.kind === "struct" && !formula.slice(last.start, last.end).endsWith("]")) problems.push("has an unclosed [ ]");
	// Paren frames: function name (or "" for grouping), argument count.
	const stack: { func: string; args: number; empty: boolean }[] = [];
	let maxDepth = 0;
	for (let k = 0; k < tokens.length; k++) {
		const t = tokens[k]!;
		const text = formula.slice(t.start, t.end);
		if (t.kind === "ws") continue;
		if (t.kind === "op" && (text === "(" || text === "{")) {
			const prev = tokens.slice(0, k).reverse().find((x) => x.kind !== "ws");
			const func = text === "(" && prev?.kind === "func" ? formula.slice(prev.start, prev.end).toUpperCase().replace(/^(_XLFN\.|_XLWS\.)+/, "") : "";
			stack.push({ func: text === "{" ? "{" : func, args: 1, empty: true });
			maxDepth = Math.max(maxDepth, stack.length);
			continue;
		}
		if (t.kind === "op" && (text === ")" || text === "}")) {
			const frame = stack.pop();
			if (!frame) {
				problems.push(`has an unmatched "${text}"`);
				continue;
			}
			if ((frame.func === "{") !== (text === "}")) problems.push("has mismatched brackets");
			const args = frame.empty ? 0 : frame.args;
			if (frame.func && frame.func !== "{") {
				if (args > 255) problems.push(`passes ${args} arguments to ${frame.func} (Excel allows 255)`);
				const arity = ARITY[frame.func];
				if (arity && (args < arity[0] || args > arity[1])) problems.push(`passes ${args} argument(s) to ${frame.func}, which takes ${arity[0] === arity[1] ? arity[0] : `${arity[0]}-${arity[1]}`}`);
			}
			continue;
		}
		const top = stack[stack.length - 1];
		if (top) {
			if (t.kind === "op" && text === "," && top.func !== "{" && top.func !== "") top.args++;
			top.empty = false;
		}
	}
	if (stack.length > 0) problems.push("has an unclosed (");
	if (maxDepth > 64) problems.push(`nests ${maxDepth} levels deep (Excel allows 64)`);
	return problems;
}
