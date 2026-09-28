export const MAX_ROW = 1_048_576;
export const MAX_COL = 16_384;

export interface CellRef {
	col: number;
	row: number;
	colAbs: boolean;
	rowAbs: boolean;
}

/** A rectangular area, 1-based and inclusive. */
export interface Area {
	c1: number;
	r1: number;
	c2: number;
	r2: number;
}

export function colToIndex(letters: string): number {
	let n = 0;
	for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
	return n;
}

export function indexToCol(n: number): string {
	let s = "";
	while (n > 0) {
		const rem = (n - 1) % 26;
		s = String.fromCharCode(65 + rem) + s;
		n = Math.floor((n - 1) / 26);
	}
	return s;
}

const CELL_RE = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})$/;

export function parseCell(text: string): CellRef | undefined {
	const m = CELL_RE.exec(text);
	if (!m) return undefined;
	const col = colToIndex(m[2] as string);
	const row = Number(m[4]);
	if (col < 1 || col > MAX_COL || row < 1 || row > MAX_ROW) return undefined;
	return { colAbs: m[1] === "$", col, rowAbs: m[3] === "$", row };
}

export function formatCell(ref: CellRef): string {
	return `${ref.colAbs ? "$" : ""}${indexToCol(ref.col)}${ref.rowAbs ? "$" : ""}${ref.row}`;
}

export function cellName(col: number, row: number): string {
	return `${indexToCol(col)}${row}`;
}

/** Parses a plain `A1` or `A1:B2` (no sheet, `$` allowed) into an area. */
export function parseArea(text: string): Area | undefined {
	const [a, b] = text.split(":");
	if (a === undefined) return undefined;
	const p = parseCell(a);
	if (!p) return undefined;
	const q = b === undefined ? p : parseCell(b);
	if (!q) return undefined;
	return {
		c1: Math.min(p.col, q.col),
		r1: Math.min(p.row, q.row),
		c2: Math.max(p.col, q.col),
		r2: Math.max(p.row, q.row),
	};
}

export function formatArea(a: Area, abs = false): string {
	const d = abs ? "$" : "";
	const first = `${d}${indexToCol(a.c1)}${d}${a.r1}`;
	if (a.c1 === a.c2 && a.r1 === a.r2) return first;
	return `${first}:${d}${indexToCol(a.c2)}${d}${a.r2}`;
}

/** Parses a space-separated sqref list. Unparseable entries are returned as undefined. */
export function parseSqref(text: string): (Area | undefined)[] {
	return text
		.trim()
		.split(/\s+/)
		.filter((s) => s.length > 0)
		.map(parseArea);
}

export function formatSqref(areas: Area[]): string {
	return areas.map((a) => formatArea(a)).join(" ");
}

/** Groups a sorted-or-not set of row numbers into inclusive [start, end] runs. */
export function rowRuns(rows: Iterable<number>): [number, number][] {
	const sorted = [...new Set(rows)].sort((a, b) => a - b);
	const runs: [number, number][] = [];
	for (const r of sorted) {
		const last = runs[runs.length - 1];
		if (last && r === last[1] + 1) last[1] = r;
		else runs.push([r, r]);
	}
	return runs;
}
