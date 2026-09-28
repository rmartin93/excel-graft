import { type CellRef, parseCell } from "./a1.js";

/**
 * A small Excel formula tokenizer, just precise enough to find every A1
 * reference with its exact span. Everything that isn't a reference is kept
 * as an opaque token so rewriting never touches it.
 */

export type RefShape = "cell" | "area" | "rows" | "cols";

export interface RefToken {
	kind: "ref";
	start: number;
	end: number;
	/** Raw prefix including the `!` (e.g. `'My Sheet'!`), or "" for none. */
	prefix: string;
	/** Decoded sheet name, if qualified. */
	sheet: string | undefined;
	/** True for `[1]Sheet!A1`-style references into another workbook, and 3-D refs. */
	foreign: boolean;
	shape: RefShape;
	/** For `rows`/`cols` shapes only `row`/`col` of a and b are meaningful. */
	a: CellRef;
	b: CellRef | undefined;
	/** `#` when this is a spill reference (`A1#`). */
	spill: string;
}

export interface OtherToken {
	kind: "str" | "struct" | "func" | "name" | "num" | "err" | "ws" | "op";
	start: number;
	end: number;
}

export type Token = RefToken | OtherToken;

const IDENT_CHAR = /[A-Za-z0-9_.\\?¡-￿]/;
const ERR_RE = /#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|GETTING_DATA|SPILL!|CALC!|FIELD!|BLOCKED!|CONNECT!|BUSY!|UNKNOWN!|PYTHON!)/iy;
const SHEET_RE = /([A-Za-z0-9_.¡-￿]+(?::[A-Za-z0-9_.¡-￿]+)?)!/y;
const AREA_RE = /(\$?[A-Za-z]{1,3}\$?\d{1,7})(?::(\$?[A-Za-z]{1,3}\$?\d{1,7}))?/y;
const ROWS_RE = /(\$?)(\d{1,7}):(\$?)(\d{1,7})/y;
const COLS_RE = /(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})/y;
const IDENT_RE = /[A-Za-z_\\¡-￿][A-Za-z0-9_.\\?¡-￿]*/y;
const NUM_RE = /(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/y;

function stickyMatch(re: RegExp, s: string, at: number): RegExpExecArray | null {
	re.lastIndex = at;
	return re.exec(s);
}

function identContinues(s: string, at: number): boolean {
	const ch = s[at];
	return ch !== undefined && (IDENT_CHAR.test(ch) || ch === "(" || ch === "[" || ch === "!");
}

function skipBrackets(s: string, at: number): number {
	let depth = 0;
	let i = at;
	while (i < s.length) {
		const ch = s[i];
		if (ch === "'") {
			i += 2;
			continue;
		}
		if (ch === "[") depth++;
		else if (ch === "]") {
			depth--;
			if (depth === 0) return i + 1;
		}
		i++;
	}
	return s.length;
}

interface RefBody {
	end: number;
	shape: RefShape;
	a: CellRef;
	b: CellRef | undefined;
	spill: string;
}

/** Tries to read a reference body (no sheet prefix) at `at`. */
function readRefBody(s: string, at: number): RefBody | undefined {
	const rows = stickyMatch(ROWS_RE, s, at);
	if (rows && !identContinues(s, at + rows[0].length)) {
		const r1 = Number(rows[2]);
		const r2 = Number(rows[4]);
		if (r1 >= 1 && r2 >= 1) {
			return {
				end: at + rows[0].length,
				shape: "rows",
				a: { col: 0, row: r1, colAbs: false, rowAbs: rows[1] === "$" },
				b: { col: 0, row: r2, colAbs: false, rowAbs: rows[3] === "$" },
				spill: "",
			};
		}
	}
	const area = stickyMatch(AREA_RE, s, at);
	if (area) {
		const end = at + area[0].length;
		const a = parseCell(area[1] as string);
		const b = area[2] === undefined ? undefined : parseCell(area[2]);
		const spill = s[end] === "#" && !stickyMatch(ERR_RE, s, end) ? "#" : "";
		if (a && (area[2] === undefined || b) && (spill !== "" || !identContinues(s, end))) {
			return { end: end + spill.length, shape: b ? "area" : "cell", a, b, spill };
		}
	}
	const cols = stickyMatch(COLS_RE, s, at);
	if (cols && !identContinues(s, at + cols[0].length)) {
		return {
			end: at + cols[0].length,
			shape: "cols",
			a: { col: 1, row: 0, colAbs: cols[1] === "$", rowAbs: false },
			b: { col: 1, row: 0, colAbs: cols[3] === "$", rowAbs: false },
			spill: "",
		};
	}
	return undefined;
}

function decodeSheet(raw: string): { name: string; foreign: boolean } {
	let name = raw;
	if (name.startsWith("'")) name = name.slice(1, -1).replace(/''/g, "'");
	const foreign = name.startsWith("[") || name.includes(":");
	return { name, foreign };
}

export function tokenize(s: string): Token[] {
	const tokens: Token[] = [];
	let i = s.startsWith("=") ? 1 : 0;
	if (i === 1) tokens.push({ kind: "op", start: 0, end: 1 });

	const pushRef = (start: number, prefix: string, body: RefBody): void => {
		const { name, foreign } = prefix === "" ? { name: undefined, foreign: false } : decodeSheet(prefix.slice(0, -1));
		tokens.push({
			kind: "ref",
			start,
			end: body.end,
			prefix,
			sheet: name,
			foreign,
			shape: body.shape,
			a: body.a,
			b: body.b,
			spill: body.spill,
		});
	};

	while (i < s.length) {
		const ch = s[i] as string;
		const start = i;

		if (ch === '"') {
			i++;
			while (i < s.length) {
				if (s[i] === '"') {
					if (s[i + 1] === '"') i += 2;
					else {
						i++;
						break;
					}
				} else i++;
			}
			tokens.push({ kind: "str", start, end: i });
			continue;
		}

		if (ch === "'") {
			let j = i + 1;
			while (j < s.length) {
				if (s[j] === "'") {
					if (s[j + 1] === "'") j += 2;
					else break;
				} else j++;
			}
			const prefixEnd = j + 1;
			if (s[prefixEnd] === "!") {
				const prefix = s.slice(i, prefixEnd + 1);
				const body = readRefBody(s, prefixEnd + 1);
				if (body) {
					pushRef(start, prefix, body);
					i = body.end;
				} else {
					i = prefixEnd + 1;
					tokens.push({ kind: "name", start, end: i });
				}
				continue;
			}
			i = prefixEnd;
			tokens.push({ kind: "op", start, end: i });
			continue;
		}

		if (ch === "#") {
			const err = stickyMatch(ERR_RE, s, i);
			i += err ? err[0].length : 1;
			tokens.push({ kind: err ? "err" : "op", start, end: i });
			continue;
		}

		if (ch === "[") {
			const close = skipBrackets(s, i);
			const sheet = stickyMatch(SHEET_RE, s, close);
			if (sheet) {
				const prefix = s.slice(i, close + sheet[0].length);
				const body = readRefBody(s, close + sheet[0].length);
				if (body) {
					pushRef(start, prefix, body);
					i = body.end;
					continue;
				}
			}
			i = close;
			tokens.push({ kind: "struct", start, end: i });
			continue;
		}

		if (/\s/.test(ch)) {
			while (i < s.length && /\s/.test(s[i] as string)) i++;
			tokens.push({ kind: "ws", start, end: i });
			continue;
		}

		if (ch === "$" || /[A-Za-z0-9_\\¡-￿.]/.test(ch)) {
			// Sheet-qualified reference?
			const sheet = stickyMatch(SHEET_RE, s, i);
			if (sheet) {
				const bodyAt = i + sheet[0].length;
				const body = readRefBody(s, bodyAt);
				if (body) {
					pushRef(start, sheet[0], body);
					i = body.end;
					continue;
				}
				// Sheet-scoped name like Sheet1!MyName, or Sheet1!#REF!
				const ident = stickyMatch(IDENT_RE, s, bodyAt) ?? stickyMatch(ERR_RE, s, bodyAt);
				i = bodyAt + (ident ? ident[0].length : 0);
				tokens.push({ kind: "name", start, end: i });
				continue;
			}
			const body = readRefBody(s, i);
			if (body) {
				pushRef(start, "", body);
				i = body.end;
				continue;
			}
			const num = stickyMatch(NUM_RE, s, i);
			if (num && /[0-9.]/.test(ch)) {
				i += num[0].length;
				tokens.push({ kind: "num", start, end: i });
				continue;
			}
			const ident = stickyMatch(IDENT_RE, s, i);
			if (ident) {
				i += ident[0].length;
				if (s[i] === "[") {
					i = skipBrackets(s, i);
					tokens.push({ kind: "struct", start, end: i });
				} else {
					tokens.push({ kind: s[i] === "(" ? "func" : "name", start, end: i });
				}
				continue;
			}
		}

		const two = s.slice(i, i + 2);
		i += two === "<=" || two === ">=" || two === "<>" ? 2 : 1;
		tokens.push({ kind: "op", start, end: i });
	}
	return tokens;
}

export interface RefContext {
	/** The reference is a whole argument of a function call (`SUM(A1, B2)`), so it may expand into several. */
	inArgList: boolean;
}

/**
 * A replacement for one reference. `text` replaces the token. When
 * `dedupeKey` is set, later references in the same argument list with the
 * same key are dropped (with their comma): `SUM(C5,C8)` over two sample
 * groups expands once, not twice.
 */
export interface RefReplacement {
	text: string;
	dedupeKey?: string;
}

export type RefMapper = (ref: RefToken, ctx: RefContext) => RefReplacement | string | undefined;

/** Rewrites references in a formula. Returns the original string when nothing changed. */
export function rewriteRefs(formula: string, map: RefMapper): string {
	const tokens = tokenize(formula);
	if (!tokens.some((t) => t.kind === "ref")) return formula;

	// Paren stack: true when the paren opens a function call.
	const callIds: (number | undefined)[] = [];
	let nextCallId = 0;
	const tokenCall: (number | undefined)[] = new Array(tokens.length);
	const argSep: boolean[] = new Array(tokens.length).fill(false);
	for (let k = 0; k < tokens.length; k++) {
		const t = tokens[k] as Token;
		const text = formula.slice(t.start, t.end);
		if (t.kind === "func") continue;
		if (text === "(" || text === "{") {
			const prev = previousSignificant(tokens, k);
			callIds.push(text === "(" && prev?.kind === "func" ? nextCallId++ : undefined);
			continue;
		}
		if (text === ")" || text === "}") {
			tokenCall[k] = callIds[callIds.length - 1];
			callIds.pop();
			continue;
		}
		tokenCall[k] = callIds[callIds.length - 1];
		if (text === "," && tokenCall[k] !== undefined) argSep[k] = true;
	}

	const out: string[] = [];
	let pos = 0;
	const seen = new Map<number, Set<string>>();
	let changed = false;

	for (let k = 0; k < tokens.length; k++) {
		const t = tokens[k] as Token;
		if (t.kind !== "ref") continue;
		const prevIdx = previousSignificantIndex(tokens, k);
		const nextIdx = nextSignificantIndex(tokens, k);
		const call = tokenCall[k];
		const prevOk =
			prevIdx !== undefined &&
			(argSep[prevIdx] || (formula.slice(tokens[prevIdx]!.start, tokens[prevIdx]!.end) === "(" && call !== undefined));
		const nextText = nextIdx === undefined ? "" : formula.slice(tokens[nextIdx]!.start, tokens[nextIdx]!.end);
		const nextOk = nextIdx !== undefined && ((nextText === "," && argSep[nextIdx]) || (nextText === ")" && tokenCall[nextIdx] === call));
		const inArgList = call !== undefined && prevOk && nextOk;

		const result = map(t, { inArgList });
		if (result === undefined) continue;
		const rep = typeof result === "string" ? { text: result } : result;

		if (rep.dedupeKey !== undefined && inArgList && call !== undefined) {
			let keys = seen.get(call);
			if (!keys) seen.set(call, (keys = new Set()));
			if (keys.has(rep.dedupeKey) && prevIdx !== undefined && argSep[prevIdx]) {
				// Drop ", ref" entirely.
				out.push(formula.slice(pos, tokens[prevIdx]!.start));
				pos = t.end;
				changed = true;
				continue;
			}
			keys.add(rep.dedupeKey);
		}
		if (rep.text !== formula.slice(t.start, t.end)) changed = true;
		out.push(formula.slice(pos, t.start), rep.text);
		pos = t.end;
	}
	if (!changed) return formula;
	out.push(formula.slice(pos));
	return out.join("");
}

function previousSignificantIndex(tokens: Token[], k: number): number | undefined {
	for (let j = k - 1; j >= 0; j--) if (tokens[j]!.kind !== "ws") return j;
	return undefined;
}

function nextSignificantIndex(tokens: Token[], k: number): number | undefined {
	for (let j = k + 1; j < tokens.length; j++) if (tokens[j]!.kind !== "ws") return j;
	return undefined;
}

function previousSignificant(tokens: Token[], k: number): Token | undefined {
	const j = previousSignificantIndex(tokens, k);
	return j === undefined ? undefined : tokens[j];
}

/** Quotes a sheet name for use as a reference prefix when needed. */
export function sheetPrefix(name: string): string {
	return /^[A-Za-z_¡-￿][A-Za-z0-9_.¡-￿]*$/.test(name) && !parseCell(name) && !/^R\d*C\d*$/i.test(name)
		? `${name}!`
		: `'${name.replace(/'/g, "''")}'!`;
}
