import { existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { expect } from "vitest";
import { readParts } from "../src/zip.js";

export const ROOT = path.join(__dirname, "..");
export const CORPUS = path.join(ROOT, "corpus", "templates");
export const GENERATED = path.join(ROOT, "corpus", "generated");

export function corpusFile(...segments: string[]): Uint8Array {
	return new Uint8Array(readFileSync(path.join(CORPUS, ...segments)));
}

/** Every `corpus/templates/<name>/*.xlsx` that is a template (not a captured broken output). */
export function corpusTemplates(): { name: string; file: string }[] {
	const out: { name: string; file: string }[] = [];
	for (const dir of readdirSync(CORPUS, { withFileTypes: true })) {
		if (!dir.isDirectory()) continue;
		for (const f of readdirSync(path.join(CORPUS, dir.name))) {
			if (!f.toLowerCase().endsWith(".xlsx") || f.startsWith("~$")) continue;
			if (/output/i.test(f)) continue;
			out.push({ name: `${dir.name}/${f}`, file: path.join(CORPUS, dir.name, f) });
		}
	}
	return out.sort((a, b) => a.name.localeCompare(b.name));
}

const decoder = new TextDecoder();

export function partText(xlsx: Uint8Array, name: string): string {
	const data = readParts(xlsx).get(name);
	if (!data) throw new Error(`no part ${name}`);
	return decoder.decode(data);
}

/**
 * CLAUDE.md invariant 2: every part the render had no business touching is
 * byte-identical. `mayChange` names the parts that are allowed to differ;
 * calcChain is the only part allowed to disappear, and nothing may appear.
 */
export function expectPassthrough(input: Uint8Array, output: Uint8Array, mayChange: (part: string) => boolean): string[] {
	const a = readParts(input);
	const b = readParts(output);
	const changed: string[] = [];
	for (const [name, bytes] of a) {
		const other = b.get(name);
		if (!other) {
			expect(name, `part "${name}" disappeared`).toMatch(/calcChain\.xml$/);
			continue;
		}
		if (Buffer.from(other).equals(Buffer.from(bytes))) continue;
		changed.push(name);
		expect(mayChange(name), `part "${name}" changed but the render had no reason to touch it`).toBe(true);
	}
	for (const name of b.keys()) expect(a.has(name), `unexpected new part "${name}"`).toBe(true);
	return changed;
}

/** Parts a render may legitimately rewrite. Everything else must pass through untouched. */
export function renderablePart(name: string): boolean {
	return (
		/^xl\/worksheets\/sheet\d+\.xml$/.test(name) ||
		/^xl\/tables\/table\d+\.xml$/.test(name) ||
		/^xl\/drawings\/drawing\d+\.xml$/.test(name) ||
		/^xl\/drawings\/vmlDrawing\d+\.vml$/.test(name) ||
		/^xl\/comments\d*\.xml$/.test(name) ||
		/^xl\/threadedComments\/threadedComment\d+\.xml$/.test(name) ||
		/^xl\/charts\/chart(Ex)?\d+\.xml$/.test(name) ||
		/^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/.test(name) ||
		name === "xl/workbook.xml" ||
		name === "xl/_rels/workbook.xml.rels" ||
		name === "[Content_Types].xml"
	);
}

export function writeGenerated(name: string, bytes: Uint8Array): string {
	if (!existsSync(GENERATED)) mkdirSync(GENERATED, { recursive: true });
	const file = path.join(GENERATED, name);
	mkdirSync(path.dirname(file), { recursive: true });
	writeFileSync(file, bytes);
	return file;
}

/** Values that have broken real spreadsheet writers. */
export const HOSTILE_STRINGS: string[] = [
	"",
	" leading and trailing spaces ",
	"emoji 🎉 👩‍👩‍👧 family and flags 🇺🇸",
	"control \u0001\u0002\u0008\u000b\u000c\u001f chars",
	"=1+1 looks like a formula",
	"+SUM(A1) also",
	"-1 dash first",
	"@mention",
	"_x0041_ literal escape",
	"<tag attr=\"x\">&amp; & < > '",
	"line one\nline two\r\nline three\ttab",
	"RTL עברית العربية",
	"combining é and zero​width",
	"lone surrogate \ud800 here",
	"￾￿ nonchars",
	"x".repeat(40_000),
	"'leading apostrophe",
	"#REF!",
	"TRUE",
	"1E+10",
	"00123 leading zeros",
];

export const HOSTILE_NUMBERS: unknown[] = [0, -0, 1, -1, 0.1 + 0.2, 1e308, -1e308, 5e-324, Number.MAX_SAFE_INTEGER, NaN, Infinity, -Infinity, 123456789012345678n, 42n];

export const HOSTILE_DATES: unknown[] = [
	new Date(Date.UTC(1900, 0, 1)),
	new Date(Date.UTC(1900, 1, 28)),
	new Date(Date.UTC(1900, 2, 1)),
	new Date(Date.UTC(1904, 0, 1)),
	new Date(Date.UTC(2026, 8, 28, 13, 45, 30)),
	new Date(Date.UTC(9999, 11, 31)),
	new Date(Date.UTC(1850, 0, 1)),
	new Date(Number.NaN),
];
