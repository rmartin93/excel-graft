import { SaxesParser } from "saxes";

/**
 * An element located by exact string offsets into the source XML. Nothing
 * is normalized: callers splice edits into the original text, so every
 * byte they don't touch survives as-is (namespaces, attribute order,
 * mc:Ignorable, whitespace).
 */
export interface XmlElement {
	name: string;
	attrs: Record<string, string>;
	/** Index of the opening `<`. */
	start: number;
	/** Index just after the start tag's `>`. */
	openEnd: number;
	/** Index of the closing `</` (equals `openEnd` when self-closing). */
	closeStart: number;
	/** Index just after the element's last `>`. */
	end: number;
	selfClosing: boolean;
	children: XmlElement[];
	parent: XmlElement | undefined;
}

/** Parses well-formed XML into an offset tree. Throws on malformed input. */
export function parseXml(xml: string): XmlElement {
	const parser = new SaxesParser({ xmlns: false, position: true });
	const root: XmlElement = {
		name: "#document",
		attrs: {},
		start: 0,
		openEnd: 0,
		closeStart: xml.length,
		end: xml.length,
		selfClosing: false,
		children: [],
		parent: undefined,
	};
	let current = root;
	let pendingStart = 0;

	parser.on("opentagstart", () => {
		pendingStart = xml.lastIndexOf("<", parser.position - 1);
	});
	parser.on("opentag", (tag) => {
		const el: XmlElement = {
			name: tag.name,
			attrs: tag.attributes as Record<string, string>,
			start: pendingStart,
			openEnd: parser.position,
			closeStart: parser.position,
			end: parser.position,
			selfClosing: tag.isSelfClosing,
			children: [],
			parent: current,
		};
		current.children.push(el);
		current = el;
	});
	parser.on("closetag", () => {
		const el = current;
		if (!el.selfClosing) {
			el.end = parser.position;
			el.closeStart = xml.lastIndexOf("</", el.end - 1);
		}
		current = el.parent ?? root;
	});
	parser.on("error", (err) => {
		throw err;
	});

	parser.write(xml).close();
	return root;
}

export function child(el: XmlElement, name: string): XmlElement | undefined {
	return el.children.find((c) => c.name === name);
}

export function children(el: XmlElement, name: string): XmlElement[] {
	return el.children.filter((c) => c.name === name);
}

/** Depth-first search for every element with this name. */
export function findAll(el: XmlElement, name: string, out: XmlElement[] = []): XmlElement[] {
	for (const c of el.children) {
		if (c.name === name) out.push(c);
		findAll(c, name, out);
	}
	return out;
}

/** Local name without prefix (`x14:sparkline` → `sparkline`). */
export function localName(name: string): string {
	const i = name.indexOf(":");
	return i === -1 ? name : name.slice(i + 1);
}

/** Decoded text content between the element's tags (no child markup expected). */
export function textOf(xml: string, el: XmlElement): string {
	return decodeXml(xml.slice(el.openEnd, el.closeStart));
}

export function rawInner(xml: string, el: XmlElement): string {
	return xml.slice(el.openEnd, el.closeStart);
}

export function rawOuter(xml: string, el: XmlElement): string {
	return xml.slice(el.start, el.end);
}

export function rawStartTag(xml: string, el: XmlElement): string {
	return xml.slice(el.start, el.openEnd);
}

const ENTITY_RE = /&(?:#x([0-9a-fA-F]+)|#(\d+)|(amp|lt|gt|quot|apos));/g;
const NAMED: Record<string, string> = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" };

export function decodeXml(s: string): string {
	if (!s.includes("&")) return s;
	return s.replace(ENTITY_RE, (_m, hex: string | undefined, dec: string | undefined, named: string | undefined) => {
		if (hex !== undefined) return String.fromCodePoint(parseInt(hex, 16));
		if (dec !== undefined) return String.fromCodePoint(parseInt(dec, 10));
		return NAMED[named as string] as string;
	});
}

export function escapeText(s: string): string {
	return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

export function escapeAttr(s: string): string {
	return escapeText(s).replace(/"/g, "&quot;").replace(/\r/g, "&#13;").replace(/\n/g, "&#10;").replace(/\t/g, "&#9;");
}

const ATTR_RE = /(\s+)([^\s=/>]+)(\s*=\s*)("[^"]*"|'[^']*')/g;

/**
 * Edits attributes in a raw start tag, leaving every other byte alone.
 * `null` removes the attribute; new attributes are appended at the end.
 */
export function setAttrs(startTag: string, updates: Record<string, string | null>): string {
	const pending = new Map(Object.entries(updates));
	let out = startTag.replace(ATTR_RE, (whole, ws: string, name: string, eq: string) => {
		if (!pending.has(name)) return whole;
		const value = pending.get(name);
		pending.delete(name);
		if (value === null || value === undefined) return "";
		return `${ws}${name}${eq}"${escapeAttr(value)}"`;
	});
	const additions = [...pending].filter(([, v]) => v !== null) as [string, string][];
	if (additions.length > 0) {
		const extra = additions.map(([n, v]) => ` ${n}="${escapeAttr(v)}"`).join("");
		const closeAt = out.endsWith("/>") ? out.length - 2 : out.length - 1;
		out = out.slice(0, closeAt) + extra + out.slice(closeAt);
	}
	return out;
}

/** Collects non-overlapping replacements against one source string and applies them in one pass. */
export class Splicer {
	private readonly edits: { start: number; end: number; text: string }[] = [];

	constructor(private readonly source: string) {}

	replace(start: number, end: number, text: string): void {
		this.edits.push({ start, end, text });
	}

	replaceElement(el: XmlElement, text: string): void {
		this.replace(el.start, el.end, text);
	}

	insert(at: number, text: string): void {
		this.edits.push({ start: at, end: at, text });
	}

	get changed(): boolean {
		return this.edits.length > 0;
	}

	apply(): string {
		if (this.edits.length === 0) return this.source;
		const sorted = [...this.edits].sort((a, b) => a.start - b.start || a.end - b.end);
		const parts: string[] = [];
		let pos = 0;
		for (const e of sorted) {
			if (e.start < pos) throw new Error(`overlapping XML edits at offset ${e.start}`);
			parts.push(this.source.slice(pos, e.start), e.text);
			pos = e.end;
		}
		parts.push(this.source.slice(pos));
		return parts.join("");
	}
}
