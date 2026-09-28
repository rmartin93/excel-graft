import { readParts, writeParts, type Parts } from "./zip.js";

export interface InspectResult {
	placeholders: string[];
	tables: string[];
}

export interface RenderOptions {
	onMissing?: "error" | "blank";
}

const SHEET_PART_RE = /^xl\/worksheets\/sheet\d+\.xml$/;
const TABLE_PART_RE = /^xl\/tables\/table\d+\.xml$/;
const PLACEHOLDER_RE = /\{\{\s*([\w.]+)\s*\}\}/g;
const TABLE_NAME_RE = /<table\b[^>]*\bname="([^"]+)"/;

const utf8Decoder = new TextDecoder("utf-8");

/**
 * Loads an .xlsx as a bag of exact-byte parts. Nothing is parsed into an
 * object model; parts this tool doesn't touch stay byte-identical on
 * output. See CLAUDE.md for why.
 */
export class Template {
	private constructor(private readonly parts: Parts) {}

	static async load(input: Uint8Array | ArrayBuffer): Promise<Template> {
		const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
		return new Template(readParts(bytes));
	}

	/**
	 * Repacks the template with no edits applied. Used to test the
	 * passthrough invariant against a template that received zero changes.
	 */
	async toBuffer(): Promise<Uint8Array> {
		return writeParts(this.parts);
	}

	/** Lists the `{{placeholders}}` and named Tables found in the sheets, for validating a template before render() exists. */
	inspect(): InspectResult {
		const placeholders = new Set<string>();
		const tables = new Set<string>();

		for (const [name, data] of this.parts) {
			if (SHEET_PART_RE.test(name)) {
				const text = utf8Decoder.decode(data);
				for (const match of text.matchAll(PLACEHOLDER_RE)) {
					const placeholder = match[1];
					if (placeholder !== undefined) placeholders.add(placeholder);
				}
			}
			if (TABLE_PART_RE.test(name)) {
				const text = utf8Decoder.decode(data);
				const nameMatch = text.match(TABLE_NAME_RE);
				if (nameMatch?.[1] !== undefined) tables.add(nameMatch[1]);
			}
		}

		return {
			placeholders: [...placeholders].sort(),
			tables: [...tables].sort(),
		};
	}

	async render(_data: Record<string, unknown>, _options: RenderOptions = {}): Promise<Uint8Array> {
		throw new Error(
			"Template.render() is not implemented yet. Phase 1 (scalar placeholders) " +
				"has not started — see PROJECT.md for the phase plan.",
		);
	}
}
