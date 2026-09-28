import { unzipSync, zipSync, type Zippable } from "fflate";

export type Parts = Map<string, Uint8Array>;

/** Decompresses every part of an .xlsx into its exact original bytes. */
export function readParts(input: Uint8Array): Parts {
	const files = unzipSync(input);
	const parts: Parts = new Map();
	for (const [name, data] of Object.entries(files)) {
		parts.set(name, data);
	}
	return parts;
}

/** Parts bigger than this (big rendered sheets) use fast deflate: ~4x quicker for ~15% more bytes. */
const FAST_DEFLATE_BYTES = 1 << 20;

/** Repacks parts into an .xlsx. Part bytes are written exactly as given. */
export function writeParts(parts: Parts): Uint8Array {
	const zippable: Zippable = {};
	for (const [name, data] of parts) {
		zippable[name] = data.length > FAST_DEFLATE_BYTES ? [data, { level: 1 }] : data;
	}
	return zipSync(zippable, { level: 6 });
}
