import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { verify } from "../src/index.js";
import { ROOT } from "./helpers.js";

/**
 * corpus/known-bad/ holds hand-corrupted workbooks (see its README for how
 * Excel reacts to each). verify() must flag every one of them — including
 * the ones this machine's Excel quietly tolerates, because another build
 * of Excel will show the repair prompt for them.
 */
const DIR = path.join(ROOT, "corpus", "known-bad");
const files = readdirSync(DIR).filter((f) => f.endsWith(".xlsx"));

describe("verify() flags every known-bad fixture", () => {
	it.each(files.filter((f) => !f.startsWith("control")))("%s", (f) => {
		expect(verify(new Uint8Array(readFileSync(path.join(DIR, f))))).not.toEqual([]);
	});

	it.each(files.filter((f) => f.startsWith("control")))("%s (clean control) passes", (f) => {
		expect(verify(new Uint8Array(readFileSync(path.join(DIR, f))))).toEqual([]);
	});
});
