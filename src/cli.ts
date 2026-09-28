import { readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { Template } from "./template.js";
import { generateTypes } from "./typegen.js";
import { verify } from "./verify.js";

const USAGE = `excel-graft — fill Excel templates without breaking them

Usage:
  excel-graft inspect <template.xlsx> [--region Key=Sheet!A7:M207]...
                                                      What the template accepts (regions, shapes, fields)
  excel-graft types <template.xlsx> [--name T] [--out file.ts] [--region ...]
                                                      Generate a TypeScript type for render() data

  --region declares a region without editing the workbook (same as the
  regions option of Template.load). Quote it: --region "Data='My Sheet'!A7:M207"
  excel-graft verify <file.xlsx>...                   Structural checks for Excel's repair prompt

Attach the inspect and verify output to bug reports.`;

function typeNameFor(file: string): string {
	const base = path.basename(file, path.extname(file)).replace(/[^A-Za-z0-9]+(.)?/g, (_m, c: string | undefined) => (c ? c.toUpperCase() : ""));
	const name = base.charAt(0).toUpperCase() + base.slice(1);
	return `${/^[A-Za-z_]/.test(name) ? name : `T${name}`}Data`;
}

function regionFlags(args: string[]): Record<string, string> {
	const out: Record<string, string> = {};
	args.forEach((a, i) => {
		if (a !== "--region") return;
		const spec = args[i + 1] ?? "";
		const eq = spec.indexOf("=");
		if (eq > 0) out[spec.slice(0, eq)] = spec.slice(eq + 1);
	});
	return out;
}

function flag(args: string[], name: string): string | undefined {
	const i = args.indexOf(name);
	return i === -1 ? undefined : args[i + 1];
}

export function main(argv: string[]): number {
	const [cmd, ...args] = argv;
	const files = args.filter((a, i) => !a.startsWith("--") && !(i > 0 && args[i - 1]?.startsWith("--")));
	if (!cmd || cmd === "-h" || cmd === "--help" || files.length === 0) {
		console.log(USAGE);
		return cmd ? 0 : 1;
	}
	if (cmd === "inspect") {
		const tpl = Template.loadSync(readFileSync(files[0] as string), { regions: regionFlags(args) });
		console.log(JSON.stringify(tpl.inspect(), null, 2));
		return 0;
	}
	if (cmd === "types") {
		const file = files[0] as string;
		const tpl = Template.loadSync(readFileSync(file), { regions: regionFlags(args) });
		const ts = generateTypes(tpl.inspect(), flag(args, "--name") ?? typeNameFor(file));
		const out = flag(args, "--out");
		if (out) writeFileSync(out, ts);
		else process.stdout.write(ts);
		return 0;
	}
	if (cmd === "verify") {
		let bad = 0;
		for (const file of files) {
			const issues = verify(new Uint8Array(readFileSync(file)));
			if (issues.length === 0) console.log(`✓ ${file}`);
			else {
				bad++;
				console.log(`✗ ${file}`);
				for (const i of issues) console.log(`    ${i.part}: ${i.message}`);
			}
		}
		return bad > 0 ? 1 : 0;
	}
	console.error(`Unknown command "${cmd}".\n\n${USAGE}`);
	return 1;
}

process.exitCode = main(process.argv.slice(2));
