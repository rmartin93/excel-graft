import { defineConfig } from "tsup";

export default defineConfig([
	{
		entry: ["src/index.ts"],
		format: ["esm", "cjs"],
		dts: true,
		sourcemap: true,
		clean: true,
		target: "node18",
	},
	{
		entry: { cli: "src/cli.ts" },
		format: ["esm"],
		sourcemap: false,
		clean: false,
		target: "node18",
		banner: { js: "#!/usr/bin/env node" },
	},
]);
