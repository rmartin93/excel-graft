// Realistic fake data shaped by a template's inferred schema, so any workbook
// (including one dropped into the page) can be exported without writing a query.

function prng(seed) {
	let a = seed >>> 0;
	return () => {
		a = (a + 0x6d2b79f5) >>> 0;
		let t = a;
		t = Math.imul(t ^ (t >>> 15), t | 1);
		t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
		return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
	};
}

const FIRST = ["Avery", "Jordan", "Riley", "Sam", "Taylor", "Morgan", "Casey", "Jamie", "Quinn", "Rowan", "Harper", "Emerson", "Skyler", "Reese", "Dakota"];
const LAST = ["Nguyen", "Garcia", "Patel", "Okafor", "Kim", "Martinez", "Schmidt", "Rossi", "Haddad", "Novak", "Silva", "Cohen", "Tanaka", "Adeyemi"];
const MERCHANTS = ["HEB", "Costco", "Target", "Shell", "Home Depot", "Whole Foods", "Walgreens", "Best Buy", "Chevron", "Trader Joe's", "IKEA", "Lowe's", "REI", "Apple Store"];
const CATEGORIES = ["Groceries", "Fuel", "Travel", "Software", "Equipment", "Meals", "Utilities", "Office Supplies", "Training", "Facilities", "Insurance", "Marketing"];
const POOLS = ["Fringe", "Overhead", "G&A", "Material Handling", "Facilities Capital"];
const PROJECTS = ["Apollo", "Borealis", "Cygnus", "Delta-7", "Everest", "Falcon", "Gemini", "Horizon"];
const LABOR = ["Engineer III", "Analyst II", "Program Manager", "Technician I", "Architect", "Scientist II", "Admin Specialist"];
const REGIONS = ["North", "South", "East", "West", "Central"];
const STATUSES = ["Open", "Paid", "Void"];
const WORDS = ["alpha", "bravo", "charlie", "delta", "echo", "foxtrot", "golf", "hotel", "india", "juliet", "kilo", "lima"];

const pick = (rand, list) => list[Math.floor(rand() * list.length)];
const money = (rand, lo, hi) => Math.round((lo + rand() * (hi - lo)) * 100) / 100;

function byName(name, type, rand, i) {
	const n = name.toLowerCase();
	if (type === "date" || /date|month|day|period|when|invoice ?date/.test(n)) {
		return new Date(Date.UTC(2026, Math.floor(rand() * 12), 1 + Math.floor(rand() * 28)));
	}
	if (type === "boolean" || /^(is|has)[A-Z_ ]|active|approved|billable/.test(name)) return rand() < 0.5;
	if (/rate|pct|percent/.test(n)) return type === "string" ? `${money(rand, 1, 60)}%` : money(rand, 0.05, 1.5);
	if (/hour|hrs/.test(n)) return Math.round(rand() * 160 * 4) / 4;
	if (/qty|quantity|units|count/.test(n)) return 1 + Math.floor(rand() * 250);
	if (/amount|cost|price|revenue|total|salary|pay|fee|budget|actual|expense|value|balance/.test(n)) {
		return type === "string" ? `$${money(rand, 5, 5000)}` : money(rand, 5, 25000);
	}
	if (type === "number") return money(rand, 1, 1000);
	if (/employee|person|name|rep|owner|preparer|prepared ?by|contact/.test(n)) return `${pick(rand, FIRST)} ${pick(rand, LAST)}`;
	if (/merchant|vendor|supplier|payee|customer|client/.test(n)) return pick(rand, MERCHANTS);
	if (/category|pool|group|type|class/.test(n)) return pick(rand, CATEGORIES);
	if (/project|contract|program/.test(n)) return pick(rand, PROJECTS);
	if (/labor|title|role|position/.test(n)) return pick(rand, LABOR);
	if (/region|territory|state/.test(n)) return pick(rand, REGIONS);
	if (/status/.test(n)) return pick(rand, STATUSES);
	if (/account|acct|gl|sku|code|id\b|number|no\b/.test(n)) return `${1000 + Math.floor(rand() * 9000)}-${String(i).padStart(4, "0")}`;
	if (/desc|note|memo|comment|detail/.test(n)) return `${pick(rand, WORDS)} ${pick(rand, WORDS)} ${pick(rand, WORDS)}`;
	return `${name} ${i + 1}`;
}

function groupLabel(field, rand, g, depth) {
	const n = field.name.toLowerCase();
	// Reuse the template's own sample group labels first (Fringe, Overhead, G&A...), so references to them resolve.
	if (field.samples?.length) {
		const base = field.samples[g % field.samples.length];
		return g < field.samples.length ? base : `${base} ${Math.floor(g / field.samples.length) + 1}`;
	}
	if (field.type === "date") return byName(field.name, "date", rand, g);
	if (field.type === "number") return g + 1;
	if (/pool/.test(n) || depth > 0) return POOLS[g % POOLS.length] + (g >= POOLS.length ? ` ${Math.floor(g / POOLS.length) + 1}` : "");
	if (/region/.test(n)) return REGIONS[g % REGIONS.length];
	return CATEGORIES[g % CATEGORIES.length] + (g >= CATEGORIES.length ? ` ${Math.floor(g / CATEGORIES.length) + 1}` : "");
}

function fillShape(shape, total, rand, depth, counter) {
	if (shape.kind === "rows") {
		return Array.from({ length: total }, () => {
			const i = counter.n++;
			const rec = {};
			for (const f of shape.fields) rec[f.name] = byName(f.name, f.type, rand, i);
			return rec;
		});
	}
	// Spread `total` detail rows over a sensible number of groups.
	const groups = Math.max(1, Math.min(2000, Math.round(Math.sqrt(total) / (depth + 1)) || 1));
	const out = [];
	let remaining = total;
	for (let g = 0; g < groups; g++) {
		const share = g === groups - 1 ? remaining : Math.max(0, Math.round((total / groups) * (0.5 + rand())));
		const n = Math.min(remaining, share);
		remaining -= n;
		const rec = {};
		for (const f of shape.fields) rec[f.name] = groupLabel(f, rand, g, depth);
		rec[shape.childKey] = fillShape(shape.child, n, rand, depth + 1, counter);
		out.push(rec);
	}
	return out;
}

/** Data for every region and scalar in the schema, with roughly `rows` detail rows per region. */
export function fakeData(schema, rows, seed = 42) {
	const rand = prng(seed);
	const data = {};
	for (const region of schema.regions) data[region.key] = fillShape(region.shape, rows, rand, 0, { n: 0 });
	for (const s of schema.scalars) {
		const [root, ...rest] = s.key.split(".");
		const leaf = rest.length ? rest[rest.length - 1] : root;
		const value = /date|as ?of|period|year ?end|fye|generated/i.test(leaf)
			? new Date(Date.UTC(2026, 8, 28))
			: /name|contractor|customer|company|client/i.test(leaf)
				? "Meraki Digital, LLC"
				: /by|preparer|owner/i.test(leaf)
					? `${pick(rand, FIRST)} ${pick(rand, LAST)}`
					: /title|report/i.test(leaf)
						? "FY2026 Report"
						: `${leaf} value`;
		if (rest.length === 0) data[root] = value;
		else {
			let cur = (data[root] ??= {});
			for (const seg of rest.slice(0, -1)) cur = cur[seg] ??= {};
			cur[leaf] = value;
		}
	}
	return data;
}
