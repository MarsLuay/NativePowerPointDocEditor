import { readFileSync } from 'node:fs';
import { mkdir, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createJiti } from 'jiti';

const stripVolatile = (manifest) => ({
	...manifest,
	generatedAt: undefined,
	operations: manifest.operations.map(({ generatedAt: _ignored, ...operation }) => operation),
});

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const jiti = createJiti(import.meta.url, { interopDefault: true });
const { buildCapabilityManifest } = jiti('../src/ai/capabilities.ts');
const { OP_IDS } = jiti('../src/ai/opCatalog.ts');
const { OP_EXAMPLES } = jiti('../src/ai/opExamples.ts');

const manifestJson = JSON.parse(readFileSync(path.join(projectRoot, 'manifest.json'), 'utf8'));
const expected = buildCapabilityManifest({
	pluginVersion: manifestJson.version,
	enabled: true,
});

const outputDir = path.join(projectRoot, 'ai');
const outputPath = path.join(outputDir, 'capabilities.json');
const mode = process.argv[2] ?? '--write';

if (mode === '--check') {
	const actual = JSON.parse(readFileSync(outputPath, 'utf8'));
	const expectedStripped = stripVolatile(expected);
	const actualStripped = stripVolatile(actual);
	if (JSON.stringify(expectedStripped) !== JSON.stringify(actualStripped)) {
		console.error('ai/capabilities.json is out of date. Run npm run ai:generate.');
		process.exit(1);
	}

	for (const opId of OP_IDS) {
		if (!OP_EXAMPLES[opId]) {
			console.error(`Missing OP_EXAMPLES entry for ${opId}.`);
			process.exit(1);
		}
	}

	console.log(`Verified ai/capabilities.json (${OP_IDS.length} operations, schema v${expected.schemaVersion}).`);
} else {
	await mkdir(outputDir, { recursive: true });
	try {
		const actual = JSON.parse(readFileSync(outputPath, 'utf8'));
		if (
			typeof actual.generatedAt === 'string'
			&& JSON.stringify(stripVolatile(actual)) === JSON.stringify(stripVolatile(expected))
		) {
			expected.generatedAt = actual.generatedAt;
		}
	} catch {
		// First generation or an invalid stale artifact: write a fresh timestamp.
	}
	await writeFile(outputPath, `${JSON.stringify(expected, null, 2)}\n`, 'utf8');
	console.log(`Wrote ${outputPath} (${expected.operations.length} operations, schema v${expected.schemaVersion}).`);
}
