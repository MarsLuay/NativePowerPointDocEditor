import test from 'node:test';
import assert from 'node:assert';
import { build } from 'esbuild';
import { mkdtemp, mkdir, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const projectRoot = path.resolve(__dirname, '..');

test('DOCX text replacement fidelity with $ variables', async () => {
    const tempDir = await mkdtemp(path.join(tmpdir(), 'npde-test-'));
    const entryFile = path.join(tempDir, 'entry.ts');

    await writeFile(entryFile, `
        import { patchRunText } from '${path.join(projectRoot, 'src/ai/docxOoxmlWrite.ts').replace(/\\/g, '/')}';
        import { replaceDocumentBodyInner } from '${path.join(projectRoot, 'src/ai/docxOoxml.ts').replace(/\\/g, '/')}';

        export function testPatchRunText(paragraphXml: string, runIndex: number, text: string) {
            return patchRunText(paragraphXml, runIndex, text);
        }

        export function testReplaceDocumentBodyInner(docXml: string, newInner: string) {
            return replaceDocumentBodyInner(docXml, newInner);
        }
    `);

    const outDir = path.join(tempDir, 'out');
    await build({
        entryPoints: [entryFile],
        bundle: true,
        outdir: outDir,
        format: 'esm',
        platform: 'node',
        target: 'node20',
        external: ['obsidian'],
    });

    const { pathToFileURL } = await import('node:url');
    const mod = await import(pathToFileURL(path.join(outDir, 'entry.js')).href);
    const { testPatchRunText, testReplaceDocumentBodyInner } = mod;

	const paragraphXml = '<w:p><w:r><w:t>Hello</w:t></w:r></w:p>';

	const result1 = testPatchRunText(paragraphXml, 0, 'Cost is $100');
	assert.strictEqual(
		result1,
		'<w:p><w:r><w:t>Cost is $100</w:t></w:r></w:p>',
		'patchRunText should not corrupt XML when text contains $1'
	);

	const result2 = testPatchRunText(paragraphXml, 0, 'Cost is $&100');
	assert.strictEqual(
		result2,
		'<w:p><w:r><w:t>Cost is $&amp;100</w:t></w:r></w:p>',
		'patchRunText should not corrupt XML when text contains $&'
	);

	const docXml = '<w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body>';
	const newInner = '<w:p><w:r><w:t>World $1</w:t></w:r></w:p>';
	const resultBody = testReplaceDocumentBodyInner(docXml, newInner);
	assert.strictEqual(
		resultBody,
		'<w:body><w:p><w:r><w:t>World $1</w:t></w:r></w:p></w:body>',
		'replaceDocumentBodyInner should not corrupt XML when inner contains $1'
	);
});
