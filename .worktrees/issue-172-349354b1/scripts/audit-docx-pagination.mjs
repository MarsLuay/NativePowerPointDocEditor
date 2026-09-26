import { execFile as execFileCallback } from 'node:child_process';
import { cp, mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { promisify } from 'node:util';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { build } from 'esbuild';
import {
	classifyPaginationDelta,
	formatPaginationAuditStatus,
} from './lib/docx-pagination-audit.mjs';

const execFile = promisify(execFileCallback);
const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const vaultRoot = path.resolve(projectRoot, '..', '..');
const outputDir = path.join(projectRoot, 'results', 'docx-pagination-audit');
const chromeBinary = process.env.CHROME_PATH
	|| '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';
const sofficeBinary = process.env.SOFFICE_PATH || '/opt/homebrew/bin/soffice';
const defaultDocuments = [
	'School/Current Classes/ENG& 235/Resources/P2 One-Pagers/Ledger Leader One-Pager.docx',
	'School/Current Classes/ENG& 235/Resources/P2 One-Pagers/Sample One-Pager Truth Builders.docx',
	'School/Current Classes/ENG& 235/Resources/P1 Quick Start Guide/P1.1 Guided Self-Analysis.docx',
	'School/Current Classes/ENG& 235/Resources/Writing Guides/Writing Clearly and Concisely Guidelines.docx',
	'Life/Financials/Marwan Luay Resume.docx',
	'.Projects/Native PowerPoint Doc Editor/test_files/demo.docx',
];

async function bundleHarness() {
	await build({
		entryPoints: [path.join(projectRoot, 'src', 'docx-page-count-entry.tsx')],
		bundle: true,
		format: 'iife',
		logLevel: 'silent',
		outfile: path.join(outputDir, 'harness.js'),
		platform: 'browser',
		target: 'es2020',
		jsx: 'automatic',
		loader: { '.css': 'text' },
	});
}

async function getPreviewPageCount(docxPath, index) {
	const base64 = (await readFile(docxPath)).toString('base64');
	const htmlPath = path.join(outputDir, `document-${index}.html`);
	await writeFile(htmlPath, `<!doctype html>
<html lang="en">
<head><meta charset="utf-8"><title>DOCX pagination audit</title></head>
<body>
<div id="root"></div>
<script>window.__DOCX_BASE64__ = ${JSON.stringify(base64)};</script>
<script src="./harness.js"></script>
</body>
</html>`);
	const { stdout, stderr } = await execFile(chromeBinary, [
		'--headless=new',
		'--disable-gpu',
		'--no-sandbox',
		'--enable-logging=stderr',
		'--virtual-time-budget=20000',
		pathToFileURL(htmlPath).href,
	], { maxBuffer: 20 * 1024 * 1024, timeout: 60000 });
	const output = `${stdout}\n${stderr}`;
	const matches = output.match(/DOCX_PAGE_COUNT_RESULT:(\{.*\})/g);
	if (!matches?.length) {
		throw new Error(`No preview page count for ${docxPath}\n${output.slice(-1200)}`);
	}
	return JSON.parse(matches[matches.length - 1].replace('DOCX_PAGE_COUNT_RESULT:', ''));
}

async function getReferencePageCount(docxPath, index) {
	const tempDir = path.join(os.tmpdir(), `docx-pagination-${process.pid}-${index}`);
	const profileDir = path.join(tempDir, 'libreoffice-profile');
	await mkdir(tempDir, { recursive: true });
	await mkdir(profileDir, { recursive: true });
	const inputPath = path.join(tempDir, `document-${index}.docx`);
	const pdfPath = path.join(tempDir, `document-${index}.pdf`);
	await cp(docxPath, inputPath);
	try {
		await execFile(sofficeBinary, [
			`-env:UserInstallation=${pathToFileURL(profileDir).href}`,
			'--headless',
			'--convert-to',
			'pdf',
			'--outdir',
			tempDir,
			inputPath,
		], { timeout: 60000 });
		const { stdout } = await execFile('pdfinfo', [pdfPath], { timeout: 30000 });
		const match = stdout.match(/^Pages:\s+(\d+)/m);
		if (!match) {
			throw new Error(`pdfinfo did not report pages for ${docxPath}`);
		}
		return Number.parseInt(match[1], 10);
	} finally {
		await rm(tempDir, { recursive: true, force: true });
	}
}

async function getSourceLayoutDiagnostics(docxPath) {
	const { stdout: documentXml } = await execFile('unzip', [
		'-p',
		docxPath,
		'word/document.xml',
	], {
		encoding: 'utf8',
		maxBuffer: 20 * 1024 * 1024,
		timeout: 30000,
	});
	const paragraphs = documentXml.match(/<w:p\b[\s\S]*?<\/w:p>/g) ?? [];
	let tabRuns = 0;
	let tabbedParagraphs = 0;
	let tabHeavyParagraphs = 0;
	let maxTabsInParagraph = 0;
	let longSpaceRuns = 0;

	for (const paragraph of paragraphs) {
		const paragraphTabRuns = (paragraph.match(/<w:tab\s*\/>/g) ?? []).length;
		tabRuns += paragraphTabRuns;
		if (paragraphTabRuns > 0) {
			tabbedParagraphs += 1;
		}
		if (paragraphTabRuns >= 2) {
			tabHeavyParagraphs += 1;
		}
		maxTabsInParagraph = Math.max(maxTabsInParagraph, paragraphTabRuns);

		for (const textMatch of paragraph.matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g)) {
			longSpaceRuns += (textMatch[1].match(/ {8,}/g) ?? []).length;
		}
	}

	return {
		tabRuns,
		tabbedParagraphs,
		tabHeavyParagraphs,
		maxTabsInParagraph,
		longSpaceRuns,
	};
}

async function main() {
	const requested = process.argv.slice(2);
	const documents = requested.length > 0 ? requested : defaultDocuments;
	await mkdir(outputDir, { recursive: true });
	await bundleHarness();
	const results = [];

	for (let index = 0; index < documents.length; index += 1) {
		const relativePath = documents[index];
		const absolutePath = path.isAbsolute(relativePath)
			? relativePath
			: path.join(vaultRoot, relativePath);
		const [referencePages, preview, sourceLayout] = await Promise.all([
			getReferencePageCount(absolutePath, index),
			getPreviewPageCount(absolutePath, index),
			getSourceLayoutDiagnostics(absolutePath),
		]);
		const delta = preview.totalPages - referencePages;
		const classification = classifyPaginationDelta(delta);
		const result = {
			file: path.relative(vaultRoot, absolutePath),
			referenceRenderer: 'LibreOffice PDF',
			referencePages,
			previewPages: preview.totalPages,
			renderedPages: preview.renderedPages,
			delta,
			classification,
			sourceLayout,
		};
		results.push(result);
		const sourceNote = sourceLayout.tabHeavyParagraphs > 0
			? ` (${sourceLayout.tabHeavyParagraphs} tab-heavy paragraphs, max ${sourceLayout.maxTabsInParagraph} tabs)`
			: '';
		console.log(
			`${formatPaginationAuditStatus(classification)} LibreOffice ${result.referencePages} -> preview ${result.previewPages} ${result.file}${sourceNote}`,
		);
	}

	await writeFile(path.join(outputDir, 'results.json'), `${JSON.stringify(results, null, 2)}\n`);
	const overPaginated = results.filter((result) => result.delta > 0);
	if (overPaginated.length > 0) {
		throw new Error(`${overPaginated.length} document(s) use more preview pages than the reference renderer`);
	}
}

await main();
