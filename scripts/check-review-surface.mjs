#!/usr/bin/env node
/**
 * Validate the public, reviewer-facing DOCX runtime surface without needing
 * the editable docx-editor source worktree.
 */
import { existsSync } from 'node:fs';
import { readFile, readdir } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import {
	createDocxEditorAliases,
	docxEditorPackages,
	resolveDocxEditorPackagesRoot,
} from './lib/docx-editor-aliases.mjs';

const scriptPath = fileURLToPath(import.meta.url);
const defaultProjectRoot = path.resolve(path.dirname(scriptPath), '..');
const BRIDGE_RELATIVE_PATH = 'src/docx/runtime/bridge.mjs';
const STYLES_RELATIVE_PATH = 'src/docx/runtime/styles.ts';
const FACADE_FILES = [
	'src/docx/runtime/contract.ts',
	BRIDGE_RELATIVE_PATH,
	'src/docx/runtime/bridge.d.mts',
	'src/docx/runtime/index.ts',
	STYLES_RELATIVE_PATH,
];
const REACT_ALIAS_KEYS = [
	'react',
	'react/jsx-runtime',
	'react/jsx-dev-runtime',
	'react-dom',
	'react-dom/client',
];
const ROOT_SCAN_IGNORED_DIRECTORIES = new Set([
	'.git',
	'.code-analysis',
	'.serena',
	'node_modules',
	'results',
	'test-results',
	'vendor',
]);
const LEGACY_LINT_MARKERS = [
	['disable', 'Type', 'Checked'].join(''),
	['catalog', 'Surface'].join(''),
	['CATALOG', 'SURFACE'].join('_'),
];
const LEGACY_MARKER_FILENAME = `${LEGACY_LINT_MARKERS[2]}.md`;
const LEGACY_TYPECHECK_BYPASS_FILENAME = [
	['typecheck', 'for', 'surface'].join('-'),
	'mjs',
].join('.');

function toRelativePath(projectRoot, filePath) {
	return path.relative(projectRoot, filePath).split(path.sep).join('/');
}

function isCodeFile(filePath) {
	return /\.(?:[cm]?js|[cm]?ts|tsx)$/i.test(filePath);
}

function isTypeScriptFile(filePath) {
	return /\.(?:[cm]?ts|tsx)$/i.test(filePath);
}

function isWithin(parent, candidate) {
	const relative = path.relative(parent, candidate);
	return relative === '' || (!relative.startsWith(`..${path.sep}`) && relative !== '..' && !path.isAbsolute(relative));
}

async function listFiles(root, ignoredDirectories = new Set()) {
	const files = [];

	async function visit(current) {
		let entries;
		try {
			entries = await readdir(current, { withFileTypes: true });
		} catch (error) {
			if (error && typeof error === 'object' && error.code === 'ENOENT') {
				return;
			}
			throw error;
		}

		for (const entry of entries) {
			const entryPath = path.join(current, entry.name);
			if (entry.isDirectory()) {
				if (!ignoredDirectories.has(entry.name)) {
					await visit(entryPath);
				}
				continue;
			}
			if (entry.isFile()) {
				files.push(entryPath);
			}
		}
	}

	await visit(root);
	return files;
}

async function readJson(filePath) {
	return JSON.parse(await readFile(filePath, 'utf8'));
}

function findTypingMetadata(value, currentPath = '') {
	if (!value || typeof value !== 'object') {
		return [];
	}
	if (Array.isArray(value)) {
		return value.flatMap((item, index) => findTypingMetadata(item, `${currentPath}[${index}]`));
	}

	const findings = [];
	for (const [key, child] of Object.entries(value)) {
		const childPath = currentPath ? `${currentPath}.${key}` : key;
		if (key === 'types' || key === 'typings') {
			findings.push(childPath);
		}
		findings.push(...findTypingMetadata(child, childPath));
	}
	return findings;
}

function push(violations, message) {
	violations.push(message);
}

async function checkFacade(projectRoot, violations) {
	const facade = {
		bridgeExists: existsSync(path.join(projectRoot, BRIDGE_RELATIVE_PATH)),
		stylesExists: existsSync(path.join(projectRoot, STYLES_RELATIVE_PATH)),
	};

	for (const relativePath of FACADE_FILES) {
		if (!existsSync(path.join(projectRoot, relativePath))) {
			push(violations, `Missing runtime facade file: ${relativePath}`);
		}
	}

	const bridgePath = path.join(projectRoot, BRIDGE_RELATIVE_PATH);
	if (facade.bridgeExists) {
		const bridge = await readFile(bridgePath, 'utf8');
		if (!bridge.includes('@npde/')) {
			push(violations, `${BRIDGE_RELATIVE_PATH} must be the explicit vendor-package bridge`);
		}
	}

	const indexPath = path.join(projectRoot, 'src/docx/runtime/index.ts');
	if (existsSync(indexPath)) {
		const index = await readFile(indexPath, 'utf8');
		if (!index.includes('./contract') || !index.includes('./bridge.mjs')) {
			push(violations, 'src/docx/runtime/index.ts must expose contract and bridge exports');
		}
	}

	const stylesPath = path.join(projectRoot, STYLES_RELATIVE_PATH);
	if (facade.stylesExists) {
		const styles = await readFile(stylesPath, 'utf8');
		for (const suffix of [
			'vendor/docx-editor-runtime/core/dist/prosemirror/editor.css',
			'vendor/docx-editor-runtime/react/dist/styles.css',
		]) {
			if (!styles.includes(suffix)) {
				push(violations, `${STYLES_RELATIVE_PATH} must import ${suffix}`);
			}
		}
	}

	return facade;
}

async function checkSourceBoundaries(projectRoot, violations) {
	const sourceRoot = path.join(projectRoot, 'src');
	const sourceFiles = await listFiles(sourceRoot);
	const cssImports = [];

	for (const filePath of sourceFiles) {
		if (!isCodeFile(filePath)) {
			continue;
		}
		const relativePath = toRelativePath(projectRoot, filePath);
		const source = await readFile(filePath, 'utf8');
		if (source.includes('@npde/') && relativePath !== BRIDGE_RELATIVE_PATH) {
			push(violations, `Direct @npde import/reference outside bridge: ${relativePath}`);
		}
		if (source.includes('vendor/docx-editor-runtime') && relativePath !== STYLES_RELATIVE_PATH) {
			push(violations, `Direct vendored runtime reference outside CSS boundary: ${relativePath}`);
		}

		for (const match of source.matchAll(/['"]([^'"]*vendor\/docx-editor-runtime\/[^'"]+\.css)['"]/g)) {
			const specifier = match[1];
			if (!specifier) {
				continue;
			}
			const resolved = path.resolve(path.dirname(filePath), specifier);
			cssImports.push({ relativePath, specifier, resolved });
			if (!existsSync(resolved)) {
				push(violations, `Missing vendored CSS import ${specifier} from ${relativePath}`);
			}
		}
	}

	if (cssImports.length === 0) {
		push(violations, `No direct vendored CSS imports found in ${STYLES_RELATIVE_PATH}`);
	}
	return cssImports;
}

async function checkRuntimeFiles(projectRoot, violations) {
	const runtimeRoot = resolveDocxEditorPackagesRoot(projectRoot);
	if (!existsSync(runtimeRoot)) {
		push(violations, `Missing vendored runtime directory: ${toRelativePath(projectRoot, runtimeRoot)}`);
		return { runtimeRoot, provenance: null, aliases: {}, aliasCount: 0, reactAliasCount: 0 };
	}

	const runtimeFiles = await listFiles(runtimeRoot);
	for (const filePath of runtimeFiles) {
		if (isTypeScriptFile(filePath)) {
			push(violations, `TypeScript/declaration file under vendor: ${toRelativePath(projectRoot, filePath)}`);
		}
	}

	for (const dirName of Object.values(docxEditorPackages)) {
		const manifestPath = path.join(runtimeRoot, dirName, 'package.json');
		if (!existsSync(manifestPath)) {
			push(violations, `Missing runtime package manifest: ${toRelativePath(projectRoot, manifestPath)}`);
			continue;
		}
		const manifest = await readJson(manifestPath);
		for (const metadataPath of findTypingMetadata(manifest)) {
			push(violations, `Typing metadata under vendor: ${toRelativePath(projectRoot, manifestPath)}#${metadataPath}`);
		}
	}

	const provenancePath = path.join(runtimeRoot, 'provenance.json');
	let provenance = null;
	if (!existsSync(provenancePath)) {
		push(violations, 'Missing vendor/docx-editor-runtime/provenance.json');
	} else {
		provenance = await readJson(provenancePath);
		if (typeof provenance.sourceRepository !== 'string' || provenance.sourceRepository.length === 0) {
			push(violations, 'Runtime provenance must include sourceRepository');
		}
		if (provenance.sourceBranch !== 'docx-editor-source') {
			push(violations, 'Runtime provenance sourceBranch must be docx-editor-source');
		}
		if (typeof provenance.sourceCommit !== 'string' || !/^[0-9a-f]{40}$/i.test(provenance.sourceCommit)) {
			push(violations, 'Runtime provenance sourceCommit must be a 40-character Git SHA');
		}
		for (const packageName of Object.values(docxEditorPackages)) {
			if (typeof provenance.packages?.[packageName] !== 'string' || provenance.packages[packageName].length === 0) {
				push(violations, `Runtime provenance must include packages.${packageName}`);
			}
		}
	}

	let aliases = {};
	try {
		aliases = await createDocxEditorAliases(runtimeRoot, projectRoot);
	} catch (error) {
		push(violations, `Could not create runtime aliases: ${error instanceof Error ? error.message : String(error)}`);
	}

	for (const [alias, target] of Object.entries(aliases)) {
		if (!existsSync(target)) {
			push(violations, `Runtime alias does not resolve: ${alias} -> ${target}`);
		}
		if (alias.startsWith('@npde/') && !isWithin(runtimeRoot, target)) {
			push(violations, `Runtime alias escapes vendor runtime: ${alias} -> ${target}`);
		}
	}

	for (const packageName of Object.keys(docxEditorPackages)) {
		if (!Object.keys(aliases).some((alias) => alias === packageName || alias.startsWith(`${packageName}/`))) {
			push(violations, `No aliases generated for ${packageName}`);
		}
	}
	for (const alias of REACT_ALIAS_KEYS) {
		if (!aliases[alias]) {
			push(violations, `Missing React deduplication alias: ${alias}`);
		}
	}

	return {
		runtimeRoot,
		provenance,
		aliases,
		aliasCount: Object.keys(aliases).length,
		reactAliasCount: REACT_ALIAS_KEYS.filter((alias) => alias in aliases).length,
	};
}

async function checkProjectConfiguration(projectRoot, violations) {
	const tsconfigPath = path.join(projectRoot, 'tsconfig.json');
	if (existsSync(tsconfigPath)) {
		const tsconfig = await readJson(tsconfigPath);
		for (const key of Object.keys(tsconfig.compilerOptions?.paths ?? {})) {
			if (key.startsWith('@npde/')) {
				push(violations, `tsconfig.json must not map ${key}`);
			}
		}
	}

	const packagePath = path.join(projectRoot, 'package.json');
	if (existsSync(packagePath)) {
		const manifest = await readJson(packagePath);
		const dependencies = { ...manifest.dependencies, ...manifest.devDependencies };
		for (const name of Object.keys(dependencies)) {
			if (name.startsWith('@npde/')) {
				push(violations, `package.json must not depend on ${name}`);
			}
		}
		if (JSON.stringify(manifest.scripts ?? {}).includes(LEGACY_TYPECHECK_BYPASS_FILENAME)) {
			push(violations, 'package.json must not invoke the legacy catalog typecheck bypass');
		}
	}

	const eslintPath = path.join(projectRoot, 'eslint.config.mts');
	if (existsSync(eslintPath)) {
		const eslint = await readFile(eslintPath, 'utf8');
		if (LEGACY_LINT_MARKERS.some((marker) => eslint.includes(marker))) {
			push(violations, 'eslint.config.mts must not contain a catalog/type-aware disable mode');
		}
	}

	const typecheckBypass = path.join(projectRoot, 'scripts', LEGACY_TYPECHECK_BYPASS_FILENAME);
	if (existsSync(typecheckBypass)) {
		push(violations, 'The legacy catalog typecheck bypass script must not exist');
	}
}

async function checkUnexpectedTypeScript(projectRoot, violations) {
	const files = await listFiles(projectRoot, ROOT_SCAN_IGNORED_DIRECTORIES);
	for (const filePath of files) {
		if (!isTypeScriptFile(filePath)) {
			continue;
		}
		const relativePath = toRelativePath(projectRoot, filePath);
		if (relativePath === 'eslint.config.mts' || relativePath.startsWith('src/')) {
			continue;
		}
		push(violations, `Unexpected TypeScript outside src/: ${relativePath}`);
	}

	for (const filePath of files) {
		if (path.basename(filePath) === LEGACY_MARKER_FILENAME) {
			push(violations, `Catalog marker must not exist: ${toRelativePath(projectRoot, filePath)}`);
		}
	}
}

export async function inspectReviewSurface({ projectRoot = defaultProjectRoot } = {}) {
	const resolvedProjectRoot = path.resolve(projectRoot);
	const violations = [];
	const facade = await checkFacade(resolvedProjectRoot, violations);
	const cssImports = await checkSourceBoundaries(resolvedProjectRoot, violations);
	const runtime = await checkRuntimeFiles(resolvedProjectRoot, violations);
	await checkProjectConfiguration(resolvedProjectRoot, violations);
	await checkUnexpectedTypeScript(resolvedProjectRoot, violations);

	return {
		projectRoot: resolvedProjectRoot,
		violations,
		facade,
		cssImports,
		provenance: runtime.provenance,
		aliasCount: runtime.aliasCount,
		reactAliasCount: runtime.reactAliasCount,
	};
}

export async function checkReviewSurface(options) {
	const report = await inspectReviewSurface(options);
	if (report.violations.length > 0) {
		throw new Error(`Review-surface validation failed:\n${report.violations.map((violation) => `- ${violation}`).join('\n')}`);
	}
	return report;
}

if (process.argv[1] && path.resolve(process.argv[1]) === scriptPath) {
	try {
		const report = await checkReviewSurface();
		console.log(`Review surface OK: ${report.aliasCount} aliases, ${report.reactAliasCount} React aliases.`);
	} catch (error) {
		console.error(error instanceof Error ? error.message : String(error));
		process.exitCode = 1;
	}
}
