import { createRequire } from 'node:module';
import { access, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);
const scriptDirectory = path.dirname(fileURLToPath(import.meta.url));
const monorepoRoot = path.resolve(scriptDirectory, '..');
const defaultPluginDirectory = path.resolve(monorepoRoot, '..', '..', 'NativePowerPointDocEditor');
const pluginDirectory = path.resolve(
	process.cwd(),
	process.env.NPDE_PLUGIN_DIR || defaultPluginDirectory,
);
const contractPath = path.join(pluginDirectory, 'src', 'docx', 'runtime', 'contract.ts');

const packageExports = [
	{
		packageDirectory: 'packages/react',
		exports: ['.'],
	},
	{
		packageDirectory: 'packages/core',
		exports: ['./layout-bridge', './prosemirror/commands', './utils', './utils/fontOptions', './plugin-api'],
	},
	{
		packageDirectory: 'packages/i18n',
		exports: ['.', './en', './he', './pl', './pt-BR', './tr', './zh-CN'],
	},
];

const requiredDeclarationFiles = [
	'packages/react/dist/index.d.ts',
	'packages/core/dist/core.d.ts',
	'packages/core/dist/layout-bridge/index.d.ts',
	'packages/core/dist/prosemirror/commands/index.d.ts',
	'packages/core/dist/utils/index.d.ts',
	'packages/core/dist/utils/fontOptions.d.ts',
	'packages/core/dist/plugin-api/index.d.ts',
	'packages/i18n/dist/index.d.ts',
	'packages/i18n/dist/en.d.ts',
	'packages/i18n/dist/he.d.ts',
	'packages/i18n/dist/pl.d.ts',
	'packages/i18n/dist/pt-BR.d.ts',
	'packages/i18n/dist/tr.d.ts',
	'packages/i18n/dist/zh-CN.d.ts',
];

function fail(message) {
	throw new Error(message);
}

function toPosixPath(filePath) {
	return filePath.split(path.sep).join('/');
}

async function assertFile(filePath, message) {
	try {
		await access(filePath);
	} catch {
		fail(`${message}: ${filePath}`);
	}
}

async function readJson(filePath) {
	try {
		return JSON.parse(await readFile(filePath, 'utf8'));
	} catch (error) {
		fail(`Could not read ${filePath}: ${error.message}`);
	}
}

function getTypeTarget(exportTarget, label) {
	if (!exportTarget || Array.isArray(exportTarget) || typeof exportTarget !== 'object') {
		fail(`${label} must declare a types target.`);
	}
	if (typeof exportTarget.types !== 'string') {
		fail(`${label} is missing its types target.`);
	}
	return exportTarget.types;
}

async function validatePublishedExports() {
	for (const packageSpec of packageExports) {
		const packageRoot = path.join(monorepoRoot, packageSpec.packageDirectory);
		const manifest = await readJson(path.join(packageRoot, 'package.json'));
		for (const exportPath of packageSpec.exports) {
			const exportTarget = manifest.exports?.[exportPath];
			const typeTarget = getTypeTarget(exportTarget, `${manifest.name} export ${exportPath}`);
			if (!typeTarget.startsWith('./dist/')) {
				fail(`${manifest.name} export ${exportPath} points outside dist: ${typeTarget}`);
			}
			await assertFile(
				path.resolve(packageRoot, typeTarget),
				`${manifest.name} export ${exportPath} has no generated declaration`,
			);
		}
	}
}

function toModuleSpecifier(fromDirectory, targetPath) {
	const relative = toPosixPath(path.relative(fromDirectory, targetPath));
	return relative.startsWith('.') ? relative : `./${relative}`;
}

function createCompatibilitySource(contractSpecifier) {
	return `import {
	DocxEditor,
	type DocxEditorProps as SourceDocxEditorProps,
	type DocxEditorRef as SourceDocxEditorRef,
} from '@npde/docx-editor-react';
import { clearParagraphMeasureCache } from '@npde/docx-editor-core/layout-bridge';
import {
	insertTable,
	setFontSize,
	setLineSpacing,
} from '@npde/docx-editor-core/prosemirror/commands';
import { loadFontFromBuffer } from '@npde/docx-editor-core/utils';
import type { FontOption as SourceFontOption } from '@npde/docx-editor-core/utils/fontOptions';
import type { RenderedDomContext as SourceRenderedDomContext } from '@npde/docx-editor-core/plugin-api';
import {
	createT,
	deepMerge,
	en,
	type TranslationKey as SourceTranslationKey,
} from '@npde/docx-editor-i18n';
import enLocale from '@npde/docx-editor-i18n/en';
import heLocale from '@npde/docx-editor-i18n/he';
import plLocale from '@npde/docx-editor-i18n/pl';
import ptBrLocale from '@npde/docx-editor-i18n/pt-BR';
import trLocale from '@npde/docx-editor-i18n/tr';
import zhCnLocale from '@npde/docx-editor-i18n/zh-CN';
import type {
	DocxCommand,
	DocxDocumentPart as ContractDocxDocumentPart,
	DocxEditorProps as ContractDocxEditorProps,
	DocxEditorRef as ContractDocxEditorRef,
	FontOption as ContractFontOption,
	RenderedDomContext as ContractRenderedDomContext,
	Translations as ContractTranslations,
} from ${JSON.stringify(contractSpecifier)};

type Assert<T extends true> = T;
type IsAssignable<From, To> = [From] extends [To] ? true : false;
type HasKeys<T, Keys extends PropertyKey> = Exclude<Keys, keyof T> extends never ? true : false;

type NpdeDocxEditorProp =
	| 'author'
	| 'className'
	| 'colorMode'
	| 'commentsSidebarOpen'
	| 'disableFindReplaceShortcuts'
	| 'documentBuffer'
	| 'documentName'
	| 'documentNameEditable'
	| 'externalPlugins'
	| 'fontFamilies'
	| 'i18n'
	| 'initialZoom'
	| 'mode'
	| 'onCommentsSidebarOpenChange'
	| 'onEditorViewReady'
	| 'onFontsLoaded'
	| 'onModeChange'
	| 'onRenderedDomContextReady'
	| 'onSelectionChange'
	| 'pluginSidebarItems'
	| 'readOnly'
	| 'showOutlineButton'
	| 'showRuler'
	| 'showToolbar'
	| 'showZoomControl';

type _ContractDeclaresNpdeProps = Assert<HasKeys<ContractDocxEditorProps, NpdeDocxEditorProp>>;
type _SourceDeclaresNpdeProps = Assert<HasKeys<SourceDocxEditorProps, NpdeDocxEditorProp>>;
type _SourceRefMethodsSatisfyContract = Assert<IsAssignable<
	Omit<SourceDocxEditorRef, 'getDocument'>,
	Omit<ContractDocxEditorRef, 'getDocument'>
>>;
type SourceDocument = NonNullable<ReturnType<SourceDocxEditorRef['getDocument']>>;
type SourceDocumentPart = SourceDocument['package']['document'];
type _SourceDocumentPartSatisfiesContract = Assert<IsAssignable<
	SourceDocumentPart,
	ContractDocxDocumentPart
>>;
type _SourceRenderedDomContextSatisfiesContract = Assert<IsAssignable<
	SourceRenderedDomContext,
	ContractRenderedDomContext
>>;
type _SourceFontOptionSatisfiesContract = Assert<IsAssignable<SourceFontOption, ContractFontOption>>;
type _SourceTranslationKeyIsString = Assert<IsAssignable<SourceTranslationKey, string>>;
type _EnglishLocaleSatisfiesContract = Assert<IsAssignable<typeof en, ContractTranslations>>;
type _LoadedEnglishLocaleSatisfiesContract = Assert<IsAssignable<typeof enLocale, ContractTranslations>>;
type _LoadedHebrewLocaleSatisfiesContract = Assert<IsAssignable<typeof heLocale, ContractTranslations>>;
type _LoadedPolishLocaleSatisfiesContract = Assert<IsAssignable<typeof plLocale, ContractTranslations>>;
type _LoadedPortugueseLocaleSatisfiesContract = Assert<IsAssignable<typeof ptBrLocale, ContractTranslations>>;
type _LoadedTurkishLocaleSatisfiesContract = Assert<IsAssignable<typeof trLocale, ContractTranslations>>;
type _LoadedChineseLocaleSatisfiesContract = Assert<IsAssignable<typeof zhCnLocale, ContractTranslations>>;

const clearCache: () => void = clearParagraphMeasureCache;
const tableCommand: (rows: number, columns: number) => DocxCommand = insertTable;
const fontSizeCommand: (size: number) => DocxCommand = setFontSize;
const lineSpacingCommand: (value: number) => DocxCommand = setLineSpacing;
const loadFont: (fontFamily: string, buffer: ArrayBuffer) => Promise<boolean> = loadFontFromBuffer;
const mergedLocale: Record<string, unknown> = deepMerge(en, undefined);
const translator = createT(en, 'en');
const translatedValue: string = translator('formattingBar.bold');

void DocxEditor;
void clearCache;
void tableCommand;
void fontSizeCommand;
void lineSpacingCommand;
void loadFont;
void mergedLocale;
void translatedValue;
`;
}

function createCompilerOptions(ts) {
	const declarationPath = (relativePath) => toPosixPath(relativePath);
	const nodeModulesPath = (relativePath) => declarationPath(path.join('node_modules', relativePath));

	return {
		allowImportingTsExtensions: true,
		allowSyntheticDefaultImports: true,
		baseUrl: monorepoRoot,
		esModuleInterop: true,
		jsx: ts.JsxEmit.ReactJSX,
		lib: ['lib.es2022.d.ts', 'lib.dom.d.ts', 'lib.dom.iterable.d.ts'],
		module: ts.ModuleKind.ESNext,
		moduleResolution: ts.ModuleResolutionKind.Bundler,
		noEmit: true,
		paths: {
			'@npde/docx-editor-react': [declarationPath('packages/react/dist/index.d.ts')],
			'@npde/docx-editor-core': [declarationPath('packages/core/dist/core.d.ts')],
			'@npde/docx-editor-core/*': [declarationPath('packages/core/dist/*')],
			'@npde/docx-editor-i18n': [declarationPath('packages/i18n/dist/index.d.ts')],
			'@npde/docx-editor-i18n/*': [declarationPath('packages/i18n/dist/*')],
			'react': [nodeModulesPath('@types/react/index.d.ts')],
			'react/*': [nodeModulesPath('@types/react/*')],
			'prosemirror-state': [nodeModulesPath('prosemirror-state/dist/index.d.ts')],
			'prosemirror-view': [nodeModulesPath('prosemirror-view/dist/index.d.ts')],
		},
		resolveJsonModule: true,
		skipLibCheck: true,
		strict: true,
		target: ts.ScriptTarget.ES2022,
	};
}

function formatDiagnostics(ts, diagnostics) {
	return ts.formatDiagnosticsWithColorAndContext(diagnostics, {
		getCanonicalFileName: (fileName) => fileName,
		getCurrentDirectory: () => monorepoRoot,
		getNewLine: () => '\n',
	});
}

async function loadTypeScript() {
	try {
		return require('typescript');
	} catch {
		fail('TypeScript is unavailable. Run `bun install` in docx-editor before `bun run check:npde-contract`.');
	}
}

async function main() {
	await assertFile(contractPath, 'Missing NPDE runtime contract');
	for (const relativePath of requiredDeclarationFiles) {
		await assertFile(
			path.join(monorepoRoot, relativePath),
			'Missing built package declaration; run `bun run build` first',
		);
	}
	await validatePublishedExports();

	const ts = await loadTypeScript();
	const temporaryDirectory = await mkdtemp(path.join(os.tmpdir(), 'npde-contract-'));
	const compatibilityPath = path.join(temporaryDirectory, 'compatibility.ts');

	try {
		await writeFile(
			compatibilityPath,
			createCompatibilitySource(toModuleSpecifier(temporaryDirectory, contractPath)),
		);
		const program = ts.createProgram([compatibilityPath], createCompilerOptions(ts));
		const diagnostics = ts.getPreEmitDiagnostics(program)
			.filter((diagnostic) => diagnostic.category === ts.DiagnosticCategory.Error);
		if (diagnostics.length > 0) {
			fail(`NPDE runtime contract is incompatible with the DOCX editor source.\n${formatDiagnostics(ts, diagnostics)}`);
		}
	} finally {
		await rm(temporaryDirectory, { recursive: true, force: true });
	}

	console.log(`NPDE DOCX contract compatibility passed against ${pluginDirectory}.`);
}

main().catch((error) => {
	console.error(`NPDE DOCX contract check failed: ${error.message}`);
	process.exitCode = 1;
});
