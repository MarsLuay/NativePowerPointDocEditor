import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import ts from 'typescript';
import { PPTX_ACTION_LOG_ALLOWLIST } from './pptx-action-log-allowlist.mjs';

const CONTROLLERS = [
	'arrangeController',
	'insertController',
	'inspectorController',
	'selectionDragController',
	'snapController',
	'textToolbarController',
	'findReplaceController',
	'historyController',
	'exportController',
	'menuBarController',
	'saveController',
	'slideFilmstripController',
];

function isPublicMethod(node) {
	if (!ts.isMethodDeclaration(node) || !node.body || !node.name) return false;
	return !node.modifiers?.some((modifier) =>
		modifier.kind === ts.SyntaxKind.PrivateKeyword || modifier.kind === ts.SyntaxKind.ProtectedKeyword,
	);
}

function methodName(node) {
	return node.name.getText();
}

function hasActionLog(sourceText, sourceFile, node) {
	if (/\b(?:debugLog|logPptxAction)\s*\(/.test(node.body.getText(sourceFile))) return true;

	const startLine = sourceFile.getLineAndCharacterOfPosition(node.end).line;
	const endLine = Math.min(startLine + 40, sourceFile.getLineStarts().length - 1);
	const start = sourceFile.getLineStarts()[startLine];
	const end = endLine + 1 < sourceFile.getLineStarts().length
		? sourceFile.getLineStarts()[endLine + 1]
		: sourceText.length;
	return /\b(?:debugLog|logPptxAction)\s*\(/.test(sourceText.slice(start, end));
}

const missing = [];
for (const controller of CONTROLLERS) {
	const relativePath = `src/powerpoint/${controller}.ts`;
	const sourceText = readFileSync(resolve(relativePath), 'utf8');
	const sourceFile = ts.createSourceFile(relativePath, sourceText, ts.ScriptTarget.Latest, true);
	const exemptions = new Set(PPTX_ACTION_LOG_ALLOWLIST[controller] ?? []);

	for (const statement of sourceFile.statements) {
		if (!ts.isClassDeclaration(statement)) continue;
		for (const member of statement.members) {
			if (!isPublicMethod(member)) continue;
			const name = methodName(member);
			if (exemptions.has(name) || hasActionLog(sourceText, sourceFile, member)) continue;
			const line = sourceFile.getLineAndCharacterOfPosition(member.getStart(sourceFile)).line + 1;
			missing.push(`${relativePath}:${line} ${name}()`);
		}
	}
}

if (missing.length > 0) {
	console.error('PPTX controller public methods missing nearby debugLog/logPptxAction:');
	for (const entry of missing) console.error(`- ${entry}`);
	process.exitCode = 1;
} else {
	console.log('PPTX controller action-log coverage passed.');
}
