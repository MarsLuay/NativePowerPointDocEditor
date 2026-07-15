import { existsSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import sdl from '@microsoft/eslint-plugin-sdl';
import obsidianmd from 'eslint-plugin-obsidianmd';
import noUnsanitizedPlugin from 'eslint-plugin-no-unsanitized';
import globals from 'globals';
import tsParser from '@typescript-eslint/parser';
import tseslint from 'typescript-eslint';
import { defineConfig, globalIgnores } from 'eslint/config';
import { obsidianLogicEslintRules } from './scripts/lib/obsidian-logic-eslint-rules.mjs';

const configDir = path.dirname(fileURLToPath(import.meta.url));
/** Public catalog mirror ships JS-only package dist (no `.d.ts`). */
const catalogSurface = existsSync(path.join(configDir, 'docx-editor', 'CATALOG_SURFACE.md'));

export default defineConfig([
	globalIgnores([
		'node_modules',
		'dist',
		'main.js',
		'package-lock.json',
		'src/vendor/**',
		'src/powerpoint/backend/pptxJsEngine.mjs',
		'docx-editor/**',
	]),
	...obsidianmd.configs.recommended,
	// Catalog: obsidianmd recommended enables type-checked typescript-eslint rules,
	// but package typings are intentionally absent — turn those rules off.
	...(catalogSurface ? [tseslint.configs.disableTypeChecked] : []),
	{
		files: ['src/**/*.ts', 'src/**/*.tsx'],
		languageOptions: {
			parser: tsParser,
			parserOptions: catalogSurface
				? {
						// Syntax-only; esbuild binds package JS at build time.
					}
				: {
						project: './tsconfig.json',
						tsconfigRootDir: import.meta.dirname,
					},
			globals: {
				...globals.browser,
			},
		},
		plugins: {
			'@typescript-eslint': tseslint.plugin,
			'@microsoft/sdl': sdl,
			'no-unsanitized': noUnsanitizedPlugin,
			obsidianmd,
		},
		rules: {
			...(catalogSurface ? {} : obsidianLogicEslintRules),
			...(catalogSurface
				? {
						// Type-aware Obsidian rules need package `.d.ts`; catalog is JS-only.
						'obsidianmd/no-plugin-as-component': 'off',
						'obsidianmd/no-unsupported-api': 'off',
						'obsidianmd/no-view-references-in-plugin': 'off',
						'obsidianmd/prefer-file-manager-trash-file': 'off',
						'obsidianmd/prefer-instanceof': 'off',
						'obsidianmd/prefer-create-el': 'off',
						'obsidianmd/prefer-active-doc': 'off',
						// Without a TS program, DOM lib globals look undefined to no-undef.
						'no-undef': 'off',
					}
				: {
						'obsidianmd/prefer-create-el': 'error',
						'obsidianmd/prefer-active-doc': 'error',
						'obsidianmd/settings-tab/prefer-setting-definitions': 'error',
					}),
			'@microsoft/sdl/no-inner-html': 'error',
			'no-unsanitized/method': 'error',
			'no-unsanitized/property': 'error',
			'no-alert': 'error',
			'no-debugger': 'error',
			'no-var': 'error',
			'prefer-const': 'warn',
			'obsidianmd/no-global-this': 'error',
			'obsidianmd/no-static-styles-assignment': 'error',
			'obsidianmd/no-tfile-tfolder-cast': 'error',
			...(catalogSurface ? {} : { 'obsidianmd/no-unsupported-api': 'error' }),
			'obsidianmd/regex-lookbehind': 'error',
			'obsidianmd/vault/iterate': 'warn',
		},
	},
]);
