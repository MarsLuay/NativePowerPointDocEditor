/**
 * Project-local ESLint logic rules so standalone clones do not depend on the
 * ObsidianNotes monorepo analyzer checkout.
 */

/** @type {import('eslint').Linter.RulesRecord} */
export const obsidianLogicEslintRules = {
	'@typescript-eslint/no-floating-promises': [
		'error',
		{
			ignoreVoid: true,
			ignoreIIFE: true,
		},
	],
	'@typescript-eslint/no-misused-promises': [
		'error',
		{
			checksVoidReturn: {
				arguments: false,
				attributes: false,
				properties: false,
				returns: false,
				variables: false,
			},
		},
	],
	'no-empty': [
		'error',
		{
			allowEmptyCatch: false,
		},
	],
};
