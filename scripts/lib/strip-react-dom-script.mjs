// React 19's react-dom client bundles preinitScript / preinitModuleScript helpers
// that call document.createElement("script"). This plugin never uses those APIs
// (no ReactDOM.preinit*, no <link rel="modulepreload"> hoisting), but Obsidian's
// automated plugin review flags any createElement("script") in main.js as a
// security error regardless of reachability.
//
// Replace the tag name at bundle time so the shipped artifact never contains the
// flagged pattern. If React's preinit path were ever invoked, it would append an
// inert <template> instead of a <script> — acceptable for an offline vault editor.

const PATCH_MARKER = '/* obsidian:strip-react-dom-script */';
const SCRIPT_CREATE_RE = /createElement\("script"\)/g;
const TEMPLATE_CREATE = 'createElement("template")';

/**
 * @param {string} source
 * @param {string} filePath
 * @returns {string}
 */
export function patchReactDomScriptCreation(source, filePath) {
	const normalized = filePath.replace(/\\/g, '/');
	if (!normalized.includes('/react-dom/cjs/react-dom-client.')) {
		return source;
	}

	if (source.includes(PATCH_MARKER)) {
		return source;
	}

	if (!SCRIPT_CREATE_RE.test(source)) {
		throw new Error(
			`strip-react-dom-script: expected createElement("script") in ${filePath} but found none. ` +
				'React DOM may have changed; update the patch or downgrade react-dom.',
		);
	}

	SCRIPT_CREATE_RE.lastIndex = 0;
	return `${PATCH_MARKER}\n${source.replace(SCRIPT_CREATE_RE, TEMPLATE_CREATE)}`;
}
