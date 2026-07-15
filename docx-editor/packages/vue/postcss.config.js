import path from 'path';
import { fileURLToPath } from 'url';
import tailwindcss from 'tailwindcss';
import autoprefixer from 'autoprefixer';

const __configDir = path.dirname(fileURLToPath(import.meta.url));

/**
 * Prepend `@tailwind utilities` only for the shared editor.css entry so the
 * core stylesheet can stay plain CSS (no unknown at-rules in source).
 *
 * @returns {import('postcss').Plugin}
 */
function prependTailwindUtilities() {
	return {
		postcssPlugin: 'prepend-tailwind-utilities',
		Once(root, { result }) {
			const from = (result.opts.from || '').replace(/\\/g, '/');
			if (!from.endsWith('/styles/editor.css')) {
				return;
			}
			const already = root.nodes?.some(
				(node) => node.type === 'atrule' && node.name === 'tailwind',
			);
			if (already) {
				return;
			}
			root.prepend({ type: 'atrule', name: 'tailwind', params: 'utilities' });
		},
	};
}
prependTailwindUtilities.postcss = true;

// Vue builds with vite, whose CSS pass auto-discovers the nearest PostCSS
// config. React instead runs the standalone `tailwindcss` CLI in a separate
// `build:css` step because tsup does not process CSS — so the equivalent for
// Vue is to wire Tailwind into vite's existing PostCSS pass here rather than
// add a separate step (which would clobber the SFC <style> CSS vite emits).
//
// The tailwind config path is absolute so it resolves regardless of the cwd
// the build runs from (see issue #340).
export default {
	plugins: [
		prependTailwindUtilities(),
		tailwindcss({ config: path.join(__configDir, 'tailwind.config.js') }),
		autoprefixer(),
	],
};
