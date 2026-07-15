import path from 'path';
import { fileURLToPath } from 'url';

const __configDir = path.dirname(fileURLToPath(import.meta.url));

/**
 * Shared color/theme palette lives in the core preset (single source of truth);
 * this config only adds the React-specific content glob + .docx-editor-root scoping.
 * @type {import('tailwindcss').Config}
 */
export default {
  presets: [require('../core/tailwind-preset.cjs')],
  // Scope all utilities under .docx-editor-root to avoid clashing with host app CSS
  important: '.docx-editor-root',
  // Only scan library source files, not demo. Absolute path so this works no
  // matter where `tailwindcss -c` is invoked from (see issue #340).
  content: [path.join(__configDir, 'src/**/*.{ts,tsx}')],
};
