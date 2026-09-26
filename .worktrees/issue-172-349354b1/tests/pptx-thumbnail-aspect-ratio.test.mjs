import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { test } from 'node:test';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

test('filmstrip preview uses the PPTX slide dimensions instead of a fixed format', async () => {
	const [controller, css] = await Promise.all([
		readFile(path.join(projectRoot, 'src/powerpoint/slideFilmstripController.ts'), 'utf8'),
		readFile(path.join(projectRoot, 'styles.css'), 'utf8'),
	]);

	assert.match(controller, /const \{ cx, cy \} = await engine\.getSlideSizeEmu\(\);/);
	assert.match(
		controller,
		/thumbnailContainer\.style\.setProperty\('--native-powerpoint-thumbnail-aspect-ratio', aspectRatio\);/,
	);
	assert.match(
		css,
		/\.native-powerpoint-thumbnail-preview\s*\{\s*aspect-ratio: var\(--native-powerpoint-thumbnail-aspect-ratio, 16 \/ 9\);/,
	);
});
