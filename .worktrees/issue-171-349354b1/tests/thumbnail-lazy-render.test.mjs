import assert from 'node:assert/strict';
import { test } from 'node:test';
import { build } from 'esbuild';
import { mkdtemp } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const require = createRequire(import.meta.url);

let modulePromise;

async function loadThumbnailLazyRenderModule() {
	modulePromise ??= (async () => {
		const outputDirectory = await mkdtemp(path.join(tmpdir(), 'npde-thumbnail-lazy-'));
		const outfile = path.join(outputDirectory, 'thumbnailLazyRender.cjs');
		await build({
			entryPoints: [path.join(projectRoot, 'src/powerpoint/thumbnailLazyRender.ts')],
			bundle: true,
			format: 'cjs',
			logLevel: 'silent',
			outfile,
			platform: 'node',
			target: 'node22',
		});
		return require(outfile);
	})();
	return modulePromise;
}

test('shouldUseLazyThumbnails gates on slide count', async () => {
	const { shouldUseLazyThumbnails, FULL_THUMBNAIL_RENDER_MAX_SLIDES } = await loadThumbnailLazyRenderModule();
	assert.equal(shouldUseLazyThumbnails(FULL_THUMBNAIL_RENDER_MAX_SLIDES), false);
	assert.equal(shouldUseLazyThumbnails(FULL_THUMBNAIL_RENDER_MAX_SLIDES + 1), true);
});

test('priorityThumbnailIndices centers on the active slide', async () => {
	const { priorityThumbnailIndices, THUMBNAIL_PRIORITY_RADIUS } = await loadThumbnailLazyRenderModule();
	assert.deepEqual(priorityThumbnailIndices(5, 20, THUMBNAIL_PRIORITY_RADIUS), [3, 4, 5, 6, 7]);
	assert.deepEqual(priorityThumbnailIndices(0, 20, THUMBNAIL_PRIORITY_RADIUS), [0, 1, 2]);
	assert.deepEqual(priorityThumbnailIndices(19, 20, THUMBNAIL_PRIORITY_RADIUS), [17, 18, 19]);
});

test('remainingThumbnailIndices skips rendered slides', async () => {
	const { remainingThumbnailIndices } = await loadThumbnailLazyRenderModule();
	assert.deepEqual(remainingThumbnailIndices(5, new Set([1, 3])), [0, 2, 4]);
});

test('sortThumbnailIndicesByProximity prefers the active slide neighborhood', async () => {
	const { sortThumbnailIndicesByProximity } = await loadThumbnailLazyRenderModule();
	assert.deepEqual(
		sortThumbnailIndicesByProximity([10, 2, 7, 5], 5),
		[5, 7, 2, 10],
	);
});
