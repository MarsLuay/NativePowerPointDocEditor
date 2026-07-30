#!/usr/bin/env node
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { emitPptxRuntimeArtifacts } from './lib/emit-pptx-runtime-artifacts.mjs';

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const minify = process.argv.includes('--minify');
const { payloads } = await emitPptxRuntimeArtifacts({
	projectRoot,
	minify,
	logLevel: 'warning',
});
console.log(
	`[emit:pptx-runtime-artifacts] ${payloads.map((payload) => payload.artifact).join(', ')}`,
);
