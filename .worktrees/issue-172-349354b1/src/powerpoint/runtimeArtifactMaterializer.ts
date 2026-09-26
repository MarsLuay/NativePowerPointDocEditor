import payloadJson from './generated/runtimeArtifactPayloads.json';
import type { PptxRuntimeArtifact } from './runtimeArtifactLoader';
import { debugLog, errorLog } from '../logger';

interface PptxRuntimeArtifactPayload {
	artifact: string;
	sha256: string;
	gzipBase64: string;
}

const PPTX_RUNTIME_ARTIFACT_PAYLOADS = payloadJson as PptxRuntimeArtifactPayload[];

export interface PptxRuntimeArtifactFs {
	exists(path: string): Promise<boolean>;
	read(path: string): Promise<ArrayBuffer>;
	write(path: string, data: ArrayBuffer): Promise<void>;
}

function base64ToBytes(base64: string): Uint8Array {
	const binary = atob(base64);
	const bytes = new Uint8Array(binary.length);
	for (let index = 0; index < binary.length; index += 1) {
		bytes[index] = binary.charCodeAt(index);
	}
	return bytes;
}

async function gunzipBytes(compressed: Uint8Array): Promise<Uint8Array> {
	if (typeof DecompressionStream !== 'function') {
		throw new Error('DecompressionStream is required to materialize PowerPoint runtime artifacts.');
	}
	const copy = Uint8Array.from(compressed);
	const stream = new Blob([copy]).stream().pipeThrough(new DecompressionStream('gzip'));
	return new Uint8Array(await new Response(stream).arrayBuffer());
}

async function sha256Hex(bytes: Uint8Array): Promise<string> {
	const copy = Uint8Array.from(bytes);
	const digest = await crypto.subtle.digest('SHA-256', copy);
	return [...new Uint8Array(digest)].map((value) => value.toString(16).padStart(2, '0')).join('');
}

async function arrayBufferMatchesSha256(buffer: ArrayBuffer, expectedSha256: string): Promise<boolean> {
	return (await sha256Hex(new Uint8Array(buffer))) === expectedSha256;
}

/**
 * Write gzip-embedded optional runtimes beside main.js when missing or stale.
 * Community installs only receive main.js / styles.css / manifest.json; vault
 * and Sync keep each extracted file under the 5 MB per-file limit.
 */
function asPptxRuntimeArtifact(artifact: string): PptxRuntimeArtifact {
	if (
		artifact === 'pptx-js-engine.mjs'
		|| artifact === 'pptx-wasm-renderer.mjs'
		|| artifact === 'heic-decode.mjs'
	) {
		return artifact;
	}
	throw new Error(`Unknown embedded PowerPoint runtime artifact: ${artifact}`);
}

export async function ensurePptxRuntimeArtifacts(
	pluginDir: string,
	fs: PptxRuntimeArtifactFs,
	joinPath: (dir: string, artifact: PptxRuntimeArtifact) => string,
): Promise<void> {
	for (const payload of PPTX_RUNTIME_ARTIFACT_PAYLOADS) {
		const artifact = asPptxRuntimeArtifact(payload.artifact);
		const path = joinPath(pluginDir, artifact);
		try {
			if (await fs.exists(path)) {
				const existing = await fs.read(path);
				if (await arrayBufferMatchesSha256(existing, payload.sha256)) {
					continue;
				}
			}
			const startedAt = Date.now();
			const bytes = await gunzipBytes(base64ToBytes(payload.gzipBase64));
			const actualSha = await sha256Hex(bytes);
			if (actualSha !== payload.sha256) {
				throw new Error(
					`Embedded ${artifact} gunzip sha256 mismatch (got ${actualSha}, expected ${payload.sha256}).`,
				);
			}
			const writable = Uint8Array.from(bytes);
			await fs.write(path, writable.buffer);
			debugLog('render', 'Materialized optional PowerPoint runtime artifact', {
				artifact,
				path,
				bytes: writable.byteLength,
				durationMs: Date.now() - startedAt,
			});
		} catch (error) {
			errorLog('render', 'Failed to materialize optional PowerPoint runtime artifact', {
				artifact,
				path,
				error,
			});
			throw error;
		}
	}
}
