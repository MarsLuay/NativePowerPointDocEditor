/**
 * Decode HEIC/HEIF into PNG for PowerPoint insert/paste.
 *
 * After raster decode, PNG keeps full 8-bit RGBA without the extra loss JPEG
 * would add — best quality for a format PowerPoint already handles well.
 */

import decodeHeic from 'heic-decode';

const HEIC_MIME_TYPES = new Set([
  'image/heic',
  'image/heif',
  'image/heic-sequence',
  'image/heif-sequence',
]);

const HEIC_EXTENSIONS = new Set(['heic', 'heif', 'heics', 'heifs']);

/** Brands at bytes 8–12 of an ISO BMFF `ftyp` box (HEIC/HEIF family). */
const HEIC_FTYP_BRANDS = new Set(['heic', 'heix', 'hevc', 'hevx', 'mif1', 'msf1']);

export function isHeicMimeType(mimeType: string | null | undefined): boolean {
  if (!mimeType) return false;
  return HEIC_MIME_TYPES.has(mimeType.toLowerCase().split(';')[0]?.trim() ?? '');
}

export function isHeicExtension(extension: string | null | undefined): boolean {
  if (!extension) return false;
  return HEIC_EXTENSIONS.has(extension.toLowerCase().replace(/^\./, ''));
}

export function looksLikeHeicBytes(bytes: Uint8Array): boolean {
  if (bytes.length < 12) return false;
  const brand = [bytes[8]!, bytes[9]!, bytes[10]!, bytes[11]!]
    .map((byte) => (byte === 0 ? ' ' : String.fromCharCode(byte)))
    .join('')
    .trim();
  return HEIC_FTYP_BRANDS.has(brand);
}

export function shouldConvertHeicToPng(
  bytes: Uint8Array,
  mimeType?: string | null,
  extensionOrFileName?: string | null,
): boolean {
  if (isHeicMimeType(mimeType)) return true;
  const extension = extensionOrFileName?.includes('.')
    ? extensionOrFileName.split('.').pop()
    : extensionOrFileName;
  if (isHeicExtension(extension)) return true;
  return looksLikeHeicBytes(bytes);
}

/**
 * heic-decode's brand check does `String.fromCharCode(...buffer.slice(8, 12))`.
 * That requires a Uint8Array/Buffer: ArrayBuffer.slice is not iterable and throws
 * "Spread syntax requires ...iterable[Symbol.iterator] to be a function".
 */
function toHeicDecodeBuffer(bytes: Uint8Array): Uint8Array {
  if (bytes.byteOffset === 0 && bytes.byteLength === bytes.buffer.byteLength) {
    return bytes;
  }
  return bytes.slice();
}

function createCanvas(width: number, height: number): HTMLCanvasElement {
  const scopedWindow = window as Window & {
    createEl?: (tag: 'canvas') => HTMLCanvasElement;
  };
  const canvas = typeof scopedWindow.createEl === 'function'
    ? scopedWindow.createEl('canvas')
    : createEl('canvas');
  canvas.width = width;
  canvas.height = height;
  return canvas;
}

/**
 * Longest edge after HEIC→PNG. Phone HEIC often decodes to 4k–8k; a full-res
 * lossless PNG is 20MB+ and makes every transform/thumbnail hit multi-second.
 * 2560 covers 1080p/2K slides at ~2× without the editor cost.
 */
export const HEIC_PNG_MAX_EDGE = 2560;

export function fitWithinMaxEdge(
  width: number,
  height: number,
  maxEdge: number = HEIC_PNG_MAX_EDGE,
): { width: number; height: number; scale: number } {
  const safeWidth = Math.max(1, Math.round(width));
  const safeHeight = Math.max(1, Math.round(height));
  const longEdge = Math.max(safeWidth, safeHeight);
  const limit = Math.max(1, Math.round(maxEdge));
  if (longEdge <= limit) {
    return { width: safeWidth, height: safeHeight, scale: 1 };
  }
  const scale = limit / longEdge;
  return {
    width: Math.max(1, Math.round(safeWidth * scale)),
    height: Math.max(1, Math.round(safeHeight * scale)),
    scale,
  };
}

/** Decode HEIC/HEIF bytes to PNG bytes (lossless relative to the decoded raster). */
export async function convertHeicBytesToPng(bytes: Uint8Array): Promise<Uint8Array> {
  const decoded = await decodeHeic({ buffer: toHeicDecodeBuffer(bytes) });
  const { width, height, data } = decoded;
  if (!width || !height || !data) {
    throw new Error('HEIC decode produced an empty image.');
  }

  const sourceCanvas = createCanvas(width, height);
  const sourceContext = sourceCanvas.getContext('2d');
  if (!sourceContext) {
    throw new Error('Could not create a canvas to encode HEIC as PNG.');
  }

  // Copy into a fresh buffer so TS/DOM ImageDataArray accepts the pixels
  // (heic-decode's view may be backed by SharedArrayBuffer-like storage).
  const rgba = new Uint8ClampedArray(width * height * 4);
  rgba.set(data);
  sourceContext.putImageData(new ImageData(rgba, width, height), 0, 0);

  const fitted = fitWithinMaxEdge(width, height);
  let encodeCanvas = sourceCanvas;
  if (fitted.scale !== 1) {
    encodeCanvas = createCanvas(fitted.width, fitted.height);
    const encodeContext = encodeCanvas.getContext('2d');
    if (!encodeContext) {
      throw new Error('Could not create a canvas to downscale HEIC as PNG.');
    }
    encodeContext.imageSmoothingEnabled = true;
    encodeContext.imageSmoothingQuality = 'high';
    encodeContext.drawImage(sourceCanvas, 0, 0, fitted.width, fitted.height);
  }

  const blob = await new Promise<Blob>((resolve, reject) => {
    encodeCanvas.toBlob(
      (result) => {
        if (result) resolve(result);
        else reject(new Error('Canvas PNG encode failed for HEIC conversion.'));
      },
      'image/png',
    );
  });

  return new Uint8Array(await blob.arrayBuffer());
}

export interface NormalizedPowerPointImage {
  bytes: Uint8Array;
  mimeType: string;
  convertedFromHeic: boolean;
}

/**
 * Ensure image bytes are a PowerPoint-friendly raster. HEIC/HEIF become PNG;
 * other formats pass through unchanged.
 */
export async function normalizeImageForPowerPoint(
  bytes: Uint8Array,
  mimeType?: string | null,
  extensionOrFileName?: string | null,
): Promise<NormalizedPowerPointImage> {
  if (!shouldConvertHeicToPng(bytes, mimeType, extensionOrFileName)) {
    return {
      bytes,
      mimeType: mimeType && mimeType.length > 0 ? mimeType : 'image/png',
      convertedFromHeic: false,
    };
  }

  const pngBytes = await convertHeicBytesToPng(bytes);
  return {
    bytes: pngBytes,
    mimeType: 'image/png',
    convertedFromHeic: true,
  };
}
