/**
 * Read a raster image from the system clipboard for PowerPoint insert.
 * Prefer HEIC/HEIF when present so callers can convert to PNG.
 */

const IMAGE_MIME_PRIORITY = [
  'image/heic',
  'image/heif',
  'image/heic-sequence',
  'image/heif-sequence',
  'image/png',
  'image/jpeg',
  'image/webp',
  'image/gif',
  'image/bmp',
] as const;

export interface ClipboardRasterImage {
  bytes: Uint8Array;
  mimeType: string;
  fileName: string | null;
}

function extensionForMime(mimeType: string): string {
  switch (mimeType.toLowerCase().split(';')[0]?.trim()) {
    case 'image/jpeg':
      return 'jpg';
    case 'image/gif':
      return 'gif';
    case 'image/webp':
      return 'webp';
    case 'image/bmp':
      return 'bmp';
    case 'image/heic':
    case 'image/heic-sequence':
      return 'heic';
    case 'image/heif':
    case 'image/heif-sequence':
      return 'heif';
    default:
      return 'png';
  }
}

function pickImageMime(types: readonly string[]): string | null {
  const normalized = types.map((type) => type.toLowerCase());
  for (const preferred of IMAGE_MIME_PRIORITY) {
    if (normalized.includes(preferred)) return preferred;
  }
  return normalized.find((type) => type.startsWith('image/')) ?? null;
}

async function blobToClipboardImage(
  blob: Blob,
  mimeType: string,
  fileName: string | null,
): Promise<ClipboardRasterImage> {
  return {
    bytes: new Uint8Array(await blob.arrayBuffer()),
    mimeType: mimeType || blob.type || 'application/octet-stream',
    fileName,
  };
}

/** Read the highest-priority image item from `navigator.clipboard`. */
export async function readClipboardRasterImage(): Promise<ClipboardRasterImage | null> {
  if (typeof navigator === 'undefined' || !navigator.clipboard?.read) {
    return null;
  }

  try {
    const items = await navigator.clipboard.read();
    for (const item of items) {
      const mimeType = pickImageMime(item.types);
      if (!mimeType) continue;
      const blob = await item.getType(mimeType);
      return blobToClipboardImage(
        blob,
        mimeType,
        `clipboard.${extensionForMime(mimeType)}`,
      );
    }
  } catch {
    // Permission denied / empty clipboard / unsupported — fall through.
  }

  return null;
}

/** Read an image from a paste/drop `DataTransfer`, including Finder HEIC files. */
export async function readDataTransferRasterImage(
  dataTransfer: DataTransfer | null | undefined,
): Promise<ClipboardRasterImage | null> {
  if (!dataTransfer) return null;

  const files = Array.from(dataTransfer.files ?? []);
  for (const file of files) {
    const mimeType = file.type || '';
    const looksImage = mimeType.startsWith('image/')
      || /\.(heic|heif|heics|heifs|png|jpe?g|gif|webp|bmp)$/i.test(file.name);
    if (!looksImage) continue;
    return blobToClipboardImage(file, mimeType || 'application/octet-stream', file.name);
  }

  const items = Array.from(dataTransfer.items ?? []);
  for (const preferred of IMAGE_MIME_PRIORITY) {
    const item = items.find((entry) => entry.kind === 'file' && entry.type.toLowerCase() === preferred);
    if (!item) continue;
    const file = item.getAsFile();
    if (!file) continue;
    return blobToClipboardImage(file, preferred, file.name || `clipboard.${extensionForMime(preferred)}`);
  }

  for (const item of items) {
    if (item.kind !== 'file' || !item.type.toLowerCase().startsWith('image/')) continue;
    const file = item.getAsFile();
    if (!file) continue;
    return blobToClipboardImage(
      file,
      item.type,
      file.name || `clipboard.${extensionForMime(item.type)}`,
    );
  }

  return null;
}
