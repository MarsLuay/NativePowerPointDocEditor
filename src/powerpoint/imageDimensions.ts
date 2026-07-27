export interface RasterImageDimensions {
  width: number;
  height: number;
}

function dimensions(width: number, height: number): RasterImageDimensions | null {
  return Number.isFinite(width) && Number.isFinite(height) && width > 0 && height > 0
    ? { width, height }
    : null;
}

function hasBytes(bytes: Uint8Array, start: number, count: number): boolean {
  return start >= 0 && count >= 0 && start + count <= bytes.length;
}

/**
 * Read common embedded-raster dimensions without a browser decoder. PowerPoint
 * images are packaged as raw media parts, so this stays available to both the
 * editor and AI paths without creating object URLs or canvases.
 */
export function readRasterImageDimensions(bytes: Uint8Array): RasterImageDimensions | null {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);

  if (
    hasBytes(bytes, 0, 24)
    && bytes[0] === 0x89
    && bytes[1] === 0x50
    && bytes[2] === 0x4e
    && bytes[3] === 0x47
  ) {
    return dimensions(view.getUint32(16), view.getUint32(20));
  }

  if (hasBytes(bytes, 0, 10) && bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46) {
    return dimensions(view.getUint16(6, true), view.getUint16(8, true));
  }

  if (hasBytes(bytes, 0, 26) && bytes[0] === 0x42 && bytes[1] === 0x4d) {
    return dimensions(view.getInt32(18, true), Math.abs(view.getInt32(22, true)));
  }

  if (hasBytes(bytes, 0, 12) && bytes[0] === 0xff && bytes[1] === 0xd8) {
    let offset = 2;
    while (offset + 9 < bytes.length) {
      if (view.getUint8(offset) !== 0xff) {
        offset += 1;
        continue;
      }
      while (offset < bytes.length && view.getUint8(offset) === 0xff) offset += 1;
      if (offset >= bytes.length) break;
      const marker = view.getUint8(offset);
      offset += 1;
      if (marker === 0x00 || marker === 0xd8 || marker === 0xd9) continue;
      if (!hasBytes(bytes, offset, 2)) break;
      const length = view.getUint16(offset);
      if (length < 2 || !hasBytes(bytes, offset, length)) break;
      if (
        (marker >= 0xc0 && marker <= 0xc3)
        || (marker >= 0xc5 && marker <= 0xc7)
        || (marker >= 0xc9 && marker <= 0xcb)
        || (marker >= 0xcd && marker <= 0xcf)
      ) {
        return dimensions(view.getUint16(offset + 5), view.getUint16(offset + 3));
      }
      offset += length;
    }
  }

  if (
    hasBytes(bytes, 0, 30)
    && bytes[0] === 0x52
    && bytes[1] === 0x49
    && bytes[2] === 0x46
    && bytes[3] === 0x46
    && bytes[8] === 0x57
    && bytes[9] === 0x45
    && bytes[10] === 0x42
    && bytes[11] === 0x50
  ) {
    const chunk = String.fromCharCode(
      view.getUint8(12),
      view.getUint8(13),
      view.getUint8(14),
      view.getUint8(15),
    );
    if (chunk === 'VP8X' && hasBytes(bytes, 0, 30)) {
      const width = 1 + view.getUint8(24) + (view.getUint8(25) << 8) + (view.getUint8(26) << 16);
      const height = 1 + view.getUint8(27) + (view.getUint8(28) << 8) + (view.getUint8(29) << 16);
      return dimensions(width, height);
    }
    if (chunk === 'VP8 ' && hasBytes(bytes, 23, 7)
      && view.getUint8(23) === 0x9d && view.getUint8(24) === 0x01 && view.getUint8(25) === 0x2a) {
      return dimensions(view.getUint16(26, true) & 0x3fff, view.getUint16(28, true) & 0x3fff);
    }
    if (chunk === 'VP8L' && hasBytes(bytes, 20, 5) && view.getUint8(20) === 0x2f) {
      const width = 1 + view.getUint8(21) + ((view.getUint8(22) & 0x3f) << 8);
      const height = 1 + (view.getUint8(22) >> 6) + (view.getUint8(23) << 2) + ((view.getUint8(24) & 0x0f) << 10);
      return dimensions(width, height);
    }
  }

  return null;
}
