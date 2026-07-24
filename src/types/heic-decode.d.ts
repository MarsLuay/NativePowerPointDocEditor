declare module 'heic-decode' {
  export interface HeicDecodedImage {
    width: number;
    height: number;
    data: Uint8ClampedArray;
  }

  export interface HeicDecodeInput {
    /** Must be Uint8Array/Buffer-like — ArrayBuffer breaks heic-decode's brand check. */
    buffer: ArrayBuffer | Uint8Array;
  }

  function decodeHeic(input: HeicDecodeInput): Promise<HeicDecodedImage>;
  export default decodeHeic;

  export function all(input: HeicDecodeInput): Promise<HeicDecodedImage[]>;
}
