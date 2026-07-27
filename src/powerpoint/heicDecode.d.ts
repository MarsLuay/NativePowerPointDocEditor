import type { HeicDecodedImage, HeicDecodeInput } from 'heic-decode';

declare function decodeHeic(input: HeicDecodeInput): Promise<HeicDecodedImage>;

export default decodeHeic;
