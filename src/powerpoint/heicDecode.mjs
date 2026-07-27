// Loading this module must stay side-effect free: libheif detects Node and
// reads `__dirname` during initialization, while Obsidian loads this artifact
// as ESM. Defer its initialization until a HEIC image actually needs decoding.
export default async function decodeHeic(input) {
  const { default: decode } = await import('heic-decode');
  return await decode(input);
}
