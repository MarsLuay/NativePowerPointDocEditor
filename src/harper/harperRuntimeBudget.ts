export const NPDE_RUNTIME_FILE_LIMIT_BYTES = 5_000_000;
export const NPDE_MAIN_BUNDLE_LIMIT_BYTES = 5_010_000;
export const HARPER_JS_VERSION = '2.10.0';
export const HARPER_SLIM_WASM_FILE = 'harper_wasm_slim_bg.wasm';

export interface HarperRuntimeContribution {
  artifact: string;
  rawBytes: number;
  gzipBytes: number;
  mainBytes: number;
  withinFileLimit: boolean;
  embedFitsMain: boolean;
}

/**
 * Measure a Harper runtime file against the same limits used for PPTX/HEIC
 * sidecars and the gzip payload embedded in main.js.
 */
export function assessHarperRuntimeContribution(input: {
  rawBytes: number;
  gzipBytes: number;
  mainBytes: number;
}): HarperRuntimeContribution {
  const projectedMainBytes = input.mainBytes + input.gzipBytes;
  return {
    artifact: HARPER_SLIM_WASM_FILE,
    rawBytes: input.rawBytes,
    gzipBytes: input.gzipBytes,
    mainBytes: input.mainBytes,
    withinFileLimit: input.rawBytes <= NPDE_RUNTIME_FILE_LIMIT_BYTES
      && input.gzipBytes <= NPDE_RUNTIME_FILE_LIMIT_BYTES,
    embedFitsMain: projectedMainBytes <= NPDE_MAIN_BUNDLE_LIMIT_BYTES,
  };
}
