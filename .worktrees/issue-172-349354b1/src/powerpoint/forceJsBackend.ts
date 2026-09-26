let testOverride = false;
let readDevOverride: (() => boolean) | null = null;

/** Headless tests and smoke scripts call this instead of touching globals. */
export function setForceJsBackendOverride(enabled: boolean): void {
  testOverride = enabled;
}

export function resetForceJsBackendOverride(): void {
  testOverride = false;
}

/** Wired from the plugin on load; devs can toggle via the plugin instance in the console. */
export function configureForceJsBackendOverrideReader(read: () => boolean): void {
  readDevOverride = read;
}

export function shouldForceJsBackend(): boolean {
  if (testOverride) {
    return true;
  }
  try {
    return readDevOverride?.() ?? false;
  } catch {
    return false;
  }
}
