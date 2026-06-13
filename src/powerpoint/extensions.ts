// PowerPoint file-extension classification and the view type identifier.
// Extracted from NativePowerPointView.ts so the public extension API lives in a
// small, dependency-free module that other code (and tests) can import directly.

export const NATIVE_POWERPOINT_VIEW_TYPE = 'native-powerpoint-view';

export const MODERN_POWERPOINT_EXTENSIONS = [
  'pptx',
  'pptm',
  'ppsx',
  'ppsm',
  'potx',
  'potm'
];

export const LEGACY_POWERPOINT_EXTENSIONS = ['ppt', 'pps', 'pot'];
export const MACRO_ENABLED_POWERPOINT_EXTENSIONS = ['pptm', 'ppsm', 'potm'];
export const EDITABLE_POWERPOINT_EXTENSIONS = ['pptx', 'ppsx', 'potx'];

export const POWERPOINT_EXTENSIONS = [
  ...MODERN_POWERPOINT_EXTENSIONS,
  ...LEGACY_POWERPOINT_EXTENSIONS
];

export function isPowerPointExtension(extension: string): boolean {
  return POWERPOINT_EXTENSIONS.includes(extension.toLowerCase());
}

export function isModernPowerPointExtension(extension: string): boolean {
  return MODERN_POWERPOINT_EXTENSIONS.includes(extension.toLowerCase());
}

export function isEditablePowerPointExtension(extension: string): boolean {
  return EDITABLE_POWERPOINT_EXTENSIONS.includes(extension.toLowerCase());
}

export function isMacroEnabledPowerPointExtension(extension: string): boolean {
  return MACRO_ENABLED_POWERPOINT_EXTENSIONS.includes(extension.toLowerCase());
}
