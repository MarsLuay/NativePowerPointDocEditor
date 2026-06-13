// Static configuration values for the PowerPoint view: text-toolbar options,
// rendering selectors, history/zoom/snap limits, and external links. Extracted
// from NativePowerPointView.ts to keep tunables in one discoverable place.

export const TEXT_TOOLBAR_FONTS = [
  'Arial',
  'Calibri',
  'Cambria',
  'Comic Sans MS',
  'Courier New',
  'Georgia',
  'Garamond',
  'Impact',
  'Tahoma',
  'Times New Roman',
  'Trebuchet MS',
  'Verdana'
];

export const TEXT_TOOLBAR_SWATCHES = [
  '000000', '434343', '666666', '999999', 'B7B7B7', 'CCCCCC', 'D9D9D9', 'FFFFFF',
  '980000', 'FF0000', 'FF9900', 'FFFF00', '00FF00', '00FFFF', '4A86E8', '0000FF',
  '9900FF', 'FF00FF', 'E6B8AF', 'FCE5CD', 'FFF2CC', 'D9EAD3', 'D0E0E3', 'C9DAF8'
];

export const TEXT_TOOLBAR_MIN_FONT_SIZE = 1;
export const TEXT_TOOLBAR_MAX_FONT_SIZE = 400;

export const GENERATED_GRID_SELECTOR =
  'g[data-ooxml-shape-type="table"], g[data-ooxml-shape-type="chart"]';

export const HISTORY_LIMIT = 20;
export const MIN_ZOOM = 0.25;
export const MAX_ZOOM = 3;
export const SNAP_THRESHOLD_PX = 6;

export const OBSIDIAN_DOWNLOAD_URL = 'https://obsidian.md/download';
