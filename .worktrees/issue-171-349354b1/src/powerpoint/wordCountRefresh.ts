import type { PptxCommand } from './commands/types';

/**
 * Whether a completed command can change the text represented by the document
 * word count. Styling-only commands intentionally return false: calculating
 * the total count renders every slide, which is prohibitively expensive for
 * image-heavy posters.
 */
export function commandMayChangePresentationWordCount(command: PptxCommand): boolean {
  switch (command.type) {
    case 'add-slide':
    case 'add-slide-with-layout':
    case 'delete-slide':
    case 'duplicate-slide':
    case 'insert-shape':
    case 'insert-text-box':
    case 'insert-table':
    case 'insert-chart':
    case 'delete-shape':
    case 'duplicate-shape':
    case 'paste-shape':
    case 'update-shape-text':
    case 'update-paragraph-text':
    case 'split-paragraph':
    case 'remove-empty-preceding-paragraph':
    case 'merge-preceding-paragraph':
    case 'update-text-run':
    case 'replace-text':
    case 'update-chart-data':
    case 'update-generated-text':
      return true;
    default:
      return false;
  }
}
