/**
 * Content-control (SDT) addressing — public surface without DocumentAgent.
 * @packageDocumentation
 * @public
 */

export {
  findContentControls,
  findContentControl,
  getContentControlText,
  setContentControlContent,
  removeContentControl,
  ContentControlNotFoundError,
  ContentControlLockedError,
  ContentControlTypeError,
  ContentControlBoundError,
  ContentControlKindError,
  type ContentControlFilter,
  type ContentControlInfo,
  type ContentControlLocation,
  type FindContentControlsOptions,
} from './contentControls';

export {
  createContentControl,
  ContentControlCreateError,
  type CreateContentControlTarget,
  type NewContentControlProps,
} from './createContentControl';

export {
  setContentControlValue,
  formatSdtDate,
  ContentControlValueError,
  type ContentControlValue,
} from './contentControlValues';

export {
  addRepeatingSectionItem,
  removeRepeatingSectionItem,
  isRepeatingSection,
  isRepeatingSectionItem,
  RepeatingSectionError,
} from './repeatingSection';

export {
  getParagraphText,
  getRunText,
  getHyperlinkText,
  getTableText,
  getBodyText,
  countWords,
  countCharacters,
  getBodyWordCount,
  getBodyCharacterCount,
  getTextBefore,
  getTextAfter,
  getFormattingAtPosition,
  isPositionInHyperlink,
  getHyperlinkAtPosition,
  isHeadingStyle,
  parseHeadingLevel,
  hasImages,
  hasHyperlinks,
  hasTables,
  getParagraphs,
  getParagraphAtIndex,
  getBlockIndexForParagraph,
} from './text-utils';
