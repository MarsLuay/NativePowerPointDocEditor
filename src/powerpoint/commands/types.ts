import type { ShapeTransform } from 'pptx-svg';
import type { ChartDataUpdate } from '../../ChartData';
import type { SlideObjectClipboard } from '../../ShapeClipboard';
import type { InsertableChartType, ParagraphListStyle } from '../../SlideInsertions';
import type {
  GeneratedTextEdit,
  ImageCrop,
  InsertableShapeGeometry,
  ParagraphAlignment,
  ParagraphTextRange,
  RunStyleChange,
  RunTarget,
  ShapeReorderMode,
  SlideLayoutKind,
  TextBoxInsertOrigin,
} from '../../PresentationEngine';

/** Intent-level commands accepted by the presentation mutation boundary. */
export type PptxCommand =
  | {
      readonly type: 'noop';
      readonly reason?: string;
    }
  | {
      readonly type: 'add-slide';
      readonly afterIndex: number;
    }
  | {
      readonly type: 'add-slide-with-layout';
      readonly afterIndex: number;
      readonly layout: SlideLayoutKind;
    }
  | { readonly type: 'delete-slide'; readonly slideIndex: number }
  | { readonly type: 'move-slide'; readonly slideIndex: number; readonly direction: -1 | 1 }
  | { readonly type: 'duplicate-slide'; readonly slideIndex: number }
  | { readonly type: 'reorder-slides'; readonly newOrder: number[] }
  | {
      readonly type: 'insert-image';
      readonly slideIndex: number;
      readonly imageData: Uint8Array;
      readonly mimeType: string;
      readonly widthPx?: number;
      readonly heightPx?: number;
    }
  | { readonly type: 'insert-shape'; readonly slideIndex: number; readonly geometry: InsertableShapeGeometry }
  | { readonly type: 'insert-text-box'; readonly slideIndex: number; readonly origin?: TextBoxInsertOrigin }
  | { readonly type: 'insert-table'; readonly slideIndex: number; readonly rows: number; readonly cols: number }
  | { readonly type: 'insert-chart'; readonly slideIndex: number; readonly chartType?: InsertableChartType }
  | {
      readonly type: 'update-shape-transform';
      readonly slideIndex: number;
      readonly shapeIndex: number;
      readonly transform: ShapeTransform;
    }
  | { readonly type: 'delete-shape'; readonly slideIndex: number; readonly shapeIndex: number }
  | {
      readonly type: 'reorder-shapes';
      readonly slideIndex: number;
      readonly shapeIndexes: number[];
      readonly mode: ShapeReorderMode;
      /** UI reorders stop at the nearest overlapping object. */
      readonly intersectingOnly?: boolean;
    }
  | { readonly type: 'group-shapes'; readonly slideIndex: number; readonly shapeIndexes: number[] }
  | { readonly type: 'ungroup-shapes'; readonly slideIndex: number; readonly shapeIndex: number }
  | { readonly type: 'duplicate-shape'; readonly slideIndex: number; readonly shapeIndex: number }
  | {
      readonly type: 'paste-shape';
      readonly clipboard: SlideObjectClipboard;
      readonly destinationSlideIndex: number;
    }
  | {
      readonly type: 'update-shape-text';
      readonly slideIndex: number;
      readonly shapeIndex: number;
      readonly text: string;
    }
  | {
      readonly type: 'update-paragraph-text';
      readonly slideIndex: number;
      readonly shapeIndex: number;
      readonly paragraphIndex: number;
      readonly text: string;
    }
  | {
      readonly type: 'split-paragraph';
      readonly slideIndex: number;
      readonly shapeIndex: number;
      readonly paragraphIndex: number;
      readonly splitOffset: number;
      readonly text?: string;
    }
  | {
      readonly type: 'remove-empty-preceding-paragraph';
      readonly slideIndex: number;
      readonly shapeIndex: number;
      readonly paragraphIndex: number;
    }
  | {
      readonly type: 'merge-preceding-paragraph';
      readonly slideIndex: number;
      readonly shapeIndex: number;
      readonly paragraphIndex: number;
      readonly text?: string;
    }
  | {
      readonly type: 'update-text-run';
      readonly slideIndex: number;
      readonly shapeIndex: number;
      readonly paragraphIndex: number;
      readonly runIndex: number;
      readonly text: string;
    }
  | {
      readonly type: 'replace-text';
      readonly query: string;
      readonly replacement: string;
      readonly matchCase?: boolean;
      readonly slideIndex?: number;
      readonly shapeIndex?: number;
    }
  | {
      readonly type: 'set-run-style';
      readonly slideIndex: number;
      readonly shapeIndex: number;
      readonly target: RunTarget | null;
      readonly change: RunStyleChange;
    }
  | {
      readonly type: 'set-run-style-range';
      readonly slideIndex: number;
      readonly shapeIndex: number;
      readonly paragraphIndex: number;
      readonly startOffset: number;
      readonly endOffset: number;
      readonly change: RunStyleChange;
    }
  | {
      readonly type: 'set-run-style-ranges';
      readonly slideIndex: number;
      readonly shapeIndex: number;
      readonly ranges: ParagraphTextRange[];
      readonly change: RunStyleChange;
    }
  | {
      readonly type: 'set-paragraph-alignment';
      readonly slideIndex: number;
      readonly shapeIndex: number;
      readonly paragraphIndex: number | null;
      readonly align: ParagraphAlignment;
    }
  | {
      readonly type: 'set-paragraph-alignment-ranges';
      readonly slideIndex: number;
      readonly shapeIndex: number;
      readonly ranges: ParagraphTextRange[];
      readonly align: ParagraphAlignment;
    }
  | {
      readonly type: 'apply-list-style';
      readonly slideIndex: number;
      readonly shapeIndex: number;
      readonly paragraphIndex: number;
      readonly style: ParagraphListStyle;
      /** Remove an imported literal bullet when the toolbar converts/toggles it. */
      readonly stripLeadingManualBullet?: boolean;
    }
  | {
      readonly type: 'apply-list-style-range';
      readonly slideIndex: number;
      readonly shapeIndex: number;
      readonly range: ParagraphTextRange;
      readonly style: ParagraphListStyle;
      /** Remove an imported literal bullet when the toolbar converts/toggles it. */
      readonly stripLeadingManualBullet?: boolean;
    }
  | {
      readonly type: 'apply-list-style-ranges';
      readonly slideIndex: number;
      readonly shapeIndex: number;
      readonly ranges: ParagraphTextRange[];
      readonly style: ParagraphListStyle;
      /** Remove imported literal bullets only where a selected range starts at the paragraph head. */
      readonly stripLeadingManualBullet?: boolean;
    }
  | { readonly type: 'set-slide-background-color'; readonly slideIndex: number; readonly hex: string }
  | {
      readonly type: 'set-shape-fill-color';
      readonly slideIndex: number;
      readonly shapeIndex: number;
      readonly hex: string;
    }
  | {
      readonly type: 'set-image-crop';
      readonly slideIndex: number;
      readonly shapeIndex: number;
      readonly crop: ImageCrop;
    }
  | { readonly type: 'reset-image'; readonly slideIndex: number; readonly shapeIndex: number }
  | {
      readonly type: 'flip-shape';
      readonly slideIndex: number;
      readonly shapeIndex: number;
      readonly axis: 'horizontal' | 'vertical';
    }
  | {
      readonly type: 'replace-image';
      readonly slideIndex: number;
      readonly shapeIndex: number;
      readonly bytes: Uint8Array;
      readonly mimeType: string;
    }
  | {
      readonly type: 'update-chart-data';
      readonly slideIndex: number;
      readonly shapeIndex: number;
      readonly update: ChartDataUpdate;
    }
  | {
      readonly type: 'update-generated-text';
      readonly slideIndex: number;
      readonly shapeIndex: number;
      readonly edit: GeneratedTextEdit;
      readonly text: string;
    };
