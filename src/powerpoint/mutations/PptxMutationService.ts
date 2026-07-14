import { debugLog, errorLog } from '../../logger';
import type { PresentationEngine } from '../../PresentationEngine';
import type { PptxCommand } from '../commands/types';
import type { MutationExecutor } from '../session/PresentationSession';

type MutationCommand = Exclude<PptxCommand, { type: 'noop' }>;
type EngineProvider = PresentationEngine | (() => PresentationEngine | null);

/**
 * The sole command-to-engine mutation boundary.
 *
 * A mutation starts from an exported lossless snapshot, delegates the intent to
 * the engine, then commits renderer/pending-slide state back into the
 * authoritative package. Failed work restores that snapshot.
 */
export class PptxMutationService implements MutationExecutor {
  constructor(private readonly engineProvider: EngineProvider) {}

  async execute(command: MutationCommand): Promise<unknown> {
    const engine = this.resolveEngine();
    const startedAt = Date.now();
    debugLog('mutate', 'PowerPoint mutation started', { op: command.type });
    const snapshot = await engine.export();

    try {
      const result = await this.apply(engine, command);
      await engine.commitMutation();
      debugLog('mutate', 'PowerPoint mutation committed', {
        op: command.type,
        ms: Date.now() - startedAt,
      });
      return result;
    } catch (error) {
      try {
        await engine.restoreSnapshot(snapshot);
      } catch (rollbackError) {
        errorLog('mutate', 'PowerPoint mutation rollback failed', {
          op: command.type,
          error: rollbackError,
        });
      }
      errorLog('mutate', 'PowerPoint mutation failed', {
        op: command.type,
        error,
      });
      throw error;
    }
  }

  private resolveEngine(): PresentationEngine {
    const engine = typeof this.engineProvider === 'function'
      ? this.engineProvider()
      : this.engineProvider;
    if (!engine) throw new Error('PowerPoint mutation requested before the presentation engine loaded.');
    return engine;
  }

  private apply(engine: PresentationEngine, command: MutationCommand): Promise<unknown> | unknown {
    switch (command.type) {
      case 'add-slide':
        return engine.addSlide(command.afterIndex);
      case 'add-slide-with-layout':
        return engine.addSlideWithLayout(command.afterIndex, command.layout);
      case 'delete-slide':
        return engine.deleteSlide(command.slideIndex);
      case 'move-slide':
        return engine.moveSlide(command.slideIndex, command.direction);
      case 'duplicate-slide':
        return engine.duplicateSlide(command.slideIndex);
      case 'reorder-slides':
        return engine.reorderSlides(command.newOrder);
      case 'insert-image':
        return engine.insertImage(
          command.slideIndex,
          command.imageData,
          command.mimeType,
          command.widthPx,
          command.heightPx,
        );
      case 'insert-shape':
        return engine.insertShapeGeometry(command.slideIndex, command.geometry);
      case 'insert-text-box':
        return engine.insertTextBox(command.slideIndex);
      case 'insert-table':
        return engine.addTable(command.slideIndex, command.rows, command.cols);
      case 'insert-chart':
        return engine.addChart(command.slideIndex);
      case 'update-shape-transform':
        return engine.updateShapeTransform(command.slideIndex, command.shapeIndex, command.transform);
      case 'delete-shape':
        return engine.deleteShape(command.slideIndex, command.shapeIndex);
      case 'reorder-shapes':
        return engine.reorderShapes(command.slideIndex, command.shapeIndexes, command.mode);
      case 'group-shapes':
        return engine.groupShapes(command.slideIndex, command.shapeIndexes);
      case 'ungroup-shapes':
        return engine.ungroupShapes(command.slideIndex, command.shapeIndex);
      case 'duplicate-shape':
        return engine.duplicateShape(command.slideIndex, command.shapeIndex);
      case 'paste-shape':
        return engine.pasteShape(command.clipboard, command.destinationSlideIndex);
      case 'update-shape-text':
        return engine.updateShapeText(command.slideIndex, command.shapeIndex, command.text);
      case 'update-paragraph-text':
        return engine.updateParagraphText(
          command.slideIndex,
          command.shapeIndex,
          command.paragraphIndex,
          command.text,
        );
      case 'update-text-run':
        return engine.updateTextRun(
          command.slideIndex,
          command.shapeIndex,
          command.paragraphIndex,
          command.runIndex,
          command.text,
        );
      case 'replace-text':
        return engine.replaceText(command.query, command.replacement, {
          matchCase: command.matchCase,
          slideIndex: command.slideIndex,
          shapeIndex: command.shapeIndex,
        });
      case 'set-run-style':
        return engine.setRunStyle(command.slideIndex, command.shapeIndex, command.target, command.change);
      case 'set-run-style-range':
        return engine.setRunStyleForRange(
          command.slideIndex,
          command.shapeIndex,
          command.paragraphIndex,
          command.startOffset,
          command.endOffset,
          command.change,
        );
      case 'set-run-style-ranges':
        return engine.setRunStyleForRanges(command.slideIndex, command.shapeIndex, command.ranges, command.change);
      case 'set-paragraph-alignment':
        return engine.setParagraphAlignment(
          command.slideIndex,
          command.shapeIndex,
          command.paragraphIndex,
          command.align,
        );
      case 'set-paragraph-alignment-ranges':
        return engine.setParagraphAlignmentForRanges(
          command.slideIndex,
          command.shapeIndex,
          command.ranges,
          command.align,
        );
      case 'apply-list-style':
        return engine.applyListStyle(
          command.slideIndex,
          command.shapeIndex,
          command.paragraphIndex,
          command.style,
        );
      case 'set-slide-background-color':
        return engine.setSlideBackgroundColor(command.slideIndex, command.hex);
      case 'set-shape-fill-color':
        return engine.setShapeFillColor(command.slideIndex, command.shapeIndex, command.hex);
      case 'set-image-crop':
        return engine.setImageCrop(command.slideIndex, command.shapeIndex, command.crop);
      case 'reset-image':
        return engine.resetImage(command.slideIndex, command.shapeIndex);
      case 'flip-shape':
        return engine.flipShape(command.slideIndex, command.shapeIndex, command.axis);
      case 'replace-image':
        return engine.replaceImage(command.slideIndex, command.shapeIndex, command.bytes, command.mimeType);
      case 'update-chart-data':
        return engine.updateChartData(command.slideIndex, command.shapeIndex, command.update);
      case 'update-generated-text':
        return engine.updateGeneratedText(command.slideIndex, command.shapeIndex, command.edit, command.text);
    }
  }
}
