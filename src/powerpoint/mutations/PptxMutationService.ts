import { debugLog, errorLog } from '../../logger';
import type { PresentationEngine } from '../../PresentationEngine';
import type { PptxCommand } from '../commands/types';
import type { MutationExecutor } from '../session/PresentationSession';

type MutationCommand = Exclude<PptxCommand, { type: 'noop' }>;
type EngineProvider = PresentationEngine | (() => PresentationEngine | null);

/**
 * Commands whose engine mutation only rewrites a single slide's XML part.
 *
 * These can roll back from a cheap synchronous `getSlideXml`/`restoreSlideXml`
 * pair instead of a full-deck `export()` snapshot (megabytes on image-heavy
 * decks), which is the dominant per-keystroke cost for inline text editing.
 * Anything that touches slide structure, ordering, media, relationships, or
 * charts must keep the lossless package snapshot.
 */
function slideLocalRollbackSlideIndex(command: MutationCommand): number | null {
  switch (command.type) {
    case 'update-paragraph-text':
    case 'split-paragraph':
    case 'remove-empty-preceding-paragraph':
    case 'merge-preceding-paragraph':
    case 'update-text-run':
      return command.slideIndex;
    default:
      return null;
  }
}

/**
 * The sole command-to-engine mutation boundary.
 *
 * A mutation snapshots enough state to undo itself, delegates the intent to the
 * engine, then commits renderer/pending-slide state back into the authoritative
 * package. Slide-local text edits snapshot only the touched slide's XML; every
 * other command exports the whole lossless package. Failed work restores from
 * whichever snapshot it took.
 */
export class PptxMutationService implements MutationExecutor {
  /** Engine export/reload transactions must not overlap. */
  private mutationQueue: Promise<void> = Promise.resolve();
  private queuedMutationCount = 0;

  constructor(private readonly engineProvider: EngineProvider) {}

  async execute(command: MutationCommand): Promise<unknown> {
    const engine = this.resolveEngine();
    const queuedAt = Date.now();
    const queueDepth = this.queuedMutationCount;
    this.queuedMutationCount += 1;
    if (queueDepth > 0) {
      debugLog('mutate', 'PowerPoint mutation queued', { op: command.type, queueDepth });
    }

    const executeTransaction = async (): Promise<unknown> => {
      try {
        return await this.executeTransaction(engine, command, Date.now() - queuedAt, queueDepth);
      } finally {
        this.queuedMutationCount -= 1;
      }
    };
    const transaction = this.mutationQueue.then(executeTransaction, executeTransaction);
    // A failed transaction must not prevent the next independent user action.
    this.mutationQueue = transaction.then(() => undefined, () => undefined);
    return transaction;
  }

  private async executeTransaction(
    engine: PresentationEngine,
    command: MutationCommand,
    queueMs: number,
    queueDepth: number,
  ): Promise<unknown> {
    const startedAt = Date.now();
    debugLog('mutate', 'PowerPoint mutation started', {
      op: command.type,
      queueMs,
      queueDepth,
    });

    // Slide-local text edits roll back from a cheap per-slide XML snapshot; only
    // fall back to the full-deck export when the engine lacks slide-XML access or
    // the command can affect more than one slide's part.
    const rollbackSlide = slideLocalRollbackSlideIndex(command);
    const useSlideXmlRollback = rollbackSlide !== null
      && typeof engine.getSlideXml === 'function'
      && typeof engine.restoreSlideXml === 'function';
    const slideXmlSnapshot = useSlideXmlRollback && rollbackSlide !== null
      ? engine.getSlideXml(rollbackSlide)
      : null;
    const snapshot = useSlideXmlRollback ? null : await engine.export();

    // The same slide-local text ops that roll back from slide XML also already
    // pushed their result into the renderer model (via `commitSlideDoc`) and
    // recorded pending lossless slide XML. The commit itself is a no-op: folding
    // pending into `currentBuffer` is deferred until a buffer reader (reorder /
    // insert / export) needs it. Eager zip-sync was the residual ~100ms+
    // keystroke tax on image-heavy decks.
    const useSlideLocalCommit = useSlideXmlRollback
      && typeof engine.commitSlideLocalMutation === 'function';

    try {
      const result = await this.apply(engine, command);
      if (useSlideLocalCommit) {
        await engine.commitSlideLocalMutation();
      } else {
        await engine.commitMutation();
      }
      debugLog('mutate', 'PowerPoint mutation committed', {
        op: command.type,
        ms: Date.now() - startedAt,
        rollback: useSlideXmlRollback ? 'slide-xml' : 'snapshot',
        commit: useSlideLocalCommit ? 'slide-local-deferred' : 'full-export',
      });
      return result;
    } catch (error) {
      try {
        if (useSlideXmlRollback && slideXmlSnapshot !== null && rollbackSlide !== null) {
          await engine.restoreSlideXml(rollbackSlide, slideXmlSnapshot);
        } else if (snapshot) {
          await engine.restoreSnapshot(snapshot);
        }
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

  private apply(engine: PresentationEngine, command: MutationCommand): unknown {
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
        return engine.insertTextBox(command.slideIndex, command.origin);
      case 'insert-table':
        return engine.addTable(command.slideIndex, command.rows, command.cols);
      case 'insert-chart':
        return engine.addChart(command.slideIndex);
      case 'update-shape-transform':
        return engine.updateShapeTransform(command.slideIndex, command.shapeIndex, command.transform);
      case 'delete-shape':
        return engine.deleteShape(command.slideIndex, command.shapeIndex);
      case 'reorder-shapes':
        return engine.reorderShapes(command.slideIndex, command.shapeIndexes, command.mode, {
          intersectingOnly: command.intersectingOnly,
        });
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
      case 'split-paragraph':
        return engine.splitParagraph(
          command.slideIndex,
          command.shapeIndex,
          command.paragraphIndex,
          command.splitOffset,
          command.text,
        );
      case 'remove-empty-preceding-paragraph':
        return engine.removeEmptyPrecedingParagraph(
          command.slideIndex,
          command.shapeIndex,
          command.paragraphIndex,
        );
      case 'merge-preceding-paragraph':
        return engine.mergePrecedingParagraph(
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
