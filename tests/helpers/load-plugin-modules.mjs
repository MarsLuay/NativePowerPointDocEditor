import { mkdtemp } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import Module, { createRequire } from "node:module";
import { fileURLToPath } from "node:url";
import { build } from "esbuild";
import { createInlinePptxSvgWasmPlugin } from "../../scripts/lib/patch-pptx-renderer.mjs";
import { createPptxRuntimeArtifactResolver } from "../../scripts/lib/pptx-runtime-artifact-test-loader.mjs";
import {
  createDocxEditorAliases,
  resolveDocxEditorPackagesRoot,
} from "../../scripts/lib/docx-editor-aliases.mjs";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const docxEditorAliases = await createDocxEditorAliases(
  resolveDocxEditorPackagesRoot(projectRoot),
);
const require = createRequire(import.meta.url);
const { DOMParser, XMLSerializer } = require("@xmldom/xmldom");
let tempDirectoryPromise;
let packageModulePromise;
let presentationEngineModulePromise;
let shapeClipboardModulePromise;
let viewModulePromise;
let docxTextExtractorModulePromise;
let docxReviewMarkupModulePromise;
let docxHiddenTextScannerModulePromise;
let docxStyleDefaultsModulePromise;
let docxParagraphLayoutRelayoutModulePromise;
let docxPlainTextInsertModulePromise;
let docxTableCellFontSizePreserverModulePromise;
let docxFloatingLayerLayoutModulePromise;
let docxEditorChromeMarkersModulePromise;
let docxSessionModulePromise;
let fakeDocxEditorAdapterModulePromise;
let docxToolbarTooltipModulePromise;
let markdownToDocxModulePromise;
let tooltipControllerModulePromise;
let powerPointToolbarTooltipTargetModulePromise;
let loggerModulePromise;
let parseRenderedSlideSvgModulePromise;
let textToolbarControllerModulePromise;
let insertControllerModulePromise;
let arrangeControllerModulePromise;
let slideExtensionPreserveModulePromise;

globalThis.DOMParser ??= DOMParser;
globalThis.XMLSerializer ??= XMLSerializer;

const inlinePptxSvgWasmPlugin = createInlinePptxSvgWasmPlugin();

export { createPptxRuntimeArtifactResolver };

function getTempDirectory() {
  tempDirectoryPromise ??= mkdtemp(path.join(tmpdir(), "native-powerpoint-tests-"));
  return tempDirectoryPromise;
}

export async function bundleSource(entry, outputName, external = [], plugins = []) {
  const outputDirectory = await getTempDirectory();
  const outfile = path.join(outputDirectory, outputName);
  await build({
    alias: docxEditorAliases,
    entryPoints: [path.join(projectRoot, entry)],
    bundle: true,
    external,
    format: "cjs",
    logLevel: "silent",
    outfile,
    platform: "node",
    plugins,
    target: "node22",
  });
  return outfile;
}

let textUtilsModulePromise;

const stubObsidianPlugin = {
  name: "stub-obsidian",
  setup(buildContext) {
    buildContext.onResolve({ filter: /^obsidian$/ }, () => ({ path: "obsidian", namespace: "stub-obsidian" }));
    buildContext.onLoad({ filter: /.*/, namespace: "stub-obsidian" }, () => ({
      contents: `
        export const Platform = { isDesktop: true, isMacOS: false, isMobile: false, isMobileApp: false };
        export const getLanguage = () => "en";
        export const normalizePath = (value) => value.replace(/\\\\/g, "/").replace(/\\/{2,}/g, "/");
        export const setIcon = () => {};
        export class Notice { constructor() {} }
        export class Modal { constructor() {} open() {} close() {} }
        export class Component {
          constructor() { this.cleanups = []; }
          register(cleanup) { this.cleanups.push(cleanup); }
          load() { this.onload?.(); }
          unload() {
            this.onunload?.();
            for (const cleanup of this.cleanups.splice(0)) cleanup();
          }
        }
      `,
      loader: "js",
    }));
  },
};

let annotateTextOffsetsModulePromise;

export function loadAnnotateTextOffsetsModule() {
  annotateTextOffsetsModulePromise ??= (async () => {
    const outputDirectory = await getTempDirectory();
    const outfile = path.join(outputDirectory, "annotate-text-offsets.cjs");
    await build({
      entryPoints: [path.join(projectRoot, "src/powerpoint/annotateTextOffsets.ts")],
      bundle: true,
      format: "cjs",
      logLevel: "silent",
      outfile,
      platform: "node",
      plugins: [stubObsidianPlugin],
      target: "node22",
    });
    return require(outfile);
  })();
  return annotateTextOffsetsModulePromise;
}

export function loadTextUtilsModule() {
  textUtilsModulePromise ??= (async () => {
    const outputDirectory = await getTempDirectory();
    const outfile = path.join(outputDirectory, "text-utils.cjs");
    await build({
      entryPoints: [path.join(projectRoot, "src/powerpoint/textUtils.ts")],
      bundle: true,
      format: "cjs",
      logLevel: "silent",
      outfile,
      platform: "node",
      plugins: [stubObsidianPlugin],
      target: "node22",
    });
    return require(outfile);
  })();
  return textUtilsModulePromise;
}

export function loadTextToolbarControllerModule() {
  textToolbarControllerModulePromise ??= bundleSource(
    "src/powerpoint/textToolbarController.ts",
    "text-toolbar-controller.cjs",
    [],
    [stubObsidianPlugin],
  ).then((outfile) => require(outfile));
  return textToolbarControllerModulePromise;
}

export function loadInsertControllerModule() {
  insertControllerModulePromise ??= bundleSource(
    "src/powerpoint/insertController.ts",
    "insert-controller.cjs",
    [],
    [stubObsidianPlugin],
  ).then((outfile) => require(outfile));
  return insertControllerModulePromise;
}

export function loadArrangeControllerModule() {
  arrangeControllerModulePromise ??= bundleSource(
    "src/powerpoint/arrangeController.ts",
    "arrange-controller.cjs",
    [],
    [stubObsidianPlugin],
  ).then((outfile) => require(outfile));
  return arrangeControllerModulePromise;
}

let inlineTextGeometryModulePromise;

export function loadInlineTextGeometryModule() {
  inlineTextGeometryModulePromise ??= bundleSource(
    "src/powerpoint/inlineTextGeometry.ts",
    "inline-text-geometry.cjs",
  ).then((outfile) => require(outfile));
  return inlineTextGeometryModulePromise;
}

export function loadParseRenderedSlideSvgModule() {
  parseRenderedSlideSvgModulePromise ??= bundleSource(
    "src/powerpoint/parseRenderedSlideSvg.ts",
    "parse-rendered-slide-svg.cjs",
  ).then((outfile) => require(outfile));
  return parseRenderedSlideSvgModulePromise;
}

export function loadPowerPointPackageModule() {
  packageModulePromise ??= bundleSource("src/PowerPointPackage.ts", "powerpoint-package.cjs").then(
    (outfile) => require(outfile),
  );
  return packageModulePromise;
}

export function loadSlideExtensionPreserveModule() {
  slideExtensionPreserveModulePromise ??= bundleSource(
    "src/powerpoint/slideExtensionPreserve.ts",
    "slide-extension-preserve.cjs",
  ).then((outfile) => require(outfile));
  return slideExtensionPreserveModulePromise;
}

export function loadPresentationEngineModule() {
  presentationEngineModulePromise ??= (async () => {
    const outputDirectory = await getTempDirectory();
    const resolveRuntimeArtifact = await createPptxRuntimeArtifactResolver({
      projectRoot,
      outputDirectory,
    });
    const outfile = path.join(outputDirectory, "presentation-engine.cjs");
    await build({
      entryPoints: [path.join(projectRoot, "src/PresentationEngine.ts")],
      bundle: true,
      format: "cjs",
      loader: { ".wasm": "binary" },
      logLevel: "silent",
      outfile,
      platform: "node",
      plugins: [inlinePptxSvgWasmPlugin],
      target: "node22",
    });
    const module = require(outfile);
    module.configurePptxRuntimeArtifactLoader(resolveRuntimeArtifact);
    return module;
  })();
  return presentationEngineModulePromise;
}

export function loadShapeClipboardModule() {
  shapeClipboardModulePromise ??= (async () => {
    const outputDirectory = await getTempDirectory();
    const outfile = path.join(outputDirectory, "shape-clipboard.cjs");
    await build({
      entryPoints: [path.join(projectRoot, "src/ShapeClipboard.ts")],
      bundle: true,
      format: "cjs",
      loader: { ".wasm": "binary" },
      logLevel: "silent",
      outfile,
      platform: "node",
      plugins: [inlinePptxSvgWasmPlugin],
      target: "node22",
    });
    return require(outfile);
  })();
  return shapeClipboardModulePromise;
}

export function loadDocxTextExtractorModule() {
  docxTextExtractorModulePromise ??= bundleSource("src/docxTextExtractor.ts", "docx-text-extractor.cjs").then(
    (outfile) => require(outfile),
  );
  return docxTextExtractorModulePromise;
}

export function loadDocxReviewMarkupModule() {
  docxReviewMarkupModulePromise ??= bundleSource("src/docxReviewMarkup.ts", "docx-review-markup.cjs").then(
    (outfile) => require(outfile),
  );
  return docxReviewMarkupModulePromise;
}

export function loadDocxHiddenTextScannerModule() {
  docxHiddenTextScannerModulePromise ??= bundleSource("src/docxHiddenTextScanner.ts", "docx-hidden-text-scanner.cjs").then(
    (outfile) => require(outfile),
  );
  return docxHiddenTextScannerModulePromise;
}

export function loadDocxStyleDefaultsModule() {
  docxStyleDefaultsModulePromise ??= bundleSource("src/docxStyleDefaults.ts", "docx-style-defaults.cjs").then(
    (outfile) => require(outfile),
  );
  return docxStyleDefaultsModulePromise;
}

export function loadDocxParagraphLayoutRelayoutModule() {
  docxParagraphLayoutRelayoutModulePromise ??= bundleSource(
    "src/docxParagraphLayoutRelayout.ts",
    "docx-paragraph-layout-relayout.cjs",
    ["prosemirror-model"],
  ).then((outfile) => require(outfile));
  return docxParagraphLayoutRelayoutModulePromise;
}

export function loadDocxPlainTextInsertModule() {
  docxPlainTextInsertModulePromise ??= bundleSource(
    "src/docxPlainTextInsert.ts",
    "docx-plain-text-insert.cjs",
    [],
  ).then((outfile) => require(outfile));
  return docxPlainTextInsertModulePromise;
}

export function loadDocxTableCellFontSizePreserverModule() {
  docxTableCellFontSizePreserverModulePromise ??= bundleSource(
    "src/docxTableCellFontSizePreserver.ts",
    "docx-table-cell-font-size-preserver.cjs",
  ).then((outfile) => require(outfile));
  return docxTableCellFontSizePreserverModulePromise;
}

export function loadDocxFloatingLayerLayoutModule() {
  docxFloatingLayerLayoutModulePromise ??= bundleSource(
    "src/docxFloatingLayerLayout.ts",
    "docx-floating-layer-layout.cjs",
  ).then((outfile) => require(outfile));
  return docxFloatingLayerLayoutModulePromise;
}

export function loadDocxEditorChromeMarkersModule() {
  docxEditorChromeMarkersModulePromise ??= bundleSource(
    "src/docxEditorChromeMarkers.ts",
    "docx-editor-chrome-markers.cjs",
  ).then((outfile) => require(outfile));
  return docxEditorChromeMarkersModulePromise;
}

export function loadDocxSessionModule() {
	docxSessionModulePromise ??= bundleSource(
		"src/docx/session/DocxSession.ts",
		"docx-session.cjs",
	).then((outfile) => require(outfile));
	return docxSessionModulePromise;
}

export function loadFakeDocxEditorAdapterModule() {
	fakeDocxEditorAdapterModulePromise ??= bundleSource(
		"src/docx/adapter/FakeDocxEditorAdapter.ts",
		"fake-docx-editor-adapter.cjs",
	).then((outfile) => require(outfile));
	return fakeDocxEditorAdapterModulePromise;
}

export function loadDocxToolbarTooltipModule() {
  docxToolbarTooltipModulePromise ??= bundleSource(
    "src/docxToolbarTooltip.ts",
    "docx-toolbar-tooltip.cjs",
    [],
    [stubObsidianPlugin],
  ).then((outfile) => require(outfile));
  return docxToolbarTooltipModulePromise;
}

export function loadMarkdownToDocxModule() {
  markdownToDocxModulePromise ??= bundleSource(
    "src/vault/markdownToDocx.ts",
    "markdown-to-docx.cjs",
    [],
    [stubObsidianPlugin],
  ).then((outfile) => require(outfile));
  return markdownToDocxModulePromise;
}

export function loadTooltipControllerModule() {
  tooltipControllerModulePromise ??= bundleSource(
    "src/ui/TooltipController.ts",
    "tooltip-controller.cjs",
  ).then((outfile) => require(outfile));
  return tooltipControllerModulePromise;
}

export function loadPowerPointToolbarTooltipTargetModule() {
  powerPointToolbarTooltipTargetModulePromise ??= bundleSource(
    "src/powerpoint/toolbarTooltipTarget.ts",
    "powerpoint-toolbar-tooltip-target.cjs",
  ).then((outfile) => require(outfile));
  return powerPointToolbarTooltipTargetModulePromise;
}

export function loadLoggerModule() {
  loggerModulePromise ??= bundleSource("src/logger.ts", "logger.cjs").then((outfile) => require(outfile));
  return loggerModulePromise;
}

export function loadNativePowerPointViewModule() {
  viewModulePromise ??= bundleSource(
    "src/NativePowerPointView.ts",
    "native-powerpoint-view.cjs",
    ["obsidian", "pptx-svg", "pptx-svg/wasm"],
  ).then((outfile) => {
    const notices = [];
    const originalLoad = Module._load;
    globalThis.activeDocument ??= {
      activeElement: null,
      addEventListener() {},
      body: { classList: { add() {}, remove() {}, toggle() {} } },
      removeEventListener() {},
    };

    class Notice {
      constructor(message, duration) {
        notices.push({ message, duration });
      }
    }

    class Component {
      load() {}
      unload() {}
      onload() {}
      onunload() {}
      register() {
        return () => {};
      }
      registerDomEvent() {
        return () => {};
      }
      registerInterval() {
        return () => {};
      }
      addChild() {
        return this;
      }
      removeChild() {}
    }

    class FileView extends Component {
      constructor(leaf = {}) {
        super();
        this.leaf = leaf;
        this.app = leaf.app ?? { vault: {} };
        this.file = null;
        this.contentEl = leaf.contentEl ?? createElementStub();
        this.containerEl = leaf.containerEl ?? { isShown: () => true };
      }
    }

    class Menu {
      addItem(callback) {
        callback?.({
          onClick() { return this; },
          setDisabled() { return this; },
          setIcon() { return this; },
          setTitle() { return this; },
        });
        return this;
      }

      addSeparator() {
        return this;
      }

      showAtMouseEvent() {}
      showAtPosition() {}
    }

    class Modal {
      constructor(app) {
        this.app = app;
        this.contentEl = createElementStub();
        this.modalEl = createElementStub();
      }

      close() {}
      open() {}
    }

    Module._load = function load(request, parent, isMain) {
      if (request === "obsidian") {
        return {
          activeDocument: {
            activeElement: null,
            addEventListener() {},
            body: { classList: { add() {}, remove() {}, toggle() {} } },
            removeEventListener() {},
          },
          Component,
          FileView,
          Menu,
          Modal,
          Notice,
          Platform: { isDesktop: true, isMacOS: false, isMobile: false, isMobileApp: false },
          normalizePath: (value) => value.replace(/\\/g, "/").replace(/\/{2,}/g, "/"),
          setIcon: () => undefined,
        };
      }

      if (request === "pptx-svg") {
        return {
          PptxRenderer: class {},
          degreesToOoxml: (value) => value,
          emuToPx: (value) => value,
          getAllShapes: () => [],
          getShapeTransform: () => ({ x: 0, y: 0, cx: 1, cy: 1, rot: 0 }),
          getSlideScale: () => 1,
          ooxmlToDegrees: (value) => value,
          pxToEmu: (value) => value,
        };
      }

      if (request === "pptx-svg/wasm") {
        return new Uint8Array();
      }

      return originalLoad.call(this, request, parent, isMain);
    };

    try {
      return { ...require(outfile), notices };
    } finally {
      Module._load = originalLoad;
    }
  });

  return viewModulePromise;
}

function createElementStub() {
  return {
    addClass() {},
    empty() {},
    removeClass() {},
  };
}
