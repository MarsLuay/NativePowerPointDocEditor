import assert from "node:assert/strict";
import { test } from "node:test";
import { bundleSource } from "./helpers/load-plugin-modules.mjs";
import Module, { createRequire } from "node:module";

const require = createRequire(import.meta.url);

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
        export class FileView extends Component {}
        export class Plugin extends Component {
            constructor(app, manifest) { super(); this.app = app; this.manifest = manifest; }
            addSettingTab() {}
            addCommand() {}
            registerView() {}
            registerEvent() {}
            addRibbonIcon() {}
        }
        export class PluginSettingTab {}
        export class MarkdownRenderChild {}
        export class TFile {}
        export class TFolder {}
        export class Menu {}
        export class Setting {}
      `,
      loader: "js",
    }));
  },
};

let mainModulePromise;
function loadMainModule() {
  mainModulePromise ??= bundleSource("src/main.ts", "main.cjs", ["obsidian", "prosemirror-model", "html2canvas", "pptx-svg", "pptx-svg/wasm", "./generated/runtimeArtifactPayloads.json"], [stubObsidianPlugin]).then((outfile) => {
      const originalLoad = Module._load;

      Module._load = function load(request, parent, isMain) {
        if (request.endsWith("generated/runtimeArtifactPayloads.json")) {
            return {};
        }
        if (request === "prosemirror-model") {
            return {
                Schema: class {
                    constructor(spec) { this.spec = spec; this.nodes = {}; this.marks = {}; }
                },
                DOMParser: class {
                    static fromSchema() { return new this(); }
                },
                DOMSerializer: class {
                    static fromSchema() { return new this(); }
                },
                Node: class {},
                Fragment: class {},
                Slice: class {}
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
        return require(outfile);
      } finally {
        Module._load = originalLoad;
      }
  });
  return mainModulePromise;
}

test("setupDevFileLog swallows errors on log append and write", async () => {
    // Reset global logs state to simulate fresh environment
    globalThis.window = {
        nativePowerPointDocEditorDebugLogging: false,
        nativePowerPointDocEditorDebugLogs: [],
    };

    const { default: NativePowerPointDocEditorPlugin } = await loadMainModule();

    let logs = [];
    let appendErrorCaught = false;
    let fallbackWrite = false;
    const app = {
        vault: {
            adapter: {
                append: async (path, data) => {
                    if (path === "fail/dev-debug.log") {
                        appendErrorCaught = true;
                        throw new Error("Simulated append error");
                    }
                    logs.push(data);
                },
                exists: async () => false,
                read: async () => "",
                write: async (path, data) => {
                    if (path === "fail/dev-debug.log" && !data.startsWith("# session ")) {
                        fallbackWrite = true;
                        throw new Error("Simulated fallback write error");
                    }
                },
                stat: async () => null,
            }
        },
        workspace: {
            on: () => {},
        }
    };

    const manifest = { dir: "fail", version: "1.0.0" };

    const plugin = new NativePowerPointDocEditorPlugin(app, manifest);
    plugin.pluginSettings = { debugLogging: true };

    // Calling setupDevFileLog will execute setup and then call infoLog internally.
    // That internally triggers the sink, which uses append.
    await plugin.setupDevFileLog("fail");

    await new Promise((resolve) => setTimeout(resolve, 50));

    assert.ok(appendErrorCaught, "Append should throw error");
    assert.ok(!fallbackWrite, "Should NOT attempt fallback write if append is defined and throws");
});

test("setupDevFileLog swallows errors on log read/write without append", async () => {
    // Reset global logs state to simulate fresh environment
    globalThis.window = {
        nativePowerPointDocEditorDebugLogging: false,
        nativePowerPointDocEditorDebugLogs: [],
    };

    const { default: NativePowerPointDocEditorPlugin } = await loadMainModule();

    let fallbackWrite = false;
    let fallbackWriteErrorCaught = false;
    const app = {
        vault: {
            adapter: {
                // omit append to force fallback
                exists: async () => false,
                read: async () => "",
                write: async (path, data) => {
                    if (path === "no-append/dev-debug.log" && !data.startsWith("# session ")) {
                        fallbackWriteErrorCaught = true;
                        throw new Error("Simulated fallback write error");
                    }
                },
                stat: async () => null,
            }
        },
        workspace: {
            on: () => {},
        }
    };

    const manifest = { dir: "no-append", version: "1.0.0" };

    const plugin = new NativePowerPointDocEditorPlugin(app, manifest);
    plugin.pluginSettings = { debugLogging: true };

    // setupDevFileLog triggers infoLog on successful initialization
    await plugin.setupDevFileLog("no-append");

    await new Promise((resolve) => setTimeout(resolve, 50));

    assert.ok(fallbackWriteErrorCaught, "Should attempt fallback write if append is undefined, and throw to be swallowed");
});
