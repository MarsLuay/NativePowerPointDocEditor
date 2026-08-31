import assert from "node:assert/strict";
import { test } from "node:test";
import { loadDocxEmbedModule } from "./helpers/load-plugin-modules.mjs";

test("processDocxEmbeds adds a scan child to the context", async () => {
  const { processDocxEmbeds } = await loadDocxEmbedModule();

  let addedChild = null;
  const ctx = {
    addChild(child) {
      addedChild = child;
    },
  };

  const el = { ownerDocument: { defaultView: { setTimeout: () => {}, clearTimeout: () => {} } } };
  const app = {};
  const getEditorLocale = () => undefined;

  processDocxEmbeds(app, el, ctx, getEditorLocale);

  assert.ok(addedChild, "Expected processDocxEmbeds to add a child to the context");
  assert.equal(addedChild.constructor.name, "DocxEmbedScanChild");
});
