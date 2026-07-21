import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { test } from "node:test";
import path from "node:path";
import { fileURLToPath } from "node:url";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

test("shipped PagedEditor dist does not replay a space key already handled by hidden ProseMirror", async () => {
  const dist = await readFile(
    path.join(projectRoot, "vendor/docx-editor-runtime/react/dist/index.js"),
    "utf8",
  );
  const handledKeyCheck = dist.indexOf("defaultPrevented");
  const spaceRelayStart = dist.indexOf('key===" "', handledKeyCheck);
  const handler = dist.slice(handledKeyCheck, spaceRelayStart);

  assert.notEqual(handledKeyCheck, -1, "PagedEditor dist must guard handled editor keys");
  assert.notEqual(spaceRelayStart, -1, "PagedEditor dist must retain its unhandled-space fallback");
  assert.match(
    handler,
    /defaultPrevented/,
    "a handled hidden-editor key must stop before the container can replay it",
  );
});

test("DOCX space routes emit bounded, text-free diagnostics", async () => {
  const source = await readFile(path.join(projectRoot, "src/DocxReactView.tsx"), "utf8");

  assert.match(source, /debugLog\('text-input', 'DOCX space input routed'/);
  assert.match(source, /isSpace: true/);
  assert.match(source, /now - lastDocxSpaceInputLogAt < 250/);
  assert.doesNotMatch(source, /data: event\.data/);
});
