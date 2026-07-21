import assert from "node:assert/strict";
import { docxEditorAliases } from './helpers/docx-esbuild-aliases.mjs';
import { test } from "node:test";
import { build } from "esbuild";
import { mkdtemp } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { createRequire } from "node:module";

const projectRoot = path.resolve(import.meta.dirname, "..");
const require = createRequire(import.meta.url);

let cachedModule;

async function loadCommentLoggingModule() {
  if (cachedModule) return cachedModule;
  const outputDirectory = await mkdtemp(path.join(tmpdir(), "npde-comment-log-"));
  const outfile = path.join(outputDirectory, "docx-comment-logging.cjs");
  await build({
		alias: docxEditorAliases,
    entryPoints: [path.join(projectRoot, "src/docxCommentLogging.ts")],
    bundle: true,
    format: "cjs",
    logLevel: "silent",
    outfile,
    platform: "node",
    target: "node22",
  });
  cachedModule = require(outfile);
  return cachedModule;
}

test("summarizeDocxComments includes replies with parentId", async () => {
  const { summarizeDocxComments } = await loadCommentLoggingModule();
  const summary = summarizeDocxComments([
    {
      id: 1,
      author: "Mars",
      date: "2026-07-15T21:18:10Z",
      content: [{ content: [{ type: "run", content: [{ type: "text", text: "make fancier" }] }] }],
    },
    {
      id: 2,
      parentId: 1,
      author: "Mars",
      content: [{ content: [{ type: "run", content: [{ type: "text", text: "add more polish" }] }] }],
    },
  ]);

  assert.equal(summary.total, 2);
  assert.equal(summary.topLevel, 1);
  assert.equal(summary.replies, 1);
  assert.equal(summary.comments[0].kind, "comment");
  assert.equal(summary.comments[0].text, "make fancier");
  assert.equal(summary.comments[1].kind, "reply");
  assert.equal(summary.comments[1].parentId, 1);
  assert.equal(summary.comments[1].text, "add more polish");
});
