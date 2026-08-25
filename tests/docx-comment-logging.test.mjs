import assert from "node:assert/strict";
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

test("extractDocxCommentPlainText handles explicit text, nested runs, edge cases, and truncation", async () => {
  const { extractDocxCommentPlainText } = await loadCommentLoggingModule();

  // Explicit text takes precedence
  assert.equal(
    extractDocxCommentPlainText({ id: 10, text: "Explicit comment text" }),
    "Explicit comment text"
  );

  // Empty comment / no content
  assert.equal(extractDocxCommentPlainText({ id: 11 }), "");
  assert.equal(extractDocxCommentPlainText({ id: 12, content: [] }), "");
  assert.equal(extractDocxCommentPlainText({ id: 13, content: [{ content: [] }] }), "");

  // Non-run and non-text child nodes ignored gracefully
  const mixedContent = [
    {
      content: [
        { type: "image", content: [] },
        {
          type: "run",
          content: [
            { type: "bold_flag", text: undefined },
            { type: "text", text: "Hello " },
            { type: "text", text: "world!" },
          ],
        },
      ],
    },
    {
      content: [
        {
          type: "run",
          content: [{ type: "text", text: " Next line." }],
        },
      ],
    },
  ];
  assert.equal(
    extractDocxCommentPlainText({ id: 14, content: mixedContent }),
    "Hello world! Next line."
  );

  // Text truncation over 500 chars
  const longStr = "A".repeat(600);
  const truncated = extractDocxCommentPlainText({ id: 15, text: longStr });
  assert.equal(truncated.length, 501); // 500 chars + '…'
  assert.ok(truncated.endsWith("…"));
});
