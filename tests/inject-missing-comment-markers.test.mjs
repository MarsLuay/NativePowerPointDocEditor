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

async function loadInjectModule() {
  if (cachedModule) return cachedModule;
  const outputDirectory = await mkdtemp(path.join(tmpdir(), "npde-comment-inject-"));
  const outfile = path.join(outputDirectory, "inject-comment-markers.cjs");
  await build({
    entryPoints: [
      path.join(
        projectRoot,
        "docx-editor/packages/core/src/docx/injectReplyRangeMarkers.ts",
      ),
    ],
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

test("injectMissingTopLevelCommentRangeMarkers wraps paragraph content", async () => {
  const {
    injectMissingTopLevelCommentRangeMarkers,
    injectReplyRangeMarkers,
  } = await loadInjectModule();

  const content = [
    {
      type: "paragraph",
      content: [{ type: "run", content: [{ type: "text", text: "Body" }] }],
    },
  ];
  const comments = [
    { id: 1, author: "Mars", content: [] },
    { id: 2, author: "Mars", content: [], parentId: 1 },
  ];

  injectMissingTopLevelCommentRangeMarkers(content, comments);
  injectReplyRangeMarkers(content, comments);

  const types = content[0].content.map(
    (item) => `${item.type}:${"id" in item ? item.id : ""}`,
  );
  assert.deepEqual(types, [
    "commentRangeStart:1",
    "commentRangeStart:2",
    "run:",
    "commentRangeEnd:1",
    "commentRangeEnd:2",
  ]);
});

test("injectMissingTopLevelCommentRangeMarkers repairs empty ranges onto text", async () => {
  const { injectMissingTopLevelCommentRangeMarkers } = await loadInjectModule();

  const content = [
    {
      type: "paragraph",
      content: [{ type: "run", content: [{ type: "text", text: "Body text" }] }],
    },
    {
      type: "paragraph",
      content: [
        { type: "run", content: [{ type: "text", text: "Commented" }] },
        { type: "commentRangeStart", id: 7 },
        { type: "run", content: [] },
        { type: "commentRangeEnd", id: 7 },
      ],
    },
  ];

  injectMissingTopLevelCommentRangeMarkers(content, [
    { id: 7, author: "Mars", content: [] },
  ]);

  const abstractTypes = content[0].content.map(
    (item) => `${item.type}:${"id" in item ? item.id : ""}`,
  );
  const emptyParaTypes = content[1].content.map(
    (item) => `${item.type}:${"id" in item ? item.id : ""}`,
  );

  // Empty range removed from para[1]; re-wrapped around last text-bearing para.
  // Only one text-bearing para after removal of markers from para1 - wait, both have text.
  // repair finds empty id 7 on para1 (markers after text, empty between), removes markers,
  // then wraps last text-bearing paragraph (para1 "Commented").
  assert.deepEqual(emptyParaTypes, [
    "commentRangeStart:7",
    "run:",
    "run:",
    "commentRangeEnd:7",
  ]);
  assert.deepEqual(abstractTypes, ["run:"]);
});

test("injectMissingTopLevelCommentRangeMarkers is a no-op when markers exist", async () => {
  const { injectMissingTopLevelCommentRangeMarkers } = await loadInjectModule();
  const content = [
    {
      type: "paragraph",
      content: [
        { type: "commentRangeStart", id: 1 },
        { type: "run", content: [{ type: "text", text: "Body" }] },
        { type: "commentRangeEnd", id: 1 },
      ],
    },
  ];
  injectMissingTopLevelCommentRangeMarkers(content, [
    { id: 1, author: "Mars", content: [] },
  ]);
  assert.equal(
    content[0].content.filter((item) => item.type === "commentRangeStart").length,
    1,
  );
});
