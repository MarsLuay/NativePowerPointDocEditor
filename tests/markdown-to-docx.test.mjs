import assert from "node:assert/strict";
import { test } from "node:test";
import JSZip from "jszip";

import {
  loadMarkdownSourceModule,
  loadMarkdownToDocxModule,
} from "./helpers/load-plugin-modules.mjs";

test("buildMarkdownDocxArrayBuffer creates an editable DOCX with Markdown structure", async () => {
  const { buildMarkdownDocxArrayBuffer } = await loadMarkdownToDocxModule();
  const buffer = await buildMarkdownDocxArrayBuffer([
    "---",
    "category: test",
    "---",
    "# Document title",
    "",
    "A **bold** and *italic* paragraph with `inline code` and [a link](https://example.com).",
    "",
    "- Bullet item",
    "1. Numbered item",
    "> Quoted text",
    "",
    "```ts",
    "const answer = 42;",
    "```",
  ].join("\n"));
  const zip = await JSZip.loadAsync(buffer);

  for (const path of [
    "[Content_Types].xml",
    "_rels/.rels",
    "word/document.xml",
    "word/_rels/document.xml.rels",
    "word/styles.xml",
    "word/numbering.xml",
  ]) {
    assert.ok(zip.file(path), `Expected DOCX part ${path}`);
  }

  const documentXml = await zip.file("word/document.xml").async("string");
  assert.match(documentXml, /<w:pStyle w:val="Heading1"\/>/);
  assert.match(documentXml, /<w:b\/><w:bCs\/>/);
  assert.match(documentXml, /<w:i\/><w:iCs\/>/);
  assert.match(documentXml, /w:ascii="Courier New"/);
  assert.match(documentXml, /<w:numId w:val="1"\/>/);
  assert.match(documentXml, /<w:numId w:val="2"\/>/);
  assert.match(documentXml, /Quoted text/);
  assert.match(documentXml, /a link/);
  assert.match(documentXml, /https:\/\/example\.com/);
  assert.doesNotMatch(documentXml, /category: test/);
});

test("buildMarkdownDocxArrayBuffer escapes XML and removes invalid control characters", async () => {
  const { buildMarkdownDocxArrayBuffer } = await loadMarkdownToDocxModule();
  const zip = await JSZip.loadAsync(await buildMarkdownDocxArrayBuffer("A\t& B < C > D \"quoted\" 'value'\u0000\u0001\u0008\u000B\u000C\u000E\u001F"));
  const documentXml = await zip.file("word/document.xml").async("string");

  assert.match(documentXml, /A\t&amp; B &lt; C &gt; D &quot;quoted&quot; &apos;value&apos;/);
  for (const character of ["\u0000", "\u0001", "\u0008", "\u000B", "\u000C", "\u000E", "\u001F"]) {
    assert.equal(documentXml.includes(character), false);
  }
});

test("resolveMarkdownDocxOutputPath creates a numbered sibling on collisions", async () => {
  const { buildMarkdownDocxCandidatePath, resolveMarkdownDocxOutputPath } = await loadMarkdownToDocxModule();
  const existing = new Set(["Notes/Plan.docx", "Notes/Plan 2.docx"]);

  assert.equal(buildMarkdownDocxCandidatePath("Notes/Plan.MD"), "Notes/Plan.docx");
  assert.equal(buildMarkdownDocxCandidatePath("Notes/Plan.MDENC"), "Notes/Plan.docx");
  assert.equal(buildMarkdownDocxCandidatePath("Notes/Plan.encrypted"), "Notes/Plan.docx");
  assert.equal(
    resolveMarkdownDocxOutputPath("Notes/Plan.md", path => existing.has(path)),
    "Notes/Plan 3.docx",
  );
  assert.throws(() => buildMarkdownDocxCandidatePath("Notes/Plan.txt"), /\.md, \.mdenc, or \.encrypted/);
});

test("Meld Encrypt source detection accepts its encrypted Markdown extensions", async () => {
  const { isMarkdownDocxSourceExtension } = await loadMarkdownSourceModule();

  assert.equal(isMarkdownDocxSourceExtension("md"), true);
  assert.equal(isMarkdownDocxSourceExtension("MDENC"), true);
  assert.equal(isMarkdownDocxSourceExtension("encrypted"), true);
  assert.equal(isMarkdownDocxSourceExtension("txt"), false);
});

test("Meld Encrypt conversion reads decrypted view data and preserves the encrypted file", async () => {
  const { convertMarkdownFileToDocx } = await loadMarkdownToDocxModule();
  const sourceFile = { path: "Private/Plan.mdenc", extension: "mdenc" };
  const plaintext = "# Private plan\n\nOnly decrypted Markdown belongs in the DOCX.";
  const encryptedPayload = "meld-ciphertext-that-must-not-be-converted";
  let createdPath;
  let createdBuffer;
  let openedOutput;
  let sourceWasModified = false;
  const encryptedView = {
    file: sourceFile,
    isSavingEnabled: true,
    getViewType: () => "meld-encrypted-view",
    getViewData: () => encryptedPayload,
    getUnencryptedViewData: () => plaintext,
  };
  const app = {
    vault: {
      getAbstractFileByPath: () => null,
      read: async () => encryptedPayload,
      createBinary: async (path, buffer) => {
        createdPath = path;
        createdBuffer = buffer;
        return { path };
      },
      modifyBinary: async () => { sourceWasModified = true; },
    },
    workspace: {
      iterateAllLeaves: (callback) => callback({ view: encryptedView }),
      getLeaf: (kind) => {
        assert.equal(kind, "tab");
        return { openFile: async file => { openedOutput = file; } };
      },
    },
  };

  const outputFile = await convertMarkdownFileToDocx(app, sourceFile);
  const zip = await JSZip.loadAsync(createdBuffer);
  const documentXml = await zip.file("word/document.xml").async("string");

  assert.equal(createdPath, "Private/Plan.docx");
  assert.equal(outputFile.path, createdPath);
  assert.equal(openedOutput, outputFile);
  assert.equal(sourceWasModified, false);
  assert.match(documentXml, /Private plan/);
  assert.match(documentXml, /Only decrypted Markdown belongs in the DOCX\./);
  assert.doesNotMatch(documentXml, /meld-ciphertext/);
});

test("Markdown conversion reads the normal vault source and opens the DOCX sibling", async () => {
  const { convertMarkdownFileToDocx } = await loadMarkdownToDocxModule();
  const sourceFile = { path: "Notes/Plan.md", extension: "md" };
  let createdPath;
  let createdBuffer;
  let openedOutput;
  const app = {
    vault: {
      getAbstractFileByPath: () => null,
      read: async file => {
        assert.equal(file, sourceFile);
        return "# Plain Markdown\n\nSaved by the normal vault read path.";
      },
      createBinary: async (path, buffer) => {
        createdPath = path;
        createdBuffer = buffer;
        return { path };
      },
    },
    workspace: {
      getLeaf: kind => {
        assert.equal(kind, "tab");
        return { openFile: async file => { openedOutput = file; } };
      },
    },
  };

  const outputFile = await convertMarkdownFileToDocx(app, sourceFile);
  const zip = await JSZip.loadAsync(createdBuffer);
  const documentXml = await zip.file("word/document.xml").async("string");

  assert.equal(createdPath, "Notes/Plan.docx");
  assert.equal(openedOutput, outputFile);
  assert.match(documentXml, /Plain Markdown/);
  assert.match(documentXml, /Saved by the normal vault read path\./);
});

test("Meld Encrypt source loading opens through its view and closes the temporary leaf", async () => {
  const { readMarkdownSourceForDocx } = await loadMarkdownSourceModule();
  const sourceFile = { path: "Private/Plan.encrypted", extension: "encrypted" };
  const plaintext = "Unlocked by Meld Encrypt";
  let openedFile;
  let detached = false;
  const temporaryLeaf = {
    view: null,
    async openFile(file, options) {
      openedFile = file;
      assert.equal(options.active, true);
      this.view = {
        file,
        isSavingEnabled: true,
        getViewType: () => "meld-encrypted-view",
        getViewData: () => "encrypted payload",
        getUnencryptedViewData: () => plaintext,
      };
    },
    detach() { detached = true; },
  };
  const app = {
    vault: { read: async () => "encrypted payload" },
    workspace: {
      iterateAllLeaves(callback) {
        if (temporaryLeaf.view) callback(temporaryLeaf);
      },
      getLeaf: kind => {
        assert.equal(kind, "tab");
        return temporaryLeaf;
      },
    },
  };

  assert.equal(await readMarkdownSourceForDocx(app, sourceFile), plaintext);
  assert.equal(openedFile, sourceFile);
  assert.equal(detached, true);
});

test("Meld Encrypt conversion refuses a locked view", async () => {
  const { readMarkdownSourceForDocx } = await loadMarkdownSourceModule();
  const sourceFile = { path: "Private/Plan.mdenc", extension: "mdenc" };
  const app = {
    vault: { read: async () => "ciphertext" },
    workspace: {
      iterateAllLeaves: callback => callback({
        view: {
          file: sourceFile,
          isSavingEnabled: false,
          getViewType: () => "meld-encrypted-view",
          getUnencryptedViewData: () => "",
        },
      }),
      getLeaf: () => { throw new Error("Should reuse the open Meld view"); },
    },
  };

  await assert.rejects(readMarkdownSourceForDocx(app, sourceFile), /Unlock this Meld Encrypt note/);
});
