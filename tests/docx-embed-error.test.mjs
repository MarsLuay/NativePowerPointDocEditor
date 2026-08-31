import assert from "node:assert/strict";
import { test } from "node:test";
import { loadDocxEmbedLoaderModule } from "./helpers/load-plugin-modules.mjs";

test("LazyDocxFileEmbed renders error message when DOCX editor chunk load fails", async () => {
	const module = await loadDocxEmbedLoaderModule();

	let registeredCreator;
	const mockPlugin = {
		app: {
			embedRegistry: {
				registerExtension: (ext, creator) => { registeredCreator = creator; }
			}
		},
		register: () => {}
	};

	module.registerDocxFileEmbed(mockPlugin, () => undefined);

	const mockContainer = module.createElementStub();
	const mockInfo = { containerEl: mockContainer };
	const mockFile = { name: "test.docx", path: "folder/test.docx" };

	const lazyEmbed = registeredCreator(mockInfo, mockFile, "");

	lazyEmbed.onload();

	// Wait a tick for promises to resolve
	await new Promise(resolve => setTimeout(resolve, 10));

	assert.equal(mockContainer.children.length, 1);
	assert.equal(mockContainer.children[0].className, "native-powerpoint-doc-editor-embed-error");
	assert.equal(mockContainer.children[0].innerText, "Could not load the DOCX editor: Simulated chunk load error");
});
