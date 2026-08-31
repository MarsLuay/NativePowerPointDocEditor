import assert from 'node:assert/strict';
import { test } from 'node:test';
import { loadDocxEmbedLoaderModule } from './helpers/load-plugin-modules.mjs';

test('registerDocxFileEmbed registers extension successfully', async () => {
    const { registerDocxFileEmbed } = await loadDocxEmbedLoaderModule();

    let extensionRegistered = false;
    const plugin = {
        app: {
            embedRegistry: {
                registerExtension: (ext, creator) => {
                    assert.equal(ext, 'docx');
                    assert.equal(typeof creator, 'function');
                    extensionRegistered = true;
                },
                unregisterExtension: () => {}
            }
        },
        register: () => {}
    };

    const result = registerDocxFileEmbed(plugin, () => undefined);
    assert.equal(result, true);
    assert.equal(extensionRegistered, true);
});

test('registerDocxFileEmbed registers extensions successfully (plural)', async () => {
    const { registerDocxFileEmbed } = await loadDocxEmbedLoaderModule();

    let extensionRegistered = false;
    const plugin = {
        app: {
            embedRegistry: {
                registerExtensions: (exts, creator) => {
                    assert.deepEqual(exts, ['docx']);
                    assert.equal(typeof creator, 'function');
                    extensionRegistered = true;
                },
                unregisterExtensions: () => {}
            }
        },
        register: () => {}
    };

    const result = registerDocxFileEmbed(plugin, () => undefined);
    assert.equal(result, true);
    assert.equal(extensionRegistered, true);
});

test('registerDocxFileEmbed handles missing registry', async () => {
    const { registerDocxFileEmbed } = await loadDocxEmbedLoaderModule();

    const plugin = {
        app: {}
    };

    const result = registerDocxFileEmbed(plugin, () => undefined);
    assert.equal(result, false);
});

test('registerDocxFileEmbed catches and logs errors during registration', async () => {
    const { registerDocxFileEmbed } = await loadDocxEmbedLoaderModule();

    const plugin = {
        app: {
            embedRegistry: {
                registerExtension: () => {
                    throw new Error('Test error');
                }
            }
        }
    };

    const result = registerDocxFileEmbed(plugin, () => undefined);
    assert.equal(result, false);
});
