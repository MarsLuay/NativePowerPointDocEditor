import assert from 'node:assert/strict';
import test, { afterEach } from 'node:test';
import { loadNotifyModule } from './helpers/load-plugin-modules.mjs';

test('showI18nNotice uses i18n translation when available', async () => {
    const { showI18nNotice, notices } = await loadNotifyModule();
    notices.length = 0;

    const mockI18n = {
        t: (key, values) => `translated: ${key} ${values?.count ?? ''}`.trim()
    };

    showI18nNotice(mockI18n, 'my.key', { count: 5 });

    assert.equal(notices.length, 1);
    assert.equal(notices[0].message, 'translated: my.key 5');
});

test('showI18nNotice falls back to key when i18n is missing', async () => {
    const { showI18nNotice, notices } = await loadNotifyModule();
    notices.length = 0;

    showI18nNotice(null, 'fallback.key');

    assert.equal(notices.length, 1);
    assert.equal(notices[0].message, 'fallback.key');
});

test('showI18nNotice falls back to key when translation returns undefined', async () => {
    const { showI18nNotice, notices } = await loadNotifyModule();
    notices.length = 0;

    const mockI18n = {
        t: () => undefined
    };

    showI18nNotice(mockI18n, 'fallback.key');

    assert.equal(notices.length, 1);
    assert.equal(notices[0].message, 'fallback.key');
});
