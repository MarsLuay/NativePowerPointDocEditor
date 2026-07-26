// Electron main process for the selection-geometry harness.
//
// Renders the generated fixture HTML inside the *same Electron/Chromium build
// Obsidian ships* (so SVG text metrics like getNumberOfChars match what the
// plugin sees), waits for the page driver to publish `document.body.dataset
// .metrics`, prints it as `HARNESS_METRICS:<json>` on stdout, and quits.
//
// Launched by scripts/smoke-selection-geometry.mjs via:
//   electron scripts/lib/electron-eval-main.cjs   (HARNESS_HTML in env)

const { app, BrowserWindow } = require('electron');

const htmlPath = process.env.HARNESS_HTML;
const apiProbe = process.env.ELECTRON_API_PROBE === '1';

// Keep the offscreen render deterministic and quiet.
app.disableHardwareAcceleration();
process.env.ELECTRON_DISABLE_SECURITY_WARNINGS = 'true';

function emit(line) {
  process.stdout.write(line + '\n');
}

app.whenReady().then(async () => {
  if (app.dock) app.dock.hide();

  if (apiProbe) {
    emit('HARNESS_API:' + typeof require('electron'));
    app.exit(0);
    return;
  }

  if (!htmlPath) {
    emit('HARNESS_ERROR:HARNESS_HTML env var was not provided');
    app.exit(1);
    return;
  }

  const win = new BrowserWindow({
    show: false,
    width: 1700,
    height: 1400,
    webPreferences: {
      offscreen: false,
      backgroundThrottling: false,
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  try {
    await win.loadFile(htmlPath);

    // The page sets dataset.metrics synchronously on `load`, but poll briefly to
    // be robust against font/layout settling.
    let metrics = null;
    for (let attempt = 0; attempt < 100; attempt++) {
      metrics = await win.webContents.executeJavaScript('document.body.dataset.metrics || null');
      if (metrics) break;
      await new Promise((resolve) => setTimeout(resolve, 100));
    }

    if (!metrics) {
      emit('HARNESS_ERROR:page never published data-metrics');
      app.exit(1);
      return;
    }

    emit('HARNESS_METRICS:' + metrics);
    app.exit(0);
  } catch (error) {
    emit('HARNESS_ERROR:' + String((error && error.stack) || error));
    app.exit(1);
  }
});
