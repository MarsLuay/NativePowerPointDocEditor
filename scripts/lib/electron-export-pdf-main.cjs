// Electron main for scripts/export-pptx-to-pdf.mjs
// Env: HARNESS_HTML, EXPORT_INPUT_PATH, EXPORT_OUTPUT_PATH
const { app, BrowserWindow } = require('electron');
const fs = require('fs');
const path = require('path');

const htmlPath = process.env.HARNESS_HTML;
const inputPath = process.env.EXPORT_INPUT_PATH;
const outputPath = process.env.EXPORT_OUTPUT_PATH;

app.disableHardwareAcceleration();
process.env.ELECTRON_DISABLE_SECURITY_WARNINGS = 'true';

function emit(line) {
	process.stdout.write(line + '\n');
}

app.whenReady().then(async () => {
	if (app.dock) app.dock.hide();
	if (!htmlPath || !inputPath || !outputPath) {
		emit('HARNESS_ERROR:missing HARNESS_HTML / EXPORT_INPUT_PATH / EXPORT_OUTPUT_PATH');
		app.exit(1);
		return;
	}

	const win = new BrowserWindow({
		show: false,
		width: 1600,
		height: 1200,
		webPreferences: {
			backgroundThrottling: false,
			contextIsolation: true,
			nodeIntegration: false,
		},
	});

	try {
		const inputBytes = fs.readFileSync(inputPath);
		const inputBase64 = inputBytes.toString('base64');
		await win.loadFile(htmlPath);

		const metricsRaw = await win.webContents.executeJavaScript(
			`window.__npdeExportPdf(${JSON.stringify(inputBase64)})`,
		);
		const metrics = typeof metricsRaw === 'string' ? JSON.parse(metricsRaw) : metricsRaw;
		if (!metrics || !metrics.ok) {
			emit('HARNESS_ERROR:' + String((metrics && metrics.error) || 'export failed'));
			app.exit(1);
			return;
		}

		fs.mkdirSync(path.dirname(outputPath), { recursive: true });
		fs.writeFileSync(outputPath, Buffer.from(metrics.pdfBase64, 'base64'));
		emit(
			'HARNESS_METRICS:' +
				JSON.stringify({
					ok: true,
					path: outputPath,
					bytes: metrics.bytes,
					slideCount: metrics.slideCount,
				}),
		);
		app.exit(0);
	} catch (error) {
		emit('HARNESS_ERROR:' + String((error && error.stack) || error));
		app.exit(1);
	}
});
