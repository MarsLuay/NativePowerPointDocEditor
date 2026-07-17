import type { TFile, Vault } from 'obsidian';
import { PresentationEngine } from '../PresentationEngine';
import {
	inspectPowerPointPackage,
	summarizePackageMessages,
	validatePowerPointExport,
	validatePowerPointExportContents,
	type PowerPointPackageInspection,
} from '../PowerPointPackage';

export async function exportValidatedPptx(
	engine: PresentationEngine,
	sourceBuffer: ArrayBuffer,
	sourcePackage: PowerPointPackageInspection,
): Promise<ArrayBuffer> {
	const output = await engine.export();
	const exportedPackage = inspectPowerPointPackage(output);
	const validation = validatePowerPointExport(sourcePackage, exportedPackage, engine.slideCount);
	if (!validation.ok) {
		throw new Error(`Export validation failed: ${summarizePackageMessages(validation.errors)}`);
	}

	const contentValidation = await validatePowerPointExportContents(sourceBuffer, output);
	if (!contentValidation.ok) {
		throw new Error(`Export validation failed: ${summarizePackageMessages(contentValidation.errors)}`);
	}

	await PresentationEngine.validateRoundTrip(output, engine.slideCount);
	return output;
}

export async function savePptxToVault(
	vault: Vault,
	file: TFile,
	engine: PresentationEngine,
	sourceBuffer: ArrayBuffer,
	sourcePackage: PowerPointPackageInspection,
): Promise<{ output: ArrayBuffer; sourcePackage: PowerPointPackageInspection }> {
	const output = await exportValidatedPptx(engine, sourceBuffer, sourcePackage);
	await vault.modifyBinary(file, output);
	return {
		output,
		sourcePackage: inspectPowerPointPackage(output),
	};
}
