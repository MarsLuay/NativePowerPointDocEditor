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
	const allowedPartRemovals = engine.getPrunedPackageParts();
	const validation = validatePowerPointExport(sourcePackage, exportedPackage, engine.slideCount, {
		allowedPartRemovals,
	});
	if (!validation.ok) {
		throw new Error(`Export validation failed: ${summarizePackageMessages(validation.errors)}`);
	}

	const contentValidation = await validatePowerPointExportContents(sourceBuffer, output, {
		allowedMarkerRemovals: engine.getProtectedSlideMarkerRemovalAllowance(),
		allowedUnknownElementRemovals: engine.getUnknownSlideElementRemovalAllowance(),
		allowedExternalRelationshipRemovals: engine.getExternalRelationshipRemovalAllowance(),
		allowedPartRemovals,
	});
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
	engine.clearProtectedSlideMarkerRemovalAllowance();
	return {
		output,
		sourcePackage: inspectPowerPointPackage(output),
	};
}
