import accessibilityArXB from '../../locales/ar-XB/accessibility.json';
import commonArXB from '../../locales/ar-XB/common.json';
import docxArXB from '../../locales/ar-XB/docx.json';
import errorsArXB from '../../locales/ar-XB/errors.json';
import powerpointArXB from '../../locales/ar-XB/powerpoint.json';
import settingsArXB from '../../locales/ar-XB/settings.json';
import accessibilityEn from '../../locales/en/accessibility.json';
import commonEn from '../../locales/en/common.json';
import docxEn from '../../locales/en/docx.json';
import errorsEn from '../../locales/en/errors.json';
import powerpointEn from '../../locales/en/powerpoint.json';
import settingsEn from '../../locales/en/settings.json';
import accessibilityEnXA from '../../locales/en-XA/accessibility.json';
import commonEnXA from '../../locales/en-XA/common.json';
import docxEnXA from '../../locales/en-XA/docx.json';
import errorsEnXA from '../../locales/en-XA/errors.json';
import powerpointEnXA from '../../locales/en-XA/powerpoint.json';
import settingsEnXA from '../../locales/en-XA/settings.json';
import accessibilityPl from '../../locales/pl/accessibility.json';
import commonPl from '../../locales/pl/common.json';
import docxPl from '../../locales/pl/docx.json';
import errorsPl from '../../locales/pl/errors.json';
import powerpointPl from '../../locales/pl/powerpoint.json';
import settingsPl from '../../locales/pl/settings.json';

export const BUNDLED_LOCALE_JSON = {
	'ar-XB': {
		accessibility: accessibilityArXB,
		common: commonArXB,
		docx: docxArXB,
		errors: errorsArXB,
		powerpoint: powerpointArXB,
		settings: settingsArXB,
	},
	en: {
		accessibility: accessibilityEn,
		common: commonEn,
		docx: docxEn,
		errors: errorsEn,
		powerpoint: powerpointEn,
		settings: settingsEn,
	},
	'en-XA': {
		accessibility: accessibilityEnXA,
		common: commonEnXA,
		docx: docxEnXA,
		errors: errorsEnXA,
		powerpoint: powerpointEnXA,
		settings: settingsEnXA,
	},
	pl: {
		accessibility: accessibilityPl,
		common: commonPl,
		docx: docxPl,
		errors: errorsPl,
		powerpoint: powerpointPl,
		settings: settingsPl,
	},
} as const;

export const BUNDLED_LOCALES = Object.keys(BUNDLED_LOCALE_JSON);
