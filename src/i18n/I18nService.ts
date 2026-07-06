import type { MessageKey } from './generated/messageKeys.generated';
import { formatMessage } from './messageFormat';
import type { LoadedLocale } from './eigenpalAdapter';
import type { PluginMessages } from './localeLoader';

export type { MessageKey };

export interface I18nService {
	readonly locale: string;
	readonly direction: 'ltr' | 'rtl';
	t(key: MessageKey | string, values?: Record<string, string | number | boolean>): string;
	formatNumber(value: number, options?: Intl.NumberFormatOptions): string;
	formatDate(value: Date | number, options?: Intl.DateTimeFormatOptions): string;
}

export class PluginI18nService implements I18nService {
	constructor(
		public readonly locale: string,
		public readonly direction: 'ltr' | 'rtl',
		private readonly messages: PluginMessages,
		private readonly baseLanguageMessages: PluginMessages | undefined,
		private readonly englishFallback: PluginMessages,
	) {}

	t(key: MessageKey | string, values?: Record<string, string | number | boolean>): string {
		const template =
			this.messages[key]
			?? this.baseLanguageMessages?.[key]
			?? this.englishFallback[key]
			?? key;

		return formatMessage(template, values, this.locale);
	}

	formatNumber(value: number, options?: Intl.NumberFormatOptions): string {
		return new Intl.NumberFormat(this.locale, options).format(value);
	}

	formatDate(value: Date | number, options?: Intl.DateTimeFormatOptions): string {
		return new Intl.DateTimeFormat(this.locale, options).format(value);
	}
}

export function createPluginI18nService(
	loaded: LoadedLocale,
	englishFallback: PluginMessages,
): PluginI18nService {
	return new PluginI18nService(
		loaded.locale,
		loaded.direction,
		loaded.pluginMessages,
		undefined,
		englishFallback,
	);
}
