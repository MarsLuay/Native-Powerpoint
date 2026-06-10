import { createT, deepMerge, en, type Translations, type TranslationKey } from '@eigenpal/docx-editor-i18n';

type TranslationVars = Record<string, string | number>;

export type EditorTranslator = (key: TranslationKey, vars?: TranslationVars, fallback?: string) => string;

function getLocaleLanguage(locale: Translations): string {
	const language = (locale as { _lang?: unknown })._lang;
	return typeof language === 'string' ? language : 'en';
}

export function createEditorTranslator(locale: Translations | undefined): EditorTranslator {
	const mergedLocale = deepMerge(en, locale) as typeof en;
	const translate = createT(mergedLocale, getLocaleLanguage(mergedLocale));

	return (key, vars, fallback = key) => {
		const value = translate(key, vars);
		return value === key ? fallback : value;
	};
}
