import en from '@eigenpal/docx-editor-i18n/en';

type PlainRecord = Record<string, unknown>;
type TranslationVars = Record<string, string | number>;

export { en };

export const locales = { en };

function isRecord(value: unknown): value is PlainRecord {
	return value !== null && typeof value === 'object' && !Array.isArray(value);
}

export function deepMerge(base: PlainRecord, override: PlainRecord | undefined): PlainRecord {
	if (!isRecord(override)) {
		return base;
	}

	const result: PlainRecord = { ...base };
	for (const key of Object.keys(override)) {
		const baseValue = base[key];
		const overrideValue = override[key];

		if (overrideValue === null) {
			continue;
		}

		if (isRecord(baseValue) && isRecord(overrideValue)) {
			result[key] = deepMerge(baseValue, overrideValue);
		} else if (overrideValue !== undefined) {
			result[key] = overrideValue;
		}
	}

	return result;
}

function lookupKey(strings: PlainRecord, path: string): string | undefined {
	let current: unknown = strings;

	for (const part of path.split('.')) {
		if (!isRecord(current)) {
			return undefined;
		}

		current = current[part];
	}

	return typeof current === 'string' ? current : undefined;
}

function parseBranches(branchString: string): Record<string, string> {
	const parsed: Record<string, string> = {};
	const regex = /(=\d+|\w+)\s*\{([^}]*)\}/g;
	let match: RegExpExecArray | null;

	while ((match = regex.exec(branchString)) !== null) {
		const branch = match[1];
		const text = match[2];
		if (branch && text !== undefined) {
			parsed[branch] = text;
		}
	}

	return parsed;
}

function getPluralCategory(count: number, language: string): string {
	const intlWithPluralRules = Intl as typeof Intl & {
		PluralRules?: new (locales?: string | string[]) => { select: (value: number) => string };
	};

	try {
		return intlWithPluralRules.PluralRules
			? new intlWithPluralRules.PluralRules(language || 'en').select(count)
			: (count === 1 ? 'one' : 'other');
	} catch {
		return count === 1 ? 'one' : 'other';
	}
}

function formatMessage(template: string, vars: TranslationVars | undefined, language: string): string {
	if (!vars) {
		return template;
	}

	const pluralized = template.replace(
		/\{(\w+),\s*plural,\s*((?:[^{}]|\{[^{}]*\})*)\}/g,
		(full, varName: string, branchString: string) => {
			const count = Number(vars[varName]);
			if (Number.isNaN(count)) {
				return full;
			}

			const parsed = parseBranches(branchString);
			const exact = parsed[`=${count}`];
			if (exact !== undefined) {
				return exact.replace(/#/g, String(count));
			}

			const category = getPluralCategory(count, language);
			const text = parsed[category] ?? parsed.other ?? '';
			return text.replace(/#/g, String(count));
		},
	);

	return pluralized.replace(/\{(\w+)\}/g, (placeholder, key: string) => {
		const value = vars[key];
		return value !== undefined ? String(value) : placeholder;
	});
}

export function createT(strings: PlainRecord, language = 'en') {
	return (key: string, vars?: TranslationVars): string => {
		const value = lookupKey(strings, key);
		return formatMessage(value ?? key, vars, language);
	};
}
