import JSZip from 'jszip';

const TEXT_PART_PATTERNS = [
	/^word\/document\.xml$/,
	/^word\/headers\/header\d+\.xml$/,
	/^word\/footers\/footer\d+\.xml$/,
	/^word\/footnotes\.xml$/,
	/^word\/endnotes\.xml$/,
	/^word\/comments\.xml$/,
];

const TEXT_TOKEN_PATTERN = /<w:t\b[^>]*>([\s\S]*?)<\/w:t>|<w:tab\b[^>]*\/>|<w:br\b[^>]*\/>|<w:cr\b[^>]*\/>|<\/w:p>/g;
const RUN_PATTERN = /<w:r\b[\s\S]*?<\/w:r>/g;
const PARAGRAPH_PATTERN = /<w:p\b[\s\S]*?<\/w:p>/g;
const STYLE_PATTERN = /<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g;
const MAX_SNIPPET_LENGTH = 500;
const TINY_FONT_SIZE_HALF_POINTS = 4;

const PROMPT_INJECTION_PATTERNS = [
	/ignore (?:all )?(?:previous|prior|above|earlier) instructions/i,
	/disregard (?:all )?(?:previous|prior|above|earlier) instructions/i,
	/you are (?:now )?(?:chatgpt|an? ai|an? assistant|a language model)/i,
	/system prompt/i,
	/developer message/i,
	/prompt injection/i,
	/do not (?:summarize|mention|reveal|tell)/i,
	/follow (?:only )?(?:these|the following) instructions/i,
];

interface TextVisibilityProps {
	hidden?: boolean;
	webHidden?: boolean;
	color?: string;
	fontSizeHalfPoints?: number;
}

interface ParsedStyle {
	id: string;
	basedOn?: string;
	props: TextVisibilityProps;
}

export interface HiddenTextFinding {
	id: string;
	partPath: string;
	partLabel: string;
	paragraphNumber: number;
	text: string;
	reasons: string[];
	promptInjectionSignals: string[];
}

export interface HiddenTextScanResult {
	findings: HiddenTextFinding[];
	partsScanned: number;
}

function decodeXmlEntities(value: string): string {
	return value
		.replace(/&lt;/g, '<')
		.replace(/&gt;/g, '>')
		.replace(/&amp;/g, '&')
		.replace(/&quot;/g, '"')
		.replace(/&apos;/g, "'")
		.replace(/&#(\d+);/g, (_match, codePoint: string) => {
			const numericCodePoint = Number(codePoint);
			return Number.isFinite(numericCodePoint) ? String.fromCodePoint(numericCodePoint) : '';
		})
		.replace(/&#x([0-9a-fA-F]+);/g, (_match, codePoint: string) => {
			const numericCodePoint = Number.parseInt(codePoint, 16);
			return Number.isFinite(numericCodePoint) ? String.fromCodePoint(numericCodePoint) : '';
		});
}

function normalizeText(value: string): string {
	return value
		.replace(/\r/g, '\n')
		.replace(/[ \t]+\n/g, '\n')
		.replace(/\n[ \t]+/g, '\n')
		.replace(/\n{3,}/g, '\n\n')
		.trim();
}

function truncateText(value: string): string {
	const normalized = normalizeText(value);
	if (normalized.length <= MAX_SNIPPET_LENGTH) {
		return normalized;
	}

	return `${normalized.slice(0, MAX_SNIPPET_LENGTH - 1).trimEnd()}…`;
}

function extractTextFromXml(xml: string): string {
	const pieces: string[] = [];
	let match: RegExpExecArray | null;

	TEXT_TOKEN_PATTERN.lastIndex = 0;

	while ((match = TEXT_TOKEN_PATTERN.exec(xml)) !== null) {
		const [token, text] = match;
		if (text !== undefined) {
			pieces.push(decodeXmlEntities(text));
		} else if (token.startsWith('<w:tab')) {
			pieces.push('\t');
		} else {
			pieces.push('\n');
		}
	}

	return normalizeText(pieces.join(''));
}

function getAttribute(xml: string, name: string): string | null {
	const pattern = new RegExp(`(?:\\w+:)?${name}="([^"]*)"`);
	return xml.match(pattern)?.[1] ?? null;
}

function getElementWithBody(xml: string, localName: string): string | null {
	const pattern = new RegExp(`<w:${localName}\\b[\\s\\S]*?</w:${localName}>|<w:${localName}\\b[^>]*/>`);
	return xml.match(pattern)?.[0] ?? null;
}

function getElementVal(xml: string, localName: string): string | null {
	const element = getElementWithBody(xml, localName);
	return element ? getAttribute(element, 'val') : null;
}

function getRunProperties(xml: string): string {
	return getElementWithBody(xml, 'rPr') ?? '';
}

function getParagraphProperties(xml: string): string {
	return getElementWithBody(xml, 'pPr') ?? '';
}

function parseToggle(xml: string, localName: string): boolean | undefined {
	const element = getElementWithBody(xml, localName);
	if (!element) {
		return undefined;
	}

	const rawValue = getAttribute(element, 'val');
	if (rawValue === null || rawValue === '') {
		return true;
	}

	return !/^(?:0|false|off|none)$/i.test(rawValue);
}

function normalizeColor(value: string | null): string | undefined {
	if (!value || /^(?:auto|automatic)$/i.test(value)) {
		return undefined;
	}

	const namedColor = value.toLowerCase() === 'white' ? 'FFFFFF' : value.replace(/^#/, '');
	if (!/^[0-9a-f]{6}$/i.test(namedColor)) {
		return undefined;
	}

	return namedColor.toUpperCase();
}

function isNearWhiteColor(value: string): boolean {
	const red = Number.parseInt(value.slice(0, 2), 16);
	const green = Number.parseInt(value.slice(2, 4), 16);
	const blue = Number.parseInt(value.slice(4, 6), 16);
	return red >= 245 && green >= 245 && blue >= 245;
}

function parseTextVisibilityProps(rPr: string): TextVisibilityProps {
	const colorElement = getElementWithBody(rPr, 'color');
	const sizeValue = getElementVal(rPr, 'sz') ?? getElementVal(rPr, 'szCs');
	const fontSizeHalfPoints = sizeValue ? Number(sizeValue) : undefined;

	return {
		hidden: parseToggle(rPr, 'vanish'),
		webHidden: parseToggle(rPr, 'webHidden'),
		color: normalizeColor(colorElement ? getAttribute(colorElement, 'val') : null),
		fontSizeHalfPoints: Number.isFinite(fontSizeHalfPoints) ? fontSizeHalfPoints : undefined,
	};
}

function mergeProps(base: TextVisibilityProps, override: TextVisibilityProps): TextVisibilityProps {
	return {
		hidden: override.hidden ?? base.hidden,
		webHidden: override.webHidden ?? base.webHidden,
		color: override.color ?? base.color,
		fontSizeHalfPoints: override.fontSizeHalfPoints ?? base.fontSizeHalfPoints,
	};
}

function parseStyles(stylesXml: string): Map<string, ParsedStyle> {
	const styles = new Map<string, ParsedStyle>();
	let match: RegExpExecArray | null;

	STYLE_PATTERN.lastIndex = 0;

	while ((match = STYLE_PATTERN.exec(stylesXml)) !== null) {
		const attrs = match[1] ?? '';
		const body = match[2] ?? '';
		const id = getAttribute(attrs, 'styleId');
		if (!id) {
			continue;
		}

		styles.set(id, {
			id,
			basedOn: getElementVal(body, 'basedOn') ?? undefined,
			props: parseTextVisibilityProps(getRunProperties(body)),
		});
	}

	return styles;
}

function resolveStyleProps(styles: Map<string, ParsedStyle>, styleId: string | null, seen = new Set<string>()): TextVisibilityProps {
	if (!styleId || seen.has(styleId)) {
		return {};
	}

	const style = styles.get(styleId);
	if (!style) {
		return {};
	}

	seen.add(styleId);
	return mergeProps(resolveStyleProps(styles, style.basedOn ?? null, seen), style.props);
}

function getPartLabel(partPath: string): string {
	if (partPath === 'word/document.xml') {
		return 'Document body';
	}

	if (partPath.startsWith('word/headers/')) {
		return 'Header';
	}

	if (partPath.startsWith('word/footers/')) {
		return 'Footer';
	}

	if (partPath === 'word/footnotes.xml') {
		return 'Footnotes';
	}

	if (partPath === 'word/endnotes.xml') {
		return 'Endnotes';
	}

	if (partPath === 'word/comments.xml') {
		return 'Comments';
	}

	return partPath;
}

function isTextPart(path: string): boolean {
	return TEXT_PART_PATTERNS.some(pattern => pattern.test(path));
}

function sortTextParts(left: string, right: string): number {
	if (left === 'word/document.xml') {
		return -1;
	}
	if (right === 'word/document.xml') {
		return 1;
	}

	return left.localeCompare(right);
}

function getPromptInjectionSignals(text: string): string[] {
	return PROMPT_INJECTION_PATTERNS
		.filter(pattern => pattern.test(text))
		.map(pattern => pattern.source.replace(/\\/g, ''));
}

function getVisibilityReasons(props: TextVisibilityProps): string[] {
	const reasons: string[] = [];
	if (props.hidden) {
		reasons.push('Hidden text property');
	}
	if (props.webHidden) {
		reasons.push('Web-hidden text property');
	}
	if (props.color && isNearWhiteColor(props.color)) {
		reasons.push(`White or near-white font color (#${props.color})`);
	}
	if (props.fontSizeHalfPoints !== undefined && props.fontSizeHalfPoints > 0 && props.fontSizeHalfPoints <= TINY_FONT_SIZE_HALF_POINTS) {
		reasons.push(`Very small font size (${props.fontSizeHalfPoints / 2} pt)`);
	}

	return reasons;
}

function scanParagraph(
	paragraphXml: string,
	partPath: string,
	paragraphNumber: number,
	styles: Map<string, ParsedStyle>,
): HiddenTextFinding[] {
	const pPr = getParagraphProperties(paragraphXml);
	const paragraphStyleProps = resolveStyleProps(styles, getElementVal(pPr, 'pStyle'));
	const findings: HiddenTextFinding[] = [];
	let runIndex = 0;
	let match: RegExpExecArray | null;

	RUN_PATTERN.lastIndex = 0;

	while ((match = RUN_PATTERN.exec(paragraphXml)) !== null) {
		runIndex += 1;
		const runXml = match[0];
		const text = truncateText(extractTextFromXml(runXml));
		if (!text) {
			continue;
		}

		const rPr = getRunProperties(runXml);
		const styleProps = resolveStyleProps(styles, getElementVal(rPr, 'rStyle'));
		const props = mergeProps(mergeProps(paragraphStyleProps, styleProps), parseTextVisibilityProps(rPr));
		const reasons = getVisibilityReasons(props);
		const promptInjectionSignals = getPromptInjectionSignals(text);

		if (reasons.length === 0) {
			continue;
		}

		findings.push({
			id: `${partPath}:${paragraphNumber}:${runIndex}`,
			partPath,
			partLabel: getPartLabel(partPath),
			paragraphNumber,
			text,
			reasons,
			promptInjectionSignals,
		});
	}

	return findings;
}

function scanTextPart(partPath: string, xml: string, styles: Map<string, ParsedStyle>): HiddenTextFinding[] {
	const findings: HiddenTextFinding[] = [];
	let paragraphNumber = 0;
	let match: RegExpExecArray | null;

	PARAGRAPH_PATTERN.lastIndex = 0;

	while ((match = PARAGRAPH_PATTERN.exec(xml)) !== null) {
		paragraphNumber += 1;
		findings.push(...scanParagraph(match[0], partPath, paragraphNumber, styles));
	}

	return findings;
}

export async function findHiddenDocxText(buffer: ArrayBuffer): Promise<HiddenTextScanResult> {
	const zip = await JSZip.loadAsync(buffer.slice(0));
	const stylesXml = await zip.file('word/styles.xml')?.async('string') ?? '';
	const styles = parseStyles(stylesXml);
	const partPaths = Object.keys(zip.files)
		.filter(isTextPart)
		.sort(sortTextParts);
	const findings: HiddenTextFinding[] = [];

	for (const partPath of partPaths) {
		const xml = await zip.file(partPath)?.async('string');
		if (!xml) {
			continue;
		}

		findings.push(...scanTextPart(partPath, xml, styles));
	}

	return {
		findings,
		partsScanned: partPaths.length,
	};
}
