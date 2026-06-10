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

function normalizeExtractedText(value: string): string {
	return value
		.replace(/\r/g, '\n')
		.replace(/[ \t]+\n/g, '\n')
		.replace(/\n[ \t]+/g, '\n')
		.replace(/\n{3,}/g, '\n\n')
		.trim();
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

	return normalizeExtractedText(pieces.join(''));
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

export async function extractDocxText(buffer: ArrayBuffer): Promise<string> {
	const zip = await JSZip.loadAsync(buffer.slice(0));
	const partPaths = Object.keys(zip.files)
		.filter(isTextPart)
		.sort(sortTextParts);
	const textParts: string[] = [];

	for (const partPath of partPaths) {
		const xml = await zip.file(partPath)?.async('string');
		if (!xml) {
			continue;
		}

		const text = extractTextFromXml(xml);
		if (text) {
			textParts.push(text);
		}
	}

	return normalizeExtractedText(textParts.join('\n\n'));
}
