import JSZip from 'jszip';

const REVIEW_MARKUP_PATTERN = /<w:(?:ins|del|moveFrom|moveTo|commentRangeStart|commentReference)\b/;
const COMMENT_PATTERN = /<w:comment\b/;

export async function hasReviewMarkup(buffer: ArrayBuffer) {
	try {
		const zip = await JSZip.loadAsync(buffer.slice(0));
		const [documentXml, commentsXml] = await Promise.all([
			zip.file('word/document.xml')?.async('string') ?? '',
			zip.file('word/comments.xml')?.async('string') ?? '',
		]);

		return REVIEW_MARKUP_PATTERN.test(documentXml) || COMMENT_PATTERN.test(commentsXml);
	} catch {
		return false;
	}
}
