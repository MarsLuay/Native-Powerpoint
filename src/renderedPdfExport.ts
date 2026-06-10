import { debugLog, infoLog, warnLog } from './logger';
import { isText } from './domGuards';

const RENDERED_PDF_EXPORT_SCALE = 2;
const RENDERED_PDF_PAGE_READY_TIMEOUT_MS = 4000;
const RENDERED_PDF_MAX_TEXT_RUNS_PER_PAGE = 4500;
const PDF_POINTS_PER_CSS_PIXEL = 72 / 96;
const SELECTED_LIST_MARKER_CLASS = 'docxidian-list-marker-selected';
const LIST_PARAGRAPH_SELECTOR = '.layout-paragraph[data-pm-start]';
const LIST_MARKER_SELECTOR = '.layout-list-marker, .docx-list-marker';
const RENDERED_PDF_PM_SPAN_SELECTOR = 'span[data-pm-start][data-pm-end]';
const RENDERED_PDF_BODY_PM_SPAN_SELECTOR = `.layout-page-content ${RENDERED_PDF_PM_SPAN_SELECTOR}`;

function getCssGeneratedContentText(element: HTMLElement) {
	const generatedContent = window.getComputedStyle(element, '::before').content;
	if (!generatedContent || generatedContent === 'none' || generatedContent === 'normal') {
		return '';
	}

	if (
		(generatedContent.startsWith('"') && generatedContent.endsWith('"'))
		|| (generatedContent.startsWith("'") && generatedContent.endsWith("'"))
	) {
		return generatedContent.slice(1, -1).replace(/\\A/g, '\n').replace(/\\"/g, '"').replace(/\\'/g, "'");
	}

	return generatedContent;
}

function getRenderedListMarkerText(marker: HTMLElement) {
	const text = marker.textContent?.trimEnd() ?? '';
	if (text.trim().length > 0) {
		return text;
	}

	const generatedContentText = getCssGeneratedContentText(marker).trimEnd();
	return generatedContentText.trim().length > 0 ? generatedContentText : '';
}

interface RenderedPdfImagePage {
	imageBytes: Uint8Array;
	imageWidth: number;
	imageHeight: number;
	pdfWidth: number;
	pdfHeight: number;
	textRuns: RenderedPdfTextRun[];
}

interface RenderedPdfTextRun {
	text: string;
	x: number;
	y: number;
	fontSize: number;
	horizontalScale: number;
}

function escapeSvgText(value: string) {
	return value
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;');
}

function toCdata(value: string) {
	return value.replace(/]]>/g, ']]]]><![CDATA[>');
}

function serializeElementForSvg(element: HTMLElement) {
	const cloneDocument = activeDocument.implementation.createHTMLDocument('docxidian-pdf-export');
	const importedElement = cloneDocument.importNode(element, true);
	importedElement.setAttribute('xmlns', 'http://www.w3.org/1999/xhtml');
	return new XMLSerializer().serializeToString(importedElement);
}

function assertSvgCanParse(svg: string) {
	const documentParser = new DOMParser();
	const parsedSvg = documentParser.parseFromString(svg, 'image/svg+xml');
	const parseError = parsedSvg.querySelector('parsererror');
	if (parseError) {
		const message = parseError.textContent?.replace(/\s+/g, ' ').trim() ?? 'Unknown SVG parse error';
		throw new Error(`Could not prepare the DOCX page for PDF export: ${message.slice(0, 240)}`);
	}
}

function collectPageExportCss() {
	let css = '';
	const styleSheets = [
		...Array.from(activeDocument.styleSheets),
		...Array.from(activeDocument.adoptedStyleSheets ?? []),
	];
	for (const sheet of styleSheets) {
		try {
			for (const rule of Array.from(sheet.cssRules)) {
				css += `${rule.cssText}\n`;
			}
		} catch {
			// Skip browser-managed stylesheets that cannot be read.
		}
	}

	return css;
}

function dataUrlToBytes(dataUrl: string) {
	const commaIndex = dataUrl.indexOf(',');
	const base64 = commaIndex >= 0 ? dataUrl.slice(commaIndex + 1) : dataUrl;
	const binary = atob(base64);
	const bytes = new Uint8Array(binary.length);
	for (let index = 0; index < binary.length; index += 1) {
		bytes[index] = binary.charCodeAt(index);
	}
	return bytes;
}

function bytesToArrayBuffer(bytes: Uint8Array) {
	const copy = new Uint8Array(bytes.byteLength);
	copy.set(bytes);
	return copy.buffer;
}

function formatPdfNumber(value: number) {
	return Number.isFinite(value)
		? String(Math.round(value * 1000) / 1000)
		: '0';
}

let renderedPdfMeasureCanvas: HTMLCanvasElement | null = null;

function measureRenderedPdfTextCssWidth(text: string, fontSizePx: number, fontFamily: string, fontWeight: string, fontStyle: string) {
	renderedPdfMeasureCanvas ??= activeDocument.createElement('canvas');
	const context = renderedPdfMeasureCanvas.getContext('2d');
	if (!context) {
		return 0;
	}

	context.font = `${fontStyle || 'normal'} ${fontWeight || '400'} ${fontSizePx}px ${fontFamily || 'Helvetica, Arial, sans-serif'}`;
	return context.measureText(text).width;
}

function encodePdfTextLiteral(value: string) {
	const winAnsiMap = new Map<string, number>([
		['•', 0x95],
		['▪', 0x95],
		['●', 0x95],
		['–', 0x96],
		['—', 0x97],
		['‘', 0x91],
		['’', 0x92],
		['“', 0x93],
		['”', 0x94],
		['…', 0x85],
		['™', 0x99],
		['©', 0xa9],
		['®', 0xae],
		['°', 0xb0],
		['±', 0xb1],
		['×', 0xd7],
		['÷', 0xf7],
	]);
	let output = '(';

	for (const char of value) {
		let code = char.codePointAt(0) ?? 0x3f;
		if (char === '\n' || char === '\r' || char === '\t') {
			code = 0x20;
		} else if (winAnsiMap.has(char)) {
			code = winAnsiMap.get(char)!;
		} else if (code < 0x20 || code > 0x7e) {
			code = 0x3f;
		}

		if (code === 0x28 || code === 0x29 || code === 0x5c) {
			output += `\\${String.fromCharCode(code)}`;
		} else if (code < 0x20 || code > 0x7e) {
			output += `\\${code.toString(8).padStart(3, '0')}`;
		} else {
			output += String.fromCharCode(code);
		}
	}

	return `${output})`;
}

function loadImageFromUrl(url: string) {
	return new Promise<HTMLImageElement>((resolve, reject) => {
		const image = new Image();
		image.onload = () => resolve(image);
		image.onerror = () => reject(new Error('Could not render the DOCX page for PDF export.'));
		image.src = url;
	});
}

function parseCssPixelValue(value: string) {
	const numericValue = Number.parseFloat(value);
	return Number.isFinite(numericValue) ? numericValue : 0;
}

function getElementExportSize(element: HTMLElement) {
	const rect = element.getBoundingClientRect();
	let width = Math.ceil(rect.width);
	let height = Math.ceil(rect.height);

	if (width <= 0) {
		width = element.offsetWidth;
	}
	if (height <= 0) {
		height = element.offsetHeight;
	}

	if (width <= 0 || height <= 0) {
		const style = window.getComputedStyle(element);
		width = width > 0 ? width : Math.ceil(parseCssPixelValue(style.width));
		height = height > 0 ? height : Math.ceil(parseCssPixelValue(style.height));
	}

	if (width <= 0 || height <= 0) {
		const content = element.querySelector<HTMLElement>('.layout-page-content');
		if (content) {
			const contentRect = content.getBoundingClientRect();
			width = width > 0 ? width : Math.ceil(contentRect.width || content.offsetWidth);
			height = height > 0 ? height : Math.ceil(contentRect.height || content.offsetHeight);
		}
	}

	return { width, height };
}

function parseCssRgbColor(value: string) {
	const match = value.match(/rgba?\(([^)]+)\)/i);
	if (!match) {
		return null;
	}

	const parts = match[1]!.split(',').map(part => part.trim());
	const red = Number.parseFloat(parts[0] ?? '');
	const green = Number.parseFloat(parts[1] ?? '');
	const blue = Number.parseFloat(parts[2] ?? '');
	const alpha = parts[3] === undefined ? 1 : Number.parseFloat(parts[3]);
	if (![red, green, blue, alpha].every(Number.isFinite)) {
		return null;
	}

	return { red, green, blue, alpha };
}

function isNearlyWhiteColor(color: { red: number; green: number; blue: number; alpha: number }) {
	return color.alpha > 0.8 && color.red >= 245 && color.green >= 245 && color.blue >= 245;
}

function hasNonWhiteBackground(element: HTMLElement, stopAt: HTMLElement) {
	let current: HTMLElement | null = element;
	while (current && stopAt.contains(current)) {
		const backgroundColor = parseCssRgbColor(window.getComputedStyle(current).backgroundColor);
		if (backgroundColor && backgroundColor.alpha > 0.1 && !isNearlyWhiteColor(backgroundColor)) {
			return true;
		}

		if (current === stopAt) {
			break;
		}
		current = current.parentElement;
	}

	return false;
}

function shouldIncludePdfTextElement(element: HTMLElement, page: HTMLElement) {
	if (element.closest('script, style, textarea, input, select')) {
		return false;
	}

	const style = window.getComputedStyle(element);
	if (
		style.display === 'none'
		|| style.visibility === 'hidden'
		|| Number.parseFloat(style.opacity || '1') <= 0.05
		|| parseCssPixelValue(style.fontSize) <= 0
	) {
		return false;
	}

	const color = parseCssRgbColor(style.color);
	if (color && color.alpha <= 0.05) {
		return false;
	}

	return !color || !isNearlyWhiteColor(color) || hasNonWhiteBackground(element, page);
}

function getPdfTextRunFromRect(text: string, rect: DOMRect, pageRect: DOMRect, page: HTMLElement, pdfWidth: number, pdfHeight: number, fontSizePx: number, fontFamily = 'Helvetica, Arial, sans-serif', fontWeight = '400', fontStyle = 'normal'): RenderedPdfTextRun | null {
	const trimmedText = text.replace(/\s+/g, ' ');
	if (!trimmedText.trim() || rect.width <= 0 || rect.height <= 0) {
		return null;
	}

	const cssPageSize = getElementExportSize(page);
	const scaleX = pdfWidth / Math.max(1, cssPageSize.width);
	const scaleY = pdfHeight / Math.max(1, cssPageSize.height);
	const left = Math.max(0, rect.left - pageRect.left);
	const top = Math.max(0, rect.top - pageRect.top);
	const fontSize = Math.max(2, Math.min(96, fontSizePx * scaleY));
	const baselineY = pdfHeight - ((top + (rect.height * 0.82)) * scaleY);
	const measuredCssWidth = measureRenderedPdfTextCssWidth(trimmedText, fontSizePx, fontFamily, fontWeight, fontStyle);
	const measuredPdfWidth = measuredCssWidth * scaleX;
	const targetPdfWidth = rect.width * scaleX;
	const horizontalScale = measuredPdfWidth > 0
		? Math.max(35, Math.min(100, (targetPdfWidth / measuredPdfWidth) * 100))
		: 100;

	return {
		text: trimmedText,
		x: left * scaleX,
		y: baselineY,
		fontSize,
		horizontalScale,
	};
}

function getCombinedClientRect(rects: DOMRect[]) {
	const left = Math.min(...rects.map(rect => rect.left));
	const top = Math.min(...rects.map(rect => rect.top));
	const right = Math.max(...rects.map(rect => rect.right));
	const bottom = Math.max(...rects.map(rect => rect.bottom));
	return new DOMRect(left, top, right - left, bottom - top);
}

function areClientRectsOnSameLine(rects: DOMRect[]) {
	if (rects.length <= 1) {
		return true;
	}

	const firstRect = rects[0]!;
	return rects.every(rect => Math.abs(rect.top - firstRect.top) <= 2 && Math.abs(rect.bottom - firstRect.bottom) <= 2);
}

function collectTextNodePdfRuns(textNode: Text, page: HTMLElement, pageRect: DOMRect, pdfWidth: number, pdfHeight: number) {
	const parent = textNode.parentElement;
	if (!parent || !shouldIncludePdfTextElement(parent, page)) {
		return [];
	}

	const text = textNode.nodeValue ?? '';
	if (!text.trim()) {
		return [];
	}

	const style = window.getComputedStyle(parent);
	const fontSizePx = parseCssPixelValue(style.fontSize) || 12;
	const runs: RenderedPdfTextRun[] = [];
	const tokenPattern = /\S+\s*/g;
	let match: RegExpExecArray | null;

	while ((match = tokenPattern.exec(text)) !== null) {
		const token = match[0];
		const start = match.index;
		const end = start + token.length;
		const range = activeDocument.createRange();
		try {
			range.setStart(textNode, start);
			range.setEnd(textNode, end);
			const rects = Array.from(range.getClientRects()).filter(rect => rect.width > 0 && rect.height > 0);
			const runRects = areClientRectsOnSameLine(rects) && rects.length > 1
				? [getCombinedClientRect(rects)]
				: rects;
			for (const rect of runRects) {
				const run = getPdfTextRunFromRect(token, rect, pageRect, page, pdfWidth, pdfHeight, fontSizePx, style.fontFamily, style.fontWeight, style.fontStyle);
				if (run) {
					runs.push(run);
				}
			}
		} finally {
			range.detach();
		}
	}

	return runs;
}

function collectWholeTextNodePdfRun(textNode: Text, page: HTMLElement, pageRect: DOMRect, pdfWidth: number, pdfHeight: number) {
	const parent = textNode.parentElement;
	if (!parent || !shouldIncludePdfTextElement(parent, page)) {
		return [];
	}

	const text = textNode.nodeValue ?? '';
	if (!text.trim()) {
		return [];
	}

	const style = window.getComputedStyle(parent);
	const fontSizePx = parseCssPixelValue(style.fontSize) || 12;
	const range = activeDocument.createRange();
	try {
		range.setStart(textNode, 0);
		range.setEnd(textNode, text.length);
		const rects = Array.from(range.getClientRects()).filter(rect => rect.width > 0 && rect.height > 0);
		if (rects.length === 0 || !areClientRectsOnSameLine(rects)) {
			return [];
		}

		const rect = rects.length === 1 ? rects[0]! : getCombinedClientRect(rects);
		const run = getPdfTextRunFromRect(text, rect, pageRect, page, pdfWidth, pdfHeight, fontSizePx, style.fontFamily, style.fontWeight, style.fontStyle);
		return run ? [run] : [];
	} finally {
		range.detach();
	}
}

function collectRenderedSpanPdfRuns(page: HTMLElement, pageRect: DOMRect, pdfWidth: number, pdfHeight: number) {
	const runs: RenderedPdfTextRun[] = [];
	const spans = page.querySelectorAll<HTMLElement>(RENDERED_PDF_BODY_PM_SPAN_SELECTOR);
	spans.forEach((span) => {
		if (runs.length >= RENDERED_PDF_MAX_TEXT_RUNS_PER_PAGE || !shouldIncludePdfTextElement(span, page)) {
			return;
		}

		const style = window.getComputedStyle(span);
		const fontSizePx = parseCssPixelValue(style.fontSize) || 12;
		if (span.classList.contains('layout-run-tab')) {
			const run = getPdfTextRunFromRect(' ', span.getBoundingClientRect(), pageRect, page, pdfWidth, pdfHeight, fontSizePx, style.fontFamily, style.fontWeight, style.fontStyle);
			if (run) {
				runs.push(run);
			}
			return;
		}

		const textNode = span.firstChild;
		if (isText(textNode)) {
			const wholeTextRuns = collectWholeTextNodePdfRun(textNode, page, pageRect, pdfWidth, pdfHeight);
			runs.push(...(wholeTextRuns.length > 0 ? wholeTextRuns : collectTextNodePdfRuns(textNode, page, pageRect, pdfWidth, pdfHeight)));
			return;
		}

		const text = span.textContent ?? '';
		if (!text.trim()) {
			return;
		}

		const run = getPdfTextRunFromRect(text, span.getBoundingClientRect(), pageRect, page, pdfWidth, pdfHeight, fontSizePx, style.fontFamily, style.fontWeight, style.fontStyle);
		if (run) {
			runs.push(run);
		}
	});

	return runs.slice(0, RENDERED_PDF_MAX_TEXT_RUNS_PER_PAGE);
}

function collectGeneratedListMarkerPdfRuns(page: HTMLElement, pageRect: DOMRect, pdfWidth: number, pdfHeight: number) {
	const runs: RenderedPdfTextRun[] = [];
	const markerElements = page.querySelectorAll<HTMLElement>(LIST_MARKER_SELECTOR);
	markerElements.forEach((marker) => {
		if (!shouldIncludePdfTextElement(marker, page) || marker.querySelector(RENDERED_PDF_PM_SPAN_SELECTOR)) {
			return;
		}

		const markerText = getRenderedListMarkerText(marker);
		const paragraph = marker.closest<HTMLElement>(LIST_PARAGRAPH_SELECTOR);
		const firstParagraphText = paragraph?.querySelector<HTMLElement>(RENDERED_PDF_PM_SPAN_SELECTOR)?.textContent?.trimStart() ?? '';
		if (
			markerText.length === 0
			|| (markerText.trim().length > 0 && firstParagraphText.startsWith(markerText.trim()))
			|| /^[•▪●◦]/.test(firstParagraphText)
		) {
			return;
		}

		const rect = marker.getBoundingClientRect();
		const style = window.getComputedStyle(marker);
		const fontSizePx = parseCssPixelValue(style.fontSize) || 12;
		const run = getPdfTextRunFromRect(markerText, rect, pageRect, page, pdfWidth, pdfHeight, fontSizePx, style.fontFamily, style.fontWeight, style.fontStyle);
		if (run) {
			runs.push(run);
		}
	});
	return runs;
}

function collectRenderedPdfTextRuns(page: HTMLElement, pdfWidth: number, pdfHeight: number) {
	const pageRect = page.getBoundingClientRect();
	const runs: RenderedPdfTextRun[] = collectRenderedSpanPdfRuns(page, pageRect, pdfWidth, pdfHeight);

	if (runs.length === 0) {
		const walker = activeDocument.createTreeWalker(page, NodeFilter.SHOW_TEXT);
		let node = walker.nextNode();

		while (node && runs.length < RENDERED_PDF_MAX_TEXT_RUNS_PER_PAGE) {
				if (isText(node)) {
					runs.push(...collectTextNodePdfRuns(node, page, pageRect, pdfWidth, pdfHeight));
			}
			node = walker.nextNode();
		}
	}

	if (runs.length < RENDERED_PDF_MAX_TEXT_RUNS_PER_PAGE) {
		runs.push(...collectGeneratedListMarkerPdfRuns(page, pageRect, pdfWidth, pdfHeight));
	}

	return runs.slice(0, RENDERED_PDF_MAX_TEXT_RUNS_PER_PAGE);
}

function getRenderedPageElements(container: HTMLElement) {
	return Array.from(container.querySelectorAll<HTMLElement>('.layout-page'))
		.filter((page) => {
			const { width, height } = getElementExportSize(page);
			return width > 0 && height > 0;
		});
}

function dedupeElements(elements: HTMLElement[]) {
	return Array.from(new Set(elements));
}

function describeExportContainer(container: HTMLElement) {
	const rect = container.getBoundingClientRect();
	return {
		className: container.className,
		childElementCount: container.childElementCount,
		isConnected: container.isConnected,
		layoutPageCount: container.querySelectorAll('.layout-page').length,
		layoutPageContentCount: container.querySelectorAll('.layout-page-content').length,
		rect: {
			width: Math.round(rect.width),
			height: Math.round(rect.height),
		},
	};
}

function waitForRenderedPdfPageElements(containers: HTMLElement[]) {
	const startedAt = performance.now();

	return new Promise<{ pages: HTMLElement[]; waitedMs: number }>((resolve) => {
		const poll = () => {
			const pages = dedupeElements(containers.flatMap(container => getRenderedPageElements(container)));
			if (pages.length > 0 || performance.now() - startedAt >= RENDERED_PDF_PAGE_READY_TIMEOUT_MS) {
				resolve({
					pages,
					waitedMs: Math.round(performance.now() - startedAt),
				});
				return;
			}

			window.setTimeout(poll, 50);
		};

		poll();
	});
}

function createRenderedPdfPageFromCanvas(canvas: HTMLCanvasElement, pageWidth: number, pageHeight: number, textRuns: RenderedPdfTextRun[]): RenderedPdfImagePage {
	return {
		imageBytes: dataUrlToBytes(canvas.toDataURL('image/jpeg', 0.96)),
		imageWidth: canvas.width,
		imageHeight: canvas.height,
		pdfWidth: Math.round(pageWidth * PDF_POINTS_PER_CSS_PIXEL * 100) / 100,
		pdfHeight: Math.round(pageHeight * PDF_POINTS_PER_CSS_PIXEL * 100) / 100,
		textRuns,
	};
}

async function renderPageElementToCanvasJpeg(page: HTMLElement) {
	const { width: pageWidth, height: pageHeight } = getElementExportSize(page);
	if (pageWidth <= 0 || pageHeight <= 0) {
		return null;
	}

	const { default: html2canvas } = await import('html2canvas');
	const canvas = await html2canvas(page, {
		allowTaint: true,
		backgroundColor: '#ffffff',
		logging: false,
		removeContainer: true,
		scale: RENDERED_PDF_EXPORT_SCALE,
		useCORS: true,
		width: pageWidth,
		height: pageHeight,
		windowWidth: Math.max(activeDocument.documentElement.clientWidth, pageWidth),
		windowHeight: Math.max(activeDocument.documentElement.clientHeight, pageHeight),
		onclone: (_clonedDocument, clonedPage) => {
			clonedPage.classList.add('docxidian-pdf-export-page');
			// Neutralize the on-screen zoom transform inline: stylesheet rules
			// (even high-specificity ones) cannot override an inline transform.
			clonedPage.setCssProps({
				transform: 'none',
				transformOrigin: 'top left',
				margin: '0',
				boxShadow: 'none',
			});
			clonedPage.querySelectorAll<HTMLElement>(`.${SELECTED_LIST_MARKER_CLASS}`).forEach((marker) => {
				marker.classList.remove(SELECTED_LIST_MARKER_CLASS);
			});
		},
	});

	if (canvas.width <= 0 || canvas.height <= 0) {
		throw new Error('Could not create a PDF export canvas.');
	}

	return createRenderedPdfPageFromCanvas(canvas, pageWidth, pageHeight, collectRenderedPdfTextRuns(page, pageWidth * PDF_POINTS_PER_CSS_PIXEL, pageHeight * PDF_POINTS_PER_CSS_PIXEL));
}

async function renderPageElementToSvgJpeg(page: HTMLElement, editorRoot: HTMLElement, cssText: string) {
	const { width: pageWidth, height: pageHeight } = getElementExportSize(page);
	if (pageWidth <= 0 || pageHeight <= 0) {
		return null;
	}

	const clone = page.cloneNode(true) as HTMLElement;
	clone.classList.add('docxidian-pdf-export-page');

	const exportCss = [
		cssText,
		'html, body { margin: 0; padding: 0; width: 100%; height: 100%; overflow: hidden; background: #ffffff; }',
		'.docxidian-pdf-export-root { margin: 0; padding: 0; background: #ffffff; color: #000000; }',
		'.docxidian-pdf-export-root .paged-editor__pages { margin: 0 !important; padding: 0 !important; display: block !important; }',
		'.docxidian-pdf-export-page { margin: 0 !important; box-shadow: none !important; transform: none !important; transform-origin: top left !important; }',
		'.docxidian-pdf-export-root * { animation: none !important; transition: none !important; caret-color: transparent !important; }',
		`.docxidian-pdf-export-root .${SELECTED_LIST_MARKER_CLASS} { background: transparent !important; outline: none !important; }`,
	].join('\n');

	const rootClassName = escapeSvgText(editorRoot.className);
	const pageHtml = serializeElementForSvg(clone);
	const svg = [
		`<svg xmlns="http://www.w3.org/2000/svg" width="${pageWidth}" height="${pageHeight}" viewBox="0 0 ${pageWidth} ${pageHeight}">`,
		'<style><![CDATA[',
		toCdata(exportCss),
		']]></style>',
		`<foreignObject x="0" y="0" width="${pageWidth}" height="${pageHeight}">`,
		`<div xmlns="http://www.w3.org/1999/xhtml" class="${rootClassName} docxidian-pdf-export-root" style="width:${pageWidth}px;height:${pageHeight}px;">`,
		'<div class="paged-editor__pages">',
		pageHtml,
		'</div>',
		'</div>',
		'</foreignObject>',
		'</svg>',
	].join('');
	assertSvgCanParse(svg);

	const svgUrl = URL.createObjectURL(new Blob([svg], { type: 'image/svg+xml;charset=utf-8' }));
	try {
		const image = await loadImageFromUrl(svgUrl);
		const canvas = activeDocument.createElement('canvas');
		canvas.width = Math.max(1, Math.round(pageWidth * RENDERED_PDF_EXPORT_SCALE));
		canvas.height = Math.max(1, Math.round(pageHeight * RENDERED_PDF_EXPORT_SCALE));
		const context = canvas.getContext('2d');
		if (!context) {
			throw new Error('Could not create a PDF export canvas.');
		}

		context.fillStyle = '#ffffff';
		context.fillRect(0, 0, canvas.width, canvas.height);
		context.drawImage(image, 0, 0, canvas.width, canvas.height);
		return createRenderedPdfPageFromCanvas(canvas, pageWidth, pageHeight, collectRenderedPdfTextRuns(page, pageWidth * PDF_POINTS_PER_CSS_PIXEL, pageHeight * PDF_POINTS_PER_CSS_PIXEL));
	} finally {
		URL.revokeObjectURL(svgUrl);
	}
}

async function renderPageElementToJpeg(page: HTMLElement, editorRoot: HTMLElement, cssText: string) {
	try {
		const canvasPage = await renderPageElementToCanvasJpeg(page);
		if (canvasPage) {
			debugLog('export', 'Rendered PDF page with html2canvas', {
				imageWidth: canvasPage.imageWidth,
				imageHeight: canvasPage.imageHeight,
				pdfWidth: canvasPage.pdfWidth,
				pdfHeight: canvasPage.pdfHeight,
			});
			return canvasPage;
		}
	} catch (canvasError) {
		warnLog('export', 'html2canvas PDF page render failed; retrying with SVG renderer', canvasError);
	}

	return renderPageElementToSvgJpeg(page, editorRoot, cssText);
}

function createRenderedPdfContentStream(page: RenderedPdfImagePage, imageName: string) {
	const lines = [
		'q',
		`${formatPdfNumber(page.pdfWidth)} 0 0 ${formatPdfNumber(page.pdfHeight)} 0 0 cm`,
		`/${imageName} Do`,
		'Q',
	];

	if (page.textRuns.length > 0) {
		lines.push('BT', '/Ftxt 1 Tf', '3 Tr');
		for (const run of page.textRuns) {
			lines.push(
				`${formatPdfNumber(run.horizontalScale)} Tz`,
				`${formatPdfNumber(run.fontSize)} 0 0 ${formatPdfNumber(run.fontSize)} ${formatPdfNumber(run.x)} ${formatPdfNumber(run.y)} Tm`,
				`${encodePdfTextLiteral(run.text)} Tj`,
			);
		}
		lines.push('100 Tz', '0 Tr', 'ET');
	}

	lines.push('');
	return lines.join('\n');
}

function createRenderedImagePdf(pages: RenderedPdfImagePage[]) {
	const encoder = new TextEncoder();
	const chunks: Uint8Array[] = [];
	const offsets: number[] = [0];
	let byteLength = 0;
	const appendBytes = (bytes: Uint8Array) => {
		chunks.push(bytes);
		byteLength += bytes.byteLength;
	};
	const appendText = (text: string) => appendBytes(encoder.encode(text));
	const addObject = (objectNumber: number, body: string) => {
		offsets[objectNumber] = byteLength;
		appendText(`${objectNumber} 0 obj\n${body}\nendobj\n`);
	};
	const addStreamObject = (objectNumber: number, dictionary: string, stream: Uint8Array) => {
		offsets[objectNumber] = byteLength;
		appendText(`${objectNumber} 0 obj\n${dictionary} /Length ${stream.byteLength} >>\nstream\n`);
		appendBytes(stream);
		appendText('\nendstream\nendobj\n');
	};

	const fontObjectNumber = 3;
	const pageObjectNumbers = pages.map((_, index) => 4 + index * 3);
	appendText('%PDF-1.4\n');
	addObject(1, '<< /Type /Catalog /Pages 2 0 R >>');
	addObject(2, `<< /Type /Pages /Count ${pages.length} /Kids [${pageObjectNumbers.map(pageObjectNumber => `${pageObjectNumber} 0 R`).join(' ')}] >>`);
	addObject(fontObjectNumber, '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>');

	pages.forEach((page, index) => {
		const pageObjectNumber = pageObjectNumbers[index]!;
		const contentObjectNumber = pageObjectNumber + 1;
		const imageObjectNumber = pageObjectNumber + 2;
		const imageName = `Im${index + 1}`;
		const content = encoder.encode(createRenderedPdfContentStream(page, imageName));

		addObject(pageObjectNumber, [
			'<< /Type /Page',
			'/Parent 2 0 R',
			`/MediaBox [0 0 ${formatPdfNumber(page.pdfWidth)} ${formatPdfNumber(page.pdfHeight)}]`,
			`/Resources << /XObject << /${imageName} ${imageObjectNumber} 0 R >> /Font << /Ftxt ${fontObjectNumber} 0 R >> >>`,
			`/Contents ${contentObjectNumber} 0 R`,
			'>>',
		].join(' '));
		addStreamObject(contentObjectNumber, '<<', content);
		addStreamObject(
			imageObjectNumber,
			`<< /Type /XObject /Subtype /Image /Width ${page.imageWidth} /Height ${page.imageHeight} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode`,
			page.imageBytes,
		);
	});

	const xrefOffset = byteLength;
	const objectCount = 4 + pages.length * 3;
	appendText(`xref\n0 ${objectCount}\n`);
	appendText('0000000000 65535 f\n');
	for (let objectNumber = 1; objectNumber < objectCount; objectNumber += 1) {
		appendText(`${String(offsets[objectNumber] ?? 0).padStart(10, '0')} 00000 n\n`);
	}
	appendText(`trailer\n<< /Size ${objectCount} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF\n`);

	const output = new Uint8Array(byteLength);
	let offset = 0;
	for (const chunk of chunks) {
		output.set(chunk, offset);
		offset += chunk.byteLength;
	}
	return bytesToArrayBuffer(output);
}

export async function exportRenderedPagesToPdf(editorRoot: HTMLElement, renderedPagesContainer?: HTMLElement | null) {
	await activeDocument.fonts?.ready;
	const pagesContainer = renderedPagesContainer?.isConnected ? renderedPagesContainer : editorRoot.querySelector<HTMLElement>('.paged-editor__pages');
	const candidateContainers = dedupeElements([
		...(pagesContainer ? [pagesContainer] : []),
		editorRoot,
	].filter((container): container is HTMLElement => container.isConnected));
	debugLog('export', 'Preparing rendered PDF export', {
		hasRenderedPagesContainer: Boolean(renderedPagesContainer),
		containers: candidateContainers.map(describeExportContainer),
	});

	const { pages: pageElements, waitedMs } = await waitForRenderedPdfPageElements(candidateContainers);

	if (pageElements.length === 0) {
		warnLog('export', 'No rendered DOCX pages found for PDF export', {
			waitedMs,
			containers: candidateContainers.map(describeExportContainer),
		});
		return null;
	}

	const cssText = collectPageExportCss();
	const renderedPages: RenderedPdfImagePage[] = [];
	for (const pageElement of pageElements) {
		const renderedPage = await renderPageElementToJpeg(pageElement, editorRoot, cssText);
		if (renderedPage) {
			renderedPages.push(renderedPage);
		}
	}

	if (renderedPages.length === 0) {
		warnLog('export', 'Rendered DOCX pages had no exportable dimensions', {
			pageCount: pageElements.length,
			pageSizes: pageElements.map(getElementExportSize),
		});
		return null;
	}

	infoLog('export', 'Rendered DOCX pages for PDF export', {
		pageCount: renderedPages.length,
		textRuns: renderedPages.reduce((total, page) => total + page.textRuns.length, 0),
		waitedMs,
	});
	return createRenderedImagePdf(renderedPages);
}
