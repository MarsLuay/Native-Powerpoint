import JSZip from 'jszip';

import {
  createRenderedImagePdf,
  dataUrlToBytes,
  type RenderedPdfImagePage
} from './renderedPdfExport';

export const SLIDE_PNG_EXPORT_SCALE = 2;
export const SLIDE_PDF_EXPORT_SCALE = 2;

const PDF_POINTS_PER_CSS_PIXEL = 72 / 96;
const FALLBACK_SLIDE_WIDTH = 1280;
const FALLBACK_SLIDE_HEIGHT = 720;

interface SlideExportSize {
  width: number;
  height: number;
}

function bytesToArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  const copy = new Uint8Array(bytes.byteLength);
  copy.set(bytes);
  return copy.buffer;
}

function getSvgExportSize(svg: SVGSVGElement): SlideExportSize {
  const viewBox = svg.viewBox?.baseVal;
  if (viewBox && viewBox.width > 0 && viewBox.height > 0) {
    return { width: viewBox.width, height: viewBox.height };
  }

  const width = Number.parseFloat(svg.getAttribute('width') ?? '');
  const height = Number.parseFloat(svg.getAttribute('height') ?? '');
  if (Number.isFinite(width) && Number.isFinite(height) && width > 0 && height > 0) {
    return { width, height };
  }

  return { width: FALLBACK_SLIDE_WIDTH, height: FALLBACK_SLIDE_HEIGHT };
}

function serializeSvgForRaster(svg: SVGSVGElement, size: SlideExportSize, pixelWidth: number, pixelHeight: number): string {
  const clone = svg.cloneNode(true) as SVGSVGElement;
  clone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
  clone.setAttribute('xmlns:xlink', 'http://www.w3.org/1999/xlink');
  clone.setAttribute('width', String(pixelWidth));
  clone.setAttribute('height', String(pixelHeight));
  if (!clone.getAttribute('viewBox')) {
    clone.setAttribute('viewBox', `0 0 ${size.width} ${size.height}`);
  }
  clone.setAttribute('preserveAspectRatio', 'xMidYMid meet');
  return new XMLSerializer().serializeToString(clone);
}

function loadImageFromUrl(url: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.addEventListener('load', () => resolve(image), { once: true });
    image.addEventListener('error', () => reject(new Error('Could not rasterize the slide for export.')), { once: true });
    image.src = url;
  });
}

async function rasterizeSlideToCanvas(
  svg: SVGSVGElement,
  ownerDocument: Document,
  scale: number
): Promise<{ canvas: HTMLCanvasElement; size: SlideExportSize }> {
  const size = getSvgExportSize(svg);
  const pixelWidth = Math.max(1, Math.round(size.width * scale));
  const pixelHeight = Math.max(1, Math.round(size.height * scale));
  const svgString = serializeSvgForRaster(svg, size, pixelWidth, pixelHeight);
  const url = URL.createObjectURL(new Blob([svgString], { type: 'image/svg+xml;charset=utf-8' }));

  try {
    const image = await loadImageFromUrl(url);
    const canvas = ownerDocument.createElement('canvas');
    canvas.width = pixelWidth;
    canvas.height = pixelHeight;
    const context = canvas.getContext('2d');
    if (!context) {
      throw new Error('Could not create a slide export canvas.');
    }

    context.fillStyle = '#ffffff';
    context.fillRect(0, 0, canvas.width, canvas.height);
    context.drawImage(image, 0, 0, canvas.width, canvas.height);
    return { canvas, size };
  } finally {
    URL.revokeObjectURL(url);
  }
}

function canvasToPngBytes(canvas: HTMLCanvasElement): Uint8Array {
  return dataUrlToBytes(canvas.toDataURL('image/png'));
}

export async function exportSlideToPng(
  svg: SVGSVGElement,
  ownerDocument: Document,
  scale: number = SLIDE_PNG_EXPORT_SCALE
): Promise<ArrayBuffer> {
  const { canvas } = await rasterizeSlideToCanvas(svg, ownerDocument, scale);
  return bytesToArrayBuffer(canvasToPngBytes(canvas));
}

export async function exportSlidesToPdf(
  svgs: SVGSVGElement[],
  ownerDocument: Document,
  scale: number = SLIDE_PDF_EXPORT_SCALE
): Promise<ArrayBuffer> {
  const pages: RenderedPdfImagePage[] = [];

  for (const svg of svgs) {
    const { canvas, size } = await rasterizeSlideToCanvas(svg, ownerDocument, scale);
    pages.push({
      imageBytes: dataUrlToBytes(canvas.toDataURL('image/jpeg', 0.95)),
      imageWidth: canvas.width,
      imageHeight: canvas.height,
      pdfWidth: Math.round(size.width * PDF_POINTS_PER_CSS_PIXEL * 100) / 100,
      pdfHeight: Math.round(size.height * PDF_POINTS_PER_CSS_PIXEL * 100) / 100,
      textRuns: []
    });
  }

  if (pages.length === 0) {
    throw new Error('There are no slides to export.');
  }

  return createRenderedImagePdf(pages);
}

export async function exportSlidesToPngZip(
  svgs: SVGSVGElement[],
  ownerDocument: Document,
  baseName: string,
  scale: number = SLIDE_PNG_EXPORT_SCALE
): Promise<ArrayBuffer> {
  if (svgs.length === 0) {
    throw new Error('There are no slides to export.');
  }

  const zip = new JSZip();
  const padLength = String(svgs.length).length;
  const safeBaseName = baseName.replace(/[\\/:*?"<>|]/g, '_') || 'presentation';

  for (let index = 0; index < svgs.length; index += 1) {
    const { canvas } = await rasterizeSlideToCanvas(svgs[index]!, ownerDocument, scale);
    const slideNumber = String(index + 1).padStart(padLength, '0');
    zip.file(`${safeBaseName}-slide-${slideNumber}.png`, canvasToPngBytes(canvas));
  }

  return zip.generateAsync({ type: 'arraybuffer' });
}
