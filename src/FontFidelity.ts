import { DEFAULT_FONT_FALLBACKS, type FontFallbackMap } from 'pptx-svg';

export interface FontSubstitution {
  requested: string;
  substitute: string;
}

interface FontResolution {
  requested: string;
  resolved: string;
  cssFamily: string;
}

interface FontFidelityOptions {
  isFontAvailable?: (fontFamily: string) => boolean;
  measureText?: (text: string, fontFamily: string, fontSizePx: number) => number;
  fontFallbacks?: FontFallbackMap;
}

type FontCategory = 'monospace' | 'sans-serif' | 'serif';

const GENERIC_FONT_FAMILIES = new Set(['cursive', 'fantasy', 'monospace', 'sans-serif', 'serif', 'system-ui']);
const FONT_DETECTION_SAMPLES = [
  'mmmmmmmmmmlli',
  'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz',
  '0123456789()[]{}'
];
const FONT_DETECTION_SIZE_PX = 72;

export const POWERPOINT_FONT_FALLBACKS: FontFallbackMap = {
  ...DEFAULT_FONT_FALLBACKS,
  Aptos: ['Arial', 'Helvetica Neue', 'Helvetica'],
  'Aptos Display': ['Arial', 'Helvetica Neue', 'Helvetica'],
  'Arial Narrow': ['Arial', 'Helvetica', 'sans-serif'],
  'Calibri Light': ['Calibri', 'Arial', 'Helvetica'],
  Consolas: ['Courier New', 'monospace'],
  Garamond: ['Georgia', 'Times New Roman', 'serif'],
  Georgia: ['Times New Roman', 'serif'],
  Helvetica: ['Arial', 'sans-serif'],
  'Helvetica Neue': ['Arial', 'Helvetica', 'sans-serif'],
  Tahoma: ['Arial', 'Helvetica', 'sans-serif'],
  'Times New Roman': ['Georgia', 'serif'],
  Verdana: ['Arial', 'Helvetica', 'sans-serif']
};

function normalizeFontFamily(fontFamily: string): string {
  const trimmed = fontFamily.trim();
  if (
    trimmed.length >= 2
    && ((trimmed.startsWith('"') && trimmed.endsWith('"'))
      || (trimmed.startsWith("'") && trimmed.endsWith("'")))
  ) {
    return trimmed.slice(1, -1).trim();
  }

  return trimmed;
}

function getComparableFontFamily(fontFamily: string): string {
  return normalizeFontFamily(fontFamily).toLowerCase();
}

function quoteCssFontFamily(fontFamily: string): string {
  const normalized = normalizeFontFamily(fontFamily);
  if (GENERIC_FONT_FAMILIES.has(normalized.toLowerCase())) {
    return normalized.toLowerCase();
  }

  return `"${normalized.replace(/\\/g, '\\\\').replace(/"/g, '\\"')}"`;
}

export function splitCssFontFamilies(fontFamilies: string): string[] {
  const families: string[] = [];
  let current = '';
  let quote = '';

  for (const character of fontFamilies) {
    if (quote) {
      current += character;
      if (character === quote) quote = '';
      continue;
    }

    if (character === '"' || character === "'") {
      quote = character;
      current += character;
    } else if (character === ',') {
      const family = normalizeFontFamily(current);
      if (family) families.push(family);
      current = '';
    } else {
      current += character;
    }
  }

  const family = normalizeFontFamily(current);
  if (family) families.push(family);
  return families;
}

function uniqueFontFamilies(fontFamilies: string[]): string[] {
  const seen = new Set<string>();
  return fontFamilies.filter((fontFamily) => {
    const comparable = getComparableFontFamily(fontFamily);
    if (!comparable || seen.has(comparable)) return false;
    seen.add(comparable);
    return true;
  });
}

function inferFontCategory(fontFamily: string): FontCategory {
  const comparable = getComparableFontFamily(fontFamily);
  if (
    comparable.includes('mono')
    || comparable.includes('courier')
    || comparable.includes('consolas')
    || comparable.includes('typewriter')
  ) {
    return 'monospace';
  }

  if (
    comparable.includes('serif')
    || comparable.includes('cambria')
    || comparable.includes('garamond')
    || comparable.includes('georgia')
    || comparable.includes('mincho')
    || comparable.includes('times')
  ) {
    return 'serif';
  }

  return 'sans-serif';
}

function getCategoryFallbacks(fontFamily: string): string[] {
  switch (inferFontCategory(fontFamily)) {
    case 'monospace':
      return ['Consolas', 'Courier New', 'monospace'];
    case 'serif':
      return ['Georgia', 'Times New Roman', 'serif'];
    default:
      return ['Arial', 'Helvetica', 'sans-serif'];
  }
}

function createCanvasContext(): CanvasRenderingContext2D | null {
  if (typeof document === 'undefined') return null;
  return document.createElement('canvas').getContext('2d');
}

function createBrowserFontAvailabilityDetector(): (fontFamily: string) => boolean {
  const context = createCanvasContext();
  if (!context) return () => true;

  const baselineWidths = new Map<string, number[]>();
  for (const baseline of ['monospace', 'sans-serif', 'serif']) {
    context.font = `${FONT_DETECTION_SIZE_PX}px ${baseline}`;
    baselineWidths.set(
      baseline,
      FONT_DETECTION_SAMPLES.map((sample) => context.measureText(sample).width)
    );
  }

  return (fontFamily: string): boolean => {
    const normalized = normalizeFontFamily(fontFamily);
    if (!normalized || GENERIC_FONT_FAMILIES.has(normalized.toLowerCase())) return true;

    return ['monospace', 'sans-serif', 'serif'].some((baseline) => {
      context.font = `${FONT_DETECTION_SIZE_PX}px ${quoteCssFontFamily(normalized)}, ${baseline}`;
      const widths = FONT_DETECTION_SAMPLES.map((sample) => context.measureText(sample).width);
      return widths.some((width, index) => Math.abs(width - (baselineWidths.get(baseline)?.[index] ?? width)) > 0.01);
    });
  };
}

function createBrowserTextMeasurer(): (text: string, fontFamily: string, fontSizePx: number) => number {
  const context = createCanvasContext();
  return (text: string, fontFamily: string, fontSizePx: number): number => {
    if (!context) return text.length * fontSizePx * 0.6;

    context.font = `${fontSizePx}px ${quoteCssFontFamily(fontFamily) || 'sans-serif'}`;
    return context.measureText(text).width;
  };
}

function replaceStyleFontFamily(style: string, cssFamily: string): string {
  return style.replace(
    /(^|;)\s*font-family\s*:\s*([^;]*)/i,
    (_match, prefix: string) => `${prefix}font-family: ${cssFamily}`
  );
}

function getStyleFontFamily(style: string): string | null {
  return style.match(/(?:^|;)\s*font-family\s*:\s*([^;]*)/i)?.[1]?.trim() || null;
}

export class FontFidelity {
  private readonly availability = new Map<string, boolean>();
  private readonly resolutions = new Map<string, FontResolution>();
  private readonly fallbackMap: FontFallbackMap;
  private readonly isFontAvailableFn: (fontFamily: string) => boolean;
  private readonly measureTextFn: (text: string, fontFamily: string, fontSizePx: number) => number;

  constructor(options: FontFidelityOptions = {}) {
    this.fallbackMap = { ...POWERPOINT_FONT_FALLBACKS, ...options.fontFallbacks };
    this.isFontAvailableFn = options.isFontAvailable ?? createBrowserFontAvailabilityDetector();
    this.measureTextFn = options.measureText ?? createBrowserTextMeasurer();
  }

  getRendererFallbacks(): FontFallbackMap {
    return { ...this.fallbackMap };
  }

  measureText(text: string, fontFamily: string, fontSizePx: number): number {
    const measured = this.measureTextFn(text, this.resolveFont(fontFamily).resolved, fontSizePx);
    return Number.isFinite(measured) && measured >= 0 ? measured : text.length * fontSizePx * 0.6;
  }

  applySvgSubstitutions(svg: Element): FontSubstitution[] {
    const substitutions = new Map<string, FontSubstitution>();
    const elements = [svg, ...Array.from(svg.getElementsByTagName('*'))];

    for (const element of elements) {
      const attributeFamily = element.getAttribute('font-family');
      if (attributeFamily) {
        const resolution = this.resolveFont(attributeFamily);
        this.applyResolution(element, resolution, substitutions);
        if (resolution.requested !== resolution.resolved) {
          element.setAttribute('font-family', resolution.cssFamily);
        }
      }

      const style = element.getAttribute('style');
      const styleFamily = style ? getStyleFontFamily(style) : null;
      if (style && styleFamily) {
        const resolution = this.resolveFont(styleFamily);
        this.applyResolution(element, resolution, substitutions);
        if (resolution.requested !== resolution.resolved) {
          element.setAttribute('style', replaceStyleFontFamily(style, resolution.cssFamily));
        }
      }
    }

    return [...substitutions.values()]
      .sort((left, right) => left.requested.localeCompare(right.requested));
  }

  private applyResolution(
    element: Element,
    resolution: FontResolution,
    substitutions: Map<string, FontSubstitution>
  ): void {
    if (resolution.requested === resolution.resolved) return;

    element.setAttribute('data-native-powerpoint-requested-font', resolution.requested);
    element.setAttribute('data-native-powerpoint-font-substitution', resolution.resolved);
    substitutions.set(getComparableFontFamily(resolution.requested), {
      requested: resolution.requested,
      substitute: resolution.resolved
    });
  }

  private resolveFont(fontFamily: string): FontResolution {
    const families = splitCssFontFamilies(fontFamily);
    const requested = families[0] || 'sans-serif';
    const comparable = getComparableFontFamily(requested);
    const cached = this.resolutions.get(comparable);
    if (cached) return cached;

    const configuredFallbacks = Object.entries(this.fallbackMap)
      .find(([source]) => getComparableFontFamily(source) === comparable)?.[1] ?? [];
    const candidates = uniqueFontFamilies([
      ...families,
      ...configuredFallbacks,
      ...getCategoryFallbacks(requested)
    ]);
    const resolved = candidates.find((candidate) => this.isFontAvailable(candidate))
      ?? getCategoryFallbacks(requested).at(-1)
      ?? 'sans-serif';
    const resolution = {
      requested,
      resolved,
      cssFamily: uniqueFontFamilies([requested, resolved, ...candidates])
        .map(quoteCssFontFamily)
        .join(', ')
    };
    this.resolutions.set(comparable, resolution);
    return resolution;
  }

  private isFontAvailable(fontFamily: string): boolean {
    const comparable = getComparableFontFamily(fontFamily);
    const cached = this.availability.get(comparable);
    if (cached !== undefined) return cached;

    const available = GENERIC_FONT_FAMILIES.has(comparable) || this.isFontAvailableFn(fontFamily);
    this.availability.set(comparable, available);
    return available;
  }
}
