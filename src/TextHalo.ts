export type GeneratedGridKind = 'chart' | 'table';

export const DEFAULT_TEXT_HALO_COLOR = 'rgba(255, 255, 255, 0.88)';

const TEXT_HALO_PROPERTY = '--native-powerpoint-text-halo';

interface RgbaColor {
  red: number;
  green: number;
  blue: number;
  alpha: number;
}

function clampChannel(value: number): number {
  return Math.max(0, Math.min(255, Math.round(value)));
}

function clampAlpha(value: number): number {
  return Math.max(0, Math.min(1, value));
}

function parseHexColor(value: string): RgbaColor | null {
  const match = value.match(/^#([\da-f]{3}|[\da-f]{6}|[\da-f]{8})$/i);
  if (!match?.[1]) return null;

  const hex = match[1].length === 3
    ? match[1].split('').map((character) => `${character}${character}`).join('')
    : match[1];
  return {
    red: Number.parseInt(hex.slice(0, 2), 16),
    green: Number.parseInt(hex.slice(2, 4), 16),
    blue: Number.parseInt(hex.slice(4, 6), 16),
    alpha: hex.length === 8 ? Number.parseInt(hex.slice(6, 8), 16) / 255 : 1
  };
}

function parseFunctionalColor(value: string): RgbaColor | null {
  const match = value.match(/^rgba?\(\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)(?:\s*,\s*([\d.]+))?\s*\)$/i);
  if (!match) return null;

  return {
    red: clampChannel(Number(match[1])),
    green: clampChannel(Number(match[2])),
    blue: clampChannel(Number(match[3])),
    alpha: clampAlpha(match[4] === undefined ? 1 : Number(match[4]))
  };
}

function parseSvgColor(value: string | null): RgbaColor | null {
  if (!value) return null;

  const normalizedValue = value.trim().toLowerCase();
  if (!normalizedValue || normalizedValue === 'none' || normalizedValue === 'transparent') {
    return null;
  }

  if (normalizedValue === 'white') {
    return { red: 255, green: 255, blue: 255, alpha: 1 };
  }

  if (normalizedValue === 'black') {
    return { red: 0, green: 0, blue: 0, alpha: 1 };
  }

  return parseHexColor(normalizedValue) ?? parseFunctionalColor(normalizedValue);
}

function compositeAgainstWhite(color: RgbaColor): string {
  const alpha = clampAlpha(color.alpha);
  const red = clampChannel(color.red * alpha + 255 * (1 - alpha));
  const green = clampChannel(color.green * alpha + 255 * (1 - alpha));
  const blue = clampChannel(color.blue * alpha + 255 * (1 - alpha));
  return `rgb(${red}, ${green}, ${blue})`;
}

function getElementChildren(element: Element): Element[] {
  return Array.from(element.childNodes)
    .filter((node): node is Element => node.nodeType === 1);
}

function getDescendants(element: Element, localName: string): Element[] {
  return Array.from(element.getElementsByTagNameNS('*', localName));
}

function getElementHaloColor(element: Element): string | null {
  const color = parseSvgColor(element.getAttribute('fill'));
  if (!color) return null;

  const fillOpacity = Number(element.getAttribute('fill-opacity') ?? 1);
  const opacity = Number(element.getAttribute('opacity') ?? 1);
  color.alpha *= Number.isFinite(fillOpacity) ? fillOpacity : 1;
  color.alpha *= Number.isFinite(opacity) ? opacity : 1;
  return compositeAgainstWhite(color);
}

function setCssProperty(element: Element, property: string, value: string): void {
  const entries = (element.getAttribute('style') ?? '')
    .split(';')
    .map((entry) => entry.trim())
    .filter(Boolean)
    .filter((entry) => !entry.startsWith(`${property}:`));
  entries.push(`${property}: ${value}`);
  element.setAttribute('style', entries.join('; '));
}

function setTextHalo(text: Element, color: string): void {
  setCssProperty(text, TEXT_HALO_PROPERTY, color);
  text.setAttribute('data-native-powerpoint-halo-color', color);
}

function applyTableTextHalos(container: Element, inheritedColor = DEFAULT_TEXT_HALO_COLOR): void {
  let currentColor = inheritedColor;

  for (const child of getElementChildren(container)) {
    if (child.localName === 'rect') {
      currentColor = getElementHaloColor(child) ?? inheritedColor;
    } else if (child.localName === 'text') {
      setTextHalo(child, currentColor);
    } else if (child.localName === 'g') {
      applyTableTextHalos(child, currentColor);
    }
  }
}

function getChartBackgroundColor(chartGroup: Element): string {
  let bestColor = DEFAULT_TEXT_HALO_COLOR;
  let bestArea = -1;

  for (const rectangle of getDescendants(chartGroup, 'rect')) {
    const color = getElementHaloColor(rectangle);
    const width = Number(rectangle.getAttribute('width'));
    const height = Number(rectangle.getAttribute('height'));
    const area = width * height;
    if (color && Number.isFinite(area) && area > bestArea) {
      bestColor = color;
      bestArea = area;
    }
  }

  return bestColor;
}

function applyChartTextHalos(chartGroup: Element): void {
  const color = getChartBackgroundColor(chartGroup);
  for (const text of getDescendants(chartGroup, 'text')) {
    setTextHalo(text, color);
  }
}

export function applyBackgroundAwareTextHalos(gridGroup: Element, kind: GeneratedGridKind): void {
  if (kind === 'table') {
    applyTableTextHalos(gridGroup);
  } else {
    applyChartTextHalos(gridGroup);
  }
}
