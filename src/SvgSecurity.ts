export interface SvgSecurityIssue {
  reason: string;
  detail: string;
}

export interface SanitizedSvg {
  svg: string;
  issues: SvgSecurityIssue[];
}

export type SvgSanitizerMode = 'strict' | 'compatibility';

export interface SanitizeSvgOptions {
  mode?: SvgSanitizerMode;
}

const ALLOWED_ELEMENTS = new Set([
  'svg',
  'g',
  'defs',
  'desc',
  'title',
  'metadata',
  'symbol',
  'use',
  'path',
  'rect',
  'circle',
  'ellipse',
  'line',
  'polyline',
  'polygon',
  'text',
  'tspan',
  'image',
  'clipPath',
  'mask',
  'pattern',
  'linearGradient',
  'radialGradient',
  'stop',
  'filter',
  'feBlend',
  'feColorMatrix',
  'feComponentTransfer',
  'feComposite',
  'feDropShadow',
  'feFlood',
  'feGaussianBlur',
  'feMerge',
  'feMergeNode',
  'feOffset',
  'feFuncR',
  'feFuncG',
  'feFuncB',
  'feFuncA'
]);
const CASE_INSENSITIVE_ALLOWED_ELEMENTS = new Set(Array.from(ALLOWED_ELEMENTS, (element) => element.toLowerCase()));

const COMPATIBILITY_ALLOWED_ELEMENTS = new Set([
  'marker',
  'switch',
  'textPath',
  'view',
  'feConvolveMatrix',
  'feDiffuseLighting',
  'feDisplacementMap',
  'feDistantLight',
  'feMorphology',
  'fePointLight',
  'feSpecularLighting',
  'feSpotLight',
  'feTile',
  'feTurbulence'
]);
const CASE_INSENSITIVE_COMPATIBILITY_ALLOWED_ELEMENTS = new Set(
  Array.from(COMPATIBILITY_ALLOWED_ELEMENTS, (element) => element.toLowerCase())
);

const PERMANENTLY_BLOCKED_ELEMENTS = new Set([
  'audio',
  'canvas',
  'embed',
  'foreignobject',
  'iframe',
  'link',
  'object',
  'script',
  'style',
  'video'
]);

const ALLOWED_ATTRS = new Set([
  'accent-height',
  'alignment-baseline',
  'aria-hidden',
  'aria-label',
  'baseline-shift',
  'class',
  'clip',
  'clip-path',
  'clip-rule',
  'clipPathUnits',
  'color',
  'color-interpolation',
  'color-interpolation-filters',
  'cx',
  'cy',
  'd',
  'direction',
  'display',
  'dominant-baseline',
  'dx',
  'dy',
  'fill',
  'fill-opacity',
  'fill-rule',
  'filter',
  'filterUnits',
  'flood-color',
  'flood-opacity',
  'font-family',
  'font-size',
  'font-stretch',
  'font-style',
  'font-variant',
  'font-weight',
  'fx',
  'fy',
  'gradientTransform',
  'gradientUnits',
  'height',
  'href',
  'id',
  'in',
  'in2',
  'letter-spacing',
  'marker-end',
  'marker-mid',
  'marker-start',
  'mask',
  'maskContentUnits',
  'maskUnits',
  'offset',
  'opacity',
  'operator',
  'overflow',
  'paint-order',
  'patternContentUnits',
  'patternTransform',
  'patternUnits',
  'points',
  'preserveAspectRatio',
  'r',
  'result',
  'role',
  'rx',
  'ry',
  'spreadMethod',
  'stdDeviation',
  'stop-color',
  'stop-opacity',
  'stroke',
  'stroke-dasharray',
  'stroke-dashoffset',
  'stroke-linecap',
  'stroke-linejoin',
  'stroke-miterlimit',
  'stroke-opacity',
  'stroke-width',
  'style',
  'text-anchor',
  'text-decoration',
  'textLength',
  'transform',
  'type',
  'vector-effect',
  'version',
  'viewBox',
  'visibility',
  'width',
  'word-spacing',
  'x',
  'x1',
  'x2',
  'xlink:href',
  'xml:space',
  'xmlns',
  'xmlns:xlink',
  'y',
  'y1',
  'y2'
]);

const CASE_INSENSITIVE_ALLOWED_ATTRS = new Set(Array.from(ALLOWED_ATTRS, (attr) => attr.toLowerCase()));
const COMPATIBILITY_ALLOWED_ATTRS = new Set([
  'aria-describedby',
  'aria-labelledby',
  'azimuth',
  'baseFrequency',
  'bias',
  'diffuseConstant',
  'divisor',
  'edgeMode',
  'elevation',
  'kernelMatrix',
  'kernelUnitLength',
  'lengthAdjust',
  'limitingConeAngle',
  'markerHeight',
  'markerUnits',
  'markerWidth',
  'method',
  'numOctaves',
  'orient',
  'pathLength',
  'pointsAtX',
  'pointsAtY',
  'pointsAtZ',
  'preserveAlpha',
  'primitiveUnits',
  'radius',
  'refX',
  'refY',
  'scale',
  'seed',
  'spacing',
  'specularConstant',
  'specularExponent',
  'startOffset',
  'stitchTiles',
  'surfaceScale',
  'targetX',
  'targetY',
  'xChannelSelector',
  'yChannelSelector'
]);
const CASE_INSENSITIVE_COMPATIBILITY_ALLOWED_ATTRS = new Set(
  Array.from(COMPATIBILITY_ALLOWED_ATTRS, (attr) => attr.toLowerCase())
);
const URL_ATTRS = new Set(['href', 'xlink:href']);

function addIssue(issues: SvgSecurityIssue[], reason: string, detail: string): void {
  issues.push({ reason, detail });
}

function isAllowedDataImage(value: string): boolean {
  return /^data:image\/(?:png|jpeg|jpg|gif|webp|bmp|tiff);base64,[a-z0-9+/=\s]+$/i.test(value);
}

function isSafeHref(elementName: string, value: string): boolean {
  const trimmed = value.trim();
  if (trimmed.startsWith('#')) return true;
  return elementName === 'image' && isAllowedDataImage(trimmed);
}

function findUnsafeUrlReference(value: string): string | null {
  const urlPattern = /url\(\s*(['"]?)(.*?)\1\s*\)/gi;
  let match: RegExpExecArray | null;
  while ((match = urlPattern.exec(value)) !== null) {
    const url = (match[2] ?? '').trim();
    if (!url.startsWith('#')) {
      return url || 'empty url()';
    }
  }
  return null;
}

function findScriptableProtocol(value: string): string | null {
  const normalized = value.replace(/[\u0000-\u0020\u007f-\u009f]/g, '');
  const match = normalized.match(/(?:java|vb)script:/i);
  return match ? match[0] : null;
}

function isUnsafeStyle(value: string): string | null {
  if (/expression\s*\(|@import|behavior\s*:|-moz-binding\s*:/i.test(value) || findScriptableProtocol(value)) {
    return 'scriptable CSS';
  }

  return findUnsafeUrlReference(value);
}

function shouldKeepAttribute(
  elementName: string,
  elementNameLower: string,
  attribute: Attr,
  issues: SvgSecurityIssue[],
  mode: SvgSanitizerMode
): boolean {
  const attrName = attribute.name;
  const attrNameLower = attrName.toLowerCase();
  const value = attribute.value;

  if (attrNameLower.startsWith('on')) {
    addIssue(issues, 'Detected risky action', `${elementName} ${attrName}`);
    return false;
  }

  if (attrNameLower === 'src' || attrNameLower === 'srcdoc') {
    addIssue(issues, 'Detected embedded source', `${elementName} ${attrName}`);
    return false;
  }

  if (attrNameLower === 'xml:base') {
    addIssue(issues, 'Detected hidden base link', `${elementName} ${attrName}`);
    return false;
  }

  const attrIsAllowed =
    CASE_INSENSITIVE_ALLOWED_ATTRS.has(attrNameLower) ||
    attrNameLower.startsWith('data-ooxml-') ||
    (
      mode === 'compatibility' &&
      (
        CASE_INSENSITIVE_COMPATIBILITY_ALLOWED_ATTRS.has(attrNameLower) ||
        attrNameLower.startsWith('aria-') ||
        attrNameLower.startsWith('data-')
      )
    );
  if (!attrIsAllowed) {
    addIssue(issues, 'Detected unsupported detail', `${elementName} ${attrName}`);
    return false;
  }

  if (findScriptableProtocol(value)) {
    addIssue(issues, 'Detected risky link or action', `${elementName} ${attrName}`);
    return false;
  }

  const unsafeUrl = findUnsafeUrlReference(value);
  if (unsafeUrl) {
    addIssue(issues, 'Detected outside link', `${elementName} ${attrName}=${unsafeUrl.slice(0, 80)}`);
    return false;
  }

  if (URL_ATTRS.has(attrNameLower) && !isSafeHref(elementNameLower, value)) {
    addIssue(issues, 'Detected outside file or link', `${elementName} ${attrName}=${value.slice(0, 80)}`);
    return false;
  }

  if (attrNameLower === 'style') {
    const unsafeStyle = isUnsafeStyle(value);
    if (unsafeStyle) {
      addIssue(issues, 'Detected risky styling', `${elementName} style references ${unsafeStyle.slice(0, 80)}`);
      return false;
    }
  }

  return true;
}

function sanitizeElement(element: Element, issues: SvgSecurityIssue[], mode: SvgSanitizerMode): void {
  const elementName = element.tagName;
  const elementNameLower = elementName.toLowerCase();
  const elementIsAllowed =
    CASE_INSENSITIVE_ALLOWED_ELEMENTS.has(elementNameLower) ||
    (mode === 'compatibility' && CASE_INSENSITIVE_COMPATIBILITY_ALLOWED_ELEMENTS.has(elementNameLower));
  if (PERMANENTLY_BLOCKED_ELEMENTS.has(elementNameLower) || !elementIsAllowed) {
    addIssue(issues, 'Detected unsupported slide element', elementName);
    element.remove();
    return;
  }

  for (const attribute of Array.from(element.attributes)) {
    if (!shouldKeepAttribute(elementName, elementNameLower, attribute, issues, mode)) {
      element.removeAttribute(attribute.name);
    }
  }

  for (const child of Array.from(element.children)) {
    sanitizeElement(child, issues, mode);
  }
}

export function sanitizeSvg(svg: string, options: SanitizeSvgOptions = {}): SanitizedSvg {
  const template = document.createElement('template');
  const issues: SvgSecurityIssue[] = [];
  const mode = options.mode || 'strict';
  template.innerHTML = svg.trim();

  for (const child of Array.from(template.content.children)) {
    sanitizeElement(child, issues, mode);
  }

  return {
    svg: template.innerHTML,
    issues
  };
}

export function summarizeSvgSecurityIssues(issues: SvgSecurityIssue[]): string[] {
  const counts = new Map<string, number>();
  for (const issue of issues) {
    counts.set(issue.reason, (counts.get(issue.reason) || 0) + 1);
  }

  return Array.from(counts.entries()).map(([reason, count]) => `${reason}: ${count}`);
}
