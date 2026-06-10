import { extractZip } from 'pptx-svg';

export interface PowerPointPackageEntry {
  name: string;
  crc32: number;
  compressedSize: number;
  uncompressedSize: number;
  method: number;
  flags: number;
  isDirectory: boolean;
}

export interface PowerPointPackageInspection {
  entries: PowerPointPackageEntry[];
  entryMap: Map<string, PowerPointPackageEntry>;
  slidePaths: string[];
  duplicateEntries: string[];
  unsafePaths: string[];
  hasVbaProject: boolean;
}

export interface PackageValidationResult {
  ok: boolean;
  errors: string[];
  warnings: string[];
}

const EOCD_SIGNATURE = 0x06054b50;
const CENTRAL_DIRECTORY_SIGNATURE = 0x02014b50;
const EOCD_MIN_LENGTH = 22;
const MAX_ZIP_COMMENT_LENGTH = 0xffff;
const UINT32_MAX = 0xffffffff;
const decoder = new TextDecoder('utf-8');

const REQUIRED_POWERPOINT_PARTS = [
  '[Content_Types].xml',
  '_rels/.rels',
  'ppt/presentation.xml',
  'ppt/_rels/presentation.xml.rels'
];

const MUTABLE_EXACT_PARTS = new Set([
  '[Content_Types].xml',
  '_rels/.rels',
  'ppt/presentation.xml',
  'ppt/_rels/presentation.xml.rels'
]);

const MUTABLE_PART_PATTERNS = [
  /^ppt\/charts\/chart\d+\.xml$/,
  /^ppt\/embeddings\/[^/]+\.xlsx$/i,
  /^ppt\/slides\/slide\d+\.xml$/,
  /^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/
];

export function inspectPowerPointPackage(buffer: ArrayBuffer): PowerPointPackageInspection {
  const bytes = new Uint8Array(buffer);
  if (bytes.byteLength < EOCD_MIN_LENGTH) {
    throw new Error('The exported file is too small to be an Open XML ZIP package.');
  }

  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const eocdOffset = findEndOfCentralDirectory(view);
  if (eocdOffset === -1) {
    throw new Error('The file is not a valid Open XML ZIP package.');
  }

  const diskNumber = view.getUint16(eocdOffset + 4, true);
  const centralDirectoryDisk = view.getUint16(eocdOffset + 6, true);
  const entriesOnDisk = view.getUint16(eocdOffset + 8, true);
  const totalEntries = view.getUint16(eocdOffset + 10, true);
  const centralDirectorySize = view.getUint32(eocdOffset + 12, true);
  const centralDirectoryOffset = view.getUint32(eocdOffset + 16, true);

  if (diskNumber !== 0 || centralDirectoryDisk !== 0 || entriesOnDisk !== totalEntries) {
    throw new Error('Multi-disk ZIP packages are not supported.');
  }

  if (
    totalEntries === 0xffff ||
    centralDirectorySize === UINT32_MAX ||
    centralDirectoryOffset === UINT32_MAX
  ) {
    throw new Error('ZIP64 PowerPoint packages need explicit validation before saving.');
  }

  if (centralDirectoryOffset + centralDirectorySize > bytes.byteLength) {
    throw new Error('The ZIP central directory points outside the file.');
  }

  const entries: PowerPointPackageEntry[] = [];
  const entryMap = new Map<string, PowerPointPackageEntry>();
  const duplicateEntries: string[] = [];
  const unsafePaths: string[] = [];
  let offset = centralDirectoryOffset;

  for (let index = 0; index < totalEntries; index++) {
    if (offset + 46 > bytes.byteLength || view.getUint32(offset, true) !== CENTRAL_DIRECTORY_SIGNATURE) {
      throw new Error('The ZIP central directory is malformed.');
    }

    const flags = view.getUint16(offset + 8, true);
    const method = view.getUint16(offset + 10, true);
    const crc32 = view.getUint32(offset + 16, true);
    const compressedSize = view.getUint32(offset + 20, true);
    const uncompressedSize = view.getUint32(offset + 24, true);
    const fileNameLength = view.getUint16(offset + 28, true);
    const extraFieldLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const nameStart = offset + 46;
    const nameEnd = nameStart + fileNameLength;

    if (nameEnd > bytes.byteLength) {
      throw new Error('The ZIP central directory contains an invalid entry name.');
    }

    if (nameEnd + extraFieldLength + commentLength > bytes.byteLength) {
      throw new Error('The ZIP central directory contains an invalid entry length.');
    }

    const name = decoder.decode(bytes.subarray(nameStart, nameEnd));
    const entry: PowerPointPackageEntry = {
      name,
      crc32,
      compressedSize,
      uncompressedSize,
      method,
      flags,
      isDirectory: name.endsWith('/')
    };

    if (!isSafeZipPath(name)) {
      unsafePaths.push(name || '(empty path)');
    }

    if (entryMap.has(name)) {
      duplicateEntries.push(name);
    } else {
      entryMap.set(name, entry);
    }

    entries.push(entry);
    offset = nameEnd + extraFieldLength + commentLength;
  }

  if (offset !== centralDirectoryOffset + centralDirectorySize) {
    throw new Error('The ZIP central directory entry count does not match its declared size.');
  }

  const slidePaths = entries
    .map((entry) => entry.name)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort(compareSlidePaths);

  return {
    entries,
    entryMap,
    slidePaths,
    duplicateEntries,
    unsafePaths,
    hasVbaProject: entryMap.has('ppt/vbaProject.bin')
  };
}

export function validatePowerPointPackageStructure(
  inspection: PowerPointPackageInspection,
  expectedSlideCount?: number
): PackageValidationResult {
  const errors: string[] = [];
  const warnings: string[] = [];

  for (const path of REQUIRED_POWERPOINT_PARTS) {
    if (!inspection.entryMap.has(path)) {
      errors.push(`Missing required package part: ${path}`);
    }
  }

  if (inspection.unsafePaths.length > 0) {
    errors.push(`Unsafe ZIP entry paths: ${inspection.unsafePaths.slice(0, 5).join(', ')}`);
  }

  if (inspection.duplicateEntries.length > 0) {
    errors.push(`Duplicate ZIP entries: ${inspection.duplicateEntries.slice(0, 5).join(', ')}`);
  }

  if (typeof expectedSlideCount === 'number' && inspection.slidePaths.length !== expectedSlideCount) {
    errors.push(`Expected ${expectedSlideCount} slide part(s), found ${inspection.slidePaths.length}.`);
  }

  if (inspection.hasVbaProject) {
    warnings.push('The package contains ppt/vbaProject.bin. Editing is disabled until macro preservation is verified.');
  }

  return {
    ok: errors.length === 0,
    errors,
    warnings
  };
}

export function validatePowerPointExport(
  original: PowerPointPackageInspection,
  exported: PowerPointPackageInspection,
  expectedSlideCount: number
): PackageValidationResult {
  const structure = validatePowerPointPackageStructure(exported, expectedSlideCount);
  const errors = [...structure.errors];
  const warnings = [...structure.warnings];

  for (const originalEntry of original.entries) {
    if (!shouldPreserveOriginalPart(originalEntry.name)) continue;

    const exportedEntry = exported.entryMap.get(originalEntry.name);
    if (!exportedEntry) {
      errors.push(`Preserved OOXML part was dropped: ${originalEntry.name}`);
      continue;
    }

    if (!sameStoredContent(originalEntry, exportedEntry)) {
      errors.push(`Preserved OOXML part changed unexpectedly: ${originalEntry.name}`);
    }
  }

  if (original.hasVbaProject) {
    const originalMacro = original.entryMap.get('ppt/vbaProject.bin');
    const exportedMacro = exported.entryMap.get('ppt/vbaProject.bin');
    if (!originalMacro || !exportedMacro || !sameStoredContent(originalMacro, exportedMacro)) {
      errors.push('Macro project bytes were not preserved exactly.');
    }
  }

  return {
    ok: errors.length === 0,
    errors: dedupeMessages(errors),
    warnings: dedupeMessages(warnings)
  };
}

export async function validatePowerPointExportContents(
  originalBuffer: ArrayBuffer,
  exportedBuffer: ArrayBuffer
): Promise<PackageValidationResult> {
  const [original, exported] = await Promise.all([
    extractZip(originalBuffer),
    extractZip(exportedBuffer)
  ]);
  const errors: string[] = [];

  for (const [path, contents] of original.textFiles) {
    if (!shouldPreserveOriginalPart(path)) continue;
    if (exported.textFiles.get(path) !== contents) {
      errors.push(`Preserved OOXML part bytes changed unexpectedly: ${path}`);
    }
  }

  for (const [path, contents] of original.binaryFiles) {
    if (!shouldPreserveOriginalPart(path)) continue;
    const exportedContents = exported.binaryFiles.get(path);
    if (!exportedContents || !sameBytes(contents, exportedContents)) {
      errors.push(`Preserved binary part bytes changed unexpectedly: ${path}`);
    }
  }

  const originalSlides = collectSlideXml(original.textFiles);
  const exportedSlides = collectSlideXml(exported.textFiles);
  if (exportedSlides.length >= originalSlides.length) {
    const originalSlideXml = originalSlides.join('\n');
    const exportedSlideXml = exportedSlides.join('\n');
    for (const marker of PROTECTED_SLIDE_MARKERS) {
      const originalCount = countMatches(originalSlideXml, marker.pattern);
      const exportedCount = countMatches(exportedSlideXml, marker.pattern);
      if (exportedCount < originalCount) {
        errors.push(`Slide edit dropped ${marker.label} markup.`);
      }
    }

    for (const elementName of collectUnknownElementNames(originalSlideXml)) {
      if (countElementName(exportedSlideXml, elementName) < countElementName(originalSlideXml, elementName)) {
        errors.push(`Slide edit dropped unknown OOXML element <${elementName}>.`);
      }
    }

    const exportedTargets = collectSlideRelationshipTargets(exported.textFiles);
    for (const target of collectSlideRelationshipTargets(original.textFiles)) {
      if (!exportedTargets.includes(target)) {
        errors.push(`Slide edit dropped external relationship ${target}.`);
      }
    }
  }

  return {
    ok: errors.length === 0,
    errors: dedupeMessages(errors),
    warnings: []
  };
}

export function summarizePackageMessages(messages: string[], limit = 4): string {
  const visible = messages.slice(0, limit);
  const hiddenCount = Math.max(0, messages.length - visible.length);
  return hiddenCount ? `${visible.join('; ')}; ${hiddenCount} more issue(s).` : visible.join('; ');
}

function findEndOfCentralDirectory(view: DataView): number {
  const minOffset = Math.max(0, view.byteLength - EOCD_MIN_LENGTH - MAX_ZIP_COMMENT_LENGTH);

  for (let offset = view.byteLength - EOCD_MIN_LENGTH; offset >= minOffset; offset--) {
    if (view.getUint32(offset, true) !== EOCD_SIGNATURE) continue;

    const commentLength = view.getUint16(offset + 20, true);
    if (offset + EOCD_MIN_LENGTH + commentLength === view.byteLength) {
      return offset;
    }
  }

  return -1;
}

function isSafeZipPath(name: string): boolean {
  if (!name || name.startsWith('/') || name.includes('\\') || /^[A-Za-z]:/.test(name)) {
    return false;
  }

  const parts = name.split('/').filter((part) => part.length > 0);
  return parts.every((part) => part !== '.' && part !== '..');
}

function compareSlidePaths(a: string, b: string): number {
  return getSlideNumber(a) - getSlideNumber(b);
}

function getSlideNumber(path: string): number {
  const match = path.match(/slide(\d+)\.xml$/);
  return match ? Number(match[1]) : 0;
}

function shouldPreserveOriginalPart(name: string): boolean {
  if (name.endsWith('/')) return false;
  if (MUTABLE_EXACT_PARTS.has(name)) return false;
  return !MUTABLE_PART_PATTERNS.some((pattern) => pattern.test(name));
}

function sameStoredContent(a: PowerPointPackageEntry, b: PowerPointPackageEntry): boolean {
  return a.crc32 === b.crc32 && a.uncompressedSize === b.uncompressedSize;
}

const PROTECTED_SLIDE_MARKERS = [
  { label: 'animation timing', pattern: /<p:timing\b/g },
  { label: 'chart', pattern: /<c:chart\b/g },
  { label: 'grouped shape', pattern: /<p:grpSp\b/g },
  { label: 'hyperlink', pattern: /<a:hlinkClick\b/g },
  { label: 'image', pattern: /<p:pic\b/g },
  { label: 'slide extension', pattern: /<p:ext\b/g },
  { label: 'table', pattern: /<a:tbl\b/g }
];

const KNOWN_OOXML_PREFIXES = new Set(['a', 'c', 'm', 'mc', 'p', 'r']);

function countMatches(contents: string, pattern: RegExp): number {
  return Array.from(contents.matchAll(pattern)).length;
}

function collectUnknownElementNames(contents: string): string[] {
  const names = new Set<string>();
  for (const match of contents.matchAll(/<([A-Za-z_][\w.-]*):([A-Za-z_][\w.-]*)\b/g)) {
    const prefix = match[1];
    const localName = match[2];
    if (prefix && localName && !KNOWN_OOXML_PREFIXES.has(prefix)) {
      names.add(`${prefix}:${localName}`);
    }
  }
  return Array.from(names);
}

function countElementName(contents: string, name: string): number {
  return countMatches(contents, new RegExp(`<${escapeRegex(name)}\\b`, 'g'));
}

function collectExternalRelationshipTargets(contents: string): string[] {
  const targets: string[] = [];
  for (const match of contents.matchAll(/<Relationship\b([^>]*)\/?>/g)) {
    const attrs = match[1] ?? '';
    const target = attrs.match(/\bTarget="([^"]+)"/)?.[1];
    if (target && /\bTargetMode="External"/.test(attrs)) {
      targets.push(target);
    }
  }
  return targets;
}

function collectSlideXml(files: Map<string, string>): string[] {
  return Array.from(files)
    .filter(([path]) => /^ppt\/slides\/slide\d+\.xml$/.test(path))
    .sort(([a], [b]) => getSlideNumber(a) - getSlideNumber(b))
    .map(([, contents]) => contents);
}

function collectSlideRelationshipTargets(files: Map<string, string>): string[] {
  const targets: string[] = [];
  for (const [path, contents] of files) {
    if (/^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/.test(path)) {
      targets.push(...collectExternalRelationshipTargets(contents));
    }
  }
  return targets;
}

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a.byteLength !== b.byteLength) return false;
  return a.every((byte, index) => byte === b[index]);
}

function escapeRegex(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function dedupeMessages(messages: string[]): string[] {
  return Array.from(new Set(messages));
}
