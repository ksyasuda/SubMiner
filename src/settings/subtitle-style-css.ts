import type { ConfigSettingsSnapshotValue } from '../types/settings';

export type SubtitleCssScope = 'primary' | 'secondary';

type LegacyCssDeclaration = {
  property: string;
  primaryPath: string;
  secondaryPath: string;
  format?: (value: unknown) => string | undefined;
};

export type SubtitleCssParseResult =
  | { ok: true; declarations: Record<string, string> }
  | { ok: false; error: string };

const LEGACY_CSS_DECLARATIONS: LegacyCssDeclaration[] = [
  {
    property: 'font-family',
    primaryPath: 'subtitleStyle.fontFamily',
    secondaryPath: 'subtitleStyle.secondary.fontFamily',
  },
  {
    property: 'font-size',
    primaryPath: 'subtitleStyle.fontSize',
    secondaryPath: 'subtitleStyle.secondary.fontSize',
    format: formatCssLengthLikeValue,
  },
  {
    property: 'font-weight',
    primaryPath: 'subtitleStyle.fontWeight',
    secondaryPath: 'subtitleStyle.secondary.fontWeight',
  },
  {
    property: 'font-style',
    primaryPath: 'subtitleStyle.fontStyle',
    secondaryPath: 'subtitleStyle.secondary.fontStyle',
  },
  {
    property: 'line-height',
    primaryPath: 'subtitleStyle.lineHeight',
    secondaryPath: 'subtitleStyle.secondary.lineHeight',
  },
  {
    property: 'letter-spacing',
    primaryPath: 'subtitleStyle.letterSpacing',
    secondaryPath: 'subtitleStyle.secondary.letterSpacing',
  },
  {
    property: 'word-spacing',
    primaryPath: 'subtitleStyle.wordSpacing',
    secondaryPath: 'subtitleStyle.secondary.wordSpacing',
  },
  {
    property: 'font-kerning',
    primaryPath: 'subtitleStyle.fontKerning',
    secondaryPath: 'subtitleStyle.secondary.fontKerning',
  },
  {
    property: 'text-rendering',
    primaryPath: 'subtitleStyle.textRendering',
    secondaryPath: 'subtitleStyle.secondary.textRendering',
  },
  {
    property: 'text-shadow',
    primaryPath: 'subtitleStyle.textShadow',
    secondaryPath: 'subtitleStyle.secondary.textShadow',
  },
  {
    property: 'backdrop-filter',
    primaryPath: 'subtitleStyle.backdropFilter',
    secondaryPath: 'subtitleStyle.secondary.backdropFilter',
  },
  {
    property: 'color',
    primaryPath: 'subtitleStyle.fontColor',
    secondaryPath: 'subtitleStyle.secondary.fontColor',
  },
  {
    property: 'background-color',
    primaryPath: 'subtitleStyle.backgroundColor',
    secondaryPath: 'subtitleStyle.secondary.backgroundColor',
  },
];

const CSS_PROPERTY_PATTERN = /^(?:--[A-Za-z0-9_-]+|-?[A-Za-z][A-Za-z0-9_-]*)$/;

export function getSubtitleCssPath(scope: SubtitleCssScope): string {
  return scope === 'primary' ? 'subtitleStyle.css' : 'subtitleStyle.secondary.css';
}

export function getSubtitleCssManagedConfigPaths(scope: SubtitleCssScope): string[] {
  return LEGACY_CSS_DECLARATIONS.map((declaration) =>
    scope === 'primary' ? declaration.primaryPath : declaration.secondaryPath,
  );
}

export function getSubtitleCssScopeForPath(path: string): SubtitleCssScope | null {
  if (path === 'subtitleStyle.css') return 'primary';
  if (path === 'subtitleStyle.secondary.css') return 'secondary';
  return null;
}

export function serializeSubtitleCssDeclarations(
  scope: SubtitleCssScope,
  values: Record<string, ConfigSettingsSnapshotValue | undefined>,
): string {
  const declarations = new Map<string, string>();

  for (const declaration of LEGACY_CSS_DECLARATIONS) {
    const path = scope === 'primary' ? declaration.primaryPath : declaration.secondaryPath;
    const formatted = (declaration.format ?? formatCssPrimitiveValue)(values[path]);
    if (formatted !== undefined) {
      declarations.set(declaration.property, formatted);
    }
  }

  const cssObject = normalizeCssDeclarationRecord(values[getSubtitleCssPath(scope)]);
  for (const [property, value] of Object.entries(cssObject)) {
    declarations.set(normalizeCssPropertyName(property), value);
  }

  return [...declarations.entries()]
    .map(([property, value]) => `${property}: ${value};`)
    .join('\n');
}

export function parseSubtitleCssDeclarations(text: string): SubtitleCssParseResult {
  const trimmed = text.trim();
  if (trimmed.length === 0) {
    return { ok: true, declarations: {} };
  }

  if (/[{}]/.test(trimmed)) {
    return {
      ok: false,
      error: 'Enter CSS declarations only, without selectors or braces.',
    };
  }

  const declarations: Record<string, string> = {};
  for (const rawDeclaration of splitCssDeclarations(trimmed)) {
    const declaration = rawDeclaration.trim();
    if (declaration.length === 0) continue;

    const colonIndex = findTopLevelColon(declaration);
    if (colonIndex <= 0) {
      return { ok: false, error: `Invalid CSS declaration: ${declaration}` };
    }

    const property = normalizeCssPropertyName(declaration.slice(0, colonIndex).trim());
    const value = declaration.slice(colonIndex + 1).trim();
    if (!CSS_PROPERTY_PATTERN.test(property)) {
      return { ok: false, error: `Invalid CSS property: ${property}` };
    }
    if (value.length === 0) {
      return { ok: false, error: `Missing CSS value for ${property}.` };
    }

    declarations[property] = value;
  }

  return { ok: true, declarations };
}

function normalizeCssDeclarationRecord(value: unknown): Record<string, string> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    return {};
  }

  const declarations: Record<string, string> = {};
  for (const [property, rawValue] of Object.entries(value)) {
    if (typeof rawValue !== 'string') continue;
    const trimmed = rawValue.trim();
    if (trimmed.length === 0) continue;
    declarations[property] = trimmed;
  }
  return declarations;
}

function normalizeCssPropertyName(property: string): string {
  const trimmed = property.trim();
  if (trimmed.startsWith('--')) return trimmed;
  if (trimmed.includes('-')) return trimmed.toLowerCase();

  const kebab = trimmed
    .replace(/([a-z0-9])([A-Z])/g, '$1-$2')
    .replace(/^Webkit-/, '-webkit-')
    .toLowerCase();
  return kebab.startsWith('webkit-') ? `-${kebab}` : kebab;
}

function formatCssLengthLikeValue(value: unknown): string | undefined {
  if (typeof value === 'number' && Number.isFinite(value)) {
    return `${value}px`;
  }
  return formatCssPrimitiveValue(value);
}

function formatCssPrimitiveValue(value: unknown): string | undefined {
  if (value === null || value === undefined || typeof value === 'object') {
    return undefined;
  }

  const text = String(value).trim();
  return text.length > 0 ? text : undefined;
}

function splitCssDeclarations(text: string): string[] {
  const declarations: string[] = [];
  let current = '';
  let quote: '"' | "'" | null = null;
  let parenDepth = 0;
  let escaping = false;

  for (const char of text) {
    if (escaping) {
      current += char;
      escaping = false;
      continue;
    }

    if (char === '\\') {
      current += char;
      escaping = true;
      continue;
    }

    if (quote) {
      current += char;
      if (char === quote) quote = null;
      continue;
    }

    if (char === '"' || char === "'") {
      current += char;
      quote = char;
      continue;
    }

    if (char === '(') {
      parenDepth += 1;
      current += char;
      continue;
    }

    if (char === ')') {
      parenDepth = Math.max(0, parenDepth - 1);
      current += char;
      continue;
    }

    if (char === ';' && parenDepth === 0) {
      declarations.push(current);
      current = '';
      continue;
    }

    current += char;
  }

  declarations.push(current);
  return declarations;
}

function findTopLevelColon(text: string): number {
  let quote: '"' | "'" | null = null;
  let parenDepth = 0;
  let escaping = false;

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    if (escaping) {
      escaping = false;
      continue;
    }
    if (char === '\\') {
      escaping = true;
      continue;
    }
    if (quote) {
      if (char === quote) quote = null;
      continue;
    }
    if (char === '"' || char === "'") {
      quote = char;
      continue;
    }
    if (char === '(') {
      parenDepth += 1;
      continue;
    }
    if (char === ')') {
      parenDepth = Math.max(0, parenDepth - 1);
      continue;
    }
    if (char === ':' && parenDepth === 0) {
      return i;
    }
  }

  return -1;
}
