import type { ConfigSettingsSnapshotValue } from '../types/settings';

export type SubtitleCssScope = 'primary' | 'secondary' | 'sidebar';

type LegacyCssDeclaration = {
  property: string;
  paths: Partial<Record<SubtitleCssScope, string>>;
  format?: (value: unknown) => string | undefined;
};

export type SubtitleCssParseResult =
  | { ok: true; declarations: Record<string, string> }
  | { ok: false; error: string };

const LEGACY_CSS_DECLARATIONS: LegacyCssDeclaration[] = [
  {
    property: 'font-family',
    paths: {
      primary: 'subtitleStyle.fontFamily',
      secondary: 'subtitleStyle.secondary.fontFamily',
      sidebar: 'subtitleSidebar.fontFamily',
    },
  },
  {
    property: 'color',
    paths: {
      primary: 'subtitleStyle.fontColor',
      secondary: 'subtitleStyle.secondary.fontColor',
      sidebar: 'subtitleSidebar.textColor',
    },
  },
  {
    property: 'background-color',
    paths: {
      primary: 'subtitleStyle.backgroundColor',
      secondary: 'subtitleStyle.secondary.backgroundColor',
      sidebar: 'subtitleSidebar.backgroundColor',
    },
  },
  {
    property: 'font-size',
    paths: {
      primary: 'subtitleStyle.fontSize',
      secondary: 'subtitleStyle.secondary.fontSize',
      sidebar: 'subtitleSidebar.fontSize',
    },
    format: formatCssLengthLikeValue,
  },
  {
    property: 'font-weight',
    paths: {
      primary: 'subtitleStyle.fontWeight',
      secondary: 'subtitleStyle.secondary.fontWeight',
    },
  },
  {
    property: 'font-style',
    paths: {
      primary: 'subtitleStyle.fontStyle',
      secondary: 'subtitleStyle.secondary.fontStyle',
    },
  },
  {
    property: 'line-height',
    paths: {
      primary: 'subtitleStyle.lineHeight',
      secondary: 'subtitleStyle.secondary.lineHeight',
    },
  },
  {
    property: 'letter-spacing',
    paths: {
      primary: 'subtitleStyle.letterSpacing',
      secondary: 'subtitleStyle.secondary.letterSpacing',
    },
  },
  {
    property: 'word-spacing',
    paths: {
      primary: 'subtitleStyle.wordSpacing',
      secondary: 'subtitleStyle.secondary.wordSpacing',
    },
  },
  {
    property: 'font-kerning',
    paths: {
      primary: 'subtitleStyle.fontKerning',
      secondary: 'subtitleStyle.secondary.fontKerning',
    },
  },
  {
    property: 'text-rendering',
    paths: {
      primary: 'subtitleStyle.textRendering',
      secondary: 'subtitleStyle.secondary.textRendering',
    },
  },
  {
    property: 'text-shadow',
    paths: {
      primary: 'subtitleStyle.textShadow',
      secondary: 'subtitleStyle.secondary.textShadow',
    },
  },
  {
    property: 'paint-order',
    paths: {
      primary: 'subtitleStyle.paintOrder',
      secondary: 'subtitleStyle.secondary.paintOrder',
    },
  },
  {
    property: '-webkit-text-stroke',
    paths: {
      primary: 'subtitleStyle.WebkitTextStroke',
      secondary: 'subtitleStyle.secondary.WebkitTextStroke',
    },
  },
  {
    property: 'backdrop-filter',
    paths: {
      primary: 'subtitleStyle.backdropFilter',
      secondary: 'subtitleStyle.secondary.backdropFilter',
    },
  },
  {
    property: '--subtitle-hover-token-color',
    paths: {
      primary: 'subtitleStyle.hoverTokenColor',
    },
  },
  {
    property: '--subtitle-hover-token-background-color',
    paths: {
      primary: 'subtitleStyle.hoverTokenBackgroundColor',
    },
  },
  {
    property: 'opacity',
    paths: {
      sidebar: 'subtitleSidebar.opacity',
    },
  },
  {
    property: '--subtitle-sidebar-max-width',
    paths: {
      sidebar: 'subtitleSidebar.maxWidth',
    },
    format: formatCssLengthLikeValue,
  },
  {
    property: '--subtitle-sidebar-timestamp-color',
    paths: {
      sidebar: 'subtitleSidebar.timestampColor',
    },
  },
  {
    property: '--subtitle-sidebar-active-line-color',
    paths: {
      sidebar: 'subtitleSidebar.activeLineColor',
    },
  },
  {
    property: '--subtitle-sidebar-active-background-color',
    paths: {
      sidebar: 'subtitleSidebar.activeLineBackgroundColor',
    },
  },
  {
    property: '--subtitle-sidebar-hover-background-color',
    paths: {
      sidebar: 'subtitleSidebar.hoverLineBackgroundColor',
    },
  },
];

const CSS_PROPERTY_PATTERN = /^(?:--[A-Za-z0-9_-]+|-?[A-Za-z][A-Za-z0-9_-]*)$/;

export function getSubtitleCssPath(scope: SubtitleCssScope): string {
  if (scope === 'primary') return 'subtitleStyle.css';
  if (scope === 'secondary') return 'subtitleStyle.secondary.css';
  return 'subtitleSidebar.css';
}

export function getSubtitleCssManagedConfigPaths(scope: SubtitleCssScope): string[] {
  return [
    ...new Set(
      LEGACY_CSS_DECLARATIONS.map((declaration) => declaration.paths[scope]).filter(
        (path): path is string => typeof path === 'string' && path.length > 0,
      ),
    ),
  ];
}

export function getSubtitleCssScopeForPath(path: string): SubtitleCssScope | null {
  if (path === 'subtitleStyle.css') return 'primary';
  if (path === 'subtitleStyle.secondary.css') return 'secondary';
  if (path === 'subtitleSidebar.css') return 'sidebar';
  return null;
}

export function serializeSubtitleCssDeclarations(
  scope: SubtitleCssScope,
  values: Record<string, ConfigSettingsSnapshotValue | undefined>,
): string {
  return Object.entries(buildSubtitleCssDeclarationObject(scope, values))
    .map(([property, value]) => `${property}: ${value};`)
    .join('\n');
}

export function buildSubtitleCssDeclarationObject(
  scope: SubtitleCssScope,
  values: Record<string, ConfigSettingsSnapshotValue | undefined>,
): Record<string, string> {
  const declarations = new Map<string, string>();

  for (const declaration of LEGACY_CSS_DECLARATIONS) {
    const path = declaration.paths[scope];
    if (typeof path !== 'string' || path.length === 0) continue;
    const formatted = (declaration.format ?? formatCssPrimitiveValue)(values[path]);
    if (formatted !== undefined) {
      declarations.set(declaration.property, formatted);
    }
  }

  const cssObject = normalizeCssDeclarationRecord(values[getSubtitleCssPath(scope)]);
  for (const [property, value] of Object.entries(cssObject)) {
    declarations.set(normalizeCssPropertyName(property), value);
  }

  return Object.fromEntries(declarations.entries());
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
