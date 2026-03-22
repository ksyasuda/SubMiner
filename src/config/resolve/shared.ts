export function isObject(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

export function asNumber(value: unknown): number | undefined {
  return typeof value === 'number' && Number.isFinite(value) ? value : undefined;
}

export function asString(value: unknown): string | undefined {
  return typeof value === 'string' ? value : undefined;
}

export function asBoolean(value: unknown): boolean | undefined {
  return typeof value === 'boolean' ? value : undefined;
}

const hexColorPattern = /^#(?:[0-9a-fA-F]{3}|[0-9a-fA-F]{4}|[0-9a-fA-F]{6}|[0-9a-fA-F]{8})$/;
const cssColorKeywords = new Set([
  'transparent',
  'currentcolor',
  'inherit',
  'initial',
  'unset',
  'revert',
  'revert-layer',
]);
const cssColorFunctionPattern = /^(?:rgba?|hsla?)\(\s*[^()]+?\s*\)$/i;

function supportsCssColor(text: string): boolean {
  const css = (globalThis as { CSS?: { supports?: (property: string, value: string) => boolean } })
    .CSS;
  return css?.supports?.('color', text) ?? false;
}

export function asColor(value: unknown): string | undefined {
  if (typeof value !== 'string') return undefined;
  const text = value.trim();
  return hexColorPattern.test(text) ? text : undefined;
}

export function asCssColor(value: unknown): string | undefined {
  if (typeof value !== 'string') return undefined;

  const text = value.trim();
  if (text.length === 0) {
    return undefined;
  }

  if (supportsCssColor(text)) {
    return text;
  }

  const normalized = text.toLowerCase();
  if (
    hexColorPattern.test(text) ||
    cssColorKeywords.has(normalized) ||
    cssColorFunctionPattern.test(text)
  ) {
    return text;
  }

  return undefined;
}

export function asFrequencyBandedColors(
  value: unknown,
): [string, string, string, string, string] | undefined {
  if (!Array.isArray(value) || value.length !== 5) {
    return undefined;
  }

  const colors = value.map((item) => asColor(item));
  if (colors.some((color) => color === undefined)) {
    return undefined;
  }

  return colors as [string, string, string, string, string];
}
