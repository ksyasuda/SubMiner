import { parse as parseJsonc, type ParseError } from 'jsonc-parser';

export function parseConfigContent(configPath: string, data: string): unknown {
  if (!configPath.endsWith('.jsonc')) {
    return JSON.parse(data);
  }

  const errors: ParseError[] = [];
  const result = parseJsonc(data, errors, {
    allowTrailingComma: true,
    disallowComments: false,
  });
  if (errors.length > 0) {
    throw new Error(`Invalid JSONC (${errors[0]?.error ?? 'unknown'})`);
  }
  return result;
}
