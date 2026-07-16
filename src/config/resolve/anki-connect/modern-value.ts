import type { ResolveContext } from '../context';
import { asNumber } from '../shared';
import { hasOwn } from './shared';

export function asIntegerInRange(value: unknown, min: number, max: number): number | undefined {
  const parsed = asNumber(value);
  return parsed !== undefined && Number.isInteger(parsed) && parsed >= min && parsed <= max
    ? parsed
    : undefined;
}

export function asNonNegativeInteger(value: unknown): number | undefined {
  const parsed = asNumber(value);
  return parsed !== undefined && Number.isInteger(parsed) && parsed >= 0 ? parsed : undefined;
}

export function asPositiveNumber(value: unknown): number | undefined {
  const parsed = asNumber(value);
  return parsed !== undefined && parsed > 0 ? parsed : undefined;
}

export function asNonNegativeNumber(value: unknown): number | undefined {
  const parsed = asNumber(value);
  return parsed !== undefined && parsed >= 0 ? parsed : undefined;
}

export function applyModernValue<T>(
  context: ResolveContext,
  source: Record<string, unknown>,
  key: string,
  path: string,
  parse: (value: unknown) => T | undefined,
  fallback: T,
  apply: (value: T) => void,
  message: string,
): void {
  if (!hasOwn(source, key)) return;
  const raw = source[key];
  const parsed = parse(raw);
  if (parsed === undefined) {
    apply(fallback);
    context.warn(path, raw, fallback, message);
    return;
  }
  apply(parsed);
}
