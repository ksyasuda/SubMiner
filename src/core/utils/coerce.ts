export function asFiniteNumber(
  value: unknown,
  fallback: number,
  min?: number,
  max?: number,
): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) return fallback;
  if (min !== undefined && value < min) return min;
  if (max !== undefined && value > max) return max;
  return value;
}

export function asString(value: unknown, fallback: string): string {
  if (typeof value !== 'string') return fallback;
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : fallback;
}

export function asBoolean(value: unknown, fallback: boolean): boolean {
  if (typeof value === 'boolean') return value;
  if (typeof value === 'string') {
    const normalized = value.trim().toLowerCase();
    if (normalized === 'yes' || normalized === 'true' || normalized === '1') {
      return true;
    }
    if (normalized === 'no' || normalized === 'false' || normalized === '0') {
      return false;
    }
  }
  return fallback;
}
