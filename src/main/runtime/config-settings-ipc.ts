import type { ConfigSettingsField, ConfigSettingsPatch } from '../../types/settings';

export function isConfigSettingsPatch(
  value: unknown,
  fields: readonly ConfigSettingsField[],
): value is ConfigSettingsPatch {
  if (!value || typeof value !== 'object') {
    return false;
  }
  const operations = (value as { operations?: unknown }).operations;
  return (
    Array.isArray(operations) &&
    operations.every((operation) => {
      if (!operation || typeof operation !== 'object') {
        return false;
      }
      const candidate = operation as { op?: unknown; path?: unknown; value?: unknown };
      const knownPath =
        typeof candidate.path === 'string' &&
        fields.some((field) => field.configPath === candidate.path);
      if (!knownPath) {
        return false;
      }
      if (candidate.op === 'set') {
        return 'value' in candidate;
      }
      return candidate.op === 'reset';
    })
  );
}
