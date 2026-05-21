import type {
  ConfigSettingsCategory,
  ConfigSettingsField,
  ConfigSettingsPatchOperation,
  ConfigSettingsSnapshotValue,
} from '../types/settings';

export interface SettingsFilter {
  category?: ConfigSettingsCategory;
  query?: string;
}

export interface SettingsDraft {
  readonly initialValues: Record<string, ConfigSettingsSnapshotValue>;
  readonly values: Record<string, ConfigSettingsSnapshotValue>;
  readonly resetPaths: Set<string>;
}

function normalizeQuery(query: string | undefined): string {
  return (query ?? '').trim().toLocaleLowerCase();
}

function searchableText(parts: Array<string | undefined>): string {
  return parts
    .filter(Boolean)
    .join(' ')
    .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
    .replace(/[^\p{L}\p{N}]+/gu, ' ')
    .toLocaleLowerCase();
}

function valuesEqual(a: unknown, b: unknown): boolean {
  return JSON.stringify(a) === JSON.stringify(b);
}

export function filterSettingsFields(
  fields: ConfigSettingsField[],
  filter: SettingsFilter,
): ConfigSettingsField[] {
  const query = normalizeQuery(filter.query);
  const terms = query.length > 0 ? searchableText([query]).split(/\s+/).filter(Boolean) : [];
  return fields.filter((field) => {
    if (field.legacyHidden || field.settingsHidden) {
      return false;
    }
    if (filter.category && field.category !== filter.category) {
      return false;
    }
    if (!query || terms.length === 0) {
      return true;
    }
    const haystack = searchableText([
      field.label,
      field.description,
      field.configPath,
      field.section,
      field.subsection ?? '',
      field.enumValues?.join(' ') ?? '',
    ]);
    return terms.every((term) => haystack.includes(term));
  });
}

export function createSettingsDraft(
  values: Record<string, ConfigSettingsSnapshotValue>,
): SettingsDraft {
  return {
    initialValues: structuredClone(values),
    values: structuredClone(values),
    resetPaths: new Set(),
  };
}

export function toSettingsDisplayValue(
  path: string,
  value: ConfigSettingsSnapshotValue,
): ConfigSettingsSnapshotValue {
  if (path === 'websocket.enabled' && typeof value === 'boolean') {
    return value ? 'true' : 'false';
  }
  if (path === 'discordPresence.updateIntervalMs' && typeof value === 'number') {
    return value / 1000;
  }
  return value;
}

export function toConfigDraftValue(
  path: string,
  value: ConfigSettingsSnapshotValue,
): ConfigSettingsSnapshotValue {
  if (path === 'websocket.enabled') {
    if (value === 'true') return true;
    if (value === 'false') return false;
  }
  if (path === 'discordPresence.updateIntervalMs' && typeof value === 'number') {
    return Math.round(value * 1000);
  }
  return value;
}

export function setDraftValue(
  draft: SettingsDraft,
  path: string,
  value: ConfigSettingsSnapshotValue,
): void {
  draft.values[path] = value;
  draft.resetPaths.delete(path);
}

export function resetDraftPath(draft: SettingsDraft, path: string, defaultValue: unknown): void {
  draft.values[path] = structuredClone(defaultValue);
  draft.resetPaths.add(path);
}

export function getDirtyOperations(draft: SettingsDraft): ConfigSettingsPatchOperation[] {
  const operations: ConfigSettingsPatchOperation[] = [];
  const paths = new Set([...Object.keys(draft.initialValues), ...Object.keys(draft.values)]);

  for (const path of [...paths].sort()) {
    if (draft.resetPaths.has(path)) {
      operations.push({ op: 'reset', path });
      continue;
    }
    if (!valuesEqual(draft.values[path], draft.initialValues[path])) {
      operations.push({
        op: 'set',
        path,
        value: draft.values[path],
      });
    }
  }

  return operations;
}
