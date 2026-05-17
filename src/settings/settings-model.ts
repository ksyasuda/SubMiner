import type {
  ConfigSettingsCategory,
  ConfigSettingsField,
  ConfigSettingsPatchOperation,
  ConfigSettingsSnapshotValue,
} from '../types/settings';

export interface SettingsFilter {
  category: ConfigSettingsCategory;
  query?: string;
}

export interface SettingsDraft {
  readonly initialValues: Record<string, ConfigSettingsSnapshotValue>;
  readonly values: Record<string, ConfigSettingsSnapshotValue>;
  readonly resetPaths: Set<string>;
}

function normalizeQuery(query: string | undefined): string {
  return (query ?? '').trim().toLowerCase();
}

function valuesEqual(a: unknown, b: unknown): boolean {
  return JSON.stringify(a) === JSON.stringify(b);
}

export function filterSettingsFields(
  fields: ConfigSettingsField[],
  filter: SettingsFilter,
): ConfigSettingsField[] {
  const query = normalizeQuery(filter.query);
  return fields.filter((field) => {
    if (field.category !== filter.category || field.legacyHidden) {
      return false;
    }
    if (!query) {
      return true;
    }
    const haystack = [
      field.label,
      field.description,
      field.configPath,
      field.section,
      field.enumValues?.join(' ') ?? '',
    ]
      .join(' ')
      .toLowerCase();
    return haystack.includes(query);
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
