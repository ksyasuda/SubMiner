import type { ConfigSettingsField, ConfigSettingsSnapshotValue } from '../types/settings';

export interface SettingsControlContext {
  setFieldError(path: string, message: string | null): void;
  resetDraftPath(path: string, defaultValue?: ConfigSettingsSnapshotValue): void;
  updateDraft(path: string, value: ConfigSettingsSnapshotValue): void;
  valueForField(field: ConfigSettingsField): ConfigSettingsSnapshotValue;
  valueForPath(path: string): ConfigSettingsSnapshotValue | undefined;
}
