import type { ConfigValidationWarning } from './config';

export type ConfigSettingsCategory =
  | 'appearance'
  | 'behavior'
  | 'mining-anki'
  | 'playback-sources'
  | 'input'
  | 'integrations'
  | 'tracking-app'
  | 'advanced';

export type ConfigSettingsControl =
  | 'boolean'
  | 'number'
  | 'text'
  | 'textarea'
  | 'select'
  | 'color'
  | 'string-list'
  | 'json'
  | 'secret'
  | 'keyboard-shortcut'
  | 'key-code'
  | 'known-words-decks'
  | 'anki-note-type'
  | 'anki-field'
  | 'mpv-keybindings'
  | 'color-list';

export type ConfigSettingsRestartBehavior = 'hot-reload' | 'restart';

export interface ConfigSettingsField {
  id: string;
  label: string;
  description: string;
  configPath: string;
  category: ConfigSettingsCategory;
  section: string;
  subsection?: string;
  control: ConfigSettingsControl;
  defaultValue: unknown;
  enumValues?: readonly string[];
  restartBehavior: ConfigSettingsRestartBehavior;
  advanced?: boolean;
  secret?: boolean;
  legacyHidden?: boolean;
}

export type ConfigSettingsSnapshotValue = unknown;

export interface ConfigSettingsSnapshot {
  configPath: string;
  fields: ConfigSettingsField[];
  values: Record<string, ConfigSettingsSnapshotValue>;
  warnings: ConfigValidationWarning[];
}

export type ConfigSettingsPatchOperation =
  | {
      op: 'set';
      path: string;
      value: unknown;
    }
  | {
      op: 'reset';
      path: string;
    };

export interface ConfigSettingsPatch {
  operations: ConfigSettingsPatchOperation[];
}

export interface ConfigSettingsSaveResult {
  ok: boolean;
  snapshot?: ConfigSettingsSnapshot;
  warnings?: ConfigValidationWarning[];
  error?: string;
  hotReloadFields: string[];
  restartRequiredFields: string[];
  restartRequiredSections?: string[];
}

export interface ConfigSettingsAPI {
  getSnapshot(): Promise<ConfigSettingsSnapshot>;
  savePatch(patch: ConfigSettingsPatch): Promise<ConfigSettingsSaveResult>;
  openSettingsFile(): Promise<boolean>;
  openSettingsWindow(): Promise<boolean>;
  getAnkiDeckNames(draftUrl?: string): Promise<ConfigSettingsAnkiListResult>;
  getAnkiDeckFieldNames(
    deckName: string,
    draftUrl?: string,
  ): Promise<ConfigSettingsAnkiListResult>;
  getAnkiModelNames(draftUrl?: string): Promise<ConfigSettingsAnkiListResult>;
  getAnkiModelFieldNames(
    modelName: string,
    draftUrl?: string,
  ): Promise<ConfigSettingsAnkiListResult>;
}

export interface ConfigSettingsAnkiListResult {
  ok: boolean;
  values: string[];
  error?: string;
}
