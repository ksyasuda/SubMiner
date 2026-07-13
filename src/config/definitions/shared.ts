import type { AnkiConnectConfig } from '../../types/anki';
import type { ResolvedConfig } from '../../types/config';
import type {
  RuntimeOptionId,
  RuntimeOptionScope,
  RuntimeOptionValue,
  RuntimeOptionValueType,
} from '../../types/runtime-options';

export type ConfigValueKind = 'boolean' | 'number' | 'string' | 'enum' | 'array' | 'object';

export interface RuntimeOptionRegistryEntry {
  id: RuntimeOptionId;
  path: string;
  label: string;
  scope: RuntimeOptionScope;
  valueType: RuntimeOptionValueType;
  allowedValues: RuntimeOptionValue[];
  defaultValue: RuntimeOptionValue;
  requiresRestart: boolean;
  formatValueForOsd: (value: RuntimeOptionValue) => string;
  toAnkiPatch: (value: RuntimeOptionValue) => Partial<AnkiConnectConfig>;
}

export interface ConfigOptionRegistryEntry {
  path: string;
  kind: ConfigValueKind;
  defaultValue: unknown;
  description: string;
  /**
   * Complete runtime-valid enum options, including legacy file-config values such as
   * `osd` and `osd-system` in NOTIFICATION_TYPE_VALUES.
   */
  enumValues?: readonly string[];
  enumLabels?: Record<string, string>;
  /**
   * Optional settings UI subset when legacy/runtime-valid enum options should remain
   * editable in config files but hidden from new UI choices, for example
   * SETTINGS_NOTIFICATION_TYPE_VALUES.
   */
  settingsEnumValues?: readonly string[];
  runtime?: RuntimeOptionRegistryEntry;
}

export interface ConfigTemplateSection {
  title: string;
  description: string[];
  key: keyof ResolvedConfig;
  notes?: string[];
}

export const SPECIAL_COMMANDS = {
  SUBSYNC_TRIGGER: '__subsync-trigger',
  RUNTIME_OPTIONS_OPEN: '__runtime-options-open',
  JIMAKU_OPEN: '__jimaku-open',
  /** @deprecated Use TSUKIHIME_OPEN. */
  ANIMETOSHO_OPEN: '__animetosho-open',
  TSUKIHIME_OPEN: '__tsukihime-open',
  RUNTIME_OPTION_CYCLE_PREFIX: '__runtime-option-cycle:',
  REPLAY_SUBTITLE: '__replay-subtitle',
  PLAY_NEXT_SUBTITLE: '__play-next-subtitle',
  YOUTUBE_PICKER_OPEN: '__youtube-picker-open',
  PLAYLIST_BROWSER_OPEN: '__playlist-browser-open',
} as const;

export const DEFAULT_KEYBINDINGS: NonNullable<ResolvedConfig['keybindings']> = [
  { key: 'Space', command: ['cycle', 'pause'] },
  { key: 'KeyF', command: ['cycle', 'fullscreen'] },
  { key: 'KeyJ', command: ['cycle', 'sid'] },
  { key: 'Shift+KeyJ', command: ['cycle', 'secondary-sid'] },
  { key: 'ArrowRight', command: ['seek', 5] },
  { key: 'ArrowLeft', command: ['seek', -5] },
  { key: 'ArrowUp', command: ['seek', 60] },
  { key: 'ArrowDown', command: ['seek', -60] },
  { key: 'Shift+KeyH', command: ['sub-seek', -1] },
  { key: 'Shift+KeyL', command: ['sub-seek', 1] },
  { key: 'Ctrl+Shift+ArrowLeft', command: ['sub-step', -1] },
  { key: 'Ctrl+Shift+ArrowRight', command: ['sub-step', 1] },
  { key: 'KeyZ', command: ['add', 'sub-delay', -0.1] },
  { key: 'Shift+KeyZ', command: ['add', 'sub-delay', 0.1] },
  { key: 'KeyX', command: ['add', 'sub-delay', 0.1] },
  { key: 'Ctrl+Alt+KeyC', command: [SPECIAL_COMMANDS.YOUTUBE_PICKER_OPEN] },
  { key: 'Ctrl+Alt+KeyP', command: [SPECIAL_COMMANDS.PLAYLIST_BROWSER_OPEN] },
  { key: 'Ctrl+Shift+KeyH', command: [SPECIAL_COMMANDS.REPLAY_SUBTITLE] },
  { key: 'Ctrl+Shift+KeyL', command: [SPECIAL_COMMANDS.PLAY_NEXT_SUBTITLE] },
  { key: 'KeyQ', command: ['quit'] },
  { key: 'Ctrl+KeyW', command: ['quit'] },
];
