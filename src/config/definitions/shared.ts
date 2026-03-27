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
  enumValues?: readonly string[];
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
  RUNTIME_OPTION_CYCLE_PREFIX: '__runtime-option-cycle:',
  REPLAY_SUBTITLE: '__replay-subtitle',
  PLAY_NEXT_SUBTITLE: '__play-next-subtitle',
  SHIFT_SUB_DELAY_TO_NEXT_SUBTITLE_START: '__sub-delay-next-line',
  SHIFT_SUB_DELAY_TO_PREVIOUS_SUBTITLE_START: '__sub-delay-prev-line',
  YOUTUBE_PICKER_OPEN: '__youtube-picker-open',
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
  { key: 'Shift+BracketRight', command: [SPECIAL_COMMANDS.SHIFT_SUB_DELAY_TO_NEXT_SUBTITLE_START] },
  {
    key: 'Shift+BracketLeft',
    command: [SPECIAL_COMMANDS.SHIFT_SUB_DELAY_TO_PREVIOUS_SUBTITLE_START],
  },
  { key: 'Ctrl+Alt+KeyC', command: [SPECIAL_COMMANDS.YOUTUBE_PICKER_OPEN] },
  { key: 'Ctrl+Shift+KeyH', command: [SPECIAL_COMMANDS.REPLAY_SUBTITLE] },
  { key: 'Ctrl+Shift+KeyL', command: [SPECIAL_COMMANDS.PLAY_NEXT_SUBTITLE] },
  { key: 'KeyQ', command: ['quit'] },
  { key: 'Ctrl+KeyW', command: ['quit'] },
];
