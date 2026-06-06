import { Config } from '../../types';

export interface ConfiguredShortcuts {
  toggleVisibleOverlayGlobal: string | null | undefined;
  copySubtitle: string | null | undefined;
  copySubtitleMultiple: string | null | undefined;
  updateLastCardFromClipboard: string | null | undefined;
  triggerFieldGrouping: string | null | undefined;
  triggerSubsync: string | null | undefined;
  mineSentence: string | null | undefined;
  mineSentenceMultiple: string | null | undefined;
  multiCopyTimeoutMs: number;
  toggleSecondarySub: string | null | undefined;
  markAudioCard: string | null | undefined;
  openCharacterDictionaryManager: string | null | undefined;
  openRuntimeOptions: string | null | undefined;
  openJimaku: string | null | undefined;
  openSessionHelp: string | null | undefined;
  openControllerSelect: string | null | undefined;
  openControllerDebug: string | null | undefined;
  toggleSubtitleSidebar: string | null | undefined;
  toggleNotificationHistory: string | null | undefined;
}

export function resolveConfiguredShortcuts(
  config: Config,
  defaultConfig: Config,
): ConfiguredShortcuts {
  const isAnkiConnectDisabled = config.ankiConnect?.enabled === false;
  type ShortcutKey = keyof Omit<ConfiguredShortcuts, 'multiCopyTimeoutMs'> &
    keyof NonNullable<Config['shortcuts']>;

  const normalizeShortcut = (value: string | null | undefined): string | null | undefined => {
    if (typeof value !== 'string') return value;
    return value.replace(/\bKey([A-Z])\b/g, '$1').replace(/\bDigit([0-9])\b/g, '$1');
  };

  const shortcutValue = (key: ShortcutKey): string | null | undefined =>
    Object.prototype.hasOwnProperty.call(config.shortcuts ?? {}, key)
      ? config.shortcuts?.[key]
      : defaultConfig.shortcuts?.[key];

  return {
    toggleVisibleOverlayGlobal: normalizeShortcut(shortcutValue('toggleVisibleOverlayGlobal')),
    copySubtitle: normalizeShortcut(shortcutValue('copySubtitle')),
    copySubtitleMultiple: normalizeShortcut(shortcutValue('copySubtitleMultiple')),
    updateLastCardFromClipboard: normalizeShortcut(
      isAnkiConnectDisabled ? null : shortcutValue('updateLastCardFromClipboard'),
    ),
    triggerFieldGrouping: normalizeShortcut(
      isAnkiConnectDisabled ? null : shortcutValue('triggerFieldGrouping'),
    ),
    triggerSubsync: normalizeShortcut(shortcutValue('triggerSubsync')),
    mineSentence: normalizeShortcut(isAnkiConnectDisabled ? null : shortcutValue('mineSentence')),
    mineSentenceMultiple: normalizeShortcut(
      isAnkiConnectDisabled ? null : shortcutValue('mineSentenceMultiple'),
    ),
    multiCopyTimeoutMs:
      config.shortcuts?.multiCopyTimeoutMs ?? defaultConfig.shortcuts?.multiCopyTimeoutMs ?? 5000,
    toggleSecondarySub: normalizeShortcut(shortcutValue('toggleSecondarySub')),
    markAudioCard: normalizeShortcut(isAnkiConnectDisabled ? null : shortcutValue('markAudioCard')),
    openCharacterDictionaryManager: normalizeShortcut(
      shortcutValue('openCharacterDictionaryManager'),
    ),
    openRuntimeOptions: normalizeShortcut(shortcutValue('openRuntimeOptions')),
    openJimaku: normalizeShortcut(shortcutValue('openJimaku')),
    openSessionHelp: normalizeShortcut(shortcutValue('openSessionHelp')),
    openControllerSelect: normalizeShortcut(shortcutValue('openControllerSelect')),
    openControllerDebug: normalizeShortcut(shortcutValue('openControllerDebug')),
    toggleSubtitleSidebar: normalizeShortcut(shortcutValue('toggleSubtitleSidebar')),
    toggleNotificationHistory: normalizeShortcut(shortcutValue('toggleNotificationHistory')),
  };
}
