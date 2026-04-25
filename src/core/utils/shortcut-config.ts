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
  openCharacterDictionary: string | null | undefined;
  openRuntimeOptions: string | null | undefined;
  openJimaku: string | null | undefined;
  openSessionHelp: string | null | undefined;
  openControllerSelect: string | null | undefined;
  openControllerDebug: string | null | undefined;
  toggleSubtitleSidebar: string | null | undefined;
}

export function resolveConfiguredShortcuts(
  config: Config,
  defaultConfig: Config,
): ConfiguredShortcuts {
  const isAnkiConnectDisabled = config.ankiConnect?.enabled === false;

  const normalizeShortcut = (value: string | null | undefined): string | null | undefined => {
    if (typeof value !== 'string') return value;
    return value.replace(/\bKey([A-Z])\b/g, '$1').replace(/\bDigit([0-9])\b/g, '$1');
  };

  return {
    toggleVisibleOverlayGlobal: normalizeShortcut(
      config.shortcuts?.toggleVisibleOverlayGlobal ??
        defaultConfig.shortcuts?.toggleVisibleOverlayGlobal,
    ),
    copySubtitle: normalizeShortcut(
      config.shortcuts?.copySubtitle ?? defaultConfig.shortcuts?.copySubtitle,
    ),
    copySubtitleMultiple: normalizeShortcut(
      config.shortcuts?.copySubtitleMultiple ?? defaultConfig.shortcuts?.copySubtitleMultiple,
    ),
    updateLastCardFromClipboard: normalizeShortcut(
      isAnkiConnectDisabled
        ? null
        : (config.shortcuts?.updateLastCardFromClipboard ??
            defaultConfig.shortcuts?.updateLastCardFromClipboard),
    ),
    triggerFieldGrouping: normalizeShortcut(
      isAnkiConnectDisabled
        ? null
        : (config.shortcuts?.triggerFieldGrouping ?? defaultConfig.shortcuts?.triggerFieldGrouping),
    ),
    triggerSubsync: normalizeShortcut(
      config.shortcuts?.triggerSubsync ?? defaultConfig.shortcuts?.triggerSubsync,
    ),
    mineSentence: normalizeShortcut(
      isAnkiConnectDisabled
        ? null
        : (config.shortcuts?.mineSentence ?? defaultConfig.shortcuts?.mineSentence),
    ),
    mineSentenceMultiple: normalizeShortcut(
      isAnkiConnectDisabled
        ? null
        : (config.shortcuts?.mineSentenceMultiple ?? defaultConfig.shortcuts?.mineSentenceMultiple),
    ),
    multiCopyTimeoutMs:
      config.shortcuts?.multiCopyTimeoutMs ?? defaultConfig.shortcuts?.multiCopyTimeoutMs ?? 5000,
    toggleSecondarySub: normalizeShortcut(
      config.shortcuts?.toggleSecondarySub ?? defaultConfig.shortcuts?.toggleSecondarySub,
    ),
    markAudioCard: normalizeShortcut(
      isAnkiConnectDisabled
        ? null
        : (config.shortcuts?.markAudioCard ?? defaultConfig.shortcuts?.markAudioCard),
    ),
    openCharacterDictionary: normalizeShortcut(
      config.shortcuts?.openCharacterDictionary ?? defaultConfig.shortcuts?.openCharacterDictionary,
    ),
    openRuntimeOptions: normalizeShortcut(
      config.shortcuts?.openRuntimeOptions ?? defaultConfig.shortcuts?.openRuntimeOptions,
    ),
    openJimaku: normalizeShortcut(
      config.shortcuts?.openJimaku ?? defaultConfig.shortcuts?.openJimaku,
    ),
    openSessionHelp: normalizeShortcut(
      config.shortcuts?.openSessionHelp ?? defaultConfig.shortcuts?.openSessionHelp,
    ),
    openControllerSelect: normalizeShortcut(
      config.shortcuts?.openControllerSelect ?? defaultConfig.shortcuts?.openControllerSelect,
    ),
    openControllerDebug: normalizeShortcut(
      config.shortcuts?.openControllerDebug ?? defaultConfig.shortcuts?.openControllerDebug,
    ),
    toggleSubtitleSidebar: normalizeShortcut(
      config.shortcuts?.toggleSubtitleSidebar ?? defaultConfig.shortcuts?.toggleSubtitleSidebar,
    ),
  };
}
