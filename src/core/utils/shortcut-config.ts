import { Config } from '../../types';

export interface ConfiguredShortcuts {
  toggleVisibleOverlayGlobal: string | null | undefined;
  toggleInvisibleOverlayGlobal: string | null | undefined;
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
  openRuntimeOptions: string | null | undefined;
  openJimaku: string | null | undefined;
}

export function resolveConfiguredShortcuts(
  config: Config,
  defaultConfig: Config,
): ConfiguredShortcuts {
  const normalizeShortcut = (value: string | null | undefined): string | null | undefined => {
    if (typeof value !== 'string') return value;
    return value.replace(/\bKey([A-Z])\b/g, '$1').replace(/\bDigit([0-9])\b/g, '$1');
  };

  return {
    toggleVisibleOverlayGlobal: normalizeShortcut(
      config.shortcuts?.toggleVisibleOverlayGlobal ??
        defaultConfig.shortcuts?.toggleVisibleOverlayGlobal,
    ),
    toggleInvisibleOverlayGlobal: normalizeShortcut(
      config.shortcuts?.toggleInvisibleOverlayGlobal ??
        defaultConfig.shortcuts?.toggleInvisibleOverlayGlobal,
    ),
    copySubtitle: normalizeShortcut(
      config.shortcuts?.copySubtitle ?? defaultConfig.shortcuts?.copySubtitle,
    ),
    copySubtitleMultiple: normalizeShortcut(
      config.shortcuts?.copySubtitleMultiple ?? defaultConfig.shortcuts?.copySubtitleMultiple,
    ),
    updateLastCardFromClipboard: normalizeShortcut(
      config.shortcuts?.updateLastCardFromClipboard ??
        defaultConfig.shortcuts?.updateLastCardFromClipboard,
    ),
    triggerFieldGrouping: normalizeShortcut(
      config.shortcuts?.triggerFieldGrouping ?? defaultConfig.shortcuts?.triggerFieldGrouping,
    ),
    triggerSubsync: normalizeShortcut(
      config.shortcuts?.triggerSubsync ?? defaultConfig.shortcuts?.triggerSubsync,
    ),
    mineSentence: normalizeShortcut(
      config.shortcuts?.mineSentence ?? defaultConfig.shortcuts?.mineSentence,
    ),
    mineSentenceMultiple: normalizeShortcut(
      config.shortcuts?.mineSentenceMultiple ?? defaultConfig.shortcuts?.mineSentenceMultiple,
    ),
    multiCopyTimeoutMs:
      config.shortcuts?.multiCopyTimeoutMs ?? defaultConfig.shortcuts?.multiCopyTimeoutMs ?? 5000,
    toggleSecondarySub: normalizeShortcut(
      config.shortcuts?.toggleSecondarySub ?? defaultConfig.shortcuts?.toggleSecondarySub,
    ),
    markAudioCard: normalizeShortcut(
      config.shortcuts?.markAudioCard ?? defaultConfig.shortcuts?.markAudioCard,
    ),
    openRuntimeOptions: normalizeShortcut(
      config.shortcuts?.openRuntimeOptions ?? defaultConfig.shortcuts?.openRuntimeOptions,
    ),
    openJimaku: normalizeShortcut(
      config.shortcuts?.openJimaku ?? defaultConfig.shortcuts?.openJimaku,
    ),
  };
}
