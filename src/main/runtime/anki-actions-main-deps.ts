import type { createRefreshKnownWordCacheHandler } from './anki-actions';

type RefreshKnownWordCacheMainDeps = Parameters<typeof createRefreshKnownWordCacheHandler>[0];

export function createBuildUpdateLastCardFromClipboardMainDepsHandler<TAnki>(deps: {
  getAnkiIntegration: () => TAnki;
  readClipboardText: () => string;
  showMpvOsd: (text: string) => void;
  updateLastCardFromClipboardCore: (options: {
    ankiIntegration: TAnki;
    readClipboardText: () => string;
    showMpvOsd: (text: string) => void;
  }) => Promise<void>;
}) {
  return () => ({
    getAnkiIntegration: () => deps.getAnkiIntegration(),
    readClipboardText: () => deps.readClipboardText(),
    showMpvOsd: (text: string) => deps.showMpvOsd(text),
    updateLastCardFromClipboardCore: (options: {
      ankiIntegration: TAnki;
      readClipboardText: () => string;
      showMpvOsd: (text: string) => void;
    }) => deps.updateLastCardFromClipboardCore(options),
  });
}

export function createBuildRefreshKnownWordCacheMainDepsHandler(
  deps: RefreshKnownWordCacheMainDeps,
) {
  return (): RefreshKnownWordCacheMainDeps => ({
    getAnkiIntegration: () => deps.getAnkiIntegration(),
    missingIntegrationMessage: deps.missingIntegrationMessage,
  });
}

export function createBuildTriggerFieldGroupingMainDepsHandler<TAnki>(deps: {
  getAnkiIntegration: () => TAnki;
  showMpvOsd: (text: string) => void;
  triggerFieldGroupingCore: (options: {
    ankiIntegration: TAnki;
    showMpvOsd: (text: string) => void;
  }) => Promise<void>;
}) {
  return () => ({
    getAnkiIntegration: () => deps.getAnkiIntegration(),
    showMpvOsd: (text: string) => deps.showMpvOsd(text),
    triggerFieldGroupingCore: (options: {
      ankiIntegration: TAnki;
      showMpvOsd: (text: string) => void;
    }) => deps.triggerFieldGroupingCore(options),
  });
}

export function createBuildMarkLastCardAsAudioCardMainDepsHandler<TAnki>(deps: {
  getAnkiIntegration: () => TAnki;
  showMpvOsd: (text: string) => void;
  markLastCardAsAudioCardCore: (options: {
    ankiIntegration: TAnki;
    showMpvOsd: (text: string) => void;
  }) => Promise<void>;
}) {
  return () => ({
    getAnkiIntegration: () => deps.getAnkiIntegration(),
    showMpvOsd: (text: string) => deps.showMpvOsd(text),
    markLastCardAsAudioCardCore: (options: {
      ankiIntegration: TAnki;
      showMpvOsd: (text: string) => void;
    }) => deps.markLastCardAsAudioCardCore(options),
  });
}

export function createBuildMineSentenceCardMainDepsHandler<TAnki, TMpv>(deps: {
  getAnkiIntegration: () => TAnki;
  getMpvClient: () => TMpv;
  showMpvOsd: (text: string) => void;
  mineSentenceCardCore: (options: {
    ankiIntegration: TAnki;
    mpvClient: TMpv;
    showMpvOsd: (text: string) => void;
  }) => Promise<boolean>;
  recordCardsMined: (count: number, noteIds?: number[]) => void;
}) {
  return () => ({
    getAnkiIntegration: () => deps.getAnkiIntegration(),
    getMpvClient: () => deps.getMpvClient(),
    showMpvOsd: (text: string) => deps.showMpvOsd(text),
    mineSentenceCardCore: (options: {
      ankiIntegration: TAnki;
      mpvClient: TMpv;
      showMpvOsd: (text: string) => void;
    }) => deps.mineSentenceCardCore(options),
    recordCardsMined: (count: number, noteIds?: number[]) => deps.recordCardsMined(count, noteIds),
  });
}
