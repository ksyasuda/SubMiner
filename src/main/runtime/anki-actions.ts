type AnkiIntegrationLike = {
  refreshKnownWordCache: () => Promise<void>;
};

export function createUpdateLastCardFromClipboardHandler<TAnki>(deps: {
  getAnkiIntegration: () => TAnki;
  readClipboardText: () => string;
  showMpvOsd: (text: string) => void;
  updateLastCardFromClipboardCore: (options: {
    ankiIntegration: TAnki;
    readClipboardText: () => string;
    showMpvOsd: (text: string) => void;
  }) => Promise<void>;
}) {
  return async (): Promise<void> => {
    await deps.updateLastCardFromClipboardCore({
      ankiIntegration: deps.getAnkiIntegration(),
      readClipboardText: deps.readClipboardText,
      showMpvOsd: deps.showMpvOsd,
    });
  };
}

export function createRefreshKnownWordCacheHandler(deps: {
  getAnkiIntegration: () => AnkiIntegrationLike | null;
  missingIntegrationMessage: string;
}) {
  return async (): Promise<void> => {
    const anki = deps.getAnkiIntegration();
    if (!anki) {
      throw new Error(deps.missingIntegrationMessage);
    }
    await anki.refreshKnownWordCache();
  };
}

export function createTriggerFieldGroupingHandler<TAnki>(deps: {
  getAnkiIntegration: () => TAnki;
  showMpvOsd: (text: string) => void;
  triggerFieldGroupingCore: (options: {
    ankiIntegration: TAnki;
    showMpvOsd: (text: string) => void;
  }) => Promise<void>;
}) {
  return async (): Promise<void> => {
    await deps.triggerFieldGroupingCore({
      ankiIntegration: deps.getAnkiIntegration(),
      showMpvOsd: deps.showMpvOsd,
    });
  };
}

export function createMarkLastCardAsAudioCardHandler<TAnki>(deps: {
  getAnkiIntegration: () => TAnki;
  showMpvOsd: (text: string) => void;
  markLastCardAsAudioCardCore: (options: {
    ankiIntegration: TAnki;
    showMpvOsd: (text: string) => void;
  }) => Promise<void>;
}) {
  return async (): Promise<void> => {
    await deps.markLastCardAsAudioCardCore({
      ankiIntegration: deps.getAnkiIntegration(),
      showMpvOsd: deps.showMpvOsd,
    });
  };
}

export function createMineSentenceCardHandler<TAnki, TMpv>(deps: {
  getAnkiIntegration: () => TAnki;
  getMpvClient: () => TMpv;
  showMpvOsd: (text: string) => void;
  mineSentenceCardCore: (options: {
    ankiIntegration: TAnki;
    mpvClient: TMpv;
    showMpvOsd: (text: string) => void;
  }) => Promise<boolean>;
  recordCardsMined: (count: number) => void;
}) {
  return async (): Promise<void> => {
    const created = await deps.mineSentenceCardCore({
      ankiIntegration: deps.getAnkiIntegration(),
      mpvClient: deps.getMpvClient(),
      showMpvOsd: deps.showMpvOsd,
    });
    if (created) {
      deps.recordCardsMined(1);
    }
  };
}
