// Creates the MeCab tokenizer on first use, then runs its availability check.
export function createMecabTokenizerAndCheckHandler<TMecab>(deps: {
  getMecabTokenizer: () => TMecab | null;
  setMecabTokenizer: (tokenizer: TMecab) => void;
  createMecabTokenizer: () => TMecab;
  checkAvailability: (tokenizer: TMecab) => Promise<unknown>;
}) {
  return async (): Promise<void> => {
    let tokenizer = deps.getMecabTokenizer();
    if (!tokenizer) {
      tokenizer = deps.createMecabTokenizer();
      deps.setMecabTokenizer(tokenizer);
    }
    await deps.checkAvailability(tokenizer);
  };
}

// Loads the JLPT and frequency dictionaries once (concurrent callers share the in-flight load),
// optionally showing a spinner OSD while loading.
export function createPrewarmSubtitleDictionariesHandler(deps: {
  ensureJlptDictionaryLookup: () => Promise<void>;
  ensureFrequencyDictionaryLookup: () => Promise<void>;
  showMpvOsd?: (message: string) => void;
  showLoadingOsd?: (message: string) => void;
  showLoadedOsd?: (message: string) => void;
  shouldShowOsdNotification?: () => boolean;
  setInterval?: (callback: () => void, delayMs: number) => unknown;
  clearInterval?: (timer: unknown) => void;
}) {
  let prewarmed = false;
  let prewarmPromise: Promise<void> | null = null;
  let loadingOsdDepth = 0;
  let loadingOsdFrame = 0;
  let loadingOsdTimer: unknown = null;
  const showMpvOsd = deps.showMpvOsd;
  const showLoadingOsd = deps.showLoadingOsd ?? showMpvOsd;
  const showLoadedOsd = deps.showLoadedOsd ?? showMpvOsd;
  const setIntervalHandler =
    deps.setInterval ??
    ((callback: () => void, delayMs: number): unknown => setInterval(callback, delayMs));
  const clearIntervalHandler =
    deps.clearInterval ??
    ((timer: unknown): void => clearInterval(timer as ReturnType<typeof setInterval>));
  const spinnerFrames = ['|', '/', '-', '\\'];

  const beginLoadingOsd = (): boolean => {
    if (!showLoadingOsd) {
      return false;
    }
    loadingOsdDepth += 1;
    if (loadingOsdDepth > 1) {
      return true;
    }

    loadingOsdFrame = 0;
    showLoadingOsd(`Loading subtitle annotations ${spinnerFrames[loadingOsdFrame]}`);
    loadingOsdFrame += 1;
    loadingOsdTimer = setIntervalHandler(() => {
      if (!showLoadingOsd) {
        return;
      }
      showLoadingOsd(
        `Loading subtitle annotations ${spinnerFrames[loadingOsdFrame % spinnerFrames.length]}`,
      );
      loadingOsdFrame += 1;
    }, 180);
    return true;
  };

  const endLoadingOsd = (): void => {
    if (!showLoadedOsd) {
      return;
    }

    loadingOsdDepth = Math.max(0, loadingOsdDepth - 1);
    if (loadingOsdDepth > 0) {
      return;
    }

    if (loadingOsdTimer) {
      clearIntervalHandler(loadingOsdTimer);
      loadingOsdTimer = null;
    }
    showLoadedOsd('Subtitle annotations loaded');
  };

  return async (options?: { showLoadingOsd?: boolean }): Promise<void> => {
    if (prewarmed) {
      return;
    }
    const shouldTrackLoadingOsd = options?.showLoadingOsd === true;
    const loadingOsdStarted = shouldTrackLoadingOsd ? beginLoadingOsd() : false;

    if (!prewarmPromise) {
      prewarmPromise = (async () => {
        try {
          await Promise.all([
            deps.ensureJlptDictionaryLookup(),
            deps.ensureFrequencyDictionaryLookup(),
          ]);
          prewarmed = true;
        } finally {
          prewarmPromise = null;
        }
      })();
    }
    try {
      await prewarmPromise;
    } finally {
      if (loadingOsdStarted) {
        endLoadingOsd();
      }
    }
  };
}
