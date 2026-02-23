export function createBuildHandleMultiCopyDigitMainDepsHandler<TSubtitleTimingTracker>(deps: {
  getSubtitleTimingTracker: () => TSubtitleTimingTracker;
  writeClipboardText: (text: string) => void;
  showMpvOsd: (text: string) => void;
  handleMultiCopyDigitCore: (
    count: number,
    options: {
      subtitleTimingTracker: TSubtitleTimingTracker;
      writeClipboardText: (text: string) => void;
      showMpvOsd: (text: string) => void;
    },
  ) => void;
}) {
  return () => ({
    getSubtitleTimingTracker: () => deps.getSubtitleTimingTracker(),
    writeClipboardText: (text: string) => deps.writeClipboardText(text),
    showMpvOsd: (text: string) => deps.showMpvOsd(text),
    handleMultiCopyDigitCore: (
      count: number,
      options: {
        subtitleTimingTracker: TSubtitleTimingTracker;
        writeClipboardText: (text: string) => void;
        showMpvOsd: (text: string) => void;
      },
    ) => deps.handleMultiCopyDigitCore(count, options),
  });
}

export function createBuildCopyCurrentSubtitleMainDepsHandler<TSubtitleTimingTracker>(deps: {
  getSubtitleTimingTracker: () => TSubtitleTimingTracker;
  writeClipboardText: (text: string) => void;
  showMpvOsd: (text: string) => void;
  copyCurrentSubtitleCore: (options: {
    subtitleTimingTracker: TSubtitleTimingTracker;
    writeClipboardText: (text: string) => void;
    showMpvOsd: (text: string) => void;
  }) => void;
}) {
  return () => ({
    getSubtitleTimingTracker: () => deps.getSubtitleTimingTracker(),
    writeClipboardText: (text: string) => deps.writeClipboardText(text),
    showMpvOsd: (text: string) => deps.showMpvOsd(text),
    copyCurrentSubtitleCore: (options: {
      subtitleTimingTracker: TSubtitleTimingTracker;
      writeClipboardText: (text: string) => void;
      showMpvOsd: (text: string) => void;
    }) => deps.copyCurrentSubtitleCore(options),
  });
}

export function createBuildHandleMineSentenceDigitMainDepsHandler<
  TSubtitleTimingTracker,
  TAnkiIntegration,
>(deps: {
  getSubtitleTimingTracker: () => TSubtitleTimingTracker;
  getAnkiIntegration: () => TAnkiIntegration;
  getCurrentSecondarySubText: () => string | undefined;
  showMpvOsd: (text: string) => void;
  logError: (message: string, err: unknown) => void;
  onCardsMined: (count: number) => void;
  handleMineSentenceDigitCore: (
    count: number,
    options: {
      subtitleTimingTracker: TSubtitleTimingTracker;
      ankiIntegration: TAnkiIntegration;
      getCurrentSecondarySubText: () => string | undefined;
      showMpvOsd: (text: string) => void;
      logError: (message: string, err: unknown) => void;
      onCardsMined: (count: number) => void;
    },
  ) => void;
}) {
  return () => ({
    getSubtitleTimingTracker: () => deps.getSubtitleTimingTracker(),
    getAnkiIntegration: () => deps.getAnkiIntegration(),
    getCurrentSecondarySubText: () => deps.getCurrentSecondarySubText(),
    showMpvOsd: (text: string) => deps.showMpvOsd(text),
    logError: (message: string, err: unknown) => deps.logError(message, err),
    onCardsMined: (count: number) => deps.onCardsMined(count),
    handleMineSentenceDigitCore: (
      count: number,
      options: {
        subtitleTimingTracker: TSubtitleTimingTracker;
        ankiIntegration: TAnkiIntegration;
        getCurrentSecondarySubText: () => string | undefined;
        showMpvOsd: (text: string) => void;
        logError: (message: string, err: unknown) => void;
        onCardsMined: (count: number) => void;
      },
    ) => deps.handleMineSentenceDigitCore(count, options),
  });
}
