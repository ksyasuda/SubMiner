export function createHandleMultiCopyDigitHandler<TSubtitleTimingTracker>(deps: {
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
  return (count: number): void => {
    deps.handleMultiCopyDigitCore(count, {
      subtitleTimingTracker: deps.getSubtitleTimingTracker(),
      writeClipboardText: deps.writeClipboardText,
      showMpvOsd: deps.showMpvOsd,
    });
  };
}

export function createCopyCurrentSubtitleHandler<TSubtitleTimingTracker>(deps: {
  getSubtitleTimingTracker: () => TSubtitleTimingTracker;
  writeClipboardText: (text: string) => void;
  showMpvOsd: (text: string) => void;
  copyCurrentSubtitleCore: (options: {
    subtitleTimingTracker: TSubtitleTimingTracker;
    writeClipboardText: (text: string) => void;
    showMpvOsd: (text: string) => void;
  }) => void;
}) {
  return (): void => {
    deps.copyCurrentSubtitleCore({
      subtitleTimingTracker: deps.getSubtitleTimingTracker(),
      writeClipboardText: deps.writeClipboardText,
      showMpvOsd: deps.showMpvOsd,
    });
  };
}

export function createHandleMineSentenceDigitHandler<
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
  return (count: number): void => {
    deps.handleMineSentenceDigitCore(count, {
      subtitleTimingTracker: deps.getSubtitleTimingTracker(),
      ankiIntegration: deps.getAnkiIntegration(),
      getCurrentSecondarySubText: deps.getCurrentSecondarySubText,
      showMpvOsd: deps.showMpvOsd,
      logError: deps.logError,
      onCardsMined: deps.onCardsMined,
    });
  };
}
