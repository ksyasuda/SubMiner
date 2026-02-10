import {
  copyCurrentSubtitleService,
  handleMineSentenceDigitService,
  handleMultiCopyDigitService,
  markLastCardAsAudioCardService,
  mineSentenceCardService,
  triggerFieldGroupingService,
  updateLastCardFromClipboardService,
} from "./mining-runtime-service";

export function createHandleMultiCopyDigitDepsRuntimeService(
  options: {
    subtitleTimingTracker: Parameters<typeof handleMultiCopyDigitService>[1]["subtitleTimingTracker"];
    writeClipboardText: Parameters<typeof handleMultiCopyDigitService>[1]["writeClipboardText"];
    showMpvOsd: Parameters<typeof handleMultiCopyDigitService>[1]["showMpvOsd"];
  },
): Parameters<typeof handleMultiCopyDigitService>[1] {
  return {
    subtitleTimingTracker: options.subtitleTimingTracker,
    writeClipboardText: options.writeClipboardText,
    showMpvOsd: options.showMpvOsd,
  };
}

export function createCopyCurrentSubtitleDepsRuntimeService(
  options: {
    subtitleTimingTracker: Parameters<typeof copyCurrentSubtitleService>[0]["subtitleTimingTracker"];
    writeClipboardText: Parameters<typeof copyCurrentSubtitleService>[0]["writeClipboardText"];
    showMpvOsd: Parameters<typeof copyCurrentSubtitleService>[0]["showMpvOsd"];
  },
): Parameters<typeof copyCurrentSubtitleService>[0] {
  return {
    subtitleTimingTracker: options.subtitleTimingTracker,
    writeClipboardText: options.writeClipboardText,
    showMpvOsd: options.showMpvOsd,
  };
}

export function createUpdateLastCardFromClipboardDepsRuntimeService(
  options: {
    ankiIntegration: Parameters<typeof updateLastCardFromClipboardService>[0]["ankiIntegration"];
    readClipboardText: Parameters<typeof updateLastCardFromClipboardService>[0]["readClipboardText"];
    showMpvOsd: Parameters<typeof updateLastCardFromClipboardService>[0]["showMpvOsd"];
  },
): Parameters<typeof updateLastCardFromClipboardService>[0] {
  return {
    ankiIntegration: options.ankiIntegration,
    readClipboardText: options.readClipboardText,
    showMpvOsd: options.showMpvOsd,
  };
}

export function createTriggerFieldGroupingDepsRuntimeService(
  options: {
    ankiIntegration: Parameters<typeof triggerFieldGroupingService>[0]["ankiIntegration"];
    showMpvOsd: Parameters<typeof triggerFieldGroupingService>[0]["showMpvOsd"];
  },
): Parameters<typeof triggerFieldGroupingService>[0] {
  return {
    ankiIntegration: options.ankiIntegration,
    showMpvOsd: options.showMpvOsd,
  };
}

export function createMarkLastCardAsAudioCardDepsRuntimeService(
  options: {
    ankiIntegration: Parameters<typeof markLastCardAsAudioCardService>[0]["ankiIntegration"];
    showMpvOsd: Parameters<typeof markLastCardAsAudioCardService>[0]["showMpvOsd"];
  },
): Parameters<typeof markLastCardAsAudioCardService>[0] {
  return {
    ankiIntegration: options.ankiIntegration,
    showMpvOsd: options.showMpvOsd,
  };
}

export function createMineSentenceCardDepsRuntimeService(
  options: {
    ankiIntegration: Parameters<typeof mineSentenceCardService>[0]["ankiIntegration"];
    mpvClient: Parameters<typeof mineSentenceCardService>[0]["mpvClient"];
    showMpvOsd: Parameters<typeof mineSentenceCardService>[0]["showMpvOsd"];
  },
): Parameters<typeof mineSentenceCardService>[0] {
  return {
    ankiIntegration: options.ankiIntegration,
    mpvClient: options.mpvClient,
    showMpvOsd: options.showMpvOsd,
  };
}

export function createHandleMineSentenceDigitDepsRuntimeService(
  options: {
    subtitleTimingTracker: Parameters<typeof handleMineSentenceDigitService>[1]["subtitleTimingTracker"];
    ankiIntegration: Parameters<typeof handleMineSentenceDigitService>[1]["ankiIntegration"];
    getCurrentSecondarySubText: Parameters<typeof handleMineSentenceDigitService>[1]["getCurrentSecondarySubText"];
    showMpvOsd: Parameters<typeof handleMineSentenceDigitService>[1]["showMpvOsd"];
    logError: Parameters<typeof handleMineSentenceDigitService>[1]["logError"];
  },
): Parameters<typeof handleMineSentenceDigitService>[1] {
  return {
    subtitleTimingTracker: options.subtitleTimingTracker,
    ankiIntegration: options.ankiIntegration,
    getCurrentSecondarySubText: options.getCurrentSecondarySubText,
    showMpvOsd: options.showMpvOsd,
    logError: options.logError,
  };
}
