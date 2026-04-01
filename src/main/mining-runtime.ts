import type { SubtitleTimingTracker } from '../subtitle-timing-tracker';
import { appendClipboardVideoToQueueRuntime } from './runtime/clipboard-queue';
import {
  createUpdateLastCardFromClipboardHandler,
  createRefreshKnownWordCacheHandler,
  createTriggerFieldGroupingHandler,
  createMarkLastCardAsAudioCardHandler,
  createMineSentenceCardHandler,
} from './runtime/anki-actions';
import {
  createHandleMultiCopyDigitHandler,
  createCopyCurrentSubtitleHandler,
  createHandleMineSentenceDigitHandler,
} from './runtime/mining-actions';

export interface MiningRuntimeInput<TAnkiIntegration = unknown, TMpvClient = unknown> {
  getSubtitleTimingTracker: () => SubtitleTimingTracker;
  getAnkiIntegration: () => TAnkiIntegration;
  getMpvClient: () => TMpvClient;
  readClipboardText: () => string;
  writeClipboardText: (text: string) => void;
  showMpvOsd: (text: string) => void;
  sendMpvCommand: (command: (string | number)[]) => void;
  updateLastCardFromClipboardCore: (options: {
    ankiIntegration: TAnkiIntegration;
    readClipboardText: () => string;
    showMpvOsd: (text: string) => void;
  }) => Promise<void>;
  triggerFieldGroupingCore: (options: {
    ankiIntegration: TAnkiIntegration;
    showMpvOsd: (text: string) => void;
  }) => Promise<void>;
  markLastCardAsAudioCardCore: (options: {
    ankiIntegration: TAnkiIntegration;
    showMpvOsd: (text: string) => void;
  }) => Promise<void>;
  mineSentenceCardCore: (options: {
    ankiIntegration: TAnkiIntegration;
    mpvClient: TMpvClient;
    showMpvOsd: (text: string) => void;
  }) => Promise<boolean>;
  handleMultiCopyDigitCore: (
    count: number,
    options: {
      subtitleTimingTracker: SubtitleTimingTracker;
      writeClipboardText: (text: string) => void;
      showMpvOsd: (text: string) => void;
    },
  ) => void;
  copyCurrentSubtitleCore: (options: {
    subtitleTimingTracker: SubtitleTimingTracker;
    writeClipboardText: (text: string) => void;
    showMpvOsd: (text: string) => void;
  }) => void;
  handleMineSentenceDigitCore: (
    count: number,
    options: {
      subtitleTimingTracker: SubtitleTimingTracker;
      ankiIntegration: TAnkiIntegration;
      getCurrentSecondarySubText: () => string | undefined;
      showMpvOsd: (text: string) => void;
      logError: (message: string, err: unknown) => void;
      onCardsMined: (count: number) => void;
    },
  ) => void;
  getCurrentSecondarySubText: () => string | undefined;
  logError: (message: string, err: unknown) => void;
  recordCardsMined: (count: number, noteIds?: number[]) => void;
}

export interface MiningRuntime {
  updateLastCardFromClipboard: () => Promise<void>;
  refreshKnownWordCache: () => Promise<void>;
  triggerFieldGrouping: () => Promise<void>;
  markLastCardAsAudioCard: () => Promise<void>;
  mineSentenceCard: () => Promise<void>;
  handleMultiCopyDigit: (count: number) => void;
  copyCurrentSubtitle: () => void;
  handleMineSentenceDigit: (count: number) => void;
  appendClipboardVideoToQueue: () => { ok: boolean; message: string };
}

export function createMiningRuntime<TAnkiIntegration, TMpvClient>(
  input: MiningRuntimeInput<TAnkiIntegration, TMpvClient>,
): MiningRuntime {
  const updateLastCardFromClipboard = createUpdateLastCardFromClipboardHandler({
    getAnkiIntegration: () => input.getAnkiIntegration(),
    readClipboardText: () => input.readClipboardText(),
    showMpvOsd: (text) => input.showMpvOsd(text),
    updateLastCardFromClipboardCore: (options) => input.updateLastCardFromClipboardCore(options),
  });

  const refreshKnownWordCache = createRefreshKnownWordCacheHandler({
    getAnkiIntegration: () =>
      input.getAnkiIntegration() as { refreshKnownWordCache: () => Promise<void> } | null,
    missingIntegrationMessage: 'AnkiConnect integration not enabled',
  });

  const triggerFieldGrouping = createTriggerFieldGroupingHandler({
    getAnkiIntegration: () => input.getAnkiIntegration(),
    showMpvOsd: (text) => input.showMpvOsd(text),
    triggerFieldGroupingCore: (options) => input.triggerFieldGroupingCore(options),
  });

  const markLastCardAsAudioCard = createMarkLastCardAsAudioCardHandler({
    getAnkiIntegration: () => input.getAnkiIntegration(),
    showMpvOsd: (text) => input.showMpvOsd(text),
    markLastCardAsAudioCardCore: (options) => input.markLastCardAsAudioCardCore(options),
  });

  const mineSentenceCard = createMineSentenceCardHandler({
    getAnkiIntegration: () => input.getAnkiIntegration(),
    getMpvClient: () => input.getMpvClient(),
    showMpvOsd: (text) => input.showMpvOsd(text),
    mineSentenceCardCore: (options) => input.mineSentenceCardCore(options),
    recordCardsMined: (count, noteIds) => input.recordCardsMined(count, noteIds),
  });

  const handleMultiCopyDigit = createHandleMultiCopyDigitHandler({
    getSubtitleTimingTracker: () => input.getSubtitleTimingTracker(),
    writeClipboardText: (text) => input.writeClipboardText(text),
    showMpvOsd: (text) => input.showMpvOsd(text),
    handleMultiCopyDigitCore: (count, options) => input.handleMultiCopyDigitCore(count, options),
  });

  const copyCurrentSubtitle = createCopyCurrentSubtitleHandler({
    getSubtitleTimingTracker: () => input.getSubtitleTimingTracker(),
    writeClipboardText: (text) => input.writeClipboardText(text),
    showMpvOsd: (text) => input.showMpvOsd(text),
    copyCurrentSubtitleCore: (options) => input.copyCurrentSubtitleCore(options),
  });

  const handleMineSentenceDigit = createHandleMineSentenceDigitHandler({
    getSubtitleTimingTracker: () => input.getSubtitleTimingTracker(),
    getAnkiIntegration: () => input.getAnkiIntegration(),
    getCurrentSecondarySubText: () => input.getCurrentSecondarySubText(),
    showMpvOsd: (text) => input.showMpvOsd(text),
    logError: (message, err) => input.logError(message, err),
    onCardsMined: (count) => input.recordCardsMined(count),
    handleMineSentenceDigitCore: (count, options) =>
      input.handleMineSentenceDigitCore(count, options),
  });

  const appendClipboardVideoToQueue = (): { ok: boolean; message: string } =>
    appendClipboardVideoToQueueRuntime({
      getMpvClient: () => input.getMpvClient() as { connected: boolean } | null,
      readClipboardText: () => input.readClipboardText(),
      showMpvOsd: (text) => input.showMpvOsd(text),
      sendMpvCommand: (command) => input.sendMpvCommand(command),
    });

  return {
    updateLastCardFromClipboard,
    refreshKnownWordCache,
    triggerFieldGrouping,
    markLastCardAsAudioCard,
    mineSentenceCard,
    handleMultiCopyDigit,
    copyCurrentSubtitle,
    handleMineSentenceDigit,
    appendClipboardVideoToQueue,
  };
}
