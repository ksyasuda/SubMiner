import type { SubtitleTimingBlock } from '../../subtitle-timing-tracker';
import { i18n } from '../../i18n/index.js';

interface SubtitleTimingTrackerLike {
  getRecentBlocks: (count: number) => string[];
  getRecentEntries?: (count: number) => SubtitleTimingBlock[];
  getCurrentSubtitle: () => string | null;
  findTiming: (text: string) => { startTime: number; endTime: number } | null;
}

interface AnkiIntegrationLike {
  updateLastAddedFromClipboard: (clipboardText: string) => Promise<void>;
  triggerFieldGroupingForLastAddedCard: () => Promise<void>;
  markLastCardAsAudioCard: () => Promise<void>;
  createSentenceCard: (
    sentence: string,
    startTime: number,
    endTime: number,
    secondarySub?: string,
  ) => Promise<boolean>;
}

interface MpvClientLike {
  connected: boolean;
  currentSubText: string;
  currentSubStart: number;
  currentSubEnd: number;
  currentSecondarySubText?: string;
  requestProperty?: (name: string) => Promise<unknown>;
}

export function handleMultiCopyDigit(
  count: number,
  deps: {
    subtitleTimingTracker: SubtitleTimingTrackerLike | null;
    writeClipboardText: (text: string) => void;
    showMpvOsd: (text: string) => void;
  },
): void {
  if (!deps.subtitleTimingTracker) return;

  const availableCount = Math.min(count, 200);
  const blocks = deps.subtitleTimingTracker.getRecentBlocks(availableCount);
  if (blocks.length === 0) {
    deps.showMpvOsd(i18n.t('osd.noSubtitleHistory'));
    return;
  }

  const actualCount = blocks.length;
  deps.writeClipboardText(blocks.join('\n\n'));
  if (actualCount < count) {
    deps.showMpvOsd(`Only ${actualCount} lines available, copied ${actualCount}`);
  } else {
    deps.showMpvOsd(`Copied ${actualCount} lines`);
  }
}

export function copyCurrentSubtitle(deps: {
  subtitleTimingTracker: SubtitleTimingTrackerLike | null;
  writeClipboardText: (text: string) => void;
  showMpvOsd: (text: string) => void;
}): void {
  if (!deps.subtitleTimingTracker) {
    deps.showMpvOsd(i18n.t('osd.noTracker'));
    return;
  }
  const currentSubtitle = deps.subtitleTimingTracker.getCurrentSubtitle();
  if (!currentSubtitle) {
    deps.showMpvOsd(i18n.t('osd.noCurrentSubtitle'));
    return;
  }
  deps.writeClipboardText(currentSubtitle);
  deps.showMpvOsd(i18n.t('osd.copiedSubtitle'));
}

function requireAnkiIntegration(
  ankiIntegration: AnkiIntegrationLike | null,
  showMpvOsd: (text: string) => void,
): AnkiIntegrationLike | null {
  if (!ankiIntegration) {
    showMpvOsd(i18n.t('osd.ankiNotEnabled'));
    return null;
  }
  return ankiIntegration;
}

function getSecondarySubTextForMinedBlocks(
  entries: SubtitleTimingBlock[] | undefined,
  getCurrentSecondarySubText: () => string | undefined,
): string | undefined {
  const secondaryBlocks = entries
    ?.map((entry) => entry.secondaryText?.trim())
    .filter((text): text is string => Boolean(text));
  if (secondaryBlocks && secondaryBlocks.length > 0) {
    return secondaryBlocks.join(' ');
  }
  return getCurrentSecondarySubText();
}

function normalizeSecondarySubText(text: unknown, primaryText: string): string | undefined {
  if (typeof text !== 'string') {
    return undefined;
  }
  const trimmed = text.trim();
  if (!trimmed || trimmed === primaryText.trim()) {
    return undefined;
  }
  return trimmed;
}

async function getCurrentSecondarySubTextForSentenceCard(
  mpvClient: MpvClientLike,
): Promise<string | undefined> {
  const primaryText = mpvClient.currentSubText;
  if (mpvClient.requestProperty) {
    try {
      const latestSecondaryText = await mpvClient.requestProperty('secondary-sub-text');
      return normalizeSecondarySubText(latestSecondaryText, primaryText);
    } catch {
      // Fall back to the cached secondary subtitle below.
    }
  }
  return normalizeSecondarySubText(mpvClient.currentSecondarySubText, primaryText);
}

export async function updateLastCardFromClipboard(deps: {
  ankiIntegration: AnkiIntegrationLike | null;
  readClipboardText: () => string;
  showMpvOsd: (text: string) => void;
}): Promise<void> {
  const anki = requireAnkiIntegration(deps.ankiIntegration, deps.showMpvOsd);
  if (!anki) return;
  await anki.updateLastAddedFromClipboard(deps.readClipboardText());
}

export async function triggerFieldGrouping(deps: {
  ankiIntegration: AnkiIntegrationLike | null;
  showMpvOsd: (text: string) => void;
}): Promise<void> {
  const anki = requireAnkiIntegration(deps.ankiIntegration, deps.showMpvOsd);
  if (!anki) return;
  await anki.triggerFieldGroupingForLastAddedCard();
}

export async function markLastCardAsAudioCard(deps: {
  ankiIntegration: AnkiIntegrationLike | null;
  showMpvOsd: (text: string) => void;
}): Promise<void> {
  const anki = requireAnkiIntegration(deps.ankiIntegration, deps.showMpvOsd);
  if (!anki) return;
  await anki.markLastCardAsAudioCard();
}

export async function mineSentenceCard(deps: {
  ankiIntegration: AnkiIntegrationLike | null;
  mpvClient: MpvClientLike | null;
  showMpvOsd: (text: string) => void;
}): Promise<boolean> {
  const anki = requireAnkiIntegration(deps.ankiIntegration, deps.showMpvOsd);
  if (!anki) return false;

  const mpvClient = deps.mpvClient;
  if (!mpvClient || !mpvClient.connected) {
    deps.showMpvOsd('MPV not connected');
    return false;
  }
  if (!mpvClient.currentSubText) {
    deps.showMpvOsd(i18n.t('osd.noCurrentSubtitle'));
    return false;
  }

  const secondarySubText = await getCurrentSecondarySubTextForSentenceCard(mpvClient);
  return await anki.createSentenceCard(
    mpvClient.currentSubText,
    mpvClient.currentSubStart,
    mpvClient.currentSubEnd,
    secondarySubText,
  );
}

export function handleMineSentenceDigit(
  count: number,
  deps: {
    subtitleTimingTracker: SubtitleTimingTrackerLike | null;
    ankiIntegration: AnkiIntegrationLike | null;
    getCurrentSecondarySubText: () => string | undefined;
    showMpvOsd: (text: string) => void;
    logError: (message: string, err: unknown) => void;
    onCardsMined?: (count: number) => void;
  },
): void {
  if (!deps.subtitleTimingTracker || !deps.ankiIntegration) return;

  const entries = deps.subtitleTimingTracker.getRecentEntries?.(count);
  const blocks =
    entries?.map((entry) => entry.displayText) ?? deps.subtitleTimingTracker.getRecentBlocks(count);
  if (blocks.length === 0) {
    deps.showMpvOsd(i18n.t('osd.noSubtitleHistory'));
    return;
  }

  const timings: { startTime: number; endTime: number }[] =
    entries ??
    blocks.flatMap((block) => {
      const timing = deps.subtitleTimingTracker?.findTiming(block);
      return timing ? [timing] : [];
    });

  if (timings.length === 0) {
    deps.showMpvOsd(i18n.t('osd.subtitleTimingNotFound'));
    return;
  }

  const rangeStart = Math.min(...timings.map((t) => t.startTime));
  const rangeEnd = Math.max(...timings.map((t) => t.endTime));
  const sentence = blocks.join(' ');
  const secondarySubText = getSecondarySubTextForMinedBlocks(
    entries,
    deps.getCurrentSecondarySubText,
  );
  const cardsToMine = 1;
  deps.ankiIntegration
    .createSentenceCard(sentence, rangeStart, rangeEnd, secondarySubText)
    .then((created) => {
      if (created) {
        deps.onCardsMined?.(cardsToMine);
      }
    })
    .catch((err) => {
      deps.logError('mineSentenceMultiple failed:', err);
      deps.showMpvOsd(`Mine sentence failed: ${(err as Error).message}`);
    });
}
