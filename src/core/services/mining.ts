interface SubtitleTimingTrackerLike {
  getRecentBlocks: (count: number) => string[];
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
    deps.showMpvOsd("No subtitle history available");
    return;
  }

  const actualCount = blocks.length;
  deps.writeClipboardText(blocks.join("\n\n"));
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
    deps.showMpvOsd("Subtitle tracker not available");
    return;
  }
  const currentSubtitle = deps.subtitleTimingTracker.getCurrentSubtitle();
  if (!currentSubtitle) {
    deps.showMpvOsd("No current subtitle");
    return;
  }
  deps.writeClipboardText(currentSubtitle);
  deps.showMpvOsd("Copied subtitle");
}

function requireAnkiIntegration(
  ankiIntegration: AnkiIntegrationLike | null,
  showMpvOsd: (text: string) => void,
): AnkiIntegrationLike | null {
  if (!ankiIntegration) {
    showMpvOsd("AnkiConnect integration not enabled");
    return null;
  }
  return ankiIntegration;
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
    deps.showMpvOsd("MPV not connected");
    return false;
  }
  if (!mpvClient.currentSubText) {
    deps.showMpvOsd("No current subtitle");
    return false;
  }

  return await anki.createSentenceCard(
    mpvClient.currentSubText,
    mpvClient.currentSubStart,
    mpvClient.currentSubEnd,
    mpvClient.currentSecondarySubText || undefined,
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

  const blocks = deps.subtitleTimingTracker.getRecentBlocks(count);
  if (blocks.length === 0) {
    deps.showMpvOsd("No subtitle history available");
    return;
  }

  const timings: { startTime: number; endTime: number }[] = [];
  for (const block of blocks) {
    const timing = deps.subtitleTimingTracker.findTiming(block);
    if (timing) timings.push(timing);
  }

  if (timings.length === 0) {
    deps.showMpvOsd("Subtitle timing not found");
    return;
  }

  const rangeStart = Math.min(...timings.map((t) => t.startTime));
  const rangeEnd = Math.max(...timings.map((t) => t.endTime));
  const sentence = blocks.join(" ");
  const cardsToMine = 1;
  deps.ankiIntegration
    .createSentenceCard(
      sentence,
      rangeStart,
      rangeEnd,
      deps.getCurrentSecondarySubText(),
    )
    .then((created) => {
      if (created) {
        deps.onCardsMined?.(cardsToMine);
      }
    })
    .catch((err) => {
      deps.logError("mineSentenceMultiple failed:", err);
      deps.showMpvOsd(`Mine sentence failed: ${(err as Error).message}`);
    });
}
