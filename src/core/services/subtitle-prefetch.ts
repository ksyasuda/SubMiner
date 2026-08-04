import type { SubtitleData } from '../../types';
import type { SubtitleCue } from '../../types';
import { normalizeSubtitleCacheKey } from './subtitle-processing-controller';

export interface SubtitlePrefetchServiceDeps {
  cues: SubtitleCue[];
  tokenizeSubtitle: (text: string) => Promise<SubtitleData | null>;
  preCacheTokenization: (text: string, data: SubtitleData) => void;
  hasCachedTokenization?: (text: string) => boolean;
  priorityWindowSize?: number;
}

export interface SubtitlePrefetchService {
  start: (currentTimeSeconds: number) => void;
  stop: () => void;
  onSeek: (newTimeSeconds: number) => void;
  pause: () => void;
  resume: () => void;
}

const DEFAULT_PRIORITY_WINDOW_SIZE = 10;

export function computePriorityWindow(
  cues: SubtitleCue[],
  currentTimeSeconds: number,
  windowSize: number,
): SubtitleCue[] {
  if (cues.length === 0) {
    return [];
  }

  // Find the first cue whose end time is after the current position.
  // This includes the currently active cue when playback starts or seeks
  // mid-line, while still skipping cues that have already finished.
  let startIndex = -1;
  for (let i = 0; i < cues.length; i += 1) {
    if (cues[i]!.endTime > currentTimeSeconds) {
      startIndex = i;
      break;
    }
  }

  if (startIndex < 0) {
    // All cues are before current time
    return [];
  }

  return cues.slice(startIndex, startIndex + windowSize);
}

export function createSubtitlePrefetchService(
  deps: SubtitlePrefetchServiceDeps,
): SubtitlePrefetchService {
  const windowSize = deps.priorityWindowSize ?? DEFAULT_PRIORITY_WINDOW_SIZE;
  let stopped = true;
  let paused = false;
  let currentRunId = 0;

  // A run is a single bounded pass over one file's cues, deduped by `warmedKeys` and by
  // `hasCachedTokenization`, so the worst case is one tokenization per cue. The cache is
  // an LRU and bounds its own memory, so a full cache is not a reason to stop warming;
  // stopping there used to leave the tail of longer media permanently uncached.
  async function tokenizeCueList(
    cuesToProcess: SubtitleCue[],
    runId: number,
    warmedKeys: Set<string>,
  ): Promise<void> {
    for (const cue of cuesToProcess) {
      if (stopped || runId !== currentRunId) {
        return;
      }

      // Wait while paused
      while (paused && !stopped && runId === currentRunId) {
        await new Promise((resolve) => setTimeout(resolve, 10));
      }

      if (stopped || runId !== currentRunId) {
        return;
      }

      const cacheKey = normalizeSubtitleCacheKey(cue.text);
      if (!cacheKey || warmedKeys.has(cacheKey) || deps.hasCachedTokenization?.(cue.text)) {
        if (cacheKey) {
          warmedKeys.add(cacheKey);
        }
        continue;
      }
      warmedKeys.add(cacheKey);

      try {
        const result = await deps.tokenizeSubtitle(cue.text);
        if (result && !stopped && runId === currentRunId) {
          deps.preCacheTokenization(cue.text, result);
        }
      } catch {
        // Skip failed cues, continue prefetching
      }

      // Yield to allow live processing to take priority
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
  }

  async function startPrefetching(currentTimeSeconds: number, runId: number): Promise<void> {
    const cues = deps.cues;
    const warmedKeys = new Set<string>();

    // Phase 1: Priority window
    const priorityCues = computePriorityWindow(cues, currentTimeSeconds, windowSize);
    await tokenizeCueList(priorityCues, runId, warmedKeys);

    if (stopped || runId !== currentRunId) {
      return;
    }

    // Phase 2: Background - remaining cues forward from current position
    const priorityTexts = new Set(priorityCues.map((c) => c.text));
    const remainingCues = cues.filter(
      (cue) => cue.startTime > currentTimeSeconds && !priorityTexts.has(cue.text),
    );
    await tokenizeCueList(remainingCues, runId, warmedKeys);

    if (stopped || runId !== currentRunId) {
      return;
    }

    // Phase 3: Background - earlier cues (for rewind support)
    const earlierCues = cues.filter(
      (cue) => cue.startTime <= currentTimeSeconds && !priorityTexts.has(cue.text),
    );
    await tokenizeCueList(earlierCues, runId, warmedKeys);
  }

  return {
    start(currentTimeSeconds: number) {
      stopped = false;
      paused = false;
      currentRunId += 1;
      const runId = currentRunId;
      void startPrefetching(currentTimeSeconds, runId);
    },

    stop() {
      stopped = true;
      currentRunId += 1;
    },

    onSeek(newTimeSeconds: number) {
      // Cancel current run and restart from new position
      currentRunId += 1;
      const runId = currentRunId;
      void startPrefetching(newTimeSeconds, runId);
    },

    pause() {
      paused = true;
    },

    resume() {
      paused = false;
    },
  };
}
