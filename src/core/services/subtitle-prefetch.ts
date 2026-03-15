import type { SubtitleCue } from './subtitle-cue-parser';
import type { SubtitleData } from '../../types';

export interface SubtitlePrefetchServiceDeps {
  cues: SubtitleCue[];
  tokenizeSubtitle: (text: string) => Promise<SubtitleData | null>;
  preCacheTokenization: (text: string, data: SubtitleData) => void;
  isCacheFull: () => boolean;
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

  // Find the first cue whose start time is >= current position.
  // This includes cues that start exactly at the current time (they haven't
  // been displayed yet and should be prefetched).
  let startIndex = -1;
  for (let i = 0; i < cues.length; i += 1) {
    if (cues[i]!.startTime >= currentTimeSeconds) {
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

  async function tokenizeCueList(
    cuesToProcess: SubtitleCue[],
    runId: number,
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

      if (deps.isCacheFull()) {
        return;
      }

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

    // Phase 1: Priority window
    const priorityCues = computePriorityWindow(cues, currentTimeSeconds, windowSize);
    await tokenizeCueList(priorityCues, runId);

    if (stopped || runId !== currentRunId) {
      return;
    }

    // Phase 2: Background - remaining cues forward from current position
    const priorityTexts = new Set(priorityCues.map((c) => c.text));
    const remainingCues = cues.filter(
      (cue) => cue.startTime > currentTimeSeconds && !priorityTexts.has(cue.text),
    );
    await tokenizeCueList(remainingCues, runId);

    if (stopped || runId !== currentRunId) {
      return;
    }

    // Phase 3: Background - earlier cues (for rewind support)
    const earlierCues = cues.filter(
      (cue) => cue.startTime <= currentTimeSeconds && !priorityTexts.has(cue.text),
    );
    await tokenizeCueList(earlierCues, runId);
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
