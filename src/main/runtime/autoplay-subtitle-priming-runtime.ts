import type { SubtitleCue, SubtitleData } from '../../types';
import { selectAutoplayStartupCue } from './autoplay-subtitle-primer';
import { primeVisibleOverlaySubtitleFromMpv } from './current-subtitle-snapshot';
import { resolveSubtitleSourcePath } from './subtitle-prefetch-source';

const AUTOPLAY_SUBTITLE_PRIME_LOOKAHEAD_SECONDS = 2;
const VISIBLE_OVERLAY_SUBTITLE_REFRESH_AFTER_FIRST_PAINT_DELAY_MS = 100;

type AutoplaySubtitlePrimingMpvClient = {
  connected?: boolean;
  requestProperty: (name: string) => Promise<unknown>;
  currentVideoPath?: string;
  currentTimePos?: number;
  currentSecondarySubText?: string;
  setCurrentSecondarySubText?: (text: string) => void;
};

type AutoplaySubtitlePrimingPrefetchService = {
  pause: () => void;
  resume: () => void;
};

export interface AutoplaySubtitlePrimingRuntimeDeps {
  getCurrentMediaPath: () => string | null | undefined;
  getMpvClient: () => AutoplaySubtitlePrimingMpvClient | null;
  setCurrentSubText: (text: string) => void;
  getCurrentSubText: () => string;
  getCurrentSubtitleData: () => SubtitleData | null;
  getActiveParsedSubtitleCues: () => SubtitleCue[];
  setActiveParsedSubtitleMediaPath: (mediaPath: string | null) => void;
  subtitleProcessingController: {
    consumeCachedSubtitle: (text: string) => SubtitleData | null;
    // Both report whether processing is pending; see pausePrefetchUntilProcessed.
    onSubtitleChange: (text: string) => boolean;
    refreshCurrentSubtitle: (text: string) => boolean;
    notePlainSubtitleEmitted: (text: string) => void;
  };
  emitSubtitlePayload: (payload: SubtitleData, options?: { resumePrefetch?: boolean }) => void;
  getSubtitlePrefetchService: () => AutoplaySubtitlePrimingPrefetchService | null;
  getLastObservedTimePos: () => number;
  getVisibleOverlayVisible: () => boolean;
  emitSecondarySubtitle: (text: string) => void;
  initSubtitlePrefetch: (
    sourcePath: string,
    currentTimePos: number,
    sourceKey?: string,
  ) => Promise<void>;
  refreshSubtitlePrefetchFromActiveTrack: () => Promise<void>;
  logDebug: (message: string) => void;
}

export function setMpvCurrentSecondarySubText(
  client: Pick<
    AutoplaySubtitlePrimingMpvClient,
    'currentSecondarySubText' | 'setCurrentSecondarySubText'
  >,
  text: string,
): void {
  if (typeof client.setCurrentSecondarySubText === 'function') {
    client.setCurrentSecondarySubText(text);
    return;
  }
  client.currentSecondarySubText = text;
}

export function createAutoplaySubtitlePrimingRuntime(deps: AutoplaySubtitlePrimingRuntimeDeps) {
  const { subtitleProcessingController, emitSubtitlePayload } = deps;

  // Prefetching is paused so the on-screen line gets the parser to itself; the
  // resume rides on the controller settling (see onProcessingSettled), not on
  // an emit, which a suppressed duplicate or a failed tokenization never sends.
  // When the controller reports it has nothing scheduled, no settle is coming
  // either, so release the pause here or prefetching idles indefinitely.
  function pausePrefetchUntilProcessed(scheduleTokenization: () => boolean): void {
    const prefetch = deps.getSubtitlePrefetchService();
    prefetch?.pause();
    if (!scheduleTokenization()) {
      prefetch?.resume();
    }
  }

  let subtitlePrefetchRefreshTimer: ReturnType<typeof setTimeout> | null = null;
  let autoplaySubtitlePrimedMediaPath: string | null = null;
  let visibleOverlaySubtitleRefreshAfterFirstPaintTimer: ReturnType<typeof setTimeout> | null =
    null;

  function getCurrentAutoplayMediaPath(): string | null {
    return (
      deps.getCurrentMediaPath()?.trim() || deps.getMpvClient()?.currentVideoPath?.trim() || null
    );
  }

  function isCurrentAutoplayMediaPath(mediaPath: string): boolean {
    return getCurrentAutoplayMediaPath() === mediaPath;
  }

  function markAutoplaySubtitlePrimeConsumed(mediaPath: string): boolean {
    if (autoplaySubtitlePrimedMediaPath === mediaPath) {
      return false;
    }
    autoplaySubtitlePrimedMediaPath = mediaPath;
    return true;
  }

  function resetAutoplaySubtitlePrime(): void {
    autoplaySubtitlePrimedMediaPath = null;
  }

  function emitAutoplayPrimedSubtitle(mediaPath: string, text: string): boolean {
    if (!text.trim() || !isCurrentAutoplayMediaPath(mediaPath)) {
      return false;
    }
    if (!markAutoplaySubtitlePrimeConsumed(mediaPath)) {
      return false;
    }

    deps.setCurrentSubText(text);
    deps.getSubtitlePrefetchService()?.pause();
    const cachedPayload = subtitleProcessingController.consumeCachedSubtitle(text);
    if (cachedPayload) {
      subtitleProcessingController.onSubtitleChange(text);
      // This emit resumes prefetching, so no pause is left outstanding.
      emitSubtitlePayload(cachedPayload);
      return true;
    }

    // Provisional raw emit: keep prefetch paused until the processing
    // controller is done with this line, and tell it this line has already been
    // painted plain so it does not broadcast the same payload again.
    emitSubtitlePayload({ text, tokens: null }, { resumePrefetch: false });
    subtitleProcessingController.notePlainSubtitleEmitted(text);
    // refreshCurrentSubtitle, not onSubtitleChange: the cache miss above can be
    // an invalidation (mining a card) on text the controller still holds, and
    // onSubtitleChange treats unchanged text as nothing to do, which would
    // leave this line permanently unannotated. refreshCurrentSubtitle also
    // re-tokenizes for a new cache generation.
    if (!subtitleProcessingController.refreshCurrentSubtitle(text)) {
      // Nothing scheduled, so no settle is coming to release the pause.
      deps.getSubtitlePrefetchService()?.resume();
    }
    return true;
  }

  async function primeCurrentSubtitleForAutoplay(mediaPath: string): Promise<void> {
    const client = deps.getMpvClient();
    if (!client?.connected || !isCurrentAutoplayMediaPath(mediaPath)) {
      return;
    }

    const subTextRaw = await client.requestProperty('sub-text').catch((error) => {
      deps.logDebug(
        `[autoplay-subtitle-prime] failed to read sub-text: ${
          error instanceof Error ? error.message : String(error)
        }`,
      );
      return null;
    });
    const text = typeof subTextRaw === 'string' ? subTextRaw : '';
    if (emitAutoplayPrimedSubtitle(mediaPath, text)) {
      return;
    }

    if (!text.trim() && isCurrentAutoplayMediaPath(mediaPath)) {
      await deps.refreshSubtitlePrefetchFromActiveTrack().catch((error) => {
        deps.logDebug(
          `[autoplay-subtitle-prime] active subtitle refresh failed after empty sub-text: ${
            error instanceof Error ? error.message : String(error)
          }`,
        );
      });
      await primeAutoplaySubtitleFromParsedCues(mediaPath, deps.getActiveParsedSubtitleCues());
    }
  }

  async function primeCurrentSubtitleForVisibleOverlay(): Promise<void> {
    await primeVisibleOverlaySubtitleFromMpv({
      getMpvClient: () => deps.getMpvClient(),
      setCurrentSubText: (text) => {
        deps.setCurrentSubText(text);
      },
      getCurrentSubtitleData: () => deps.getCurrentSubtitleData(),
      consumeCachedSubtitle: (text) => subtitleProcessingController.consumeCachedSubtitle(text),
      onSubtitleChange: (text) => {
        pausePrefetchUntilProcessed(() => subtitleProcessingController.onSubtitleChange(text));
      },
      refreshCurrentSubtitle: (text) => {
        pausePrefetchUntilProcessed(() =>
          subtitleProcessingController.refreshCurrentSubtitle(text),
        );
      },
      deferUncachedRefresh: true,
      emitSubtitle: (payload) => emitSubtitlePayload(payload),
      setCurrentSecondarySubText: (text) => {
        const client = deps.getMpvClient();
        if (client) {
          setMpvCurrentSecondarySubText(client, text);
        }
      },
      emitSecondarySubtitle: (text) => {
        deps.emitSecondarySubtitle(text);
      },
      logDebug: (message) => {
        deps.logDebug(message);
      },
    });
  }

  function cancelVisibleOverlaySubtitleRefreshAfterFirstPaint(): void {
    if (!visibleOverlaySubtitleRefreshAfterFirstPaintTimer) {
      return;
    }
    clearTimeout(visibleOverlaySubtitleRefreshAfterFirstPaintTimer);
    visibleOverlaySubtitleRefreshAfterFirstPaintTimer = null;
  }

  function scheduleVisibleOverlaySubtitleRefreshAfterFirstPaint(): void {
    if (visibleOverlaySubtitleRefreshAfterFirstPaintTimer) {
      return;
    }
    if (!deps.getVisibleOverlayVisible() || !deps.getCurrentSubText().trim()) {
      return;
    }

    visibleOverlaySubtitleRefreshAfterFirstPaintTimer = setTimeout(() => {
      visibleOverlaySubtitleRefreshAfterFirstPaintTimer = null;
      if (!deps.getVisibleOverlayVisible()) {
        return;
      }
      const text = deps.getCurrentSubText();
      if (!text.trim()) {
        return;
      }
      pausePrefetchUntilProcessed(() => subtitleProcessingController.refreshCurrentSubtitle(text));
    }, VISIBLE_OVERLAY_SUBTITLE_REFRESH_AFTER_FIRST_PAINT_DELAY_MS);
    visibleOverlaySubtitleRefreshAfterFirstPaintTimer.unref?.();
  }

  async function primeAutoplaySubtitleFromParsedCues(
    mediaPath: string,
    cues: SubtitleCue[],
  ): Promise<void> {
    if (
      cues.length === 0 ||
      autoplaySubtitlePrimedMediaPath === mediaPath ||
      !isCurrentAutoplayMediaPath(mediaPath)
    ) {
      return;
    }

    const client = deps.getMpvClient();
    const timePosRaw = await client?.requestProperty('time-pos').catch(() => null);
    const currentTimeSeconds = Number(
      timePosRaw ?? client?.currentTimePos ?? deps.getLastObservedTimePos() ?? 0,
    );
    const cue = selectAutoplayStartupCue(
      cues,
      Number.isFinite(currentTimeSeconds) ? currentTimeSeconds : 0,
      AUTOPLAY_SUBTITLE_PRIME_LOOKAHEAD_SECONDS,
    );
    if (!cue) {
      return;
    }

    emitAutoplayPrimedSubtitle(mediaPath, cue.text);
  }

  function clearScheduledSubtitlePrefetchRefresh(): void {
    if (subtitlePrefetchRefreshTimer) {
      clearTimeout(subtitlePrefetchRefreshTimer);
      subtitlePrefetchRefreshTimer = null;
    }
  }

  async function refreshSubtitleSidebarFromSource(
    sourcePath: string,
    mediaPath?: string,
  ): Promise<void> {
    const normalizedSourcePath = resolveSubtitleSourcePath(sourcePath.trim());
    if (!normalizedSourcePath) {
      return;
    }
    const nextMediaPath = mediaPath?.trim() || getCurrentAutoplayMediaPath();
    await deps.initSubtitlePrefetch(
      normalizedSourcePath,
      deps.getLastObservedTimePos(),
      normalizedSourcePath,
    );
    deps.setActiveParsedSubtitleMediaPath(nextMediaPath);
  }

  function scheduleSubtitlePrefetchRefresh(delayMs = 0): void {
    clearScheduledSubtitlePrefetchRefresh();
    subtitlePrefetchRefreshTimer = setTimeout(() => {
      subtitlePrefetchRefreshTimer = null;
      void deps.refreshSubtitlePrefetchFromActiveTrack().catch((error) => {
        deps.logDebug(
          `[autoplay-subtitle-prime] subtitle prefetch refresh failed: ${
            error instanceof Error ? error.message : String(error)
          }`,
        );
      });
    }, delayMs);
  }

  return {
    getCurrentAutoplayMediaPath,
    resetAutoplaySubtitlePrime,
    primeCurrentSubtitleForAutoplay,
    primeCurrentSubtitleForVisibleOverlay,
    cancelVisibleOverlaySubtitleRefreshAfterFirstPaint,
    scheduleVisibleOverlaySubtitleRefreshAfterFirstPaint,
    primeAutoplaySubtitleFromParsedCues,
    clearScheduledSubtitlePrefetchRefresh,
    refreshSubtitleSidebarFromSource,
    scheduleSubtitlePrefetchRefresh,
  };
}
