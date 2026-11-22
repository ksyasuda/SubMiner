import type { AniSkipMetadata } from './aniskip-metadata';
import {
  ANISKIP_CHAPTER_PREFIX,
  classifyChapterSkipIntervals,
  type SkipInterval,
} from './aniskip-chapters';

/** mpv script-message that skips whichever of the intro or outro is playing. */
export const ANISKIP_SKIP_INTRO_MESSAGE = 'subminer-skip-intro';
export const ANISKIP_REFRESH_MESSAGE = 'subminer-aniskip-refresh';

const DEFAULT_ANISKIP_BUTTON_KEY = 'TAB';
const LEGACY_ANISKIP_BUTTON_KEY = 'y-k';
const SKIP_WINDOW_EPSILON_SECONDS = 0.35;
const PROMPT_WINDOW_SECONDS = 3;
const PROMPT_OSD_DURATION_MS = 3000;
export interface AniSkipRuntimeConfig {
  aniskipEnabled: boolean;
  aniskipButtonKey: string;
}

export interface AniSkipRuntimeDeps {
  getAniSkipConfig: () => AniSkipRuntimeConfig;
  resolveMetadataForFile: (mediaPath: string, episodeLength: number) => Promise<AniSkipMetadata>;
  sendMpvCommand: (command: unknown[]) => void;
  requestMpvProperty: (name: string) => Promise<unknown>;
  isMpvConnected: () => boolean;
  getCurrentTimePos: () => number;
  showMpvOsd: (text: string, durationMs: number) => void;
  showPlaybackFeedback?: (text: string) => void;
  logInfo: (message: string) => void;
  logWarn: (message: string, error?: unknown) => void;
  logDebug: (message: string) => void;
}

export type SkipWindowKind = 'intro' | 'outro';

export interface AniSkipWindow extends SkipInterval {
  /** `chapters` = the file's own chapters; `aniskip` = the AniSkip database. */
  source: 'chapters' | 'aniskip';
}

export interface AniSkipWindows {
  intro: AniSkipWindow | null;
  outro: AniSkipWindow | null;
}

interface AniSkipLookupResult {
  intro: SkipInterval | null;
  outro: SkipInterval | null;
  malId: number | null;
}

type MpvChapter = { time?: unknown; title?: unknown };

const KIND_LABEL: Record<SkipWindowKind, { chapter: string; word: string }> = {
  intro: { chapter: 'Intro', word: 'intro' },
  outro: { chapter: 'Outro', word: 'outro' },
};

export function isRemoteMediaPath(mediaPath: string): boolean {
  return /^[a-zA-Z][\w+.-]*:\/\//.test(mediaPath.trim());
}

function isAniSkipChapter(chapter: MpvChapter): boolean {
  return typeof chapter?.title === 'string' && chapter.title.startsWith(ANISKIP_CHAPTER_PREFIX);
}

function isUsableDuration(value: unknown): value is number {
  return typeof value === 'number' && Number.isFinite(value) && value > 0;
}

/**
 * Finds the intro and outro for each loaded file (file chapters first, then
 * AniSkip for whatever the chapters lack), marks them as chapters, prompts when
 * one starts, and skips the active one when the skip key is pressed.
 */
export function createAniSkipRuntime(deps: AniSkipRuntimeDeps) {
  let requestGeneration = 0;
  let currentMediaPath = '';
  let currentDuration: number | null = null;
  // True once resolution for the current media has started (or finished).
  let resolutionStarted = false;
  // True when resolution for the current media hit a retryable failure (the
  // chapter-list read or the AniSkip lookup), so a reload or duration change
  // runs it again.
  let retryResolution = false;
  let windows: AniSkipWindows = { intro: null, outro: null };
  const promptShown: Record<SkipWindowKind, boolean> = { intro: false, outro: false };
  let boundButtonKey: string | null = null;
  let legacyFallbackBound = false;
  // AniSkip results per media path. Only definitive answers are stored.
  const lookupCache = new Map<string, AniSkipLookupResult>();

  function resolveButtonKey(): string {
    const key = deps.getAniSkipConfig().aniskipButtonKey.trim();
    return key || DEFAULT_ANISKIP_BUTTON_KEY;
  }

  function showPlaybackFeedback(text: string, durationMs = PROMPT_OSD_DURATION_MS): void {
    if (deps.showPlaybackFeedback) {
      deps.showPlaybackFeedback(text);
      return;
    }
    deps.showMpvOsd(text, durationMs);
  }

  function bindSkipKeys(): void {
    if (!deps.isMpvConnected()) return;
    const enabled = deps.getAniSkipConfig().aniskipEnabled;
    const key = resolveButtonKey();
    const wantLegacyFallback =
      enabled && key !== LEGACY_ANISKIP_BUTTON_KEY && key !== DEFAULT_ANISKIP_BUTTON_KEY;

    if (boundButtonKey && (!enabled || boundButtonKey !== key)) {
      deps.sendMpvCommand(['keybind', boundButtonKey, '']);
      boundButtonKey = null;
    }
    if (legacyFallbackBound && !wantLegacyFallback) {
      deps.sendMpvCommand(['keybind', LEGACY_ANISKIP_BUTTON_KEY, '']);
      legacyFallbackBound = false;
    }
    if (!enabled) return;

    if (boundButtonKey !== key) {
      deps.sendMpvCommand(['keybind', key, `script-message ${ANISKIP_SKIP_INTRO_MESSAGE}`]);
      boundButtonKey = key;
    }
    if (wantLegacyFallback && !legacyFallbackBound) {
      deps.sendMpvCommand([
        'keybind',
        LEGACY_ANISKIP_BUTTON_KEY,
        `script-message ${ANISKIP_SKIP_INTRO_MESSAGE}`,
      ]);
      legacyFallbackBound = true;
    }
  }

  async function readChapterList(): Promise<MpvChapter[] | null> {
    try {
      const chapterList = await deps.requestMpvProperty('chapter-list');
      return Array.isArray(chapterList) ? (chapterList as MpvChapter[]) : [];
    } catch {
      // chapter-list may be unavailable mid-load
      return null;
    }
  }

  /** True while the given request generation still owns the given media. */
  function ownsPlayback(generation: number, mediaPath: string): boolean {
    return generation === requestGeneration && mediaPath === currentMediaPath;
  }

  /** Replaces SubMiner's marker chapters with ones for the AniSkip-sourced windows. */
  async function setSkipChapters(): Promise<void> {
    const generation = requestGeneration;
    const mediaPath = currentMediaPath;
    const markers: MpvChapter[] = [];
    for (const kind of ['intro', 'outro'] as const) {
      const window = windows[kind];
      if (window?.source !== 'aniskip') continue;
      const label = KIND_LABEL[kind].chapter;
      markers.push({ time: window.start, title: `${ANISKIP_CHAPTER_PREFIX}${label} Start` });
      markers.push({ time: window.end, title: `${ANISKIP_CHAPTER_PREFIX}${label} End` });
    }
    if (markers.length === 0) return;
    const existing = await readChapterList();
    // A failed read must not wipe the file's chapters, and a read that finished
    // after the media changed must not write into the next file.
    if (!existing || !ownsPlayback(generation, mediaPath)) return;
    const chapters = existing.filter((chapter) => !isAniSkipChapter(chapter)).concat(markers);
    chapters.sort((a, b) => {
      const aTime = typeof a.time === 'number' ? a.time : 0;
      const bTime = typeof b.time === 'number' ? b.time : 0;
      return aTime - bTime;
    });
    deps.sendMpvCommand(['set_property', 'chapter-list', chapters]);
  }

  async function removeSkipChapters(): Promise<void> {
    if (!deps.isMpvConnected()) return;
    const generation = requestGeneration;
    const mediaPath = currentMediaPath;
    const existing = await readChapterList();
    if (!existing || !ownsPlayback(generation, mediaPath)) return;
    const filtered = existing.filter((chapter) => !isAniSkipChapter(chapter));
    if (filtered.length !== existing.length) {
      deps.sendMpvCommand(['set_property', 'chapter-list', filtered]);
    }
  }

  function clearState(): void {
    requestGeneration += 1;
    resolutionStarted = false;
    retryResolution = false;
    windows = { intro: null, outro: null };
    promptShown.intro = false;
    promptShown.outro = false;
  }

  function logWindow(kind: SkipWindowKind, window: AniSkipWindow, malId: number | null): void {
    deps.logInfo(
      `AniSkip ${kind} window ${window.start.toFixed(3)} -> ${window.end.toFixed(3)} (source=${
        window.source
      }, MAL ${malId ?? '-'})`,
    );
  }

  /** Installs the file-chapter windows, resetting the prompts. */
  function applyChapterWindows(next: AniSkipWindows): void {
    windows = next;
    promptShown.intro = false;
    promptShown.outro = false;
    for (const kind of ['intro', 'outro'] as const) {
      const window = next[kind];
      if (window) logWindow(kind, window, null);
    }
  }

  /**
   * Fills the windows the chapters lacked from AniSkip, leaving the chapter
   * windows (and whether their prompt was shown) untouched.
   */
  async function mergeAniSkipWindows(lookup: AniSkipLookupResult): Promise<void> {
    const added: SkipWindowKind[] = [];
    for (const kind of ['intro', 'outro'] as const) {
      const interval = lookup[kind];
      if (windows[kind] || !interval) continue;
      windows = { ...windows, [kind]: { ...interval, source: 'aniskip' } };
      promptShown[kind] = false;
      added.push(kind);
    }
    if (added.length === 0) return;
    for (const kind of added) {
      const window = windows[kind];
      if (window) logWindow(kind, window, lookup.malId);
    }
    await setSkipChapters();
  }

  /** Looks up AniSkip, using the cache for definitive answers. Null means a retryable failure. */
  async function lookUpAniSkip(
    mediaPath: string,
    duration: number,
    generation: number,
  ): Promise<AniSkipLookupResult | null> {
    const cached = lookupCache.get(mediaPath);
    if (cached) return cached;

    let metadata: AniSkipMetadata;
    try {
      metadata = await deps.resolveMetadataForFile(mediaPath, duration);
    } catch (error) {
      deps.logWarn('AniSkip metadata lookup failed', error);
      return null;
    }
    const result: AniSkipLookupResult = {
      intro: metadata.intro,
      outro: metadata.outro,
      malId: metadata.malId,
    };
    if (metadata.lookupStatus === 'lookup_failed') {
      deps.logInfo(`AniSkip lookup failed for "${metadata.title}"; will retry on next load`);
      return null;
    }
    // Every other status is definitive (a 404 or an unmatched title), so cache it,
    // unless a refresh or media change superseded this lookup while it ran.
    if (generation === requestGeneration) lookupCache.set(mediaPath, result);
    if (metadata.lookupStatus !== 'ready') {
      deps.logInfo(
        `AniSkip: no skip times for "${metadata.title}" (status=${metadata.lookupStatus ?? 'unknown'})`,
      );
    }
    return result;
  }

  async function resolveForMedia(mediaPath: string, duration: number): Promise<void> {
    if (!deps.getAniSkipConfig().aniskipEnabled) return;
    const generation = requestGeneration;
    const isStale = () => generation !== requestGeneration || mediaPath !== currentMediaPath;

    const chapterList = await readChapterList();
    if (isStale()) return;
    // A failed read is retried later; a successful empty list is a final answer.
    if (chapterList === null) retryResolution = true;
    const fromChapters = classifyChapterSkipIntervals(chapterList ?? [], duration);
    const next: AniSkipWindows = {
      intro: fromChapters.intro ? { ...fromChapters.intro, source: 'chapters' } : null,
      outro: fromChapters.outro ? { ...fromChapters.outro, source: 'chapters' } : null,
    };
    // Chapter windows work right away; AniSkip only fills what they lack.
    applyChapterWindows(next);
    if (next.intro && next.outro) return;
    if (isRemoteMediaPath(mediaPath)) {
      deps.logDebug('AniSkip lookup skipped for remote media; using file chapters only');
      return;
    }

    const lookup = await lookUpAniSkip(mediaPath, duration, generation);
    if (isStale()) return;
    if (lookup === null) retryResolution = true;
    if (lookup) await mergeAniSkipWindows(lookup);
  }

  /** Starts resolution once the current media has both a path and a known duration. */
  function maybeStartResolution(): void {
    if (resolutionStarted || !currentMediaPath || currentDuration === null) return;
    if (!deps.getAniSkipConfig().aniskipEnabled) return;
    resolutionStarted = true;
    void resolveForMedia(currentMediaPath, currentDuration).catch((error) => {
      deps.logWarn('AniSkip media resolution failed', error);
    });
  }

  function requestCurrentDuration(): void {
    const generation = requestGeneration;
    void deps
      .requestMpvProperty('duration')
      .then((value) => {
        if (generation !== requestGeneration || !isUsableDuration(value)) return;
        currentDuration ??= value;
        maybeStartResolution();
      })
      .catch(() => {
        // duration is unavailable until the file is loaded; duration-change will follow
      });
  }

  function getActiveWindow(now: number): SkipWindowKind | null {
    for (const kind of ['intro', 'outro'] as const) {
      const window = windows[kind];
      if (
        window &&
        now >= window.start - SKIP_WINDOW_EPSILON_SECONDS &&
        now <= window.end + SKIP_WINDOW_EPSILON_SECONDS
      ) {
        return kind;
      }
    }
    return null;
  }

  function skipNow(): void {
    if (!deps.getAniSkipConfig().aniskipEnabled) return;
    if (!windows.intro && !windows.outro) {
      showPlaybackFeedback('Intro/outro skip unavailable');
      return;
    }
    const now = deps.getCurrentTimePos();
    if (!Number.isFinite(now)) {
      showPlaybackFeedback('Skip unavailable');
      return;
    }
    const kind = getActiveWindow(now);
    const window = kind ? windows[kind] : null;
    if (!kind || !window) {
      showPlaybackFeedback('Skip only during intro or outro');
      return;
    }
    deps.sendMpvCommand(['set_property', 'time-pos', window.end]);
    showPlaybackFeedback(`Skipped ${KIND_LABEL[kind].word}`);
  }

  function handleTimePosChange({ time }: { time: number }): void {
    if (!deps.getAniSkipConfig().aniskipEnabled) return;
    for (const kind of ['intro', 'outro'] as const) {
      const window = windows[kind];
      if (!window || promptShown[kind]) continue;
      const promptWindowEnd = Math.min(window.start + PROMPT_WINDOW_SECONDS, window.end);
      if (time >= window.start && time < promptWindowEnd) {
        promptShown[kind] = true;
        showPlaybackFeedback(
          `You can skip the ${KIND_LABEL[kind].word} by pressing ${resolveButtonKey()}`,
        );
      }
    }
  }

  function handleDurationChange({ duration }: { duration: number }): void {
    if (!isUsableDuration(duration) || !currentMediaPath) return;
    currentDuration = duration;
    if (retryResolution) clearState();
    maybeStartResolution();
  }

  function handleMediaPathChange({ path }: { path: string }): void {
    const nextPath = typeof path === 'string' ? path : '';
    if (nextPath === currentMediaPath && resolutionStarted) {
      if (retryResolution) {
        // The last resolution failed transiently; a reload is the retry.
        clearState();
        maybeStartResolution();
        return;
      }
      // Same-media reload: mpv rebuilt the chapter list, so re-apply markers.
      void setSkipChapters().catch(() => {});
      return;
    }
    currentMediaPath = nextPath;
    currentDuration = null;
    clearState();
    if (!nextPath || !deps.getAniSkipConfig().aniskipEnabled) return;
    requestCurrentDuration();
  }

  function handleConnectionChange({ connected }: { connected: boolean }): void {
    if (!connected) {
      boundButtonKey = null;
      legacyFallbackBound = false;
      clearState();
      currentMediaPath = '';
      currentDuration = null;
      return;
    }
    bindSkipKeys();
  }

  function handleClientMessage({ args }: { args: string[] }): void {
    const messageName = args[0];
    if (messageName === ANISKIP_SKIP_INTRO_MESSAGE) {
      skipNow();
      return;
    }
    if (messageName === ANISKIP_REFRESH_MESSAGE) {
      if (!currentMediaPath) return;
      lookupCache.delete(currentMediaPath);
      clearState();
      void removeSkipChapters().catch(() => {});
      if (currentDuration === null) {
        requestCurrentDuration();
        return;
      }
      maybeStartResolution();
    }
  }

  function applyConfigChange(): void {
    bindSkipKeys();
    if (!deps.getAniSkipConfig().aniskipEnabled) {
      clearState();
      void removeSkipChapters().catch(() => {});
      return;
    }
    if (!currentMediaPath || resolutionStarted) return;
    if (currentDuration === null) {
      requestCurrentDuration();
      return;
    }
    maybeStartResolution();
  }

  return {
    handleConnectionChange,
    handleMediaPathChange,
    handleDurationChange,
    handleTimePosChange,
    handleClientMessage,
    applyConfigChange,
    skipNow,
    getSkipWindows: (): AniSkipWindows => windows,
  };
}

export type AniSkipRuntime = ReturnType<typeof createAniSkipRuntime>;
