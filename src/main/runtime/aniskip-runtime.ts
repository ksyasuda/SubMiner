import type { AniSkipMetadata } from './aniskip-metadata';

export const ANISKIP_SKIP_INTRO_MESSAGE = 'subminer-skip-intro';
export const ANISKIP_REFRESH_MESSAGE = 'subminer-aniskip-refresh';

const DEFAULT_ANISKIP_BUTTON_KEY = 'TAB';
const LEGACY_ANISKIP_BUTTON_KEY = 'y-k';
const ANISKIP_CHAPTER_PREFIX = 'AniSkip ';
const SKIP_WINDOW_EPSILON_SECONDS = 0.35;
const PROMPT_WINDOW_SECONDS = 3;
const PROMPT_OSD_DURATION_MS = 3000;
export interface AniSkipRuntimeConfig {
  aniskipEnabled: boolean;
  aniskipButtonKey: string;
}

export interface AniSkipRuntimeDeps {
  getAniSkipConfig: () => AniSkipRuntimeConfig;
  resolveMetadataForFile: (mediaPath: string) => Promise<AniSkipMetadata>;
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

interface AniSkipIntroWindow {
  start: number;
  end: number;
  malId: number | null;
}

type MpvChapter = { time?: unknown; title?: unknown };

export function isRemoteMediaPath(mediaPath: string): boolean {
  return /^[a-zA-Z][\w+.-]*:\/\//.test(mediaPath.trim());
}

export function createAniSkipRuntime(deps: AniSkipRuntimeDeps) {
  let requestGeneration = 0;
  let currentMediaPath = '';
  let introWindow: AniSkipIntroWindow | null = null;
  let promptShown = false;
  let boundButtonKey: string | null = null;
  let legacyFallbackBound = false;
  const introWindowCache = new Map<string, AniSkipIntroWindow | null>();

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

  async function setIntroChapters(introStart: number, introEnd: number): Promise<void> {
    let existing: MpvChapter[] = [];
    try {
      const chapterList = await deps.requestMpvProperty('chapter-list');
      if (Array.isArray(chapterList)) {
        existing = chapterList as MpvChapter[];
      }
    } catch {
      // chapter-list may be unavailable mid-load; fall back to AniSkip chapters only
    }
    const chapters = existing.filter(
      (chapter) =>
        typeof chapter?.title !== 'string' || !chapter.title.startsWith(ANISKIP_CHAPTER_PREFIX),
    );
    chapters.push({ time: introStart, title: 'AniSkip Intro Start' });
    chapters.push({ time: introEnd, title: 'AniSkip Intro End' });
    chapters.sort((a, b) => {
      const aTime = typeof a.time === 'number' ? a.time : 0;
      const bTime = typeof b.time === 'number' ? b.time : 0;
      return aTime - bTime;
    });
    deps.sendMpvCommand(['set_property', 'chapter-list', chapters]);
  }

  async function removeIntroChapters(): Promise<void> {
    if (!deps.isMpvConnected()) return;
    let existing: MpvChapter[] = [];
    try {
      const chapterList = await deps.requestMpvProperty('chapter-list');
      if (Array.isArray(chapterList)) {
        existing = chapterList as MpvChapter[];
      }
    } catch {
      return;
    }
    const filtered = existing.filter(
      (chapter) =>
        typeof chapter?.title !== 'string' || !chapter.title.startsWith(ANISKIP_CHAPTER_PREFIX),
    );
    if (filtered.length !== existing.length) {
      deps.sendMpvCommand(['set_property', 'chapter-list', filtered]);
    }
  }

  function clearState(): void {
    requestGeneration += 1;
    introWindow = null;
    promptShown = false;
  }

  async function applyIntroWindow(window: AniSkipIntroWindow): Promise<void> {
    introWindow = window;
    promptShown = false;
    await setIntroChapters(window.start, window.end);
    deps.logInfo(
      `AniSkip intro window ${window.start.toFixed(3)} -> ${window.end.toFixed(3)} (MAL ${
        window.malId ?? '-'
      })`,
    );
  }

  async function resolveForMedia(mediaPath: string, options?: { force?: boolean }): Promise<void> {
    if (!deps.getAniSkipConfig().aniskipEnabled) return;
    if (!mediaPath || isRemoteMediaPath(mediaPath)) {
      deps.logDebug('AniSkip lookup skipped: no local media path');
      return;
    }

    if (options?.force) {
      introWindowCache.delete(mediaPath);
    }
    const generation = requestGeneration;
    const cached = introWindowCache.get(mediaPath);
    if (cached !== undefined) {
      if (cached) {
        await applyIntroWindow(cached);
      }
      return;
    }

    let metadata: AniSkipMetadata;
    try {
      metadata = await deps.resolveMetadataForFile(mediaPath);
    } catch (error) {
      deps.logWarn('AniSkip metadata lookup failed', error);
      return;
    }
    if (generation !== requestGeneration || mediaPath !== currentMediaPath) {
      return;
    }

    if (
      metadata.lookupStatus !== 'ready' ||
      metadata.introStart === null ||
      metadata.introEnd === null ||
      metadata.introEnd <= metadata.introStart
    ) {
      // Only definitive "no skip window exists" results are cached; transient
      // lookup failures stay retryable on the next load or refresh.
      if (metadata.lookupStatus !== 'lookup_failed') {
        introWindowCache.set(mediaPath, null);
      }
      deps.logInfo(
        `AniSkip: no intro window for "${metadata.title}" (status=${metadata.lookupStatus ?? 'unknown'})`,
      );
      return;
    }

    const window: AniSkipIntroWindow = {
      start: metadata.introStart,
      end: metadata.introEnd,
      malId: metadata.malId,
    };
    introWindowCache.set(mediaPath, window);
    await applyIntroWindow(window);
  }

  function skipIntroNow(): void {
    if (!deps.getAniSkipConfig().aniskipEnabled) return;
    if (!introWindow) {
      showPlaybackFeedback('Intro skip unavailable');
      return;
    }
    const now = deps.getCurrentTimePos();
    if (!Number.isFinite(now)) {
      showPlaybackFeedback('Skip unavailable');
      return;
    }
    if (
      now < introWindow.start - SKIP_WINDOW_EPSILON_SECONDS ||
      now > introWindow.end + SKIP_WINDOW_EPSILON_SECONDS
    ) {
      showPlaybackFeedback('Skip intro only during intro');
      return;
    }
    deps.sendMpvCommand(['set_property', 'time-pos', introWindow.end]);
    showPlaybackFeedback('Skipped intro');
  }

  function handleTimePosChange({ time }: { time: number }): void {
    if (!introWindow || promptShown) return;
    if (!deps.getAniSkipConfig().aniskipEnabled) return;
    const promptWindowEnd = Math.min(introWindow.start + PROMPT_WINDOW_SECONDS, introWindow.end);
    if (time >= introWindow.start && time < promptWindowEnd) {
      promptShown = true;
      showPlaybackFeedback(`You can skip by pressing ${resolveButtonKey()}`);
    }
  }

  function handleMediaPathChange({ path }: { path: string }): void {
    const nextPath = typeof path === 'string' ? path : '';
    if (nextPath === currentMediaPath && introWindow) {
      // Same-media reload: mpv rebuilt the chapter list, so re-apply markers.
      void setIntroChapters(introWindow.start, introWindow.end).catch(() => {});
      return;
    }
    currentMediaPath = nextPath;
    clearState();
    if (!nextPath) return;
    void resolveForMedia(nextPath).catch((error) => {
      deps.logWarn('AniSkip media resolution failed', error);
    });
  }

  function handleConnectionChange({ connected }: { connected: boolean }): void {
    if (!connected) {
      boundButtonKey = null;
      legacyFallbackBound = false;
      clearState();
      currentMediaPath = '';
      return;
    }
    bindSkipKeys();
  }

  function handleClientMessage({ args }: { args: string[] }): void {
    const messageName = args[0];
    if (messageName === ANISKIP_SKIP_INTRO_MESSAGE) {
      skipIntroNow();
      return;
    }
    if (messageName === ANISKIP_REFRESH_MESSAGE) {
      const mediaPath = currentMediaPath;
      if (!mediaPath) return;
      clearState();
      void removeIntroChapters().catch(() => {});
      void resolveForMedia(mediaPath, { force: true }).catch((error) => {
        deps.logWarn('AniSkip refresh failed', error);
      });
    }
  }

  function applyConfigChange(): void {
    bindSkipKeys();
    const enabled = deps.getAniSkipConfig().aniskipEnabled;
    if (!enabled) {
      clearState();
      void removeIntroChapters().catch(() => {});
      return;
    }
    if (!introWindow && currentMediaPath) {
      void resolveForMedia(currentMediaPath).catch((error) => {
        deps.logWarn('AniSkip media resolution failed', error);
      });
    }
  }

  return {
    handleConnectionChange,
    handleMediaPathChange,
    handleTimePosChange,
    handleClientMessage,
    applyConfigChange,
    skipIntroNow,
    getIntroWindow: () => introWindow,
  };
}

export type AniSkipRuntime = ReturnType<typeof createAniSkipRuntime>;
