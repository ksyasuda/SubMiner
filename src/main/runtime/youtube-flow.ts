import os from 'node:os';
import path from 'node:path';
import type {
  YoutubePickerOpenPayload,
  YoutubePickerResolveRequest,
  YoutubePickerResolveResult,
} from '../../types';
import type {
  YoutubeTrackOption,
  YoutubeTrackProbeResult,
} from '../../core/services/youtube/track-probe';
import {
  chooseDefaultYoutubeTrackIds,
  normalizeYoutubeTrackSelection,
} from '../../core/services/youtube/track-selection';
import {
  acquireYoutubeSubtitleTrack,
  acquireYoutubeSubtitleTracks,
} from '../../core/services/youtube/generate';
import { resolveSubtitleSourcePath } from './subtitle-prefetch-source';

type YoutubeFlowOpenPicker = (payload: YoutubePickerOpenPayload) => Promise<boolean>;
type YoutubeFlowMode = 'download' | 'generate';

type YoutubeFlowDeps = {
  probeYoutubeTracks: (url: string) => Promise<YoutubeTrackProbeResult>;
  acquireYoutubeSubtitleTrack: typeof acquireYoutubeSubtitleTrack;
  acquireYoutubeSubtitleTracks: typeof acquireYoutubeSubtitleTracks;
  retimeYoutubePrimaryTrack: (input: {
    targetUrl: string;
    primaryTrack: YoutubeTrackOption;
    primaryPath: string;
    secondaryTrack: YoutubeTrackOption | null;
    secondaryPath: string | null;
  }) => Promise<string>;
  openPicker: YoutubeFlowOpenPicker;
  pauseMpv: () => void;
  resumeMpv: () => void;
  sendMpvCommand: (command: Array<string | number>) => void;
  requestMpvProperty: (name: string) => Promise<unknown>;
  refreshCurrentSubtitle: (text: string) => void;
  refreshSubtitleSidebarSource?: (sourcePath: string) => Promise<void>;
  startTokenizationWarmups: () => Promise<void>;
  waitForTokenizationReady: () => Promise<void>;
  waitForAnkiReady: () => Promise<void>;
  wait: (ms: number) => Promise<void>;
  waitForPlaybackWindowReady: () => Promise<void>;
  waitForOverlayGeometryReady: () => Promise<void>;
  focusOverlayWindow: () => void;
  showMpvOsd: (text: string) => void;
  reportSubtitleFailure: (message: string) => void;
  warn: (message: string) => void;
  log: (message: string) => void;
  getYoutubeOutputDir: () => string;
};

type YoutubeFlowSession = {
  sessionId: string;
  resolve: (request: YoutubePickerResolveRequest) => void;
  reject: (error: Error) => void;
};

const YOUTUBE_PICKER_SETTLE_DELAY_MS = 150;
const YOUTUBE_SECONDARY_RETRY_DELAY_MS = 350;

function createSessionId(): string {
  return `yt-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

function getTrackById(tracks: YoutubeTrackOption[], id: string | null): YoutubeTrackOption | null {
  if (!id) return null;
  return tracks.find((track) => track.id === id) ?? null;
}

function normalizeOutputPath(value: string): string {
  const trimmed = value.trim();
  return trimmed || path.join(os.tmpdir(), 'subminer-youtube-subs');
}

function createYoutubeFlowOsdProgress(showMpvOsd: (text: string) => void) {
  const frames = ['|', '/', '-', '\\'];
  let timer: ReturnType<typeof setInterval> | null = null;
  let frame = 0;

  const stop = (): void => {
    if (!timer) {
      return;
    }
    clearInterval(timer);
    timer = null;
  };

  const setMessage = (message: string): void => {
    stop();
    frame = 0;
    showMpvOsd(message);
    timer = setInterval(() => {
      showMpvOsd(`${message} ${frames[frame % frames.length]}`);
      frame += 1;
    }, 180);
  };

  return {
    setMessage,
    stop,
  };
}

function releasePlaybackGate(deps: YoutubeFlowDeps): void {
  deps.sendMpvCommand(['script-message', 'subminer-autoplay-ready']);
  deps.resumeMpv();
}

function suppressYoutubeSubtitleState(deps: YoutubeFlowDeps): void {
  deps.sendMpvCommand(['set_property', 'sub-auto', 'no']);
  deps.sendMpvCommand(['set_property', 'sid', 'no']);
  deps.sendMpvCommand(['set_property', 'secondary-sid', 'no']);
  deps.sendMpvCommand(['set_property', 'sub-visibility', 'no']);
  deps.sendMpvCommand(['set_property', 'secondary-sub-visibility', 'no']);
}

function restoreOverlayInputFocus(deps: YoutubeFlowDeps): void {
  deps.focusOverlayWindow();
}

function parseTrackId(value: unknown): number | null {
  if (typeof value === 'number' && Number.isInteger(value)) {
    return value;
  }
  if (typeof value === 'string') {
    const parsed = Number(value.trim());
    return Number.isInteger(parsed) ? parsed : null;
  }
  return null;
}

async function ensureSubtitleTrackSelection(input: {
  deps: YoutubeFlowDeps;
  property: 'sid' | 'secondary-sid';
  targetId: number;
}): Promise<void> {
  input.deps.sendMpvCommand(['set_property', input.property, input.targetId]);
  for (let attempt = 0; attempt < 4; attempt += 1) {
    const currentId = parseTrackId(await input.deps.requestMpvProperty(input.property));
    if (currentId === input.targetId) {
      return;
    }
    await input.deps.wait(100);
    input.deps.sendMpvCommand(['set_property', input.property, input.targetId]);
  }
}

function normalizeTrackListEntry(track: Record<string, unknown>): {
  id: number | null;
  lang: string;
  title: string;
  external: boolean;
  externalFilename: string | null;
} {
  const externalFilenameRaw =
    typeof track['external-filename'] === 'string'
      ? track['external-filename']
      : typeof track.external_filename === 'string'
        ? track.external_filename
        : '';
  const externalFilename = externalFilenameRaw.trim()
    ? resolveSubtitleSourcePath(externalFilenameRaw.trim())
    : null;
  return {
    id: parseTrackId(track.id),
    lang: String(track.lang || '').trim(),
    title: String(track.title || '').trim(),
    external: track.external === true,
    externalFilename,
  };
}

function matchExternalTrackId(
  trackListRaw: unknown,
  filePath: string,
  trackOption: YoutubeTrackOption,
  excludeId: number | null = null,
): number | null {
  if (!Array.isArray(trackListRaw)) {
    return null;
  }

  const normalizedFilePath = resolveSubtitleSourcePath(filePath);
  const basename = path.basename(normalizedFilePath);
  const externalTracks = trackListRaw
    .filter(
      (track): track is Record<string, unknown> => Boolean(track) && typeof track === 'object',
    )
    .filter((track) => track.type === 'sub')
    .map(normalizeTrackListEntry)
    .filter((track) => track.external && track.id !== null && track.id !== excludeId);

  const exactPathMatch = externalTracks.find(
    (track) => track.externalFilename === normalizedFilePath,
  );
  if (exactPathMatch?.id !== null && exactPathMatch?.id !== undefined) {
    return exactPathMatch.id;
  }

  const basenameMatch = externalTracks.find(
    (track) => track.externalFilename && path.basename(track.externalFilename) === basename,
  );
  if (basenameMatch?.id !== null && basenameMatch?.id !== undefined) {
    return basenameMatch.id;
  }

  const languageMatch = externalTracks.find((track) => track.lang === trackOption.sourceLanguage);
  if (languageMatch?.id !== null && languageMatch?.id !== undefined) {
    return languageMatch.id;
  }

  const normalizedLanguageMatch = externalTracks.find(
    (track) => track.lang === trackOption.language,
  );
  if (normalizedLanguageMatch?.id !== null && normalizedLanguageMatch?.id !== undefined) {
    return normalizedLanguageMatch.id;
  }

  return null;
}

async function injectDownloadedSubtitles(
  deps: YoutubeFlowDeps,
  primaryTrack: YoutubeTrackOption,
  primaryPath: string,
  secondaryTrack: YoutubeTrackOption | null,
  secondaryPath: string | null,
): Promise<boolean> {
  deps.sendMpvCommand(['set_property', 'sub-delay', 0]);
  deps.sendMpvCommand(['set_property', 'sid', 'no']);
  deps.sendMpvCommand(['set_property', 'secondary-sid', 'no']);
  deps.sendMpvCommand([
    'sub-add',
    primaryPath,
    'select',
    path.basename(primaryPath),
    primaryTrack.sourceLanguage,
  ]);
  if (secondaryPath && secondaryTrack) {
    deps.sendMpvCommand([
      'sub-add',
      secondaryPath,
      'cached',
      path.basename(secondaryPath),
      secondaryTrack.sourceLanguage,
    ]);
  }

  let trackListRaw: unknown = null;
  let primaryTrackId: number | null = null;
  let secondaryTrackId: number | null = null;
  for (let attempt = 0; attempt < 12; attempt += 1) {
    await deps.wait(attempt === 0 ? 150 : 100);
    trackListRaw = await deps.requestMpvProperty('track-list');
    primaryTrackId = matchExternalTrackId(trackListRaw, primaryPath, primaryTrack);
    secondaryTrackId =
      secondaryPath && secondaryTrack
        ? matchExternalTrackId(trackListRaw, secondaryPath, secondaryTrack, primaryTrackId)
        : null;
    if (primaryTrackId !== null && (!secondaryPath || secondaryTrackId !== null)) {
      break;
    }
  }

  if (primaryTrackId !== null) {
    await ensureSubtitleTrackSelection({
      deps,
      property: 'sid',
      targetId: primaryTrackId,
    });
    deps.sendMpvCommand(['set_property', 'sub-visibility', 'yes']);
  } else {
    deps.warn(
      `Unable to bind downloaded primary subtitle track in mpv: ${path.basename(primaryPath)}`,
    );
  }
  if (secondaryPath && secondaryTrack) {
    if (secondaryTrackId !== null) {
      await ensureSubtitleTrackSelection({
        deps,
        property: 'secondary-sid',
        targetId: secondaryTrackId,
      });
      deps.sendMpvCommand(['set_property', 'secondary-sub-visibility', 'yes']);
    } else {
      deps.warn(
        `Unable to bind downloaded secondary subtitle track in mpv: ${path.basename(secondaryPath)}`,
      );
    }
  }

  if (primaryTrackId === null) {
    return false;
  }

  const currentSubText = await deps.requestMpvProperty('sub-text');
  if (typeof currentSubText === 'string' && currentSubText.trim().length > 0) {
    deps.refreshCurrentSubtitle(currentSubText);
  }

  deps.showMpvOsd(
    secondaryPath && secondaryTrack ? 'Primary and secondary subtitles loaded.' : 'Subtitles loaded.',
  );
  return true;
}

export function createYoutubeFlowRuntime(deps: YoutubeFlowDeps) {
  let activeSession: YoutubeFlowSession | null = null;

  const acquireSelectedTracks = async (input: {
    targetUrl: string;
    outputDir: string;
    primaryTrack: YoutubeTrackOption;
    secondaryTrack: YoutubeTrackOption | null;
    secondaryFailureLabel: string;
  }): Promise<{ primaryPath: string; secondaryPath: string | null }> => {
    if (!input.secondaryTrack) {
      const primaryPath = (
        await deps.acquireYoutubeSubtitleTrack({
          targetUrl: input.targetUrl,
          outputDir: input.outputDir,
          track: input.primaryTrack,
        })
      ).path;
      return { primaryPath, secondaryPath: null };
    }

    try {
      const batchResult = await deps.acquireYoutubeSubtitleTracks({
        targetUrl: input.targetUrl,
        outputDir: input.outputDir,
        tracks: [input.primaryTrack, input.secondaryTrack],
      });
      const primaryPath = batchResult.get(input.primaryTrack.id) ?? null;
      const secondaryPath = batchResult.get(input.secondaryTrack.id) ?? null;
      if (primaryPath) {
        if (secondaryPath) {
          return { primaryPath, secondaryPath };
        }

        deps.log(
          `${
            input.secondaryFailureLabel
          }: No subtitle file was downloaded for ${input.secondaryTrack.sourceLanguage}; retrying secondary separately after delay.`,
        );
        await deps.wait(YOUTUBE_SECONDARY_RETRY_DELAY_MS);
        try {
          const retriedSecondaryPath = (
            await deps.acquireYoutubeSubtitleTrack({
              targetUrl: input.targetUrl,
              outputDir: input.outputDir,
              track: input.secondaryTrack,
            })
          ).path;
          return { primaryPath, secondaryPath: retriedSecondaryPath };
        } catch (error) {
          deps.warn(
            `${input.secondaryFailureLabel}: ${
              error instanceof Error ? error.message : String(error)
            }`,
          );
          return { primaryPath, secondaryPath: null };
        }
      }
    } catch {
      // fall through to primary-only recovery
    }

    try {
      const primaryPath = (
        await deps.acquireYoutubeSubtitleTrack({
          targetUrl: input.targetUrl,
          outputDir: input.outputDir,
          track: input.primaryTrack,
        })
      ).path;
      return { primaryPath, secondaryPath: null };
    } catch (error) {
      throw error;
    }
  };

  const resolveActivePicker = async (
    request: YoutubePickerResolveRequest,
  ): Promise<YoutubePickerResolveResult> => {
    if (!activeSession || activeSession.sessionId !== request.sessionId) {
      return { ok: false, message: 'No active YouTube subtitle picker session.' };
    }
    activeSession.resolve(request);
    return { ok: true, message: 'Picker selection accepted.' };
  };

  const cancelActivePicker = (): boolean => {
    if (!activeSession) {
      return false;
    }
    activeSession.resolve({
      sessionId: activeSession.sessionId,
      action: 'continue-without-subtitles',
      primaryTrackId: null,
      secondaryTrackId: null,
    });
    return true;
  };

  const createPickerSelectionPromise = (sessionId: string): Promise<YoutubePickerResolveRequest> =>
    new Promise<YoutubePickerResolveRequest>((resolve, reject) => {
      activeSession = { sessionId, resolve, reject };
    }).finally(() => {
      activeSession = null;
    });

  const reportPrimarySubtitleFailure = (): void => {
    deps.reportSubtitleFailure(
      'Primary subtitles failed to load. Use the YouTube subtitle picker to try manually.',
    );
  };

  const buildOpenPayload = (
    input: {
      url: string;
    },
    probe: YoutubeTrackProbeResult,
  ): YoutubePickerOpenPayload => {
    const defaults = chooseDefaultYoutubeTrackIds(probe.tracks);
    return {
      sessionId: createSessionId(),
      url: input.url,
      tracks: probe.tracks,
      defaultPrimaryTrackId: defaults.primaryTrackId,
      defaultSecondaryTrackId: defaults.secondaryTrackId,
      hasTracks: probe.tracks.length > 0,
    };
  };

  const loadTracksIntoMpv = async (input: {
    url: string;
    mode: YoutubeFlowMode;
    outputDir: string;
    primaryTrack: YoutubeTrackOption;
    secondaryTrack: YoutubeTrackOption | null;
    secondaryFailureLabel: string;
    tokenizationWarmupPromise?: Promise<void>;
    showDownloadProgress: boolean;
  }): Promise<boolean> => {
    const osdProgress = input.showDownloadProgress
      ? createYoutubeFlowOsdProgress(deps.showMpvOsd)
      : null;
    if (osdProgress) {
      osdProgress.setMessage('Downloading subtitles...');
    }
    try {
      const acquired = await acquireSelectedTracks({
        targetUrl: input.url,
        outputDir: input.outputDir,
        primaryTrack: input.primaryTrack,
        secondaryTrack: input.secondaryTrack,
        secondaryFailureLabel: input.secondaryFailureLabel,
      });
      const resolvedPrimaryPath = await deps.retimeYoutubePrimaryTrack({
        targetUrl: input.url,
        primaryTrack: input.primaryTrack,
        primaryPath: acquired.primaryPath,
        secondaryTrack: input.secondaryTrack,
        secondaryPath: acquired.secondaryPath,
      });
      deps.showMpvOsd('Loading subtitles...');
      const refreshedActiveSubtitle = await injectDownloadedSubtitles(
        deps,
        input.primaryTrack,
        resolvedPrimaryPath,
        input.secondaryTrack,
        acquired.secondaryPath,
      );
      if (!refreshedActiveSubtitle) {
        return false;
      }
      try {
        await deps.refreshSubtitleSidebarSource?.(resolvedPrimaryPath);
      } catch (error) {
        deps.warn(
          `Failed to refresh parsed subtitle cues for sidebar: ${
            error instanceof Error ? error.message : String(error)
          }`,
        );
      }
      if (input.tokenizationWarmupPromise) {
        await input.tokenizationWarmupPromise;
      }
      await deps.waitForTokenizationReady();
      await deps.waitForAnkiReady();
      return true;
    } finally {
      osdProgress?.stop();
    }
  };

  const openManualPicker = async (input: {
    url: string;
    mode?: YoutubeFlowMode;
  }): Promise<void> => {
    let probe: YoutubeTrackProbeResult;
    try {
      probe = await deps.probeYoutubeTracks(input.url);
    } catch (error) {
      deps.warn(
        `Failed to probe YouTube subtitle tracks: ${
          error instanceof Error ? error.message : String(error)
        }`,
      );
      reportPrimarySubtitleFailure();
      restoreOverlayInputFocus(deps);
      return;
    }

    const openPayload = buildOpenPayload(input, probe);
    await deps.waitForPlaybackWindowReady();
    await deps.waitForOverlayGeometryReady();
    await deps.wait(YOUTUBE_PICKER_SETTLE_DELAY_MS);
    const pickerSelection = createPickerSelectionPromise(openPayload.sessionId);
    void pickerSelection.catch(() => undefined);

    let opened = false;
    try {
      opened = await deps.openPicker(openPayload);
    } catch (error) {
      activeSession?.reject(error instanceof Error ? error : new Error(String(error)));
      deps.warn(
        `Unable to open YouTube subtitle picker: ${
          error instanceof Error ? error.message : String(error)
        }`,
      );
      restoreOverlayInputFocus(deps);
      return;
    }
    if (!opened) {
      activeSession?.reject(new Error('Unable to open YouTube subtitle picker.'));
      activeSession = null;
      deps.warn('Unable to open YouTube subtitle picker.');
      restoreOverlayInputFocus(deps);
      return;
    }

    const request = await pickerSelection;
    if (request.action === 'continue-without-subtitles') {
      restoreOverlayInputFocus(deps);
      return;
    }

    const primaryTrack = getTrackById(probe.tracks, request.primaryTrackId);
    if (!primaryTrack) {
      deps.warn('No primary YouTube subtitle track selected.');
      restoreOverlayInputFocus(deps);
      return;
    }

    const selected = normalizeYoutubeTrackSelection({
      primaryTrackId: primaryTrack.id,
      secondaryTrackId: request.secondaryTrackId,
    });
    const secondaryTrack = getTrackById(probe.tracks, selected.secondaryTrackId);

    try {
      deps.showMpvOsd('Getting subtitles...');
      const loaded = await loadTracksIntoMpv({
        url: input.url,
        mode: input.mode ?? 'download',
        outputDir: normalizeOutputPath(deps.getYoutubeOutputDir()),
        primaryTrack,
        secondaryTrack,
        secondaryFailureLabel: 'Failed to download secondary YouTube subtitle track',
        showDownloadProgress: true,
      });
      if (!loaded) {
        reportPrimarySubtitleFailure();
      }
    } catch (error) {
      deps.warn(
        `Failed to download primary YouTube subtitle track: ${
          error instanceof Error ? error.message : String(error)
        }`,
      );
      reportPrimarySubtitleFailure();
    } finally {
      restoreOverlayInputFocus(deps);
    }
  };

  async function runYoutubePlaybackFlow(input: {
    url: string;
    mode: YoutubeFlowMode;
  }): Promise<void> {
    deps.showMpvOsd('Opening YouTube video');
    const tokenizationWarmupPromise = deps.startTokenizationWarmups().catch((error) => {
      deps.warn(
        `Failed to warm subtitle tokenization prerequisites: ${
          error instanceof Error ? error.message : String(error)
        }`,
      );
    });

    deps.pauseMpv();
    suppressYoutubeSubtitleState(deps);
    const outputDir = normalizeOutputPath(deps.getYoutubeOutputDir());

    let probe: YoutubeTrackProbeResult;
    try {
      probe = await deps.probeYoutubeTracks(input.url);
    } catch (error) {
      deps.warn(
        `Failed to probe YouTube subtitle tracks: ${
          error instanceof Error ? error.message : String(error)
        }`,
      );
      reportPrimarySubtitleFailure();
      releasePlaybackGate(deps);
      restoreOverlayInputFocus(deps);
      return;
    }

    const defaults = chooseDefaultYoutubeTrackIds(probe.tracks);
    const primaryTrack = getTrackById(probe.tracks, defaults.primaryTrackId);
    const secondaryTrack = getTrackById(probe.tracks, defaults.secondaryTrackId);
    if (!primaryTrack) {
      reportPrimarySubtitleFailure();
      releasePlaybackGate(deps);
      restoreOverlayInputFocus(deps);
      return;
    }

    try {
      deps.showMpvOsd('Getting subtitles...');
      const loaded = await loadTracksIntoMpv({
        url: input.url,
        mode: input.mode,
        outputDir,
        primaryTrack,
        secondaryTrack,
        secondaryFailureLabel:
          input.mode === 'generate'
            ? 'Failed to generate secondary YouTube subtitle track'
            : 'Failed to download secondary YouTube subtitle track',
        tokenizationWarmupPromise,
        showDownloadProgress: false,
      });
      if (!loaded) {
        reportPrimarySubtitleFailure();
      }
    } catch (error) {
      deps.warn(
        `Failed to ${
          input.mode === 'generate' ? 'generate' : 'download'
        } primary YouTube subtitle track: ${
          error instanceof Error ? error.message : String(error)
        }`,
      );
      reportPrimarySubtitleFailure();
    } finally {
      releasePlaybackGate(deps);
      restoreOverlayInputFocus(deps);
    }
  }

  return {
    runYoutubePlaybackFlow,
    openManualPicker,
    resolveActivePicker,
    cancelActivePicker,
    hasActiveSession: () => Boolean(activeSession),
  };
}
