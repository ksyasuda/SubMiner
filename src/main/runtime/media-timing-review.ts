import { randomUUID } from 'crypto';
import type {
  MediaTimingReviewActionResult,
  MediaTimingReviewContextLine,
  MediaTimingReviewDecision,
  MediaTimingReviewOpenPayload,
  MediaTimingReviewPreviewRequest,
  MediaTimingReviewRequest,
  MediaTimingReviewResolveRequest,
  MediaTimingReviewWaveformRequest,
  MediaTimingReviewWaveformResult,
} from '../../types/anki';
import type { SpeechWaveformOptions } from '../../core/services/media-timing-waveform';
import {
  isRemoteMediaWindowSourcePath,
  type RemoteMediaWindow,
  type RemoteMediaWindowRange,
  type RemoteMediaWindowSource,
} from '../../core/services/remote-media-window-cache';
import type { MediaInput, MediaInputOptions } from '../../media-input';

const INITIAL_TIMELINE_MARGIN_SECONDS = 2;
const REVIEW_DECISION_TIMEOUT_MS = 5 * 60_000;
const CONTEXT_LINE_LIMIT = 12;
const CONTEXT_LINE_EPSILON_SECONDS = 0.05;

interface ReviewMpvClient {
  connected: boolean;
  currentVideoPath: string;
  currentAudioStreamIndex?: number | null;
  requestProperty?: (name: string) => Promise<unknown>;
  send: (payload: { command: Array<string | number> }) => void;
}

interface PreviewSession {
  start(options: {
    mediaPath: string;
    executablePath?: string;
    audioTrackId?: number;
    volume?: number;
    absoluteTimestamps?: boolean;
  }): Promise<void>;
  play(startTime: number, endTime: number): Promise<void>;
  stop(): Promise<void>;
  /** Fires when the player reaches the end of the clip started by play(). */
  onPlaybackEnded(listener: () => void): void;
  dispose(): void;
}

interface ReviewMediaSource {
  path: string;
  inputOptions?: MediaInputOptions;
  singleResolvedStream?: boolean;
}

interface ActiveReview {
  payload: MediaTimingReviewOpenPayload;
  /** What the hidden mpv preview plays when no cached window is available. */
  mediaPath: string;
  /** What the waveform reads when no cached window is available. */
  waveformMedia: MediaInput;
  audioStreamIndex?: number;
  /** Remote source to download windows of; null for local media or without a cache. */
  windowSource: RemoteMediaWindowSource | null;
  /** Latest window returned for this review; reused while it still covers the request. */
  window: RemoteMediaWindow | null;
  windowRequest: (RemoteMediaWindowRange & { promise: Promise<RemoteMediaWindow | null> }) | null;
  windowFailed: boolean;
  previewOptions: { executablePath?: string; audioTrackId?: number; volume?: number };
  preview: { path: string; session: Promise<PreviewSession> } | null;
  mpvClient: ReviewMpvClient;
  restorePlayback: boolean;
  resolve: (decision: MediaTimingReviewDecision) => void;
}

export interface MediaTimingReviewRuntimeDeps {
  getMpvClient: () => ReviewMpvClient | null;
  getCurrentMediaPath: () => string | null;
  getMpvExecutablePath: () => string;
  createPreviewSession: () => PreviewSession;
  generateWaveform: (options: SpeechWaveformOptions) => Promise<number[]>;
  /** Resolves the FFmpeg-readable stream URL and headers behind the current media path. */
  resolveMediaSource?: () => Promise<ReviewMediaSource | null>;
  /** Downloads (or reuses) a local window of a remote source covering the range. */
  acquireMediaWindow?: (
    source: RemoteMediaWindowSource,
    range: RemoteMediaWindowRange,
  ) => Promise<RemoteMediaWindow>;
  getSubtitleContextLines?: (range: { startTime: number; endTime: number }) => {
    previous: MediaTimingReviewContextLine[];
    next: MediaTimingReviewContextLine[];
  };
  decisionTimeoutMs?: number;
  openModal: (payload: MediaTimingReviewOpenPayload) => Promise<boolean>;
  /** Tells the modal that the hidden player finished the previewed clip. */
  onPreviewEnded?: (reviewId: string) => void;
  showStatus: (message: string) => void;
}

function finiteNumber(value: unknown): number | null {
  return typeof value === 'number' && Number.isFinite(value) ? value : null;
}

function booleanProperty(value: unknown): boolean | null {
  if (typeof value === 'boolean') return value;
  if (value === 'yes' || value === 1) return true;
  if (value === 'no' || value === 0) return false;
  return null;
}

/**
 * Picks the subtitle lines adjacent to the mined range that the review modal can pull
 * onto the card. Parsed cues cover both directions; when none are loaded (e.g. the
 * active track was never parsed) the timing tracker's history still provides the
 * lines that already played, so only "next" is unavailable.
 */
export function collectMediaTimingContextLines(options: {
  cues: readonly { text: string; startTime: number; endTime: number }[];
  fallbackPrevious?: readonly { displayText: string; startTime: number; endTime: number }[];
  startTime: number;
  endTime: number;
}): { previous: MediaTimingReviewContextLine[]; next: MediaTimingReviewContextLine[] } {
  const usable = options.cues
    .filter(
      (cue) =>
        cue.text.trim().length > 0 &&
        Number.isFinite(cue.startTime) &&
        Number.isFinite(cue.endTime) &&
        cue.endTime > cue.startTime,
    )
    .sort((a, b) => a.startTime - b.startTime || a.endTime - b.endTime);

  let previous = usable
    .filter((cue) => cue.endTime <= options.startTime + CONTEXT_LINE_EPSILON_SECONDS)
    .slice(-CONTEXT_LINE_LIMIT)
    .map(({ text, startTime, endTime }) => ({ text: text.trim(), startTime, endTime }));
  const next = usable
    .filter((cue) => cue.startTime >= options.endTime - CONTEXT_LINE_EPSILON_SECONDS)
    .slice(0, CONTEXT_LINE_LIMIT)
    .map(({ text, startTime, endTime }) => ({ text: text.trim(), startTime, endTime }));

  if (previous.length === 0 && options.fallbackPrevious) {
    previous = options.fallbackPrevious
      .filter(
        (entry) =>
          entry.displayText.trim().length > 0 &&
          Number.isFinite(entry.startTime) &&
          Number.isFinite(entry.endTime) &&
          entry.endTime > entry.startTime &&
          entry.endTime <= options.startTime + CONTEXT_LINE_EPSILON_SECONDS,
      )
      .slice(-CONTEXT_LINE_LIMIT)
      .map((entry) => ({
        text: entry.displayText.trim(),
        startTime: entry.startTime,
        endTime: entry.endTime,
      }));
  }
  return { previous, next };
}

function isValidMediaTimingRange(
  payload: MediaTimingReviewOpenPayload,
  startTime: number,
  endTime: number,
): boolean {
  return (
    Number.isFinite(startTime) &&
    Number.isFinite(endTime) &&
    startTime >= 0 &&
    endTime > startTime &&
    (payload.maxMediaDuration <= 0 || endTime - startTime <= payload.maxMediaDuration + 0.001) &&
    (payload.mediaDuration === undefined || endTime <= payload.mediaDuration + 0.001)
  );
}

export function buildMediaTimingReviewPayload(
  request: MediaTimingReviewRequest,
  options: {
    reviewId: string;
    mediaDuration?: number;
    contextLines?: {
      previous: MediaTimingReviewContextLine[];
      next: MediaTimingReviewContextLine[];
    };
  },
): MediaTimingReviewOpenPayload {
  const duration = finiteNumber(options.mediaDuration);
  const maxTime = duration !== null && duration > 0 ? duration : Number.POSITIVE_INFINITY;
  const paddedStart = Math.max(0, request.startTime - request.audioPadding);
  let paddedEnd = Math.min(maxTime, request.endTime + request.audioPadding);
  const maxMediaDuration = Math.max(0, request.maxMediaDuration);
  if (maxMediaDuration > 0 && paddedEnd - paddedStart > maxMediaDuration) {
    paddedEnd = paddedStart + maxMediaDuration;
  }
  if (paddedEnd <= paddedStart) {
    paddedEnd = Math.min(maxTime, paddedStart + 0.1);
  }

  const timelineStartTime = Math.max(0, paddedStart - INITIAL_TIMELINE_MARGIN_SECONDS);
  const timelineEndTime = Math.max(
    paddedEnd,
    Math.min(maxTime, paddedEnd + INITIAL_TIMELINE_MARGIN_SECONDS),
  );

  return {
    reviewId: options.reviewId,
    kind: request.kind,
    text: request.text,
    previousLines: options.contextLines?.previous ?? [],
    nextLines: options.contextLines?.next ?? [],
    ...(request.noteId !== undefined ? { noteId: request.noteId } : {}),
    originalStartTime: request.startTime,
    originalEndTime: request.endTime,
    selectionStartTime: paddedStart,
    selectionEndTime: paddedEnd,
    timelineStartTime,
    timelineEndTime,
    ...(duration !== null && duration > 0 ? { mediaDuration: duration } : {}),
    maxMediaDuration,
  };
}

export function createMediaTimingReviewRuntime(deps: MediaTimingReviewRuntimeDeps) {
  let active: ActiveReview | null = null;
  let reviewInProgress = false;
  let pendingPauseRestore: ReviewMpvClient | null = null;

  function restorePendingPlayback(): void {
    const mpvClient = pendingPauseRestore;
    pendingPauseRestore = null;
    if (mpvClient?.connected) {
      mpvClient.send({ command: ['set_property', 'pause', 'no'] });
    }
  }

  function ensureWindow(
    review: ActiveReview,
    range: RemoteMediaWindowRange,
  ): Promise<RemoteMediaWindow | null> {
    const { windowSource } = review;
    if (!windowSource || review.windowFailed || !deps.acquireMediaWindow) {
      return Promise.resolve(null);
    }
    const coversRange = (candidate: RemoteMediaWindowRange): boolean =>
      candidate.startTime <= range.startTime && candidate.endTime >= range.endTime;
    if (review.window && coversRange(review.window)) return Promise.resolve(review.window);
    const inFlight = review.windowRequest;
    if (inFlight && coversRange(inFlight)) return inFlight.promise;

    const request = {
      startTime: range.startTime,
      endTime: range.endTime,
      promise: Promise.resolve<RemoteMediaWindow | null>(null),
    };
    request.promise = deps
      .acquireMediaWindow(windowSource, { startTime: range.startTime, endTime: range.endTime })
      .then((window) => {
        review.window = window;
        return window;
      })
      .catch(() => {
        // Fall back to the remote source for the rest of this review instead of retrying.
        review.windowFailed = true;
        return null;
      })
      .finally(() => {
        if (review.windowRequest === request) review.windowRequest = null;
      });
    review.windowRequest = request;
    return request.promise;
  }

  /**
   * Returns the preview player for the range, restarting it when the range needs a
   * different file (the first cached window, or a wider one after the timeline grew).
   */
  async function previewFor(
    review: ActiveReview,
    range: RemoteMediaWindowRange,
  ): Promise<PreviewSession> {
    const window = await ensureWindow(review, range);
    if (active !== review) {
      // The review ended during the download; do not start a player nobody will dispose.
      throw new Error('This timing review is no longer active.');
    }
    const mediaPath = window?.path ?? review.mediaPath;
    if (review.preview?.path === mediaPath) return review.preview.session;

    const previous = review.preview;
    const session = deps.createPreviewSession();
    session.onPlaybackEnded(() => {
      if (active === review && review.preview?.session === started) {
        deps.onPreviewEnded?.(review.payload.reviewId);
      }
    });
    const { audioTrackId, ...previewOptions } = review.previewOptions;
    const started = session
      .start({
        mediaPath,
        ...previewOptions,
        // A cached window keeps one audio stream, so mpv's track id from the source no longer applies.
        ...(window
          ? { absoluteTimestamps: true }
          : audioTrackId !== undefined
            ? { audioTrackId }
            : {}),
      })
      .then(() => session)
      .catch((error) => {
        session.dispose();
        throw error;
      });
    review.preview = { path: mediaPath, session: started };
    void started.catch(() => {});
    if (previous) void previous.session.then((old) => old.dispose()).catch(() => {});
    return started;
  }

  async function runReview(request: MediaTimingReviewRequest): Promise<MediaTimingReviewDecision> {
    const mpvClient = deps.getMpvClient();
    const mediaPath =
      deps.getCurrentMediaPath()?.trim() || mpvClient?.currentVideoPath?.trim() || '';
    if (!mpvClient?.connected || !mediaPath) {
      deps.showStatus('Timing review unavailable. Using the original subtitle timing.');
      return { action: 'use-original' };
    }

    const [pauseRaw, durationRaw, audioTrackRaw, volumeRaw, resolvedSource] = await Promise.all([
      mpvClient.requestProperty?.('pause').catch(() => null) ?? null,
      mpvClient.requestProperty?.('duration').catch(() => null) ?? null,
      mpvClient.requestProperty?.('aid').catch(() => null) ?? null,
      mpvClient.requestProperty?.('volume').catch(() => null) ?? null,
      deps.resolveMediaSource?.().catch(() => null) ?? null,
    ]);
    const pauseState = booleanProperty(pauseRaw);
    mpvClient.send({ command: ['set_property', 'pause', 'yes'] });
    pendingPauseRestore = pauseState === false ? mpvClient : null;

    let contextLines: ReturnType<NonNullable<typeof deps.getSubtitleContextLines>> | undefined;
    try {
      contextLines = deps.getSubtitleContextLines?.({
        startTime: request.startTime,
        endTime: request.endTime,
      });
    } catch {
      contextLines = undefined;
    }
    const payload = buildMediaTimingReviewPayload(request, {
      reviewId: randomUUID(),
      mediaDuration: finiteNumber(durationRaw) ?? undefined,
      ...(contextLines ? { contextLines } : {}),
    });
    const sourcePath = resolvedSource?.path.trim() || mediaPath;
    const inputOptions = resolvedSource?.inputOptions;
    const audioStreamIndex =
      resolvedSource?.singleResolvedStream || mpvClient.currentAudioStreamIndex == null
        ? undefined
        : mpvClient.currentAudioStreamIndex;
    const windowSource: RemoteMediaWindowSource | null =
      deps.acquireMediaWindow && isRemoteMediaWindowSourcePath(sourcePath)
        ? {
            path: sourcePath,
            ...(inputOptions ? { inputOptions } : {}),
            audioStreamIndex: audioStreamIndex ?? null,
          }
        : null;

    let resolveDecision!: (decision: MediaTimingReviewDecision) => void;
    const decisionPromise = new Promise<MediaTimingReviewDecision>((resolve) => {
      resolveDecision = resolve;
    });
    const review: ActiveReview = {
      payload,
      mediaPath,
      waveformMedia: inputOptions ? { path: sourcePath, inputOptions } : sourcePath,
      ...(audioStreamIndex !== undefined ? { audioStreamIndex } : {}),
      windowSource,
      window: null,
      windowRequest: null,
      windowFailed: false,
      previewOptions: {
        executablePath: deps.getMpvExecutablePath(),
        audioTrackId: finiteNumber(audioTrackRaw) ?? undefined,
        volume: finiteNumber(volumeRaw) ?? undefined,
      },
      preview: null,
      mpvClient,
      restorePlayback: pendingPauseRestore === mpvClient,
      resolve: resolveDecision,
    };
    active = review;
    pendingPauseRestore = null;
    // Download the visible timeline once now; the waveform and preview both wait on it.
    void previewFor(review, {
      startTime: payload.timelineStartTime,
      endTime: payload.timelineEndTime,
    }).catch(() => {});

    const opened = await deps.openModal(payload).catch(() => false);
    if (!opened) {
      await cleanupActiveReview();
      deps.showStatus('Timing review could not open. Using the original subtitle timing.');
      return { action: 'use-original' };
    }

    const decisionWatchdog = setTimeout(
      () => resolveDecision({ action: 'use-original' }),
      Math.max(0, deps.decisionTimeoutMs ?? REVIEW_DECISION_TIMEOUT_MS),
    );
    let decision: MediaTimingReviewDecision;
    try {
      decision = await decisionPromise;
    } finally {
      clearTimeout(decisionWatchdog);
    }
    await cleanupActiveReview();
    return decision;
  }

  async function requestReview(
    request: MediaTimingReviewRequest,
  ): Promise<MediaTimingReviewDecision> {
    if (active || reviewInProgress) {
      deps.showStatus('Finish the current timing review before mining another card.');
      return { action: 'use-original' };
    }
    reviewInProgress = true;
    try {
      return await runReview(request);
    } catch {
      await cleanupActiveReview();
      restorePendingPlayback();
      deps.showStatus('Timing review failed. Using the original subtitle timing.');
      return { action: 'use-original' };
    } finally {
      reviewInProgress = false;
    }
  }

  async function previewRange(
    request: MediaTimingReviewPreviewRequest,
  ): Promise<MediaTimingReviewActionResult> {
    const current = active;
    if (!current || request.reviewId !== current.payload.reviewId) {
      return { ok: false, message: 'This timing review is no longer active.' };
    }
    if (!isValidMediaTimingRange(current.payload, request.startTime, request.endTime)) {
      return { ok: false, message: 'The selected preview range is invalid.' };
    }
    try {
      const previewSession = await previewFor(current, request);
      if (active !== current) {
        return { ok: false, message: 'This timing review is no longer active.' };
      }
      await previewSession.play(request.startTime, request.endTime);
      return { ok: true };
    } catch (error) {
      return {
        ok: false,
        message: `Audio preview unavailable: ${error instanceof Error ? error.message : String(error)}`,
      };
    }
  }

  async function getWaveform(
    request: MediaTimingReviewWaveformRequest,
  ): Promise<MediaTimingReviewWaveformResult> {
    const current = active;
    if (!current || request.reviewId !== current.payload.reviewId) {
      return { ok: false, message: 'This timing review is no longer active.' };
    }
    if (
      !Number.isFinite(request.startTime) ||
      !Number.isFinite(request.endTime) ||
      request.startTime < 0 ||
      request.endTime <= request.startTime ||
      (current.payload.mediaDuration !== undefined &&
        request.endTime > current.payload.mediaDuration + 0.001)
    ) {
      return { ok: false, message: 'The waveform range is invalid.' };
    }

    try {
      const window = await ensureWindow(current, request);
      if (active !== current) {
        return { ok: false, message: 'This timing review is no longer active.' };
      }
      const peaks = await deps.generateWaveform({
        mediaPath: window?.media ?? current.waveformMedia,
        startTime: request.startTime,
        endTime: request.endTime,
        ...(!window && current.audioStreamIndex !== undefined
          ? { audioStreamIndex: current.audioStreamIndex }
          : {}),
      });
      if (active !== current || peaks.length < 2 || peaks.some((peak) => !Number.isFinite(peak))) {
        return { ok: false, message: 'Timing waveform is unavailable.' };
      }
      return { ok: true, peaks };
    } catch {
      return { ok: false, message: 'Timing waveform is unavailable.' };
    }
  }

  async function stopPreview(reviewId: string): Promise<MediaTimingReviewActionResult> {
    const current = active;
    if (!current || reviewId !== current.payload.reviewId) {
      return { ok: false, message: 'This timing review is no longer active.' };
    }
    try {
      const previewSession = current.preview ? await current.preview.session : null;
      await previewSession?.stop();
      return { ok: true };
    } catch (error) {
      return {
        ok: false,
        message: `Could not stop preview: ${error instanceof Error ? error.message : String(error)}`,
      };
    }
  }

  function resolveReview(request: MediaTimingReviewResolveRequest): MediaTimingReviewActionResult {
    const current = active;
    if (!current || request.reviewId !== current.payload.reviewId) {
      return { ok: false, message: 'This timing review is no longer active.' };
    }
    if (request.decision.action === 'confirm') {
      const { startTime, endTime, text } = request.decision;
      if (!isValidMediaTimingRange(current.payload, startTime, endTime)) {
        return { ok: false, message: 'The selected timing range is invalid.' };
      }
      if (text !== undefined && (typeof text !== 'string' || text.trim().length === 0)) {
        return { ok: false, message: 'The combined sentence text is invalid.' };
      }
    }
    current.resolve(request.decision);
    return { ok: true };
  }

  async function cleanupActiveReview(): Promise<void> {
    const current = active;
    active = null;
    if (!current) return;
    void current.preview?.session.then((session) => session.dispose()).catch(() => {});
    if (current.restorePlayback && current.mpvClient.connected) {
      current.mpvClient.send({ command: ['set_property', 'pause', 'no'] });
    }
  }

  async function dispose(): Promise<void> {
    active?.resolve({ action: 'use-original' });
    await cleanupActiveReview();
    restorePendingPlayback();
  }

  return {
    requestReview,
    previewRange,
    getWaveform,
    stopPreview,
    resolveReview,
    dispose,
  };
}
