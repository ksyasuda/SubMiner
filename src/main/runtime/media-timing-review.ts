import { randomUUID } from 'crypto';
import type {
  MediaTimingReviewActionResult,
  MediaTimingReviewContextLine,
  MediaTimingReviewDecision,
  MediaTimingReviewOpenPayload,
  MediaTimingReviewPreviewRequest,
  MediaTimingReviewRequest,
  MediaTimingReviewResolveRequest,
  MediaTimingReviewFrameRequest,
  MediaTimingReviewFrameResult,
  MediaTimingReviewWaveformRequest,
  MediaTimingReviewWaveformResult,
} from '../../types/anki';
import type { SpeechWaveformOptions } from '../../core/services/media-timing-waveform';
import type { MediaTimingFrameOptions } from '../../core/services/media-timing-frame';
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
  videoSource: ReviewMediaSource | null;
  frameInFlight: boolean;
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

interface ReviewRequestLifecycle {
  signal: AbortSignal;
  cancelled: Promise<void>;
  settled: Promise<void>;
  isCancelled(): boolean;
  cancel(): void;
  markSettled(): void;
}

export interface MediaTimingReviewRuntimeDeps {
  getMpvClient: () => ReviewMpvClient | null;
  getCurrentMediaPath: () => string | null;
  getMpvExecutablePath: () => string;
  createPreviewSession: () => PreviewSession;
  generateWaveform: (options: SpeechWaveformOptions) => Promise<number[]>;
  /** Resolves the FFmpeg-readable stream URL and headers behind the current media path. */
  resolveMediaSource?: () => Promise<ReviewMediaSource | null>;
  resolveVideoSource?: () => Promise<ReviewMediaSource | null>;
  generateFrame?: (
    options: MediaTimingFrameOptions,
  ) => Promise<{ dataUrl: string; timestamp: number }>;
  clearFrameCache?: () => void;
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
  openModal: (payload: MediaTimingReviewOpenPayload, signal: AbortSignal) => Promise<boolean>;
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

function createReviewRequestLifecycle(): ReviewRequestLifecycle {
  const controller = new AbortController();
  let resolveCancellation: (() => void) | null = null;
  let resolveSettled: (() => void) | null = null;
  const cancellation = new Promise<void>((resolve) => {
    resolveCancellation = resolve;
  });
  const settled = new Promise<void>((resolve) => {
    resolveSettled = resolve;
  });
  return {
    signal: controller.signal,
    cancelled: cancellation,
    settled,
    isCancelled: () => controller.signal.aborted,
    cancel: () => {
      if (controller.signal.aborted) return;
      controller.abort();
      resolveCancellation?.();
    },
    markSettled: () => {
      resolveSettled?.();
      resolveSettled = null;
    },
  };
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

/**
 * Result for requests that name a review main has already resolved or disposed (decision
 * watchdog, overlay teardown, duplicate modal). The renderer closes on it instead of
 * leaving the user with controls that can never succeed.
 */
function staleReviewResult(): MediaTimingReviewActionResult {
  return { ok: false, stale: true, message: 'This timing review is no longer active.' };
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
    ...(request.screenshotEnabled !== undefined
      ? { screenshotEnabled: request.screenshotEnabled }
      : {}),
  };
}

export function createMediaTimingReviewRuntime(deps: MediaTimingReviewRuntimeDeps) {
  let active: ActiveReview | null = null;
  let currentRequest: ReviewRequestLifecycle | null = null;
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
    const { audioTrackId, ...previewOptions } = review.previewOptions;
    const startSession = async (): Promise<PreviewSession> => {
      if (active !== review) {
        throw new Error('This timing review is no longer active.');
      }
      await session.start({
        mediaPath,
        ...previewOptions,
        // A cached window keeps one audio stream, so mpv's track id from the source no longer applies.
        ...(window
          ? { absoluteTimestamps: true }
          : audioTrackId !== undefined
            ? { audioTrackId }
            : {}),
      });
      return session;
    };
    const started = Promise.resolve()
      .then(startSession)
      .catch((error: unknown) => {
        session.dispose();
        throw error;
      });
    review.preview = { path: mediaPath, session: started };
    session.onPlaybackEnded(() => {
      if (active === review && review.preview?.session === started) {
        deps.onPreviewEnded?.(review.payload.reviewId);
      }
    });
    void started.catch(() => {});
    if (previous) void previous.session.then((old) => old.dispose()).catch(() => {});
    return started;
  }

  async function runReview(
    request: MediaTimingReviewRequest,
    lifecycle: ReviewRequestLifecycle,
  ): Promise<MediaTimingReviewDecision> {
    const mpvClient = deps.getMpvClient();
    const mediaPath =
      deps.getCurrentMediaPath()?.trim() || mpvClient?.currentVideoPath?.trim() || '';
    if (!mpvClient?.connected || !mediaPath) {
      deps.showStatus('Timing review unavailable. Using the original subtitle timing.');
      return { action: 'use-original' };
    }

    const setupPromise = Promise.all([
      mpvClient.requestProperty?.('pause').catch(() => null) ?? null,
      mpvClient.requestProperty?.('duration').catch(() => null) ?? null,
      mpvClient.requestProperty?.('aid').catch(() => null) ?? null,
      mpvClient.requestProperty?.('volume').catch(() => null) ?? null,
      deps.resolveMediaSource?.().catch(() => null) ?? null,
      request.screenshotEnabled ? (deps.resolveVideoSource?.().catch(() => null) ?? null) : null,
    ]);
    const setup = await Promise.race([
      setupPromise.then((values) => ({ kind: 'ready' as const, values })),
      lifecycle.cancelled.then(() => ({ kind: 'cancelled' as const })),
    ]);
    if (setup.kind === 'cancelled' || lifecycle.isCancelled()) {
      return { action: 'use-original' };
    }
    const [pauseRaw, durationRaw, audioTrackRaw, volumeRaw, resolvedSource, videoSource] =
      setup.values;
    const pauseState = booleanProperty(pauseRaw);
    pendingPauseRestore = pauseState === false ? mpvClient : null;
    mpvClient.send({ command: ['set_property', 'pause', 'yes'] });
    if (lifecycle.isCancelled()) {
      restorePendingPlayback();
      return { action: 'use-original' };
    }

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
      videoSource,
      frameInFlight: false,
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

    if (lifecycle.isCancelled() || active !== review) {
      await cleanupActiveReview(review);
      return { action: 'use-original' };
    }
    const openModal = deps.openModal(payload, lifecycle.signal).catch(() => false);
    const openResult = await Promise.race([
      openModal.then((opened) => ({ kind: 'opened' as const, opened })),
      lifecycle.cancelled.then(() => ({ kind: 'cancelled' as const })),
    ]);
    if (openResult.kind === 'cancelled' || lifecycle.isCancelled() || active !== review) {
      await cleanupActiveReview(review);
      return { action: 'use-original' };
    }
    const { opened } = openResult;
    if (!opened) {
      await cleanupActiveReview(review);
      deps.showStatus('Timing review could not open. Using the original subtitle timing.');
      return { action: 'use-original' };
    }

    const decisionWatchdog = setTimeout(
      () => resolveDecision({ action: 'use-original' }),
      Math.max(0, deps.decisionTimeoutMs ?? REVIEW_DECISION_TIMEOUT_MS),
    );
    let decision: MediaTimingReviewDecision = { action: 'use-original' };
    try {
      const decisionResult = await Promise.race([
        decisionPromise.then((value) => ({ kind: 'decided' as const, value })),
        lifecycle.cancelled.then(() => ({ kind: 'cancelled' as const })),
      ]);
      if (decisionResult.kind === 'decided') {
        decision = decisionResult.value;
      }
    } finally {
      clearTimeout(decisionWatchdog);
    }
    await cleanupActiveReview(review);
    return decision;
  }

  async function requestReview(
    request: MediaTimingReviewRequest,
  ): Promise<MediaTimingReviewDecision> {
    if (active || currentRequest) {
      deps.showStatus('Finish the current timing review before mining another card.');
      return { action: 'use-original' };
    }
    const lifecycle = createReviewRequestLifecycle();
    currentRequest = lifecycle;
    try {
      return await runReview(request, lifecycle);
    } catch {
      await cleanupActiveReview();
      restorePendingPlayback();
      deps.showStatus('Timing review failed. Using the original subtitle timing.');
      return { action: 'use-original' };
    } finally {
      if (currentRequest === lifecycle) {
        currentRequest = null;
      }
      lifecycle.markSettled();
    }
  }

  async function previewRange(
    request: MediaTimingReviewPreviewRequest,
  ): Promise<MediaTimingReviewActionResult> {
    const current = active;
    if (!current || request.reviewId !== current.payload.reviewId) {
      return staleReviewResult();
    }
    if (!isValidMediaTimingRange(current.payload, request.startTime, request.endTime)) {
      return { ok: false, message: 'The selected preview range is invalid.' };
    }
    try {
      const previewSession = await previewFor(current, request);
      if (active !== current) {
        return staleReviewResult();
      }
      await previewSession.play(request.startTime, request.endTime);
      // Playback spans the whole clip, so the review can end (watchdog, teardown) while
      // it runs; reporting success would leave the modal open on a dead review.
      if (active !== current) {
        return staleReviewResult();
      }
      return { ok: true };
    } catch (error) {
      if (active !== current) {
        return staleReviewResult();
      }
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
      return staleReviewResult();
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
        return staleReviewResult();
      }
      const peaks = await deps.generateWaveform({
        mediaPath: window?.media ?? current.waveformMedia,
        startTime: request.startTime,
        endTime: request.endTime,
        ...(!window && current.audioStreamIndex !== undefined
          ? { audioStreamIndex: current.audioStreamIndex }
          : {}),
      });
      // ffmpeg decoding runs long enough for the review to end underneath it.
      if (active !== current) {
        return staleReviewResult();
      }
      if (peaks.length < 2 || peaks.some((peak) => !Number.isFinite(peak))) {
        return { ok: false, message: 'Timing waveform is unavailable.' };
      }
      return { ok: true, peaks };
    } catch {
      if (active !== current) {
        return staleReviewResult();
      }
      return { ok: false, message: 'Timing waveform is unavailable.' };
    }
  }

  async function getFrame(
    request: MediaTimingReviewFrameRequest,
  ): Promise<MediaTimingReviewFrameResult> {
    const current = active;
    if (!current || request.reviewId !== current.payload.reviewId) return staleReviewResult();
    if (!current.payload.screenshotEnabled || !deps.generateFrame || !current.videoSource) {
      return { ok: false, message: 'Screenshot preview is unavailable for this media.' };
    }
    if (
      !Number.isFinite(request.timestamp) ||
      request.timestamp < 0 ||
      (current.payload.mediaDuration !== undefined &&
        request.timestamp >= current.payload.mediaDuration) ||
      (request.direction !== undefined && request.direction !== -1 && request.direction !== 1)
    ) {
      return { ok: false, message: 'The screenshot time is invalid.' };
    }
    if (current.frameInFlight)
      return { ok: false, message: 'A screenshot preview is already loading.' };
    current.frameInFlight = true;
    try {
      // Reuse the audio window only when it contains this same video source (not split streams).
      const range = {
        startTime: Math.max(0, Math.min(current.payload.timelineStartTime, request.timestamp - 2)),
        endTime: Math.min(
          current.payload.mediaDuration ?? Infinity,
          Math.max(current.payload.timelineEndTime, request.timestamp + 2),
        ),
      };
      const window =
        current.windowSource?.path === current.videoSource.path
          ? await ensureWindow(current, range)
          : null;
      if (active !== current) return staleReviewResult();
      const frame = await deps.generateFrame({
        media: window?.media ?? current.videoSource,
        timestamp: request.timestamp,
        ...(request.direction !== undefined ? { direction: request.direction } : {}),
      });
      if (active !== current) return staleReviewResult();
      return { ok: true, ...frame };
    } catch {
      if (active !== current) return staleReviewResult();
      return {
        ok: false,
        message: 'Screenshot preview unavailable. Try another time or reset to the midpoint.',
      };
    } finally {
      current.frameInFlight = false;
    }
  }

  async function stopPreview(reviewId: string): Promise<MediaTimingReviewActionResult> {
    const current = active;
    if (!current || reviewId !== current.payload.reviewId) {
      return staleReviewResult();
    }
    try {
      const previewSession = current.preview ? await current.preview.session : null;
      if (active !== current) {
        return staleReviewResult();
      }
      await previewSession?.stop();
      if (active !== current) {
        return staleReviewResult();
      }
      return { ok: true };
    } catch (error) {
      if (active !== current) {
        return staleReviewResult();
      }
      return {
        ok: false,
        message: `Could not stop preview: ${error instanceof Error ? error.message : String(error)}`,
      };
    }
  }

  function resolveReview(request: MediaTimingReviewResolveRequest): MediaTimingReviewActionResult {
    const current = active;
    if (!current || request.reviewId !== current.payload.reviewId) {
      return staleReviewResult();
    }
    if (request.decision.action === 'confirm') {
      const { startTime, endTime, text, screenshotTime } = request.decision;
      if (!isValidMediaTimingRange(current.payload, startTime, endTime)) {
        return { ok: false, message: 'The selected timing range is invalid.' };
      }
      if (text !== undefined && (typeof text !== 'string' || text.trim().length === 0)) {
        return { ok: false, message: 'The combined sentence text is invalid.' };
      }
      if (
        screenshotTime !== undefined &&
        (!current.payload.screenshotEnabled ||
          !Number.isFinite(screenshotTime) ||
          screenshotTime < 0 ||
          (current.payload.mediaDuration !== undefined &&
            screenshotTime >= current.payload.mediaDuration))
      ) {
        return { ok: false, message: 'The screenshot time is invalid.' };
      }
    }
    current.resolve(request.decision);
    return { ok: true };
  }

  async function cleanupActiveReview(expected?: ActiveReview): Promise<void> {
    const current = active;
    if (expected && current !== expected) return;
    active = null;
    if (!current) return;
    deps.clearFrameCache?.();
    void current.preview?.session.then((session) => session.dispose()).catch(() => {});
    if (current.restorePlayback && current.mpvClient.connected) {
      current.mpvClient.send({ command: ['set_property', 'pause', 'no'] });
    }
  }

  async function dispose(): Promise<void> {
    const request = currentRequest;
    request?.cancel();
    active?.resolve({ action: 'use-original' });
    await cleanupActiveReview();
    restorePendingPlayback();
    await request?.settled;
  }

  return {
    requestReview,
    previewRange,
    getWaveform,
    getFrame,
    stopPreview,
    resolveReview,
    dispose,
  };
}
