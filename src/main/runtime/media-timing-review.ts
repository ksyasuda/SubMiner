import { randomUUID } from 'crypto';
import type {
  MediaTimingReviewActionResult,
  MediaTimingReviewDecision,
  MediaTimingReviewOpenPayload,
  MediaTimingReviewPreviewRequest,
  MediaTimingReviewRequest,
  MediaTimingReviewResolveRequest,
  MediaTimingReviewWaveformRequest,
  MediaTimingReviewWaveformResult,
} from '../../types/anki';
import type { SpeechWaveformOptions } from '../../core/services/media-timing-waveform';

const INITIAL_TIMELINE_MARGIN_SECONDS = 2;

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
  }): Promise<void>;
  play(startTime: number, endTime: number): Promise<void>;
  stop(): Promise<void>;
  dispose(): void;
}

interface ActiveReview {
  payload: MediaTimingReviewOpenPayload;
  mediaPath: string;
  audioStreamIndex?: number;
  mpvClient: ReviewMpvClient;
  restorePlayback: boolean;
  preview: Promise<PreviewSession>;
  resolve: (decision: MediaTimingReviewDecision) => void;
}

export interface MediaTimingReviewRuntimeDeps {
  getMpvClient: () => ReviewMpvClient | null;
  getCurrentMediaPath: () => string | null;
  getMpvExecutablePath: () => string;
  createPreviewSession: () => PreviewSession;
  generateWaveform: (options: SpeechWaveformOptions) => Promise<number[]>;
  openModal: (payload: MediaTimingReviewOpenPayload) => Promise<boolean>;
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

export function buildMediaTimingReviewPayload(
  request: MediaTimingReviewRequest,
  options: { reviewId: string; mediaDuration?: number },
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

  async function runReview(request: MediaTimingReviewRequest): Promise<MediaTimingReviewDecision> {
    const mpvClient = deps.getMpvClient();
    const mediaPath =
      deps.getCurrentMediaPath()?.trim() || mpvClient?.currentVideoPath?.trim() || '';
    if (!mpvClient?.connected || !mediaPath) {
      deps.showStatus('Timing review unavailable. Using the original subtitle timing.');
      return { action: 'use-original' };
    }

    const [pauseRaw, durationRaw, audioTrackRaw, volumeRaw] = await Promise.all([
      mpvClient.requestProperty?.('pause').catch(() => null) ?? null,
      mpvClient.requestProperty?.('duration').catch(() => null) ?? null,
      mpvClient.requestProperty?.('aid').catch(() => null) ?? null,
      mpvClient.requestProperty?.('volume').catch(() => null) ?? null,
    ]);
    const pauseState = booleanProperty(pauseRaw);
    mpvClient.send({ command: ['set_property', 'pause', 'yes'] });
    pendingPauseRestore = pauseState === false ? mpvClient : null;

    const payload = buildMediaTimingReviewPayload(request, {
      reviewId: randomUUID(),
      mediaDuration: finiteNumber(durationRaw) ?? undefined,
    });
    const previewSession = deps.createPreviewSession();
    const preview = previewSession
      .start({
        mediaPath,
        executablePath: deps.getMpvExecutablePath(),
        audioTrackId: finiteNumber(audioTrackRaw) ?? undefined,
        volume: finiteNumber(volumeRaw) ?? undefined,
      })
      .then(() => previewSession)
      .catch((error) => {
        previewSession.dispose();
        throw error;
      });
    void preview.catch(() => {});

    let resolveDecision!: (decision: MediaTimingReviewDecision) => void;
    const decisionPromise = new Promise<MediaTimingReviewDecision>((resolve) => {
      resolveDecision = resolve;
    });
    active = {
      payload,
      mediaPath,
      ...(mpvClient.currentAudioStreamIndex !== null &&
      mpvClient.currentAudioStreamIndex !== undefined
        ? { audioStreamIndex: mpvClient.currentAudioStreamIndex }
        : {}),
      mpvClient,
      restorePlayback: pendingPauseRestore === mpvClient,
      preview,
      resolve: resolveDecision,
    };
    pendingPauseRestore = null;

    const opened = await deps.openModal(payload).catch(() => false);
    if (!opened) {
      await cleanupActiveReview();
      deps.showStatus('Timing review could not open. Using the original subtitle timing.');
      return { action: 'use-original' };
    }

    const decision = await decisionPromise;
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
    if (
      !Number.isFinite(request.startTime) ||
      !Number.isFinite(request.endTime) ||
      request.startTime < 0 ||
      request.endTime <= request.startTime ||
      (current.payload.maxMediaDuration > 0 &&
        request.endTime - request.startTime > current.payload.maxMediaDuration + 0.001) ||
      (current.payload.mediaDuration !== undefined &&
        request.endTime > current.payload.mediaDuration + 0.001)
    ) {
      return { ok: false, message: 'The selected preview range is invalid.' };
    }
    try {
      const previewSession = await current.preview;
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
      const peaks = await deps.generateWaveform({
        mediaPath: current.mediaPath,
        startTime: request.startTime,
        endTime: request.endTime,
        ...(current.audioStreamIndex !== undefined
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
      const previewSession = await current.preview;
      await previewSession.stop();
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
      const { startTime, endTime } = request.decision;
      if (
        !Number.isFinite(startTime) ||
        !Number.isFinite(endTime) ||
        startTime < 0 ||
        endTime <= startTime ||
        (current.payload.maxMediaDuration > 0 &&
          endTime - startTime > current.payload.maxMediaDuration + 0.001) ||
        (current.payload.mediaDuration !== undefined &&
          endTime > current.payload.mediaDuration + 0.001)
      ) {
        return { ok: false, message: 'The selected timing range is invalid.' };
      }
    }
    current.resolve(request.decision);
    return { ok: true };
  }

  async function cleanupActiveReview(): Promise<void> {
    const current = active;
    active = null;
    if (!current) return;
    void current.preview.then((session) => session.dispose()).catch(() => {});
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
