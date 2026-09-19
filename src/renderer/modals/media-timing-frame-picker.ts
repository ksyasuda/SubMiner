import type { MediaTimingReviewFrameRequest, MediaTimingReviewFrameResult } from '../../types/anki';

export interface MediaTimingFramePickerState {
  enabled: boolean;
  manual: boolean;
  loading: boolean;
  timestamp?: number;
  requestedTime?: number;
  dataUrl?: string;
  message: string;
  blockConfirm: boolean;
}

/** Coalesces scrubbing into one active extraction and the latest requested frame. */
export function createMediaTimingFramePicker(options: {
  load: (request: MediaTimingReviewFrameRequest) => Promise<MediaTimingReviewFrameResult>;
  onChange: (state: MediaTimingFramePickerState) => void;
  onStale: () => void;
  debounceMs?: number;
}) {
  let reviewId: string | null = null;
  let midpoint = 0;
  let sequence = 0;
  let inFlight = false;
  let pending: (MediaTimingReviewFrameRequest & { sequence: number }) | null = null;
  let timer: ReturnType<typeof setTimeout> | null = null;
  let state: MediaTimingFramePickerState = emptyState();

  function emptyState(): MediaTimingFramePickerState {
    return { enabled: false, manual: false, loading: false, message: '', blockConfirm: false };
  }

  function publish(): void {
    options.onChange({ ...state });
  }

  async function drain(): Promise<void> {
    if (inFlight || !pending) return;
    const request = pending;
    pending = null;
    inFlight = true;
    try {
      const result = await options.load(request);
      if (reviewId !== request.reviewId || sequence !== request.sequence) return;
      if (result.stale) {
        options.onStale();
        return;
      }
      if (
        !result.ok ||
        !Number.isFinite(result.timestamp) ||
        !result.dataUrl?.startsWith('data:image/jpeg;base64,')
      ) {
        throw new Error(result.message ?? 'Screenshot preview is unavailable.');
      }
      state.timestamp = result.timestamp;
      state.dataUrl = result.dataUrl;
      state.blockConfirm = false;
      state.message = state.manual
        ? 'Selected frame stays fixed when you trim audio.'
        : 'Following the audio midpoint.';
    } catch (error) {
      if (reviewId !== request.reviewId || sequence !== request.sequence) return;
      state.message = error instanceof Error ? error.message : 'Screenshot preview is unavailable.';
      state.blockConfirm = state.manual;
    } finally {
      inFlight = false;
      if (reviewId === request.reviewId && sequence === request.sequence) {
        state.loading = false;
        publish();
      }
      if (timer === null) void drain();
    }
  }

  function request(timestamp: number, direction?: -1 | 1): void {
    if (!reviewId || !state.enabled) return;
    sequence += 1;
    pending = { reviewId, timestamp, ...(direction ? { direction } : {}), sequence };
    state.requestedTime = timestamp;
    state.loading = true;
    state.blockConfirm = state.manual;
    state.message = 'Loading screenshot…';
    publish();
    if (timer !== null) clearTimeout(timer);
    timer = setTimeout(() => {
      timer = null;
      void drain();
    }, options.debounceMs ?? 120);
  }

  function close(): void {
    sequence += 1;
    reviewId = null;
    pending = null;
    if (timer !== null) clearTimeout(timer);
    timer = null;
    state = emptyState();
    publish();
  }

  return {
    open(id: string, enabled: boolean, time: number) {
      close();
      reviewId = id;
      midpoint = time;
      state.enabled = enabled;
      publish();
      if (enabled) request(time);
    },
    updateMidpoint(time: number) {
      if (midpoint === time) return;
      midpoint = time;
      if (!state.manual) request(time);
    },
    choose(time: number, direction?: -1 | 1) {
      if (!Number.isFinite(time) || time < 0) return;
      state.manual = true;
      request(time, direction);
    },
    reset() {
      state.manual = false;
      request(midpoint);
    },
    close,
    getState: () => ({ ...state }),
    getScreenshotTime: () => (state.manual && !state.blockConfirm ? state.timestamp : undefined),
  };
}
