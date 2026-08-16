import type { MediaTimingReviewDecision, MediaTimingReviewOpenPayload } from '../../types/anki';
import type { ModalStateReader, RendererContext } from '../context';

const MINIMUM_CLIP_SECONDS = 0.1;
const FINE_ADJUST_SECONDS = 0.1;
const COARSE_ADJUST_SECONDS = 0.5;
const TIMELINE_EXPANSION_SECONDS = 5;

function clamp(value: number, minimum: number, maximum: number): number {
  return Math.min(maximum, Math.max(minimum, value));
}

export function formatMediaTimingTimestamp(seconds: number, includeMilliseconds = true): string {
  const unitsPerSecond = includeMilliseconds ? 1_000 : 1;
  const totalUnits = Math.max(0, Math.round(seconds * unitsPerSecond));
  const unitsPerMinute = 60 * unitsPerSecond;
  const minutes = Math.floor(totalUnits / unitsPerMinute);
  const remainingUnits = totalUnits % unitsPerMinute;
  const remaining = includeMilliseconds
    ? (remainingUnits / unitsPerSecond).toFixed(3)
    : String(remainingUnits);
  return `${String(minutes).padStart(2, '0')}:${remaining.padStart(includeMilliseconds ? 6 : 2, '0')}`;
}

export function createMediaTimingPreviewRequestGuard() {
  let sequence = 0;
  let activeRequestId: number | null = null;

  return {
    begin(): number | null {
      if (activeRequestId !== null) return null;
      sequence += 1;
      activeRequestId = sequence;
      return activeRequestId;
    },
    invalidate(): void {
      sequence += 1;
      activeRequestId = null;
    },
    isCurrent(requestId: number): boolean {
      return requestId === activeRequestId;
    },
    finish(requestId: number): void {
      if (activeRequestId === requestId) activeRequestId = null;
    },
    isInFlight(): boolean {
      return activeRequestId !== null;
    },
  };
}

export function buildMediaTimingWaveformPath(peaks: number[]): string {
  if (peaks.length < 2) return '';
  const points = peaks.map((peak, index) => ({
    x: (index / (peaks.length - 1)) * 1_000,
    amplitude: clamp(Number.isFinite(peak) ? peak : 0, 0, 1) * 44,
  }));
  const upper = points.map(({ x, amplitude }) => `${x.toFixed(2)} ${(50 - amplitude).toFixed(2)}`);
  const lower = [...points]
    .reverse()
    .map(({ x, amplitude }) => `${x.toFixed(2)} ${(50 + amplitude).toFixed(2)}`);
  return `M ${upper.join(' L ')} L ${lower.join(' L ')} Z`;
}

export function constrainMediaTimingSelection(options: {
  nextStart: number;
  nextEnd: number;
  currentStart: number;
  timelineStart: number;
  timelineEnd: number;
  mediaEnd: number;
  maxMediaDuration: number;
}): { start: number; end: number } {
  let start = clamp(options.nextStart, options.timelineStart, options.timelineEnd);
  let end = clamp(options.nextEnd, options.timelineStart, options.timelineEnd);
  const startMoved = options.nextStart !== options.currentStart;
  if (end - start < MINIMUM_CLIP_SECONDS) {
    if (startMoved) start = end - MINIMUM_CLIP_SECONDS;
    else end = start + MINIMUM_CLIP_SECONDS;
  }
  if (options.maxMediaDuration > 0 && end - start > options.maxMediaDuration) {
    if (startMoved) start = end - options.maxMediaDuration;
    else end = start + options.maxMediaDuration;
  }
  return {
    start: Math.max(0, start),
    end: Math.min(options.mediaEnd, end),
  };
}

/** Maps a pointer position over the timeline track back onto a media timestamp. */
export function mediaTimingTimeFromPointer(options: {
  clientX: number;
  trackLeft: number;
  trackWidth: number;
  timelineStart: number;
  timelineEnd: number;
}): number {
  if (options.trackWidth <= 0) return options.timelineStart;
  const ratio = clamp((options.clientX - options.trackLeft) / options.trackWidth, 0, 1);
  return options.timelineStart + ratio * (options.timelineEnd - options.timelineStart);
}

/** Slides the selection without resizing it, keeping the whole clip inside the visible timeline. */
export function slideMediaTimingSelection(options: {
  nextStart: number;
  span: number;
  timelineStart: number;
  timelineEnd: number;
  mediaEnd: number;
}): { start: number; end: number } {
  const latestEnd = Math.min(options.timelineEnd, options.mediaEnd);
  const maxStart = Math.max(options.timelineStart, latestEnd - options.span);
  const start = clamp(options.nextStart, options.timelineStart, maxStart);
  return { start, end: start + options.span };
}

export function createMediaTimingReviewModal(
  ctx: RendererContext,
  options: {
    modalStateReader: Pick<ModalStateReader, 'isAnyModalOpen'>;
    syncSettingsModalSubtitleSuppression: () => void;
  },
) {
  let payload: MediaTimingReviewOpenPayload | null = null;
  let selectionStart = 0;
  let selectionEnd = 0;
  let timelineStart = 0;
  let timelineEnd = 0;
  let resolveInFlight = false;
  let previewPlaying = false;
  let previewTimer: ReturnType<typeof setTimeout> | null = null;
  let waveformTimer: ReturnType<typeof setTimeout> | null = null;
  let waveformSequence = 0;
  let drag: {
    edge: 'start' | 'end' | 'both';
    pointerId: number;
    trackLeft: number;
    trackWidth: number;
    grabOffset: number;
  } | null = null;
  const previewRequest = createMediaTimingPreviewRequestGuard();

  function setStatus(message: string, isError = false): void {
    ctx.dom.mediaTimingReviewStatus.textContent = message;
    ctx.dom.mediaTimingReviewStatus.classList.toggle('is-error', isError);
  }

  function clearPreviewTimer(): void {
    if (previewTimer !== null) clearTimeout(previewTimer);
    previewTimer = null;
  }

  /** Drives the play button label plus the playhead sweep that mirrors the hidden audio player. */
  function setPreviewPlaying(playing: boolean): void {
    previewPlaying = playing;
    ctx.dom.mediaTimingReviewPlayLabel.textContent = playing ? 'Stop preview' : 'Play selection';
    ctx.dom.mediaTimingReviewPlay.classList.toggle('is-playing', playing);
    clearPreviewTimer();
    const track = ctx.dom.mediaTimingReviewSelectionTrack;
    track.classList.remove('is-previewing');
    if (!playing) return;
    const clipSeconds = Math.max(MINIMUM_CLIP_SECONDS, selectionEnd - selectionStart);
    track.style.setProperty('--playhead-duration', `${clipSeconds}s`);
    void track.offsetWidth;
    track.classList.add('is-previewing');
    previewTimer = setTimeout(() => stopPreview(), clipSeconds * 1000);
  }

  /** Callers that need to report a failure set their own status after stopping the preview. */
  function stopPreview(): void {
    previewRequest.invalidate();
    setPreviewPlaying(false);
    setStatus('');
    if (payload) {
      void window.electronAPI.stopMediaTimingReviewPreview(payload.reviewId).catch(() => {});
    }
  }

  function renderHandle(handle: HTMLDivElement, value: number): void {
    handle.setAttribute('aria-valuemin', timelineStart.toFixed(3));
    handle.setAttribute('aria-valuemax', timelineEnd.toFixed(3));
    handle.setAttribute('aria-valuenow', value.toFixed(3));
    handle.setAttribute('aria-valuetext', formatMediaTimingTimestamp(value));
  }

  function renderSelection(): void {
    const span = Math.max(MINIMUM_CLIP_SECONDS, timelineEnd - timelineStart);
    const startPercent = ((selectionStart - timelineStart) / span) * 100;
    const endPercent = ((selectionEnd - timelineStart) / span) * 100;
    const originalStartPercent = payload
      ? ((payload.originalStartTime - timelineStart) / span) * 100
      : 0;
    const originalEndPercent = payload
      ? ((payload.originalEndTime - timelineStart) / span) * 100
      : 0;
    ctx.dom.mediaTimingReviewSelectionTrack.style.setProperty(
      '--selection-start',
      `${clamp(startPercent, 0, 100)}%`,
    );
    ctx.dom.mediaTimingReviewSelectionTrack.style.setProperty(
      '--selection-end',
      `${clamp(endPercent, 0, 100)}%`,
    );
    ctx.dom.mediaTimingReviewSelectionTrack.style.setProperty(
      '--original-start',
      `${clamp(originalStartPercent, 0, 100)}%`,
    );
    ctx.dom.mediaTimingReviewSelectionTrack.style.setProperty(
      '--original-end',
      `${clamp(originalEndPercent, 0, 100)}%`,
    );
    renderHandle(ctx.dom.mediaTimingReviewStartHandle, selectionStart);
    renderHandle(ctx.dom.mediaTimingReviewEndHandle, selectionEnd);
    ctx.dom.mediaTimingReviewStartValue.textContent = formatMediaTimingTimestamp(selectionStart);
    ctx.dom.mediaTimingReviewEndValue.textContent = formatMediaTimingTimestamp(selectionEnd);
    ctx.dom.mediaTimingReviewDuration.textContent = `${(selectionEnd - selectionStart).toFixed(2)}s`;
    ctx.dom.mediaTimingReviewTimelineStart.textContent = formatMediaTimingTimestamp(
      timelineStart,
      false,
    );
    ctx.dom.mediaTimingReviewTimelineEnd.textContent = formatMediaTimingTimestamp(
      timelineEnd,
      false,
    );
    ctx.dom.mediaTimingReviewShowEarlier.disabled = timelineStart <= 0;
    ctx.dom.mediaTimingReviewShowLater.disabled =
      payload?.mediaDuration !== undefined && timelineEnd >= payload.mediaDuration;
  }

  function setWaveformState(state: 'loading' | 'ready' | 'unavailable'): void {
    ctx.dom.mediaTimingReviewSelectionTrack.classList.toggle('is-loading', state === 'loading');
    ctx.dom.mediaTimingReviewSelectionTrack.classList.toggle(
      'is-waveform-unavailable',
      state === 'unavailable',
    );
    ctx.dom.mediaTimingReviewWaveformLabel.textContent =
      state === 'loading'
        ? 'isolating dialogue...'
        : state === 'ready'
          ? 'speech-weighted waveform'
          : 'waveform unavailable';
  }

  async function loadWaveform(): Promise<void> {
    if (!payload || !ctx.state.mediaTimingReviewModalOpen) return;
    waveformSequence += 1;
    const sequence = waveformSequence;
    const reviewId = payload.reviewId;
    const startTime = timelineStart;
    const endTime = timelineEnd;
    setWaveformState('loading');
    ctx.dom.mediaTimingReviewWaveformPath.setAttribute('d', '');
    try {
      const result = await window.electronAPI.getMediaTimingReviewWaveform({
        reviewId,
        startTime,
        endTime,
      });
      if (
        sequence !== waveformSequence ||
        payload?.reviewId !== reviewId ||
        timelineStart !== startTime ||
        timelineEnd !== endTime
      ) {
        return;
      }
      const path = result.ok ? buildMediaTimingWaveformPath(result.peaks ?? []) : '';
      if (!path) {
        setWaveformState('unavailable');
        return;
      }
      ctx.dom.mediaTimingReviewWaveformPath.setAttribute('d', path);
      setWaveformState('ready');
    } catch {
      if (sequence === waveformSequence && payload?.reviewId === reviewId) {
        setWaveformState('unavailable');
      }
    }
  }

  function queueWaveformLoad(delayMs = 0): void {
    if (waveformTimer !== null) clearTimeout(waveformTimer);
    waveformSequence += 1;
    if (payload && ctx.state.mediaTimingReviewModalOpen) {
      ctx.dom.mediaTimingReviewWaveformPath.setAttribute('d', '');
      setWaveformState('loading');
    }
    waveformTimer = setTimeout(() => {
      waveformTimer = null;
      void loadWaveform();
    }, delayMs);
  }

  function updateSelection(nextStart: number, nextEnd: number): void {
    if (!payload) return;
    const mediaEnd = payload.mediaDuration ?? Number.POSITIVE_INFINITY;
    const nextSelection = constrainMediaTimingSelection({
      nextStart,
      nextEnd,
      currentStart: selectionStart,
      timelineStart,
      timelineEnd,
      mediaEnd,
      maxMediaDuration: payload.maxMediaDuration,
    });
    selectionStart = nextSelection.start;
    selectionEnd = nextSelection.end;
    if (previewPlaying || previewRequest.isInFlight()) stopPreview();
    setStatus('');
    renderSelection();
  }

  /** Shifts the whole clip without changing its length. */
  function moveSelection(nextStart: number): void {
    if (!payload) return;
    const slid = slideMediaTimingSelection({
      nextStart,
      span: selectionEnd - selectionStart,
      timelineStart,
      timelineEnd,
      mediaEnd: payload.mediaDuration ?? Number.POSITIVE_INFINITY,
    });
    updateSelection(slid.start, slid.end);
  }

  function setEdge(edge: 'start' | 'end', value: number): void {
    if (edge === 'start') updateSelection(value, selectionEnd);
    else updateSelection(selectionStart, value);
  }

  function applyDrag(clientX: number): void {
    if (!drag) return;
    const pointerTime = mediaTimingTimeFromPointer({
      clientX,
      trackLeft: drag.trackLeft,
      trackWidth: drag.trackWidth,
      timelineStart,
      timelineEnd,
    });
    const time = pointerTime - drag.grabOffset;
    if (drag.edge === 'both') moveSelection(time);
    else setEdge(drag.edge, time);
  }

  /**
   * Grabbing a handle drags that edge, grabbing the highlighted clip slides the whole selection,
   * and pressing anywhere else snaps the nearest edge to that point and keeps dragging it.
   */
  function beginDrag(event: PointerEvent): void {
    if (!payload || resolveInFlight || drag !== null || event.button !== 0) return;
    const track = ctx.dom.mediaTimingReviewSelectionTrack;
    const rect = track.getBoundingClientRect();
    const pointerTime = mediaTimingTimeFromPointer({
      clientX: event.clientX,
      trackLeft: rect.left,
      trackWidth: rect.width,
      timelineStart,
      timelineEnd,
    });
    const target = event.target;
    let edge: 'start' | 'end' | 'both';
    let grabOffset: number;
    if (target === ctx.dom.mediaTimingReviewStartHandle) {
      edge = 'start';
      grabOffset = pointerTime - selectionStart;
    } else if (target === ctx.dom.mediaTimingReviewEndHandle) {
      edge = 'end';
      grabOffset = pointerTime - selectionEnd;
    } else if (target === ctx.dom.mediaTimingReviewSelectedRange) {
      edge = 'both';
      grabOffset = pointerTime - selectionStart;
    } else {
      edge =
        Math.abs(pointerTime - selectionStart) <= Math.abs(pointerTime - selectionEnd)
          ? 'start'
          : 'end';
      grabOffset = 0;
    }
    event.preventDefault();
    drag = {
      edge,
      pointerId: event.pointerId,
      trackLeft: rect.left,
      trackWidth: rect.width,
      grabOffset,
    };
    track.setPointerCapture(event.pointerId);
    ctx.dom.mediaTimingReviewModal.classList.add(edge === 'both' ? 'is-sliding' : 'is-scrubbing');
    if (edge !== 'both') {
      const handle =
        edge === 'start'
          ? ctx.dom.mediaTimingReviewStartHandle
          : ctx.dom.mediaTimingReviewEndHandle;
      handle.focus();
      if (grabOffset === 0) applyDrag(event.clientX);
    }
  }

  function endDrag(event: PointerEvent): void {
    if (!drag || drag.pointerId !== event.pointerId) return;
    const track = ctx.dom.mediaTimingReviewSelectionTrack;
    if (track.hasPointerCapture(event.pointerId)) track.releasePointerCapture(event.pointerId);
    drag = null;
    ctx.dom.mediaTimingReviewModal.classList.remove('is-scrubbing', 'is-sliding');
  }

  function cancelDrag(): void {
    drag = null;
    ctx.dom.mediaTimingReviewModal.classList.remove('is-scrubbing', 'is-sliding');
  }

  function handleEdgeKeydown(event: KeyboardEvent, edge: 'start' | 'end'): void {
    const step = event.shiftKey ? COARSE_ADJUST_SECONDS : FINE_ADJUST_SECONDS;
    const current = edge === 'start' ? selectionStart : selectionEnd;
    if (event.key === 'ArrowLeft' || event.key === 'ArrowDown') setEdge(edge, current - step);
    else if (event.key === 'ArrowRight' || event.key === 'ArrowUp') setEdge(edge, current + step);
    else if (event.key === 'Home') setEdge(edge, timelineStart);
    else if (event.key === 'End') setEdge(edge, timelineEnd);
    else return;
    event.preventDefault();
  }

  function showEditor(): void {
    ctx.dom.mediaTimingReviewCancelStep.classList.add('hidden');
    ctx.dom.mediaTimingReviewEditor.classList.remove('hidden');
    ctx.dom.mediaTimingReviewCancel.classList.remove('hidden');
  }

  function requestCancel(): void {
    if (!ctx.state.mediaTimingReviewModalOpen || !payload || resolveInFlight) return;
    stopPreview();
    ctx.dom.mediaTimingReviewEditor.classList.add('hidden');
    ctx.dom.mediaTimingReviewCancel.classList.add('hidden');
    ctx.dom.mediaTimingReviewCancelStep.classList.remove('hidden');
    ctx.dom.mediaTimingReviewCancelMessage.textContent =
      payload.noteId !== undefined
        ? 'Keep editing, finish this card with its original timing, or delete the card.'
        : 'Keep editing, finish this card with its original timing, or do not create it.';
    ctx.dom.mediaTimingReviewDiscard.textContent =
      payload.noteId !== undefined ? 'Delete card' : "Don't create card";
    ctx.dom.mediaTimingReviewCancelBack.focus();
  }

  function closeResolvedReview(): void {
    if (!ctx.state.mediaTimingReviewModalOpen) return;
    cancelDrag();
    clearPreviewTimer();
    if (waveformTimer !== null) clearTimeout(waveformTimer);
    waveformTimer = null;
    waveformSequence += 1;
    previewPlaying = false;
    ctx.state.mediaTimingReviewModalOpen = false;
    ctx.dom.mediaTimingReviewModal.classList.add('hidden');
    ctx.dom.mediaTimingReviewModal.setAttribute('aria-hidden', 'true');
    window.electronAPI.notifyOverlayModalClosed('media-timing-review');
    options.syncSettingsModalSubtitleSuppression();
    payload = null;
    if (!options.modalStateReader.isAnyModalOpen()) {
      ctx.dom.overlay.classList.remove('interactive');
      if (ctx.platform.shouldToggleMouseIgnore) {
        window.electronAPI.setIgnoreMouseEvents(true, { forward: true });
      }
    }
  }

  async function resolveReview(decision: MediaTimingReviewDecision): Promise<void> {
    if (!payload || resolveInFlight) return;
    resolveInFlight = true;
    stopPreview();
    const controls = ctx.dom.mediaTimingReviewModal.querySelectorAll<HTMLButtonElement>('button');
    controls.forEach((button) => {
      button.disabled = true;
    });
    try {
      const result = await window.electronAPI.resolveMediaTimingReview({
        reviewId: payload.reviewId,
        decision,
      });
      if (!result.ok) {
        setStatus(result.message ?? 'The timing review could not be resolved.', true);
        showEditor();
        return;
      }
      closeResolvedReview();
    } catch (error) {
      setStatus(error instanceof Error ? error.message : String(error), true);
      showEditor();
    } finally {
      resolveInFlight = false;
      controls.forEach((button) => {
        button.disabled = false;
      });
      renderSelection();
    }
  }

  async function togglePreview(): Promise<void> {
    if (!payload || resolveInFlight) return;
    if (previewPlaying) {
      stopPreview();
      return;
    }
    const requestId = previewRequest.begin();
    if (requestId === null) return;
    const requestedReviewId = payload.reviewId;
    setStatus('Starting audio preview...');
    try {
      const result = await window.electronAPI.previewMediaTimingReview({
        reviewId: requestedReviewId,
        startTime: selectionStart,
        endTime: selectionEnd,
      });
      if (!previewRequest.isCurrent(requestId) || payload?.reviewId !== requestedReviewId) {
        if (result.ok && !previewRequest.isInFlight() && !previewPlaying) {
          void window.electronAPI.stopMediaTimingReviewPreview(requestedReviewId).catch(() => {});
        }
        return;
      }
      if (!result.ok) {
        setStatus(result.message ?? 'Audio preview is unavailable.', true);
        setPreviewPlaying(false);
        return;
      }
      setStatus('Previewing in the hidden audio player.');
      setPreviewPlaying(true);
    } catch (error) {
      if (!previewRequest.isCurrent(requestId) || payload?.reviewId !== requestedReviewId) return;
      setStatus(
        `Audio preview unavailable: ${error instanceof Error ? error.message : String(error)}`,
        true,
      );
      setPreviewPlaying(false);
    } finally {
      previewRequest.finish(requestId);
    }
  }

  function openMediaTimingReviewModal(nextPayload: MediaTimingReviewOpenPayload): void {
    previewRequest.invalidate();
    cancelDrag();
    payload = nextPayload;
    selectionStart = nextPayload.selectionStartTime;
    selectionEnd = nextPayload.selectionEndTime;
    timelineStart = nextPayload.timelineStartTime;
    timelineEnd = nextPayload.timelineEndTime;
    resolveInFlight = false;
    setPreviewPlaying(false);
    ctx.dom.mediaTimingReviewKind.textContent =
      nextPayload.kind === 'word'
        ? 'Word card'
        : nextPayload.kind === 'audio'
          ? 'Audio card'
          : 'Sentence card';
    ctx.dom.mediaTimingReviewKind.dataset.kind = nextPayload.kind;
    ctx.dom.mediaTimingReviewText.textContent = nextPayload.text;
    ctx.dom.mediaTimingReviewDiscard.textContent =
      nextPayload.noteId !== undefined ? 'Delete card' : "Don't create card";
    setStatus('');
    showEditor();
    renderSelection();
    ctx.state.mediaTimingReviewModalOpen = true;
    options.syncSettingsModalSubtitleSuppression();
    ctx.dom.overlay.classList.add('interactive');
    if (ctx.platform.shouldToggleMouseIgnore) window.electronAPI.setIgnoreMouseEvents(false);
    ctx.dom.mediaTimingReviewModal.classList.remove('hidden');
    ctx.dom.mediaTimingReviewModal.setAttribute('aria-hidden', 'false');
    window.electronAPI.notifyOverlayModalOpened('media-timing-review');
    ctx.dom.mediaTimingReviewStartHandle.focus();
    queueWaveformLoad();
  }

  function expandTimeline(direction: 'earlier' | 'later'): void {
    if (!payload) return;
    if (direction === 'earlier') {
      timelineStart = Math.max(0, timelineStart - TIMELINE_EXPANSION_SECONDS);
    } else {
      timelineEnd = Math.min(
        payload.mediaDuration ?? Number.POSITIVE_INFINITY,
        timelineEnd + TIMELINE_EXPANSION_SECONDS,
      );
    }
    renderSelection();
    queueWaveformLoad(120);
  }

  function handleMediaTimingReviewKeydown(event: KeyboardEvent): boolean {
    if (!ctx.state.mediaTimingReviewModalOpen) return false;
    if (event.key === 'Escape') {
      event.preventDefault();
      if (!ctx.dom.mediaTimingReviewCancelStep.classList.contains('hidden')) showEditor();
      else requestCancel();
      return true;
    }
    if (
      event.code === 'Space' &&
      event.target instanceof Element &&
      !event.target.closest('button')
    ) {
      event.preventDefault();
      void togglePreview();
      return true;
    }
    if (
      event.key === 'Enter' &&
      ctx.dom.mediaTimingReviewCancelStep.classList.contains('hidden') &&
      !(event.target instanceof Element && event.target.closest('button'))
    ) {
      event.preventDefault();
      void resolveReview({ action: 'confirm', startTime: selectionStart, endTime: selectionEnd });
      return true;
    }
    return false;
  }

  function wireDomEvents(): void {
    const track = ctx.dom.mediaTimingReviewSelectionTrack;
    track.addEventListener('pointerdown', beginDrag);
    track.addEventListener('pointermove', (event) => {
      if (drag?.pointerId === event.pointerId) applyDrag(event.clientX);
    });
    track.addEventListener('pointerup', endDrag);
    track.addEventListener('pointercancel', endDrag);
    ctx.dom.mediaTimingReviewStartHandle.addEventListener('keydown', (event) =>
      handleEdgeKeydown(event, 'start'),
    );
    ctx.dom.mediaTimingReviewEndHandle.addEventListener('keydown', (event) =>
      handleEdgeKeydown(event, 'end'),
    );
    ctx.dom.mediaTimingReviewShowEarlier.addEventListener('click', () => expandTimeline('earlier'));
    ctx.dom.mediaTimingReviewShowLater.addEventListener('click', () => expandTimeline('later'));
    ctx.dom.mediaTimingReviewStartBack.addEventListener('click', () =>
      updateSelection(selectionStart - FINE_ADJUST_SECONDS, selectionEnd),
    );
    ctx.dom.mediaTimingReviewStartForward.addEventListener('click', () =>
      updateSelection(selectionStart + FINE_ADJUST_SECONDS, selectionEnd),
    );
    ctx.dom.mediaTimingReviewEndBack.addEventListener('click', () =>
      updateSelection(selectionStart, selectionEnd - FINE_ADJUST_SECONDS),
    );
    ctx.dom.mediaTimingReviewEndForward.addEventListener('click', () =>
      updateSelection(selectionStart, selectionEnd + FINE_ADJUST_SECONDS),
    );
    ctx.dom.mediaTimingReviewPlay.addEventListener('click', () => void togglePreview());
    ctx.dom.mediaTimingReviewReset.addEventListener('click', () => {
      if (!payload) return;
      updateSelection(payload.selectionStartTime, payload.selectionEndTime);
    });
    ctx.dom.mediaTimingReviewCancel.addEventListener('click', requestCancel);
    ctx.dom.mediaTimingReviewCancelBack.addEventListener('click', showEditor);
    ctx.dom.mediaTimingReviewUseOriginal.addEventListener(
      'click',
      () => void resolveReview({ action: 'use-original' }),
    );
    ctx.dom.mediaTimingReviewDiscard.addEventListener(
      'click',
      () => void resolveReview({ action: 'discard' }),
    );
    ctx.dom.mediaTimingReviewConfirm.addEventListener(
      'click',
      () =>
        void resolveReview({ action: 'confirm', startTime: selectionStart, endTime: selectionEnd }),
    );
  }

  return {
    openMediaTimingReviewModal,
    requestCancel,
    handleMediaTimingReviewKeydown,
    wireDomEvents,
  };
}
