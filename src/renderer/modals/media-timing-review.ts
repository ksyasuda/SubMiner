import type { MediaTimingReviewDecision, MediaTimingReviewOpenPayload } from '../../types/anki';
import type { ModalStateReader, RendererContext } from '../context';

const MINIMUM_CLIP_SECONDS = 0.1;
const FINE_ADJUST_SECONDS = 0.1;
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
  const previewRequest = createMediaTimingPreviewRequestGuard();

  function setStatus(message: string, isError = false): void {
    ctx.dom.mediaTimingReviewStatus.textContent = message;
    ctx.dom.mediaTimingReviewStatus.classList.toggle('is-error', isError);
  }

  function clearPreviewTimer(): void {
    if (previewTimer !== null) clearTimeout(previewTimer);
    previewTimer = null;
  }

  function setPreviewPlaying(playing: boolean): void {
    previewPlaying = playing;
    ctx.dom.mediaTimingReviewPlay.textContent = playing ? 'Stop preview' : 'Play selection';
    ctx.dom.mediaTimingReviewPlay.classList.toggle('is-playing', playing);
    clearPreviewTimer();
    if (playing) {
      previewTimer = setTimeout(() => stopPreview(), (selectionEnd - selectionStart) * 1000);
    }
  }

  function stopPreview(): void {
    previewRequest.invalidate();
    setPreviewPlaying(false);
    if (payload) {
      void window.electronAPI.stopMediaTimingReviewPreview(payload.reviewId).catch(() => {});
    }
  }

  function renderSelection(): void {
    const span = Math.max(MINIMUM_CLIP_SECONDS, timelineEnd - timelineStart);
    const startPercent = ((selectionStart - timelineStart) / span) * 100;
    const endPercent = ((selectionEnd - timelineStart) / span) * 100;
    ctx.dom.mediaTimingReviewSelectionTrack.style.setProperty(
      '--selection-start',
      `${clamp(startPercent, 0, 100)}%`,
    );
    ctx.dom.mediaTimingReviewSelectionTrack.style.setProperty(
      '--selection-end',
      `${clamp(endPercent, 0, 100)}%`,
    );
    ctx.dom.mediaTimingReviewStartRange.min = String(timelineStart);
    ctx.dom.mediaTimingReviewStartRange.max = String(timelineEnd);
    ctx.dom.mediaTimingReviewStartRange.value = String(selectionStart);
    ctx.dom.mediaTimingReviewEndRange.min = String(timelineStart);
    ctx.dom.mediaTimingReviewEndRange.max = String(timelineEnd);
    ctx.dom.mediaTimingReviewEndRange.value = String(selectionEnd);
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
    clearPreviewTimer();
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
    ctx.dom.mediaTimingReviewPlay.focus();
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
    ctx.dom.mediaTimingReviewStartRange.addEventListener('input', () => {
      updateSelection(Number(ctx.dom.mediaTimingReviewStartRange.value), selectionEnd);
    });
    ctx.dom.mediaTimingReviewEndRange.addEventListener('input', () => {
      updateSelection(selectionStart, Number(ctx.dom.mediaTimingReviewEndRange.value));
    });
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
