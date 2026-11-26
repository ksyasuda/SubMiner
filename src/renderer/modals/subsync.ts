import type {
  SubsyncEngine,
  SubsyncManualPayload,
  SubsyncManualRunRequest,
  SubsyncReferenceCandidate,
} from '../../types';
import type { ModalStateReader, RendererContext } from '../context';

const VIDEO_REFERENCE_VALUE = 'video';
const TRACK_PREFIX = 'track:';
const FILE_PREFIX = 'file:';

interface SubsyncSelection {
  engine: SubsyncEngine;
  referenceValue: string;
  targetTrackId: number | null;
  stripSigns: boolean;
  allowSplits: boolean;
  offsetMs: number;
}

export function createSubsyncModal(
  ctx: RendererContext,
  options: {
    modalStateReader: Pick<ModalStateReader, 'isAnyModalOpen'>;
    syncSettingsModalSubtitleSuppression: () => void;
  },
) {
  let currentPayload: SubsyncManualPayload | null = null;
  let customFileCandidates: SubsyncReferenceCandidate[] = [];

  function setSubsyncStatus(message: string, isError = false): void {
    ctx.dom.subsyncStatus.textContent = message;
    ctx.dom.subsyncStatus.classList.toggle('error', isError);
  }

  function hasBuiltinEngineUi(): boolean {
    return Boolean(ctx.dom.subsyncEngineBuiltin);
  }

  function hasBuiltinReference(): boolean {
    if (!currentPayload) return false;
    if (currentPayload.builtinAvailable === false) return false;
    return (
      currentPayload.audioReferenceAvailable === true ||
      currentPayload.videoReferenceAvailable ||
      currentPayload.subtitleTracks.length > 1 ||
      (currentPayload.referenceCandidates?.length ?? 0) > 0 ||
      customFileCandidates.length > 0 ||
      currentPayload.canPickReferenceFile === true
    );
  }

  function hasAlassReference(): boolean {
    if (!currentPayload) return false;
    if (currentPayload.alassAvailable === false) return false;
    return (
      currentPayload.videoReferenceAvailable ||
      currentPayload.subtitleTracks.length > 1 ||
      (currentPayload.referenceCandidates?.some((c) => c.kind !== 'video') ?? false) ||
      customFileCandidates.length > 0
    );
  }

  function hasOffsetEngine(): boolean {
    return Boolean(ctx.dom.subsyncEngineOffset) && ctx.state.subsyncSubtitleTracks.length > 0;
  }

  function getSelectedEngine(): SubsyncEngine | null {
    if (ctx.dom.subsyncEngineBuiltin?.checked) return 'builtin';
    if (ctx.dom.subsyncEngineAlass.checked) return 'alass';
    if (ctx.dom.subsyncEngineFfsubsync.checked) return 'ffsubsync';
    if (ctx.dom.subsyncEngineOffset?.checked) return 'offset';
    return null;
  }

  function setSelectedEngine(engine: SubsyncEngine | null): void {
    if (ctx.dom.subsyncEngineBuiltin) {
      ctx.dom.subsyncEngineBuiltin.checked = engine === 'builtin';
    }
    ctx.dom.subsyncEngineAlass.checked = engine === 'alass';
    ctx.dom.subsyncEngineFfsubsync.checked = engine === 'ffsubsync';
    if (ctx.dom.subsyncEngineOffset) {
      ctx.dom.subsyncEngineOffset.checked = engine === 'offset';
    }
  }

  function updateSubsyncFieldVisibility(): void {
    const engine = getSelectedEngine();
    const usesReference = engine === 'builtin' || engine === 'alass';
    const isOffset = engine === 'offset';

    ctx.dom.subsyncReferenceLabel.classList.toggle('hidden', !usesReference);
    ctx.dom.subsyncTargetLabel.classList.toggle(
      'hidden',
      ctx.state.subsyncSubtitleTracks.length === 0,
    );
    ctx.dom.subsyncOptionsGroup?.classList.toggle('hidden', !usesReference);
    ctx.dom.subsyncAllowSplitsLabel?.classList.toggle('hidden', engine !== 'builtin');
    ctx.dom.subsyncOffsetLabel?.classList.toggle('hidden', !isOffset);
    ctx.dom.subsyncBrowseReferenceButton?.classList.toggle(
      'hidden',
      !usesReference || !currentPayload?.canPickReferenceFile,
    );
  }

  function appendOption(select: HTMLSelectElement, value: string, label: string): void {
    const option = document.createElement('option');
    option.value = value;
    option.textContent = label;
    select.appendChild(option);
  }

  function getSelectedTargetTrackId(): number | null {
    const raw = Number.parseInt(ctx.dom.subsyncTargetSelect.value, 10);
    return Number.isFinite(raw) ? raw : null;
  }

  function renderTargetTracks(preferredTrackId: number | null): void {
    const select = ctx.dom.subsyncTargetSelect;
    select.innerHTML = '';
    select.value = '';
    for (const track of ctx.state.subsyncSubtitleTracks) {
      appendOption(select, String(track.id), track.label);
    }
    select.disabled = ctx.state.subsyncSubtitleTracks.length === 0;

    const preferred = ctx.state.subsyncSubtitleTracks.find(
      (track) => track.id === preferredTrackId,
    );
    const fallback = ctx.state.subsyncSubtitleTracks[0];
    const selected = preferred ?? fallback;
    if (selected) {
      select.value = String(selected.id);
    }
  }

  function formatCandidateOptionLabel(candidate: SubsyncReferenceCandidate): string {
    const prefix = candidate.recommended ? '★ ' : '';
    const detail = candidate.detail ? ` — ${candidate.detail}` : '';
    return `${prefix}${candidate.label}${detail}`;
  }

  function buildReferenceOptions(): Array<{ value: string; label: string }> {
    const targetTrackId = getSelectedTargetTrackId();
    const options: Array<{ value: string; label: string }> = [];
    const seen = new Set<string>();

    // When rich candidates are supplied, preserve their diagnostic ranking while
    // ensuring any non-target track in subtitleTracks is also reachable.
    if (currentPayload?.referenceCandidates && currentPayload.referenceCandidates.length > 0) {
      for (const candidate of [...customFileCandidates, ...currentPayload.referenceCandidates]) {
        if (candidate.kind === 'track' && candidate.trackId === targetTrackId) continue;
        if (seen.has(candidate.key)) continue;
        seen.add(candidate.key);
        options.push({
          value: candidate.key,
          label: formatCandidateOptionLabel(candidate),
        });
      }
      for (const track of ctx.state.subsyncSubtitleTracks) {
        if (track.id === targetTrackId) continue;
        const key = `${TRACK_PREFIX}${track.id}`;
        if (seen.has(key)) continue;
        seen.add(key);
        options.push({ value: key, label: track.label });
      }
      return options;
    }

    // Legacy / simplified payload format (used by minimal payloads & existing tests).
    for (const candidate of customFileCandidates) {
      if (seen.has(candidate.key)) continue;
      seen.add(candidate.key);
      options.push({
        value: candidate.key,
        label: formatCandidateOptionLabel(candidate),
      });
    }
    for (const track of ctx.state.subsyncSubtitleTracks) {
      if (track.id === targetTrackId) continue;
      const value = String(track.id);
      if (seen.has(value)) continue;
      seen.add(value);
      options.push({ value, label: track.label });
    }
    if (currentPayload?.videoReferenceAvailable) {
      options.push({ value: VIDEO_REFERENCE_VALUE, label: 'Video file (audio reference)' });
    }
    return options;
  }

  function renderReferenceTracks(preferredValue: string | null): void {
    const select = ctx.dom.subsyncReferenceSelect;
    const items = buildReferenceOptions();
    const values = items.map((item) => item.value);

    select.innerHTML = '';
    select.value = '';
    for (const item of items) {
      appendOption(select, item.value, item.label);
    }
    select.disabled = values.length === 0;

    const preferred = preferredValue && values.includes(preferredValue) ? preferredValue : null;
    const defaultKey =
      currentPayload?.defaultReferenceKey && values.includes(currentPayload.defaultReferenceKey)
        ? currentPayload.defaultReferenceKey
        : null;
    const defaultTrackValue =
      currentPayload?.defaultReferenceTrackId !== null &&
      currentPayload?.defaultReferenceTrackId !== undefined
        ? values.includes(`${TRACK_PREFIX}${currentPayload.defaultReferenceTrackId}`)
          ? `${TRACK_PREFIX}${currentPayload.defaultReferenceTrackId}`
          : values.includes(String(currentPayload.defaultReferenceTrackId))
            ? String(currentPayload.defaultReferenceTrackId)
            : null
        : null;
    const fallback = defaultKey ?? defaultTrackValue ?? values[0];
    const selected = preferred ?? fallback;
    if (selected) {
      select.value = selected;
    }
  }

  function describeSubsyncState(): string {
    if (!currentPayload) return '';
    const builtinReady = hasBuiltinEngineUi() && hasBuiltinReference();
    const alassReady = hasAlassReference();
    if (builtinReady) {
      return 'Choose reference and options, or switch engines, then run.';
    }
    if (alassReady && currentPayload.ffsubsyncAvailable) {
      return 'Choose engine, reference and out-of-sync subtitle, then run.';
    }
    if (alassReady) {
      return 'Choose the alass reference and out-of-sync subtitle, then run.';
    }
    if (currentPayload.ffsubsyncAvailable) {
      return 'No reference available for alass. Use ffsubsync.';
    }
    if (hasOffsetEngine()) {
      return 'Apply a manual offset or browse for a reference subtitle file.';
    }
    return 'No sync engine available for current media.';
  }

  function closeSubsyncModal(): void {
    if (!ctx.state.subsyncModalOpen) return;

    ctx.state.subsyncModalOpen = false;
    options.syncSettingsModalSubtitleSuppression();

    ctx.dom.subsyncModal.classList.add('hidden');
    ctx.dom.subsyncModal.setAttribute('aria-hidden', 'true');
    window.electronAPI.notifyOverlayModalClosed('subsync');

    if (!ctx.state.isOverSubtitle && !options.modalStateReader.isAnyModalOpen()) {
      ctx.dom.overlay.classList.remove('interactive');
    }
  }

  function chooseInitialEngine(
    payload: SubsyncManualPayload,
    selection?: SubsyncSelection,
  ): SubsyncEngine | null {
    const builtinReady = hasBuiltinEngineUi() && hasBuiltinReference();
    const alassReady = hasAlassReference();
    const ffsubsyncReady = payload.ffsubsyncAvailable;
    const offsetReady = hasOffsetEngine();

    const isEngineAvailable = (engine: SubsyncEngine): boolean => {
      if (engine === 'builtin') return builtinReady;
      if (engine === 'alass') return alassReady;
      if (engine === 'ffsubsync') return ffsubsyncReady;
      if (engine === 'offset') return offsetReady;
      return false;
    };

    if (selection && isEngineAvailable(selection.engine)) {
      return selection.engine;
    }
    if (payload.defaultEngine && isEngineAvailable(payload.defaultEngine)) {
      return payload.defaultEngine;
    }
    if (builtinReady) return 'builtin';
    if (alassReady) return 'alass';
    if (ffsubsyncReady) return 'ffsubsync';
    if (offsetReady) return 'offset';
    return null;
  }

  function openSubsyncModal(payload: SubsyncManualPayload, selection?: SubsyncSelection): void {
    ctx.state.subsyncSubmitting = false;
    ctx.state.subsyncSubtitleTracks = payload.subtitleTracks;
    currentPayload = payload;
    if (!selection) {
      customFileCandidates = [];
    }

    const builtinReady = hasBuiltinEngineUi() && hasBuiltinReference();
    const alassReady = hasAlassReference();
    const offsetReady = hasOffsetEngine();
    const chosenEngine = chooseInitialEngine(payload, selection);

    setSelectedEngine(chosenEngine);
    if (ctx.dom.subsyncEngineBuiltin) {
      ctx.dom.subsyncEngineBuiltin.disabled = !builtinReady;
    }
    ctx.dom.subsyncEngineAlass.disabled = !alassReady;
    ctx.dom.subsyncEngineFfsubsync.disabled = !payload.ffsubsyncAvailable;
    if (ctx.dom.subsyncEngineOffset) {
      ctx.dom.subsyncEngineOffset.disabled = !offsetReady;
    }
    ctx.dom.subsyncRunButton.disabled = chosenEngine === null;

    if (ctx.dom.subsyncGuidance) {
      const guidance = payload.guidanceMessage?.trim() ?? '';
      ctx.dom.subsyncGuidance.textContent = guidance;
      ctx.dom.subsyncGuidance.classList.toggle('hidden', guidance.length === 0);
    }

    if (ctx.dom.subsyncStripSignsCheckbox) {
      ctx.dom.subsyncStripSignsCheckbox.checked = selection ? selection.stripSigns : true;
    }
    if (ctx.dom.subsyncAllowSplitsCheckbox) {
      ctx.dom.subsyncAllowSplitsCheckbox.checked = selection ? selection.allowSplits : true;
    }
    if (ctx.dom.subsyncOffsetInput) {
      ctx.dom.subsyncOffsetInput.value = String(selection ? selection.offsetMs : 0);
    }

    renderTargetTracks(
      selection
        ? (selection.targetTrackId ?? payload.defaultTargetTrackId)
        : payload.defaultTargetTrackId,
    );
    renderReferenceTracks(selection?.referenceValue ?? null);
    updateSubsyncFieldVisibility();
    setSubsyncStatus(describeSubsyncState(), false);

    ctx.state.subsyncModalOpen = true;
    options.syncSettingsModalSubtitleSuppression();

    ctx.dom.overlay.classList.add('interactive');
    ctx.dom.subsyncModal.classList.remove('hidden');
    ctx.dom.subsyncModal.setAttribute('aria-hidden', 'false');
  }

  function reopenSubsyncModalWithError(
    payload: SubsyncManualPayload,
    selection: SubsyncSelection,
    message: string,
  ): void {
    openSubsyncModal(payload, selection);
    setSubsyncStatus(message, true);
    window.electronAPI.notifyOverlayModalOpened('subsync');
  }

  async function browseReferenceFile(): Promise<void> {
    if (!window.electronAPI.pickSubsyncReferenceFile) return;
    const picked = await window.electronAPI.pickSubsyncReferenceFile();
    if (!picked) return;
    const key = `${FILE_PREFIX}${picked}`;
    const filename = picked.split(/[\\/]/).pop() ?? picked;
    customFileCandidates = [
      {
        key,
        kind: 'file',
        label: `File - ${filename}`,
        detail: 'selected from disk',
        filePath: picked,
      },
      ...customFileCandidates.filter((c) => c.key !== key),
    ];
    if (ctx.dom.subsyncEngineBuiltin && ctx.dom.subsyncEngineBuiltin.disabled) {
      ctx.dom.subsyncEngineBuiltin.disabled = false;
      setSelectedEngine('builtin');
      ctx.dom.subsyncRunButton.disabled = false;
    }
    renderReferenceTracks(key);
    updateSubsyncFieldVisibility();
  }

  function parseReferenceValue(referenceValue: string): {
    referenceMode: 'track' | 'video' | 'file';
    referenceTrackId: number | null;
    referenceFilePath: string | null;
  } {
    if (referenceValue === VIDEO_REFERENCE_VALUE) {
      return { referenceMode: 'video', referenceTrackId: null, referenceFilePath: null };
    }
    if (referenceValue.startsWith(FILE_PREFIX)) {
      return {
        referenceMode: 'file',
        referenceTrackId: null,
        referenceFilePath: referenceValue.slice(FILE_PREFIX.length),
      };
    }
    const rawId = referenceValue.startsWith(TRACK_PREFIX)
      ? referenceValue.slice(TRACK_PREFIX.length)
      : referenceValue;
    return {
      referenceMode: 'track',
      referenceTrackId: Number.parseInt(rawId, 10),
      referenceFilePath: null,
    };
  }

  async function runSubsyncManualFromModal(): Promise<void> {
    if (ctx.state.subsyncSubmitting) return;
    if (ctx.dom.subsyncRunButton.disabled) return;
    if (!currentPayload) return;

    const engine = getSelectedEngine();
    if (!engine) {
      setSubsyncStatus('No sync engine available for current media.', true);
      return;
    }

    const referenceValue = ctx.dom.subsyncReferenceSelect.value;
    const targetTrackId = getSelectedTargetTrackId();
    const stripSigns = ctx.dom.subsyncStripSignsCheckbox?.checked ?? true;
    const allowSplits = ctx.dom.subsyncAllowSplitsCheckbox?.checked ?? true;
    const offsetMs = Number.parseInt(ctx.dom.subsyncOffsetInput?.value ?? '0', 10) || 0;

    if (targetTrackId === null) {
      setSubsyncStatus('Select the out-of-sync subtitle track to retime.', true);
      return;
    }
    if ((engine === 'builtin' || engine === 'alass') && !referenceValue) {
      setSubsyncStatus(
        engine === 'builtin' ? 'Select a reference source.' : 'Select a reference for alass.',
        true,
      );
      return;
    }
    if (engine === 'offset' && offsetMs === 0) {
      setSubsyncStatus('Enter a non-zero offset in milliseconds.', true);
      return;
    }

    const request: SubsyncManualRunRequest = {
      engine,
      targetTrackId,
    };
    if (engine === 'builtin' || engine === 'alass') {
      const parsed = parseReferenceValue(referenceValue);
      request.referenceMode = parsed.referenceMode;
      request.referenceTrackId = parsed.referenceTrackId;
      if (parsed.referenceFilePath !== null) {
        request.referenceFilePath = parsed.referenceFilePath;
      }
      if (ctx.dom.subsyncStripSignsCheckbox) {
        request.stripSigns = stripSigns;
      }
      if (engine === 'builtin' && ctx.dom.subsyncAllowSplitsCheckbox) {
        request.allowSplits = allowSplits;
      }
    } else if (engine === 'offset') {
      request.offsetMs = offsetMs;
    }

    const payloadSnapshot: SubsyncManualPayload = {
      ...currentPayload,
      subtitleTracks: currentPayload.subtitleTracks.map((track) => ({ ...track })),
      referenceCandidates: currentPayload.referenceCandidates?.map((candidate) => ({
        ...candidate,
      })),
    };
    const selection: SubsyncSelection = {
      engine,
      referenceValue,
      targetTrackId,
      stripSigns,
      allowSplits,
      offsetMs,
    };

    ctx.state.subsyncSubmitting = true;
    ctx.dom.subsyncRunButton.disabled = true;
    closeSubsyncModal();

    try {
      const result = await window.electronAPI.runSubsyncManual(request);
      if (result.ok) return;
      reopenSubsyncModalWithError(payloadSnapshot, selection, result.message);
    } catch (error) {
      reopenSubsyncModalWithError(
        payloadSnapshot,
        selection,
        `Subsync failed: ${(error as Error).message}`,
      );
    } finally {
      ctx.state.subsyncSubmitting = false;
      ctx.dom.subsyncRunButton.disabled = false;
    }
  }

  function handleSubsyncKeydown(e: KeyboardEvent): boolean {
    if (e.key === 'Escape') {
      e.preventDefault();
      closeSubsyncModal();
      return true;
    }

    if (e.key === 'Enter') {
      e.preventDefault();
      void runSubsyncManualFromModal();
      return true;
    }

    return true;
  }

  function wireDomEvents(): void {
    ctx.dom.subsyncCloseButton.addEventListener('click', () => {
      closeSubsyncModal();
    });
    ctx.dom.subsyncEngineBuiltin?.addEventListener('change', () => {
      updateSubsyncFieldVisibility();
    });
    ctx.dom.subsyncEngineAlass.addEventListener('change', () => {
      updateSubsyncFieldVisibility();
    });
    ctx.dom.subsyncEngineFfsubsync.addEventListener('change', () => {
      updateSubsyncFieldVisibility();
    });
    ctx.dom.subsyncEngineOffset?.addEventListener('change', () => {
      updateSubsyncFieldVisibility();
    });
    ctx.dom.subsyncBrowseReferenceButton?.addEventListener('click', () => {
      void browseReferenceFile();
    });
    ctx.dom.subsyncTargetSelect.addEventListener('change', () => {
      renderReferenceTracks(ctx.dom.subsyncReferenceSelect.value || null);
    });
    ctx.dom.subsyncRunButton.addEventListener('click', () => {
      void runSubsyncManualFromModal();
    });
  }

  return {
    closeSubsyncModal,
    handleSubsyncKeydown,
    openSubsyncModal,
    wireDomEvents,
  };
}
