import type {
  SubtitleCue,
  SubtitleData,
  SubtitleSidebarSnapshot,
} from '../../types';
import type { ModalStateReader, RendererContext } from '../context';

const MANUAL_SCROLL_HOLD_MS = 1500;
const ACTIVE_CUE_LOOKAHEAD_SEC = 0.18;
const CLICK_SEEK_OFFSET_SEC = 0.08;
const SNAPSHOT_POLL_INTERVAL_MS = 80;
const EMBEDDED_SIDEBAR_MIN_WIDTH_PX = 240;
const EMBEDDED_SIDEBAR_MAX_RATIO = 0.45;

function subtitleCueListsEqual(a: SubtitleCue[], b: SubtitleCue[]): boolean {
  if (a.length !== b.length) {
    return false;
  }

  for (let i = 0; i < a.length; i += 1) {
    const left = a[i]!;
    const right = b[i]!;
    if (
      left.startTime !== right.startTime ||
      left.endTime !== right.endTime ||
      left.text !== right.text
    ) {
      return false;
    }
  }

  return true;
}

function normalizeCueText(text: string): string {
  return text.replace(/\r\n/g, '\n').trim();
}

function formatCueTimestamp(seconds: number): string {
  const totalSeconds = Math.max(0, Math.floor(seconds));
  const mins = Math.floor(totalSeconds / 60);
  const secs = totalSeconds % 60;
  return `${String(mins).padStart(2, '0')}:${String(secs).padStart(2, '0')}`;
}

export function findActiveSubtitleCueIndex(
  cues: SubtitleCue[],
  current: { text: string; startTime?: number | null } | null,
  currentTimeSec: number | null = null,
  preferredCueIndex: number = -1,
): number {
  if (cues.length === 0) {
    return -1;
  }

  if (typeof currentTimeSec === 'number' && Number.isFinite(currentTimeSec)) {
    const activeOrUpcomingCue = cues.findIndex(
      (cue) =>
        cue.endTime > currentTimeSec &&
        cue.startTime <= currentTimeSec + ACTIVE_CUE_LOOKAHEAD_SEC,
    );
    if (activeOrUpcomingCue >= 0) {
      return activeOrUpcomingCue;
    }

    const nextCue = cues.findIndex((cue) => cue.endTime > currentTimeSec);
    if (nextCue >= 0) {
      return nextCue;
    }
  }

  if (!current) {
    return -1;
  }

  if (typeof current.startTime === 'number' && Number.isFinite(current.startTime)) {
    const timingMatch = cues.findIndex(
      (cue) => current.startTime! >= cue.startTime && current.startTime! < cue.endTime,
    );
    if (timingMatch >= 0) {
      return timingMatch;
    }
  }

  const normalizedText = normalizeCueText(current.text);
  if (!normalizedText) {
    return -1;
  }

  const matchingIndices: number[] = [];
  for (const [index, cue] of cues.entries()) {
    if (normalizeCueText(cue.text) === normalizedText) {
      matchingIndices.push(index);
    }
  }

  if (matchingIndices.length === 0) {
    return -1;
  }

  const hasTiming =
    typeof current.startTime === 'number' && Number.isFinite(current.startTime);

  if (preferredCueIndex >= 0) {
    if (!hasTiming && currentTimeSec === null) {
      const forwardMatches = matchingIndices.filter((index) => index >= preferredCueIndex);
      if (forwardMatches.length > 0) {
        return forwardMatches[0]!;
      }
      return preferredCueIndex;
    }

    let nearestIndex = matchingIndices[0]!;
    let nearestDistance = Math.abs(nearestIndex - preferredCueIndex);
    for (const matchIndex of matchingIndices) {
      const distance = Math.abs(matchIndex - preferredCueIndex);
      if (distance < nearestDistance) {
        nearestIndex = matchIndex;
        nearestDistance = distance;
      }
    }
    return nearestIndex;
  }

  return matchingIndices[0]!;
}

export function createSubtitleSidebarModal(
  ctx: RendererContext,
  options: {
    modalStateReader: Pick<ModalStateReader, 'isAnyModalOpen'>;
  },
) {
  let snapshotPollInterval: ReturnType<typeof setInterval> | null = null;
  let lastAppliedVideoMarginRatio: number | null = null;

  function setStatus(message: string): void {
    ctx.dom.subtitleSidebarStatus.textContent = message;
  }

  function getReservedSidebarWidthPx(): number {
    const config = ctx.state.subtitleSidebarConfig;
    if (!config || config.layout !== 'embedded' || !ctx.state.subtitleSidebarModalOpen) {
      return 0;
    }

    const measuredWidth = ctx.dom.subtitleSidebarContent.getBoundingClientRect().width;
    if (Number.isFinite(measuredWidth) && measuredWidth > 0) {
      return measuredWidth;
    }

    return Math.max(EMBEDDED_SIDEBAR_MIN_WIDTH_PX, config.maxWidth);
  }

  function syncEmbeddedSidebarLayout(): void {
    const config = ctx.state.subtitleSidebarConfig;
    const reservedWidthPx = getReservedSidebarWidthPx();
    const embedded = Boolean(config && config.layout === 'embedded' && reservedWidthPx > 0);

    if (embedded) {
      ctx.dom.subtitleSidebarContent.classList.add('subtitle-sidebar-content-embedded');
      ctx.dom.subtitleSidebarModal.classList.add('subtitle-sidebar-modal-embedded');
      document.body.classList.add('subtitle-sidebar-embedded-open');
    } else {
      ctx.dom.subtitleSidebarContent.classList.remove('subtitle-sidebar-content-embedded');
      ctx.dom.subtitleSidebarModal.classList.remove('subtitle-sidebar-modal-embedded');
      document.body.classList.remove('subtitle-sidebar-embedded-open');
    }
    document.documentElement.style.setProperty(
      '--subtitle-sidebar-reserved-width',
      `${Math.max(0, Math.round(reservedWidthPx))}px`,
    );

    const viewportWidth = window.innerWidth;
    const ratio =
      embedded && Number.isFinite(viewportWidth) && viewportWidth > 0
        ? Math.min(EMBEDDED_SIDEBAR_MAX_RATIO, reservedWidthPx / viewportWidth)
        : 0;

    if (
      lastAppliedVideoMarginRatio !== null &&
      Math.abs(ratio - lastAppliedVideoMarginRatio) < 0.0001
    ) {
      return;
    }

    lastAppliedVideoMarginRatio = ratio;
    window.electronAPI.sendMpvCommand([
      'set_property',
      'video-margin-ratio-right',
      Number(ratio.toFixed(4)),
    ]);
    if (ratio === 0) {
      window.electronAPI.sendMpvCommand(['set_property', 'video-pan-x', 0]);
    }
  }

  function applyConfig(snapshot: SubtitleSidebarSnapshot): void {
    ctx.state.subtitleSidebarConfig = snapshot.config;
    ctx.state.subtitleSidebarToggleKey = snapshot.config.toggleKey;
    ctx.state.subtitleSidebarPauseVideoOnHover = snapshot.config.pauseVideoOnHover;
    ctx.state.subtitleSidebarAutoScroll = snapshot.config.autoScroll;
    const style = ctx.dom.subtitleSidebarModal.style;
    style.setProperty('--subtitle-sidebar-max-width', `${snapshot.config.maxWidth}px`);
    style.setProperty('--subtitle-sidebar-opacity', String(snapshot.config.opacity));
    style.setProperty('--subtitle-sidebar-background-color', snapshot.config.backgroundColor);
    style.setProperty('--subtitle-sidebar-text-color', snapshot.config.textColor);
    style.setProperty('--subtitle-sidebar-font-family', snapshot.config.fontFamily);
    style.setProperty('--subtitle-sidebar-font-size', `${snapshot.config.fontSize}px`);
    style.setProperty('--subtitle-sidebar-timestamp-color', snapshot.config.timestampColor);
    style.setProperty('--subtitle-sidebar-active-line-color', snapshot.config.activeLineColor);
    style.setProperty(
      '--subtitle-sidebar-active-background-color',
      snapshot.config.activeLineBackgroundColor,
    );
    style.setProperty(
      '--subtitle-sidebar-hover-background-color',
      snapshot.config.hoverLineBackgroundColor,
    );
  }

  function seekToCue(cue: SubtitleCue): void {
    const targetTime = Math.min(cue.endTime - 0.01, cue.startTime + CLICK_SEEK_OFFSET_SEC);
    window.electronAPI.sendMpvCommand([
      'seek',
      Math.max(cue.startTime, targetTime),
      'absolute+exact',
    ]);
  }

  function maybeAutoScrollActiveCue(
    previousActiveCueIndex: number,
    behavior: ScrollBehavior = 'smooth',
    force = false,
  ): void {
    if (
      !ctx.state.subtitleSidebarAutoScroll ||
      ctx.state.subtitleSidebarActiveCueIndex < 0 ||
      (!force && ctx.state.subtitleSidebarActiveCueIndex === previousActiveCueIndex) ||
      Date.now() < ctx.state.subtitleSidebarManualScrollUntilMs
    ) {
      return;
    }

    const list = ctx.dom.subtitleSidebarList;
    const active = list.children[ctx.state.subtitleSidebarActiveCueIndex] as HTMLElement | undefined;
    if (!active) {
      return;
    }

    const targetScrollTop =
      active.offsetTop - (list.clientHeight - active.clientHeight) / 2;
    list.scrollTo({
      top: Math.max(0, targetScrollTop),
      behavior,
    });
  }

  function renderCueList(): void {
    ctx.dom.subtitleSidebarList.innerHTML = '';
    for (const [index, cue] of ctx.state.subtitleSidebarCues.entries()) {
      const row = document.createElement('li');
      row.className = 'subtitle-sidebar-item';
      row.classList.toggle('active', index === ctx.state.subtitleSidebarActiveCueIndex);
      row.dataset.index = String(index);

      const timestamp = document.createElement('div');
      timestamp.className = 'subtitle-sidebar-timestamp';
      timestamp.textContent = formatCueTimestamp(cue.startTime);

      const text = document.createElement('div');
      text.className = 'subtitle-sidebar-text';
      text.textContent = cue.text;

      row.appendChild(timestamp);
      row.appendChild(text);

      ctx.dom.subtitleSidebarList.appendChild(row);
    }
  }

  function syncActiveCueClasses(previousActiveCueIndex: number): void {
    if (previousActiveCueIndex >= 0) {
      const previous = ctx.dom.subtitleSidebarList.children[previousActiveCueIndex] as
        | HTMLElement
        | undefined;
      previous?.classList.remove('active');
    }

    if (ctx.state.subtitleSidebarActiveCueIndex >= 0) {
      const current = ctx.dom.subtitleSidebarList.children[ctx.state.subtitleSidebarActiveCueIndex] as
        | HTMLElement
        | undefined;
      current?.classList.add('active');
    }
  }

  function updateActiveCue(
    current: { text: string; startTime?: number | null } | null,
    currentTimeSec: number | null = null,
  ): void {
    const previousActiveCueIndex = ctx.state.subtitleSidebarActiveCueIndex;
    ctx.state.subtitleSidebarActiveCueIndex = findActiveSubtitleCueIndex(
      ctx.state.subtitleSidebarCues,
      current,
      currentTimeSec,
      previousActiveCueIndex,
    );
    if (ctx.state.subtitleSidebarModalOpen) {
      syncActiveCueClasses(previousActiveCueIndex);
      maybeAutoScrollActiveCue(previousActiveCueIndex);
    }
  }

  async function refreshSnapshot(): Promise<SubtitleSidebarSnapshot> {
    const snapshot = await window.electronAPI.getSubtitleSidebarSnapshot();
    applyConfig(snapshot);
    const cuesChanged = !subtitleCueListsEqual(ctx.state.subtitleSidebarCues, snapshot.cues);
    if (cuesChanged) {
      ctx.state.subtitleSidebarCues = snapshot.cues;
      if (ctx.state.subtitleSidebarModalOpen) {
        renderCueList();
      }
    }
    updateActiveCue(snapshot.currentSubtitle, snapshot.currentTimeSec ?? null);
    syncEmbeddedSidebarLayout();
    return snapshot;
  }

  function startSnapshotPolling(): void {
    if (snapshotPollInterval) {
      clearInterval(snapshotPollInterval);
    }
    snapshotPollInterval = setInterval(() => {
      void refreshSnapshot();
    }, SNAPSHOT_POLL_INTERVAL_MS);
  }

  function stopSnapshotPolling(): void {
    if (!snapshotPollInterval) {
      return;
    }
    clearInterval(snapshotPollInterval);
    snapshotPollInterval = null;
  }

  async function openSubtitleSidebarModal(): Promise<void> {
    const snapshot = await refreshSnapshot();
    ctx.dom.subtitleSidebarList.innerHTML = '';
    if (!snapshot.config.enabled) {
      setStatus('Subtitle sidebar disabled in config.');
    } else if (snapshot.cues.length === 0) {
      setStatus('No parsed subtitle cues available.');
    } else {
      setStatus(`${snapshot.cues.length} parsed subtitle lines`);
    }

    ctx.state.subtitleSidebarModalOpen = true;
    ctx.dom.subtitleSidebarModal.classList.remove('hidden');
    ctx.dom.subtitleSidebarModal.setAttribute('aria-hidden', 'false');
    renderCueList();
    syncActiveCueClasses(-1);
    maybeAutoScrollActiveCue(-1, 'auto', true);
    startSnapshotPolling();
    syncEmbeddedSidebarLayout();
  }

  function closeSubtitleSidebarModal(): void {
    if (!ctx.state.subtitleSidebarModalOpen) {
      return;
    }
    ctx.state.subtitleSidebarModalOpen = false;
    ctx.dom.subtitleSidebarModal.classList.add('hidden');
    ctx.dom.subtitleSidebarModal.setAttribute('aria-hidden', 'true');
    stopSnapshotPolling();
    syncEmbeddedSidebarLayout();
    if (!ctx.state.isOverSubtitle && !options.modalStateReader.isAnyModalOpen()) {
      ctx.dom.overlay.classList.remove('interactive');
    }
  }

  async function toggleSubtitleSidebarModal(): Promise<void> {
    if (ctx.state.subtitleSidebarModalOpen) {
      closeSubtitleSidebarModal();
      return;
    }
    await openSubtitleSidebarModal();
  }

  function handleSubtitleUpdated(data: SubtitleData): void {
    if (ctx.state.subtitleSidebarModalOpen) {
      return;
    }

    updateActiveCue(
      { text: data.text, startTime: data.startTime },
      data.startTime ?? null,
    );
  }

  function wireDomEvents(): void {
    ctx.dom.subtitleSidebarClose.addEventListener('click', () => {
      closeSubtitleSidebarModal();
    });
    ctx.dom.subtitleSidebarList.addEventListener('click', (event) => {
      const target = event.target;
      if (!(target instanceof Element)) {
        return;
      }
      const row = target.closest<HTMLElement>('.subtitle-sidebar-item');
      if (!row) {
        return;
      }
      const index = Number.parseInt(row.dataset.index ?? '', 10);
      if (!Number.isInteger(index) || index < 0 || index >= ctx.state.subtitleSidebarCues.length) {
        return;
      }
      const cue = ctx.state.subtitleSidebarCues[index];
      if (!cue) {
        return;
      }
      seekToCue(cue);
    });
    ctx.dom.subtitleSidebarList.addEventListener('wheel', () => {
      ctx.state.subtitleSidebarManualScrollUntilMs = Date.now() + MANUAL_SCROLL_HOLD_MS;
    });
    ctx.dom.subtitleSidebarModal.addEventListener('mouseenter', async () => {
      if (!ctx.state.subtitleSidebarPauseVideoOnHover || ctx.state.subtitleSidebarPausedByHover) {
        return;
      }
      const paused = await window.electronAPI.getPlaybackPaused();
      if (paused === false) {
        window.electronAPI.sendMpvCommand(['set_property', 'pause', 'yes']);
        ctx.state.subtitleSidebarPausedByHover = true;
      }
    });
    ctx.dom.subtitleSidebarModal.addEventListener('mouseleave', () => {
      if (!ctx.state.subtitleSidebarPausedByHover) {
        return;
      }
      ctx.state.subtitleSidebarPausedByHover = false;
      window.electronAPI.sendMpvCommand(['set_property', 'pause', 'no']);
    });
    window.addEventListener('resize', () => {
      if (!ctx.state.subtitleSidebarModalOpen) {
        return;
      }
      syncEmbeddedSidebarLayout();
    });
  }

  return {
    openSubtitleSidebarModal,
    closeSubtitleSidebarModal,
    toggleSubtitleSidebarModal,
    refreshSubtitleSidebarSnapshot: refreshSnapshot,
    wireDomEvents,
    handleSubtitleUpdated,
    seekToCue,
  };
}
