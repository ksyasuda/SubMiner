import type {
  SubtitleCue,
  SubtitleData,
  SubtitleMiningContext,
  SubtitleSidebarSnapshot,
} from '../../types';
import type { ModalStateReader, RendererContext } from '../context';
import { i18n } from '../../i18n/index.js';
import { syncOverlayMouseIgnoreState } from '../overlay-mouse-ignore.js';
import {
  YOMITAN_POPUP_HIDDEN_EVENT,
  YOMITAN_POPUP_SHOWN_EVENT,
  isYomitanPopupVisible,
} from '../yomitan-popup.js';

const MANUAL_SCROLL_HOLD_MS = 1500;
const ACTIVE_CUE_LOOKAHEAD_SEC = 0.18;
const CLICK_SEEK_OFFSET_SEC = 0.08;
const SNAPSHOT_POLL_INTERVAL_MS = 80;
const EMBEDDED_SIDEBAR_MIN_WIDTH_PX = 240;
const EMBEDDED_SIDEBAR_MAX_RATIO = 0.45;
const appliedSidebarCssKeys = new WeakMap<HTMLElement, Set<string>>();

function nowForUiTiming(): number {
  if (typeof performance !== 'undefined' && typeof performance.now === 'function') {
    return performance.now();
  }
  return Date.now();
}

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
  const hours = Math.floor(totalSeconds / 3600);
  const mins = Math.floor((totalSeconds % 3600) / 60);
  const secs = totalSeconds % 60;
  if (hours > 0) {
    return [
      String(hours).padStart(2, '0'),
      String(mins).padStart(2, '0'),
      String(secs).padStart(2, '0'),
    ].join(':');
  }
  return `${String(mins).padStart(2, '0')}:${String(secs).padStart(2, '0')}`;
}

export function applySidebarCssDeclarations(
  target: HTMLElement,
  declarations: Record<string, string>,
): void {
  const targetStyle = (target as HTMLElement & { style?: CSSStyleDeclaration }).style;
  if (!targetStyle) return;
  const styleTarget = targetStyle as unknown as Record<string, string>;
  const previousKeys = appliedSidebarCssKeys.get(target) ?? new Set<string>();
  const nextKeys = new Set<string>();

  for (const property of previousKeys) {
    if (Object.prototype.hasOwnProperty.call(declarations, property)) continue;
    if (property.includes('-')) {
      targetStyle.removeProperty(property);
    } else {
      styleTarget[property] = '';
    }
  }

  for (const [property, rawValue] of Object.entries(declarations)) {
    const value = rawValue.trim();
    if (value.length === 0) {
      if (property.includes('-')) {
        targetStyle.removeProperty(property);
      } else {
        styleTarget[property] = '';
      }
      continue;
    }
    if (property.includes('-')) {
      targetStyle.setProperty(property, value);
    } else {
      styleTarget[property] = value;
    }
    nextKeys.add(property);
  }

  appliedSidebarCssKeys.set(target, nextKeys);
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

  const hasCurrentTiming =
    typeof current?.startTime === 'number' && Number.isFinite(current.startTime);

  if (hasCurrentTiming) {
    const timingMatch = cues.findIndex(
      (cue) => current.startTime! >= cue.startTime && current.startTime! < cue.endTime,
    );
    if (timingMatch >= 0) {
      return timingMatch;
    }
  }

  if (typeof currentTimeSec === 'number' && Number.isFinite(currentTimeSec)) {
    const activeOrUpcomingCue = cues.findIndex(
      (cue) =>
        cue.endTime > currentTimeSec && cue.startTime <= currentTimeSec + ACTIVE_CUE_LOOKAHEAD_SEC,
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

  const hasTiming = typeof current.startTime === 'number' && Number.isFinite(current.startTime);

  if (preferredCueIndex >= 0) {
    if (!hasTiming && currentTimeSec === null) {
      const forwardMatches = matchingIndices.filter((index) => index >= preferredCueIndex);
      if (forwardMatches.length > 0) {
        return forwardMatches[0]!;
      }
      if (matchingIndices.includes(preferredCueIndex)) {
        return preferredCueIndex;
      }
      return matchingIndices[matchingIndices.length - 1] ?? -1;
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
    onVisibilityChanged?: (visible: boolean) => void;
    shouldRestoreOpenOnStartup?: () => Promise<boolean>;
  },
) {
  let snapshotPollInterval: ReturnType<typeof setTimeout> | null = null;
  let lastAppliedVideoMarginRatio: number | null = null;
  let subtitleSidebarHoverRequestId = 0;
  let disposeDomEvents: (() => void) | null = null;
  let subtitleSidebarHovered = false;
  let subtitleSidebarFocusedWithin = false;
  let subtitleSidebarYomitanPopupVisible = false;
  let subtitleSidebarPauseHeldByYomitanPopup = false;
  let lastSubtitleSidebarLookupCueIndex = -1;

  function restoreEmbeddedSidebarPassthrough(): void {
    syncOverlayMouseIgnoreState(ctx);
  }

  function syncSidebarInteractionState(): void {
    ctx.state.isOverSubtitleSidebar = subtitleSidebarHovered || subtitleSidebarFocusedWithin;
  }

  function clearSidebarInteractionState(): void {
    subtitleSidebarHovered = false;
    subtitleSidebarFocusedWithin = false;
    lastSubtitleSidebarLookupCueIndex = -1;
    syncSidebarInteractionState();
  }

  function findCueIndexFromNode(node: Node | null): number | null {
    if (!node || typeof Element === 'undefined') {
      return null;
    }
    const element = node instanceof Element ? node : node.parentElement;
    const row = element?.closest<HTMLElement>('.subtitle-sidebar-item') ?? null;
    if (!row) {
      return null;
    }
    const index = Number.parseInt(row.dataset.index ?? '', 10);
    if (!Number.isInteger(index) || index < 0 || index >= ctx.state.subtitleSidebarCues.length) {
      return null;
    }
    return index;
  }

  function rememberLookupCueFromTarget(target: EventTarget | null): void {
    if (typeof Node === 'undefined') {
      return;
    }
    if (!(target instanceof Node)) {
      return;
    }
    const index = findCueIndexFromNode(target);
    if (index === null) {
      return;
    }
    lastSubtitleSidebarLookupCueIndex = index;
  }

  function getSubtitleSidebarMiningContext(): SubtitleMiningContext | null {
    if (!ctx.state.subtitleSidebarModalOpen) {
      return null;
    }

    const selection = window.getSelection?.() ?? null;
    const selectionIndex =
      findCueIndexFromNode(selection?.anchorNode ?? null) ??
      findCueIndexFromNode(selection?.focusNode ?? null);
    const index =
      selectionIndex ??
      (lastSubtitleSidebarLookupCueIndex >= 0 ? lastSubtitleSidebarLookupCueIndex : null);
    if (index === null) {
      return null;
    }

    const cue = ctx.state.subtitleSidebarCues[index];
    if (
      !cue ||
      !Number.isFinite(cue.startTime) ||
      !Number.isFinite(cue.endTime) ||
      cue.endTime <= cue.startTime
    ) {
      return null;
    }

    return {
      source: 'subtitle-sidebar',
      text: cue.text,
      startTime: cue.startTime,
      endTime: cue.endTime,
      capturedAtMs: Date.now(),
    };
  }

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
    const wantsEmbedded = Boolean(
      config && config.layout === 'embedded' && ctx.state.subtitleSidebarModalOpen,
    );

    if (wantsEmbedded) {
      ctx.dom.subtitleSidebarContent.classList.add('subtitle-sidebar-content-embedded');
      ctx.dom.subtitleSidebarModal.classList.add('subtitle-sidebar-modal-embedded');
      document.body.classList.add('subtitle-sidebar-embedded-open');
    } else {
      ctx.dom.subtitleSidebarContent.classList.remove('subtitle-sidebar-content-embedded');
      ctx.dom.subtitleSidebarModal.classList.remove('subtitle-sidebar-modal-embedded');
      document.body.classList.remove('subtitle-sidebar-embedded-open');
    }
    const reservedWidthPx = wantsEmbedded ? getReservedSidebarWidthPx() : 0;
    const embedded = wantsEmbedded && reservedWidthPx > 0;
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
    window.electronAPI.sendMpvCommand(['set_property', 'osd-align-x', 'left']);
    window.electronAPI.sendMpvCommand(['set_property', 'osd-align-y', 'top']);
    window.electronAPI.sendMpvCommand([
      'set_property',
      'user-data/osc/margins',
      JSON.stringify({
        l: 0,
        r: Number(ratio.toFixed(4)),
        t: 0,
        b: 0,
      }),
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
    applySidebarCssDeclarations(ctx.dom.subtitleSidebarContent, snapshot.config.css ?? {});
  }

  function seekToCue(cue: SubtitleCue): void {
    const targetTime = Math.min(cue.endTime - 0.01, cue.startTime + CLICK_SEEK_OFFSET_SEC);
    window.electronAPI.sendMpvCommand([
      'seek',
      Math.max(cue.startTime, targetTime),
      'absolute+exact',
    ]);
  }

  function getCueRowLabel(cue: SubtitleCue): string {
    return i18n.t('subtitleSidebar.jumpToCue', { time: formatCueTimestamp(cue.startTime) });
  }

  function isYomitanPopupVisibleForSidebar(): boolean {
    if (subtitleSidebarYomitanPopupVisible || ctx.state.yomitanPopupVisible) {
      return true;
    }
    if (typeof document === 'undefined') {
      return false;
    }
    return isYomitanPopupVisible(document);
  }

  function shouldHoldSidebarPauseForYomitanPopup(): boolean {
    return (
      ctx.state.autoPauseVideoOnYomitanPopup &&
      ctx.state.subtitleSidebarPausedByHover &&
      isYomitanPopupVisibleForSidebar()
    );
  }

  function resumeSubtitleSidebarHoverPause(): void {
    subtitleSidebarHoverRequestId += 1;
    if (!ctx.state.subtitleSidebarPausedByHover) {
      subtitleSidebarPauseHeldByYomitanPopup = false;
      restoreEmbeddedSidebarPassthrough();
      return;
    }

    if (shouldHoldSidebarPauseForYomitanPopup()) {
      subtitleSidebarPauseHeldByYomitanPopup = true;
      restoreEmbeddedSidebarPassthrough();
      return;
    }

    subtitleSidebarPauseHeldByYomitanPopup = false;
    ctx.state.subtitleSidebarPausedByHover = false;
    window.electronAPI.sendMpvCommand(['set_property', 'pause', 'no']);
    restoreEmbeddedSidebarPassthrough();
  }

  function handleYomitanPopupShown(): void {
    subtitleSidebarYomitanPopupVisible = true;
    if (ctx.state.autoPauseVideoOnYomitanPopup && ctx.state.subtitleSidebarPausedByHover) {
      subtitleSidebarPauseHeldByYomitanPopup = true;
    }
  }

  function handleYomitanPopupHidden(): void {
    subtitleSidebarYomitanPopupVisible = false;
    if (!subtitleSidebarPauseHeldByYomitanPopup) {
      return;
    }

    subtitleSidebarPauseHeldByYomitanPopup = false;
    if (ctx.state.isOverSubtitleSidebar) {
      restoreEmbeddedSidebarPassthrough();
      return;
    }
    resumeSubtitleSidebarHoverPause();
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
      nowForUiTiming() < ctx.state.subtitleSidebarManualScrollUntilMs
    ) {
      return;
    }

    const list = ctx.dom.subtitleSidebarList;
    const active = list.children[ctx.state.subtitleSidebarActiveCueIndex] as
      | HTMLElement
      | undefined;
    if (!active) {
      return;
    }

    const targetScrollTop = active.offsetTop - (list.clientHeight - active.clientHeight) / 2;
    const nextScrollTop = Math.max(0, targetScrollTop);
    if (previousActiveCueIndex < 0) {
      list.scrollTop = nextScrollTop;
      return;
    }

    list.scrollTo({
      top: nextScrollTop,
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
      row.tabIndex = 0;
      row.setAttribute('role', 'button');
      row.setAttribute('aria-label', getCueRowLabel(cue));
      row.addEventListener('keydown', (event: KeyboardEvent) => {
        if (event.key !== 'Enter' && event.key !== ' ') {
          return;
        }
        event.preventDefault();
        seekToCue(cue);
      });

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
      const current = ctx.dom.subtitleSidebarList.children[
        ctx.state.subtitleSidebarActiveCueIndex
      ] as HTMLElement | undefined;
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
    if (!snapshot.config.enabled) {
      resumeSubtitleSidebarHoverPause();
      clearSidebarInteractionState();
      ctx.state.subtitleSidebarCues = [];
      ctx.state.subtitleSidebarModalOpen = false;
      ctx.dom.subtitleSidebarModal.classList.add('hidden');
      ctx.dom.subtitleSidebarModal.setAttribute('aria-hidden', 'true');
      stopSnapshotPolling();
      updateActiveCue(null, snapshot.currentTimeSec ?? null);
      setStatus(i18n.t('subtitleSidebar.disabledInConfig'));
      syncEmbeddedSidebarLayout();
      if (!ctx.state.isOverSubtitle && !options.modalStateReader.isAnyModalOpen()) {
        ctx.dom.overlay.classList.remove('interactive');
      }
      restoreEmbeddedSidebarPassthrough();
      return snapshot;
    }

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
    stopSnapshotPolling();

    const pollOnce = async (): Promise<void> => {
      try {
        await refreshSnapshot();
      } catch {
        // Keep polling; a transient IPC failure should not stop updates.
      }

      if (!ctx.state.subtitleSidebarModalOpen) {
        snapshotPollInterval = null;
        return;
      }

      snapshotPollInterval = setTimeout(pollOnce, SNAPSHOT_POLL_INTERVAL_MS);
    };

    snapshotPollInterval = setTimeout(pollOnce, SNAPSHOT_POLL_INTERVAL_MS);
  }

  function stopSnapshotPolling(): void {
    if (!snapshotPollInterval) {
      return;
    }
    clearTimeout(snapshotPollInterval);
    snapshotPollInterval = null;
  }

  async function openSubtitleSidebarModal(): Promise<void> {
    const snapshot = await refreshSnapshot();
    if (!snapshot.config.enabled) {
      setStatus(i18n.t('subtitleSidebar.disabledInConfig'));
      return;
    }

    ctx.dom.subtitleSidebarList.innerHTML = '';
    if (snapshot.cues.length === 0) {
      setStatus(i18n.t('subtitleSidebar.noSubtitles'));
    } else {
      setStatus(i18n.t('subtitleSidebar.parsedCueCount', { count: snapshot.cues.length }));
    }

    ctx.state.subtitleSidebarModalOpen = true;
    clearSidebarInteractionState();
    ctx.dom.subtitleSidebarModal.classList.remove('hidden');
    ctx.dom.subtitleSidebarModal.setAttribute('aria-hidden', 'false');
    renderCueList();
    syncActiveCueClasses(-1);
    maybeAutoScrollActiveCue(-1, 'auto', true);
    startSnapshotPolling();
    syncEmbeddedSidebarLayout();
    restoreEmbeddedSidebarPassthrough();
    window.electronAPI.notifyOverlayModalOpened?.('subtitle-sidebar');
    options.onVisibilityChanged?.(true);
  }

  async function autoOpenSubtitleSidebarOnStartup(): Promise<void> {
    const snapshot = await refreshSnapshot();
    const shouldRestoreOpen = (await options.shouldRestoreOpenOnStartup?.()) === true;
    if (
      !snapshot.config.enabled ||
      (!snapshot.config.autoOpen && !shouldRestoreOpen) ||
      ctx.state.subtitleSidebarModalOpen
    ) {
      return;
    }
    await openSubtitleSidebarModal();
  }

  function closeSubtitleSidebarModal(): void {
    if (!ctx.state.subtitleSidebarModalOpen) {
      return;
    }
    resumeSubtitleSidebarHoverPause();
    clearSidebarInteractionState();
    ctx.state.subtitleSidebarModalOpen = false;
    ctx.dom.subtitleSidebarModal.classList.add('hidden');
    ctx.dom.subtitleSidebarModal.setAttribute('aria-hidden', 'true');
    stopSnapshotPolling();
    syncEmbeddedSidebarLayout();
    if (!ctx.state.isOverSubtitle && !options.modalStateReader.isAnyModalOpen()) {
      ctx.dom.overlay.classList.remove('interactive');
    }
    restoreEmbeddedSidebarPassthrough();
    window.electronAPI.notifyOverlayModalClosed?.('subtitle-sidebar');
    options.onVisibilityChanged?.(false);
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

    updateActiveCue({ text: data.text, startTime: data.startTime }, data.startTime ?? null);
  }

  function wireDomEvents(): void {
    if (disposeDomEvents) {
      return;
    }

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
      ctx.state.subtitleSidebarManualScrollUntilMs = nowForUiTiming() + MANUAL_SCROLL_HOLD_MS;
    });
    ctx.dom.subtitleSidebarList.addEventListener('pointerover', (event) => {
      rememberLookupCueFromTarget(event.target);
    });
    ctx.dom.subtitleSidebarList.addEventListener('focusin', (event) => {
      rememberLookupCueFromTarget(event.target);
    });
    ctx.dom.subtitleSidebarContent.addEventListener('mouseenter', async () => {
      subtitleSidebarHovered = true;
      syncSidebarInteractionState();
      restoreEmbeddedSidebarPassthrough();
      if (!ctx.state.subtitleSidebarPauseVideoOnHover || ctx.state.subtitleSidebarPausedByHover) {
        return;
      }
      const requestId = ++subtitleSidebarHoverRequestId;
      let paused: boolean | null | undefined;
      try {
        paused = await window.electronAPI.getPlaybackPaused();
      } catch {
        paused = undefined;
      }
      if (requestId !== subtitleSidebarHoverRequestId) {
        return;
      }
      if (paused === false) {
        window.electronAPI.sendMpvCommand(['set_property', 'pause', 'yes']);
        ctx.state.subtitleSidebarPausedByHover = true;
      }
    });
    ctx.dom.subtitleSidebarContent.addEventListener('mouseleave', () => {
      subtitleSidebarHovered = false;
      if (!subtitleSidebarFocusedWithin) {
        lastSubtitleSidebarLookupCueIndex = -1;
      }
      syncSidebarInteractionState();
      if (ctx.state.isOverSubtitleSidebar) {
        restoreEmbeddedSidebarPassthrough();
        return;
      }
      resumeSubtitleSidebarHoverPause();
    });
    ctx.dom.subtitleSidebarContent.addEventListener('focusin', () => {
      subtitleSidebarFocusedWithin = true;
      syncSidebarInteractionState();
      restoreEmbeddedSidebarPassthrough();
    });
    ctx.dom.subtitleSidebarContent.addEventListener('focusout', (event: FocusEvent) => {
      const relatedTarget = event.relatedTarget;
      if (
        typeof Node !== 'undefined' &&
        relatedTarget instanceof Node &&
        ctx.dom.subtitleSidebarContent.contains(relatedTarget)
      ) {
        return;
      }

      subtitleSidebarFocusedWithin = false;
      lastSubtitleSidebarLookupCueIndex = -1;
      syncSidebarInteractionState();
      if (ctx.state.isOverSubtitleSidebar) {
        restoreEmbeddedSidebarPassthrough();
        return;
      }
      resumeSubtitleSidebarHoverPause();
    });
    const resizeHandler = () => {
      if (!ctx.state.subtitleSidebarModalOpen) {
        return;
      }
      syncEmbeddedSidebarLayout();
    };
    window.addEventListener('resize', resizeHandler);
    window.addEventListener(YOMITAN_POPUP_SHOWN_EVENT, handleYomitanPopupShown);
    window.addEventListener(YOMITAN_POPUP_HIDDEN_EVENT, handleYomitanPopupHidden);
    disposeDomEvents = () => {
      window.removeEventListener('resize', resizeHandler);
      window.removeEventListener(YOMITAN_POPUP_SHOWN_EVENT, handleYomitanPopupShown);
      window.removeEventListener(YOMITAN_POPUP_HIDDEN_EVENT, handleYomitanPopupHidden);
      disposeDomEvents = null;
    };
  }

  return {
    autoOpenSubtitleSidebarOnStartup,
    openSubtitleSidebarModal,
    closeSubtitleSidebarModal,
    toggleSubtitleSidebarModal,
    refreshSubtitleSidebarSnapshot: refreshSnapshot,
    wireDomEvents,
    disposeDomEvents: () => {
      disposeDomEvents?.();
    },
    handleSubtitleUpdated,
    seekToCue,
    getSubtitleSidebarMiningContext,
  };
}
