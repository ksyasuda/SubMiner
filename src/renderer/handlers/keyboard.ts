import type { CompiledSessionBinding, ShortcutsConfig } from '../../types';
import type { RendererContext } from '../context';
import {
  YOMITAN_POPUP_HIDDEN_EVENT,
  YOMITAN_POPUP_SHOWN_EVENT,
  YOMITAN_POPUP_COMMAND_EVENT,
  isYomitanPopupVisible,
  isYomitanPopupIframe,
} from '../yomitan-popup.js';

export function createKeyboardHandlers(
  ctx: RendererContext,
  options: {
    handleRuntimeOptionsKeydown: (e: KeyboardEvent) => boolean;
    handleSubsyncKeydown: (e: KeyboardEvent) => boolean;
    handleKikuKeydown: (e: KeyboardEvent) => boolean;
    handleJimakuKeydown: (e: KeyboardEvent) => boolean;
    handleYoutubePickerKeydown: (e: KeyboardEvent) => boolean;
    handlePlaylistBrowserKeydown: (e: KeyboardEvent) => boolean;
    handleControllerSelectKeydown: (e: KeyboardEvent) => boolean;
    handleControllerDebugKeydown: (e: KeyboardEvent) => boolean;
    handleSessionHelpKeydown: (e: KeyboardEvent) => boolean;
    openSessionHelpModal: (opening: {
      bindingKey: 'KeyH' | 'KeyK';
      fallbackUsed: boolean;
      fallbackUnavailable: boolean;
    }) => void;
    appendClipboardVideoToQueue: () => void;
    getPlaybackPaused: () => Promise<boolean | null>;
    openControllerSelectModal: () => void;
    openControllerDebugModal: () => void;
    toggleSubtitleSidebarModal?: () => void;
  },
) {
  // Timeout for the modal chord capture window (e.g. Y followed by H/K).
  const CHORD_TIMEOUT_MS = 1000;
  const KEYBOARD_SELECTED_WORD_CLASS = 'keyboard-selected';
  let pendingSelectionAnchorAfterSubtitleSeek: 'start' | 'end' | null = null;
  let pendingLookupRefreshAfterSubtitleSeek = false;
  let resetSelectionToStartOnNextSubtitleSync = false;
  let lookupScanFallbackTimer: ReturnType<typeof setTimeout> | null = null;
  let pendingNumericSelection:
    | {
        actionId: 'copySubtitleMultiple' | 'mineSentenceMultiple';
        timeout: ReturnType<typeof setTimeout> | null;
      }
    | null = null;

  const CHORD_MAP = new Map<
    string,
    { type: 'mpv' | 'electron'; command?: string[]; action?: () => void }
  >([
    ['KeyS', { type: 'mpv', command: ['script-message', 'subminer-start'] }],
    ['Shift+KeyS', { type: 'mpv', command: ['script-message', 'subminer-stop'] }],
    ['KeyT', { type: 'mpv', command: ['script-message', 'subminer-toggle'] }],
    ['KeyO', { type: 'mpv', command: ['script-message', 'subminer-options'] }],
    ['KeyR', { type: 'mpv', command: ['script-message', 'subminer-restart'] }],
    ['KeyC', { type: 'mpv', command: ['script-message', 'subminer-status'] }],
    ['KeyY', { type: 'mpv', command: ['script-message', 'subminer-menu'] }],
    ['KeyD', { type: 'electron', action: () => window.electronAPI.toggleDevTools() }],
  ]);

  function isInteractiveTarget(target: EventTarget | null): boolean {
    if (!(target instanceof Element)) return false;
    if (target.closest('.modal')) return true;
    if (ctx.dom.subtitleContainer.contains(target)) return true;
    if (isYomitanPopupIframe(target)) return true;
    if (target.closest && target.closest('iframe.yomitan-popup, iframe[id^="yomitan-popup"]'))
      return true;
    return false;
  }

  function keyEventToString(e: KeyboardEvent): string {
    const parts: string[] = [];
    if (e.ctrlKey) parts.push('Ctrl');
    if (e.altKey) parts.push('Alt');
    if (e.shiftKey) parts.push('Shift');
    if (e.metaKey) parts.push('Meta');
    parts.push(e.code);
    return parts.join('+');
  }

  function updateConfiguredShortcuts(shortcuts: Required<ShortcutsConfig>): void {
    ctx.state.sessionActionTimeoutMs = shortcuts.multiCopyTimeoutMs;
  }

  async function refreshConfiguredShortcuts(): Promise<void> {
    updateConfiguredShortcuts(await window.electronAPI.getConfiguredShortcuts());
  }

  function updateSessionBindings(bindings: CompiledSessionBinding[]): void {
    ctx.state.sessionBindings = bindings;
    ctx.state.sessionBindingMap = new Map(
      bindings.map((binding) => [keyEventToStringFromBinding(binding), binding]),
    );
  }

  function keyEventToStringFromBinding(binding: CompiledSessionBinding): string {
    const parts: string[] = [];
    for (const modifier of binding.key.modifiers) {
      if (modifier === 'ctrl') parts.push('Ctrl');
      else if (modifier === 'alt') parts.push('Alt');
      else if (modifier === 'shift') parts.push('Shift');
      else if (modifier === 'meta') parts.push('Meta');
    }
    parts.push(binding.key.code);
    return parts.join('+');
  }

  function isTextEntryTarget(target: EventTarget | null): boolean {
    if (!target || typeof target !== 'object' || !('closest' in target)) return false;
    const element = target as { closest: (selector: string) => unknown };
    if (element.closest('[contenteditable="true"]')) return true;
    return Boolean(element.closest('input, textarea, select'));
  }

  function showSessionSelectionMessage(message: string): void {
    window.electronAPI.sendMpvCommand(['show-text', message, '3000']);
  }

  function cancelPendingNumericSelection(showCancelled: boolean): void {
    if (!pendingNumericSelection) return;
    if (pendingNumericSelection.timeout !== null) {
      clearTimeout(pendingNumericSelection.timeout);
    }
    pendingNumericSelection = null;
    if (showCancelled) {
      showSessionSelectionMessage('Cancelled');
    }
  }

  function startPendingNumericSelection(
    actionId: 'copySubtitleMultiple' | 'mineSentenceMultiple',
  ): void {
    cancelPendingNumericSelection(false);
    const timeoutMessage = actionId === 'copySubtitleMultiple' ? 'Copy timeout' : 'Mine timeout';
    const promptMessage =
      actionId === 'copySubtitleMultiple'
        ? 'Copy how many lines? Press 1-9 (Esc to cancel)'
        : 'Mine how many lines? Press 1-9 (Esc to cancel)';
    pendingNumericSelection = {
      actionId,
      timeout: setTimeout(() => {
        pendingNumericSelection = null;
        showSessionSelectionMessage(timeoutMessage);
      }, ctx.state.sessionActionTimeoutMs),
    };
    showSessionSelectionMessage(promptMessage);
  }

  function beginSessionNumericSelection(
    actionId: 'copySubtitleMultiple' | 'mineSentenceMultiple',
  ): void {
    startPendingNumericSelection(actionId);
  }

  function handlePendingNumericSelection(e: KeyboardEvent): boolean {
    if (!pendingNumericSelection) return false;
    if (e.key === 'Escape') {
      e.preventDefault();
      cancelPendingNumericSelection(true);
      return true;
    }

    if (!/^[1-9]$/.test(e.key) || e.ctrlKey || e.metaKey || e.altKey) {
      return false;
    }

    e.preventDefault();
    const count = Number(e.key);
    const actionId = pendingNumericSelection.actionId;
    cancelPendingNumericSelection(false);
    void window.electronAPI.dispatchSessionAction(actionId, { count });
    return true;
  }

  function dispatchSessionBinding(binding: CompiledSessionBinding): void {
    if (
      binding.actionType === 'session-action' &&
      (binding.actionId === 'copySubtitleMultiple' || binding.actionId === 'mineSentenceMultiple')
    ) {
      startPendingNumericSelection(binding.actionId);
      return;
    }

    if (binding.actionType === 'mpv-command') {
      dispatchConfiguredMpvCommand(binding.command);
      return;
    }

    void window.electronAPI.dispatchSessionAction(binding.actionId, binding.payload);
  }

  function dispatchYomitanPopupKeydown(
    key: string,
    code: string,
    modifiers: Array<'alt' | 'ctrl' | 'shift' | 'meta'>,
    repeat: boolean,
  ) {
    window.dispatchEvent(
      new CustomEvent(YOMITAN_POPUP_COMMAND_EVENT, {
        detail: {
          type: 'forwardKeyDown',
          key,
          code,
          modifiers,
          repeat,
        },
      }),
    );
  }

  function dispatchYomitanPopupVisibility(visible: boolean) {
    window.dispatchEvent(
      new CustomEvent(YOMITAN_POPUP_COMMAND_EVENT, {
        detail: {
          type: 'setVisible',
          visible,
        },
      }),
    );
  }

  function dispatchYomitanPopupMineSelected() {
    window.dispatchEvent(
      new CustomEvent(YOMITAN_POPUP_COMMAND_EVENT, {
        detail: {
          type: 'mineSelected',
        },
      }),
    );
  }

  function dispatchYomitanPopupCycleAudioSource(direction: -1 | 1) {
    window.dispatchEvent(
      new CustomEvent(YOMITAN_POPUP_COMMAND_EVENT, {
        detail: {
          type: 'cycleAudioSource',
          direction,
        },
      }),
    );
  }

  function dispatchYomitanPopupPlayCurrentAudio() {
    window.dispatchEvent(
      new CustomEvent(YOMITAN_POPUP_COMMAND_EVENT, {
        detail: {
          type: 'playCurrentAudio',
        },
      }),
    );
  }

  function dispatchYomitanPopupScrollBy(deltaX: number, deltaY: number) {
    window.dispatchEvent(
      new CustomEvent(YOMITAN_POPUP_COMMAND_EVENT, {
        detail: {
          type: 'scrollBy',
          deltaX,
          deltaY,
        },
      }),
    );
  }

  function dispatchYomitanFrontendScanSelectedText() {
    window.dispatchEvent(
      new CustomEvent(YOMITAN_POPUP_COMMAND_EVENT, {
        detail: {
          type: 'scanSelectedText',
        },
      }),
    );
  }

  function dispatchYomitanFrontendClearActiveTextSource() {
    window.dispatchEvent(
      new CustomEvent(YOMITAN_POPUP_COMMAND_EVENT, {
        detail: {
          type: 'clearActiveTextSource',
        },
      }),
    );
  }

  function isPrimaryModifierPressed(e: KeyboardEvent): boolean {
    return e.ctrlKey || e.metaKey;
  }

  function isKeyboardDrivenModeToggle(e: KeyboardEvent): boolean {
    const isYKey = e.code === 'KeyY' || e.key.toLowerCase() === 'y';
    return isPrimaryModifierPressed(e) && !e.altKey && e.shiftKey && isYKey && !e.repeat;
  }

  function isLookupWindowToggle(e: KeyboardEvent): boolean {
    const isYKey = e.code === 'KeyY' || e.key.toLowerCase() === 'y';
    return isPrimaryModifierPressed(e) && !e.altKey && !e.shiftKey && isYKey && !e.repeat;
  }

  function isControllerModalShortcut(e: KeyboardEvent): boolean {
    return !e.ctrlKey && !e.metaKey && e.altKey && !e.repeat && e.code === 'KeyC';
  }

  function isSubtitleSidebarToggle(e: KeyboardEvent): boolean {
    const toggleKey = ctx.state.subtitleSidebarToggleKey;
    if (!toggleKey) return false;
    const isBackslashConfigured = toggleKey === 'Backslash' || toggleKey === '\\';
    const isBackslashLikeCode = ['Backslash', 'IntlBackslash', 'IntlYen'].includes(e.code);
    const keyMatches =
      toggleKey === e.code ||
      (isBackslashConfigured && isBackslashLikeCode) ||
      (isBackslashConfigured && e.key === '\\') ||
      (toggleKey.length === 1 && e.key === toggleKey);

    return keyMatches && !e.ctrlKey && !e.altKey && !e.metaKey && !e.repeat;
  }

  function isStatsOverlayToggle(e: KeyboardEvent): boolean {
    return (
      e.code === ctx.state.statsToggleKey &&
      !e.ctrlKey &&
      !e.altKey &&
      !e.metaKey &&
      !e.shiftKey &&
      !e.repeat
    );
  }

  function isMarkWatchedKey(e: KeyboardEvent): boolean {
    return (
      e.code === ctx.state.markWatchedKey &&
      !e.ctrlKey &&
      !e.altKey &&
      !e.metaKey &&
      !e.shiftKey &&
      !e.repeat
    );
  }

  async function handleMarkWatched(): Promise<void> {
    const marked = await window.electronAPI.markActiveVideoWatched();
    if (marked) {
      window.electronAPI.sendMpvCommand(['show-text', 'Marked as watched', '1500']);
      window.electronAPI.sendMpvCommand(['playlist-next', 'force']);
    }
  }

  function getSubtitleWordNodes(): HTMLElement[] {
    return Array.from(
      ctx.dom.subtitleRoot.querySelectorAll<HTMLElement>('.word[data-token-index]'),
    );
  }

  function clearKeyboardSelectedWordClasses(
    wordNodes: HTMLElement[] = getSubtitleWordNodes(),
  ): void {
    for (const wordNode of wordNodes) {
      wordNode.classList.remove(KEYBOARD_SELECTED_WORD_CLASS);
    }
  }

  function clearNativeSubtitleSelection(): void {
    window.getSelection()?.removeAllRanges();
    ctx.dom.subtitleRoot.classList.remove('has-selection');
  }

  function syncKeyboardTokenSelection(): void {
    const wordNodes = getSubtitleWordNodes();
    clearKeyboardSelectedWordClasses(wordNodes);

    if (!ctx.state.keyboardDrivenModeEnabled || wordNodes.length === 0) {
      ctx.state.keyboardSelectedWordIndex = null;
      ctx.state.keyboardSelectionVisible = false;
      if (!ctx.state.keyboardDrivenModeEnabled) {
        pendingSelectionAnchorAfterSubtitleSeek = null;
        pendingLookupRefreshAfterSubtitleSeek = false;
        resetSelectionToStartOnNextSubtitleSync = false;
        clearNativeSubtitleSelection();
      }
      return;
    }

    if (pendingSelectionAnchorAfterSubtitleSeek) {
      ctx.state.keyboardSelectedWordIndex =
        pendingSelectionAnchorAfterSubtitleSeek === 'start' ? 0 : wordNodes.length - 1;
      ctx.state.keyboardSelectionVisible = true;
      pendingSelectionAnchorAfterSubtitleSeek = null;
      resetSelectionToStartOnNextSubtitleSync = false;
      const shouldRefreshLookup =
        pendingLookupRefreshAfterSubtitleSeek &&
        (ctx.state.yomitanPopupVisible || isYomitanPopupVisible(document));
      pendingLookupRefreshAfterSubtitleSeek = false;
      if (shouldRefreshLookup) {
        queueMicrotask(() => {
          triggerLookupForSelectedWord();
        });
      }
    }

    if (resetSelectionToStartOnNextSubtitleSync) {
      ctx.state.keyboardSelectedWordIndex = 0;
      ctx.state.keyboardSelectionVisible = true;
      resetSelectionToStartOnNextSubtitleSync = false;
    }

    const selectedIndex = Math.min(
      Math.max(ctx.state.keyboardSelectedWordIndex ?? 0, 0),
      wordNodes.length - 1,
    );
    ctx.state.keyboardSelectedWordIndex = selectedIndex;
    const selectedWordNode = wordNodes[selectedIndex];
    if (selectedWordNode && ctx.state.keyboardSelectionVisible) {
      selectedWordNode.classList.add(KEYBOARD_SELECTED_WORD_CLASS);
    }
  }

  function setKeyboardDrivenModeEnabled(enabled: boolean): void {
    ctx.state.keyboardDrivenModeEnabled = enabled;
    ctx.state.keyboardSelectionVisible = enabled;
    if (!enabled) {
      ctx.state.keyboardSelectedWordIndex = null;
      pendingSelectionAnchorAfterSubtitleSeek = null;
      pendingLookupRefreshAfterSubtitleSeek = false;
      resetSelectionToStartOnNextSubtitleSync = false;
      clearNativeSubtitleSelection();
    }
    syncKeyboardTokenSelection();
  }

  function toggleKeyboardDrivenMode(): void {
    setKeyboardDrivenModeEnabled(!ctx.state.keyboardDrivenModeEnabled);
  }

  function moveKeyboardSelection(
    delta: -1 | 1,
  ): 'moved' | 'start-boundary' | 'end-boundary' | 'no-words' {
    const wordNodes = getSubtitleWordNodes();
    if (wordNodes.length === 0) {
      ctx.state.keyboardSelectedWordIndex = null;
      syncKeyboardTokenSelection();
      return 'no-words';
    }

    const currentIndex = Math.min(
      Math.max(ctx.state.keyboardSelectedWordIndex ?? 0, 0),
      wordNodes.length - 1,
    );
    if (delta < 0 && currentIndex <= 0) {
      return 'start-boundary';
    }
    if (delta > 0 && currentIndex >= wordNodes.length - 1) {
      return 'end-boundary';
    }

    const nextIndex = currentIndex + delta;
    ctx.state.keyboardSelectedWordIndex = nextIndex;
    ctx.state.keyboardSelectionVisible = true;
    syncKeyboardTokenSelection();
    return 'moved';
  }

  function seekAdjacentSubtitleAndQueueSelection(delta: -1 | 1, popupVisible: boolean): void {
    pendingSelectionAnchorAfterSubtitleSeek = delta > 0 ? 'start' : 'end';
    pendingLookupRefreshAfterSubtitleSeek = popupVisible;
    void options
      .getPlaybackPaused()
      .then((paused) => {
        window.electronAPI.sendMpvCommand(['sub-seek', delta]);
        if (paused !== false) {
          window.electronAPI.sendMpvCommand(['set_property', 'pause', 'yes']);
        }
      })
      .catch(() => {
        window.electronAPI.sendMpvCommand(['sub-seek', delta]);
      });
  }

  function isSubtitleSeekCommand(
    command: (string | number)[] | undefined,
  ): command is [string, number] {
    return Array.isArray(command) && command[0] === 'sub-seek' && typeof command[1] === 'number';
  }

  function dispatchConfiguredMpvCommand(command: (string | number)[]): void {
    if (!isSubtitleSeekCommand(command)) {
      window.electronAPI.sendMpvCommand(command);
      return;
    }

    void options
      .getPlaybackPaused()
      .then((paused) => {
        window.electronAPI.sendMpvCommand(command);
        if (paused !== false) {
          window.electronAPI.sendMpvCommand(['set_property', 'pause', 'yes']);
        }
      })
      .catch(() => {
        window.electronAPI.sendMpvCommand(command);
      });
  }

  type ScanModifierState = {
    shiftKey?: boolean;
    ctrlKey?: boolean;
    altKey?: boolean;
    metaKey?: boolean;
  };

  function emitSyntheticScanEvents(
    target: Element,
    clientX: number,
    clientY: number,
    modifiers: ScanModifierState = {},
  ): void {
    if (typeof PointerEvent === 'function') {
      const pointerEventInit = {
        bubbles: true,
        cancelable: true,
        composed: true,
        clientX,
        clientY,
        pointerType: 'mouse',
        isPrimary: true,
        button: 0,
        buttons: 0,
        shiftKey: modifiers.shiftKey ?? false,
        ctrlKey: modifiers.ctrlKey ?? false,
        altKey: modifiers.altKey ?? false,
        metaKey: modifiers.metaKey ?? false,
      } satisfies PointerEventInit;

      target.dispatchEvent(new PointerEvent('pointerover', pointerEventInit));
      target.dispatchEvent(new PointerEvent('pointermove', pointerEventInit));
      target.dispatchEvent(new PointerEvent('pointerdown', { ...pointerEventInit, buttons: 1 }));
      target.dispatchEvent(new PointerEvent('pointerup', pointerEventInit));
    }

    if (typeof MouseEvent === 'function') {
      const mouseEventInit = {
        bubbles: true,
        cancelable: true,
        composed: true,
        clientX,
        clientY,
        button: 0,
        buttons: 0,
        shiftKey: modifiers.shiftKey ?? false,
        ctrlKey: modifiers.ctrlKey ?? false,
        altKey: modifiers.altKey ?? false,
        metaKey: modifiers.metaKey ?? false,
      } satisfies MouseEventInit;
      target.dispatchEvent(new MouseEvent('mousemove', mouseEventInit));
      target.dispatchEvent(new MouseEvent('mousedown', { ...mouseEventInit, buttons: 1 }));
      target.dispatchEvent(new MouseEvent('mouseup', mouseEventInit));
      target.dispatchEvent(new MouseEvent('click', mouseEventInit));
    }
  }

  function emitLookupScanFallback(target: Element, clientX: number, clientY: number): void {
    emitSyntheticScanEvents(target, clientX, clientY, {});
  }

  function selectWordNodeText(wordNode: HTMLElement): void {
    const selection = window.getSelection();
    if (!selection) return;
    const range = document.createRange();
    range.selectNodeContents(wordNode);
    selection.removeAllRanges();
    selection.addRange(range);
    ctx.dom.subtitleRoot.classList.add('has-selection');
  }

  function triggerLookupForSelectedWord(): boolean {
    const wordNodes = getSubtitleWordNodes();
    if (wordNodes.length === 0) {
      return false;
    }

    const selectedIndex = Math.min(
      Math.max(ctx.state.keyboardSelectedWordIndex ?? 0, 0),
      wordNodes.length - 1,
    );
    ctx.state.keyboardSelectedWordIndex = selectedIndex;
    const selectedWordNode = wordNodes[selectedIndex];
    if (!selectedWordNode) return false;

    ctx.state.keyboardSelectionVisible = true;
    syncKeyboardTokenSelection();
    selectWordNodeText(selectedWordNode);

    const rect = selectedWordNode.getBoundingClientRect();
    const clientX = rect.left + rect.width / 2;
    const clientY = rect.top + rect.height / 2;

    dispatchYomitanFrontendScanSelectedText();
    if (ctx.state.keyboardDrivenModeEnabled) {
      // Keep overlay as the keyboard focus owner so token navigation can continue
      // while the popup is visible.
      queueMicrotask(() => {
        scheduleOverlayFocusReclaim(8);
      });
    }
    // Fallback only if the explicit scan path did not open popup quickly.
    if (lookupScanFallbackTimer !== null) clearTimeout(lookupScanFallbackTimer);
    lookupScanFallbackTimer = setTimeout(() => {
      lookupScanFallbackTimer = null;
      if (ctx.state.yomitanPopupVisible || isYomitanPopupVisible(document)) {
        return;
      }
      // Dispatch directly on the selected token span; when overlay pointer-events are disabled,
      // elementFromPoint may resolve to the underlying video surface instead.
      emitLookupScanFallback(selectedWordNode, clientX, clientY);
    }, 60);
    return true;
  }

  function handleKeyboardModeToggleRequested(): void {
    toggleKeyboardDrivenMode();
  }

  function handleSubtitleContentUpdated(): void {
    if (!ctx.state.keyboardDrivenModeEnabled) {
      dispatchYomitanFrontendClearActiveTextSource();
      clearNativeSubtitleSelection();
      return;
    }
    if (pendingSelectionAnchorAfterSubtitleSeek) {
      return;
    }
    resetSelectionToStartOnNextSubtitleSync = true;
  }

  function handleLookupWindowToggleRequested(): void {
    if (ctx.state.yomitanPopupVisible || isYomitanPopupVisible(document)) {
      closeLookupWindow();
      return;
    }
    triggerLookupForSelectedWord();
  }

  function closeLookupWindow(): boolean {
    if (!ctx.state.yomitanPopupVisible && !isYomitanPopupVisible(document)) {
      return false;
    }

    if (lookupScanFallbackTimer !== null) {
      clearTimeout(lookupScanFallbackTimer);
      lookupScanFallbackTimer = null;
    }
    dispatchYomitanPopupVisibility(false);
    dispatchYomitanFrontendClearActiveTextSource();
    clearNativeSubtitleSelection();
    if (ctx.state.keyboardDrivenModeEnabled) {
      queueMicrotask(() => {
        restoreOverlayKeyboardFocus();
      });
    }
    return true;
  }

  function moveSelectionForController(delta: -1 | 1): boolean {
    if (!ctx.state.keyboardDrivenModeEnabled) {
      return false;
    }

    const popupVisible = ctx.state.yomitanPopupVisible || isYomitanPopupVisible(document);
    const result = moveKeyboardSelection(delta);
    if (result === 'no-words') {
      seekAdjacentSubtitleAndQueueSelection(delta, popupVisible);
      return true;
    }

    if (result === 'start-boundary' || result === 'end-boundary') {
      seekAdjacentSubtitleAndQueueSelection(delta, popupVisible);
    } else if (popupVisible && result === 'moved') {
      triggerLookupForSelectedWord();
    }

    return true;
  }

  function forwardPopupKeydownForController(
    key: string,
    code: string,
    repeat: boolean = true,
  ): boolean {
    if (!ctx.state.yomitanPopupVisible && !isYomitanPopupVisible(document)) {
      return false;
    }
    dispatchYomitanPopupKeydown(key, code, [], repeat);
    return true;
  }

  function mineSelectedFromController(): boolean {
    if (!ctx.state.yomitanPopupVisible && !isYomitanPopupVisible(document)) {
      return false;
    }
    dispatchYomitanPopupMineSelected();
    return true;
  }

  function cyclePopupAudioSourceForController(direction: -1 | 1): boolean {
    if (!ctx.state.yomitanPopupVisible && !isYomitanPopupVisible(document)) {
      return false;
    }
    dispatchYomitanPopupCycleAudioSource(direction);
    return true;
  }

  function playCurrentAudioForController(): boolean {
    if (!ctx.state.yomitanPopupVisible && !isYomitanPopupVisible(document)) {
      return false;
    }
    dispatchYomitanPopupPlayCurrentAudio();
    return true;
  }

  function scrollPopupByController(deltaX: number, deltaY: number): boolean {
    if (!ctx.state.yomitanPopupVisible && !isYomitanPopupVisible(document)) {
      return false;
    }
    dispatchYomitanPopupScrollBy(deltaX, deltaY);
    return true;
  }

  function restoreOverlayKeyboardFocus(): void {
    void window.electronAPI.focusMainWindow();
    window.focus();
    ctx.dom.overlay.focus({ preventScroll: true });
  }

  function scheduleOverlayFocusReclaim(attempts: number = 0): void {
    if (!ctx.state.keyboardDrivenModeEnabled) {
      return;
    }
    restoreOverlayKeyboardFocus();
    if (attempts <= 0) {
      return;
    }

    let remaining = attempts;
    const reclaim = () => {
      if (!ctx.state.keyboardDrivenModeEnabled) {
        return;
      }
      if (!ctx.state.yomitanPopupVisible && !isYomitanPopupVisible(document)) {
        return;
      }
      restoreOverlayKeyboardFocus();
      remaining -= 1;
      if (remaining > 0) {
        setTimeout(reclaim, 25);
      }
    };

    setTimeout(reclaim, 25);
  }

  function handleKeyboardDrivenModeNavigation(e: KeyboardEvent): boolean {
    if (e.ctrlKey || e.metaKey || e.altKey || e.shiftKey) {
      return false;
    }

    const key = e.code;
    if (key === 'ArrowLeft') {
      const result = moveKeyboardSelection(-1);
      if (result === 'start-boundary' || result === 'no-words') {
        seekAdjacentSubtitleAndQueueSelection(-1, false);
      }
      return true;
    }
    if (key === 'ArrowRight' || key === 'KeyL') {
      const result = moveKeyboardSelection(1);
      if (result === 'end-boundary' || result === 'no-words') {
        seekAdjacentSubtitleAndQueueSelection(1, false);
      }
      return true;
    }
    return false;
  }

  function handleKeyboardDrivenModeLookupControls(e: KeyboardEvent): boolean {
    if (!ctx.state.keyboardDrivenModeEnabled) {
      return false;
    }
    if (e.ctrlKey || e.metaKey || e.altKey || e.shiftKey) {
      return false;
    }

    const key = e.code;
    const popupVisible = ctx.state.yomitanPopupVisible || isYomitanPopupVisible(document);
    if (key === 'ArrowLeft' || key === 'KeyH') {
      const result = moveKeyboardSelection(-1);
      if (result === 'start-boundary' || result === 'no-words') {
        seekAdjacentSubtitleAndQueueSelection(-1, popupVisible);
      } else if (popupVisible && result === 'moved') {
        triggerLookupForSelectedWord();
      }
      return true;
    }

    if (key === 'ArrowRight' || key === 'KeyL') {
      const result = moveKeyboardSelection(1);
      if (result === 'end-boundary' || result === 'no-words') {
        seekAdjacentSubtitleAndQueueSelection(1, popupVisible);
      } else if (popupVisible && result === 'moved') {
        triggerLookupForSelectedWord();
      }
      return true;
    }

    return false;
  }

  function handleYomitanPopupKeybind(e: KeyboardEvent): boolean {
    const modifierOnlyCodes = new Set([
      'ShiftLeft',
      'ShiftRight',
      'ControlLeft',
      'ControlRight',
      'AltLeft',
      'AltRight',
      'MetaLeft',
      'MetaRight',
    ]);
    if (modifierOnlyCodes.has(e.code)) return false;

    const keyString = keyEventToString(e);
    if (ctx.state.sessionBindingMap.has(keyString)) {
      return false;
    }

    if (!e.ctrlKey && !e.metaKey && !e.altKey && e.code === 'KeyM') {
      if (e.repeat) return false;
      dispatchYomitanPopupMineSelected();
      return true;
    }

    const modifiers: Array<'alt' | 'ctrl' | 'shift' | 'meta'> = [];
    if (e.altKey) modifiers.push('alt');
    if (e.ctrlKey) modifiers.push('ctrl');
    if (e.shiftKey) modifiers.push('shift');
    if (e.metaKey) modifiers.push('meta');

    dispatchYomitanPopupKeydown(e.key, e.code, modifiers, e.repeat);
    return true;
  }

  function resolveSessionHelpChordBinding(): {
    bindingKey: 'KeyH' | 'KeyK';
    fallbackUsed: boolean;
    fallbackUnavailable: boolean;
  } {
    const firstChoice = 'KeyH';
    if (!ctx.state.sessionBindingMap.has('KeyH')) {
      return {
        bindingKey: firstChoice,
        fallbackUsed: false,
        fallbackUnavailable: false,
      };
    }

    if (!ctx.state.sessionBindingMap.has('KeyK')) {
      return {
        bindingKey: 'KeyK',
        fallbackUsed: true,
        fallbackUnavailable: false,
      };
    }

    return {
      bindingKey: 'KeyK',
      fallbackUsed: true,
      fallbackUnavailable: true,
    };
  }

  function applySessionHelpChordBinding(): void {
    CHORD_MAP.delete('KeyH');
    CHORD_MAP.delete('KeyK');
    const info = resolveSessionHelpChordBinding();
    CHORD_MAP.set(info.bindingKey, {
      type: 'electron',
      action: () => {
        options.openSessionHelpModal(info);
      },
    });
  }

  function resetChord(): void {
    ctx.state.chordPending = false;
    if (ctx.state.chordTimeout !== null) {
      clearTimeout(ctx.state.chordTimeout);
      ctx.state.chordTimeout = null;
    }
  }

  async function setupMpvInputForwarding(): Promise<void> {
    const [sessionBindings, shortcuts, statsToggleKey, markWatchedKey] = await Promise.all([
      window.electronAPI.getSessionBindings(),
      window.electronAPI.getConfiguredShortcuts(),
      window.electronAPI.getStatsToggleKey(),
      window.electronAPI.getMarkWatchedKey(),
    ]);
    updateSessionBindings(sessionBindings);
    updateConfiguredShortcuts(shortcuts);
    ctx.state.statsToggleKey = statsToggleKey;
    ctx.state.markWatchedKey = markWatchedKey;
    syncKeyboardTokenSelection();

    const subtitleMutationObserver = new MutationObserver(() => {
      syncKeyboardTokenSelection();
    });
    subtitleMutationObserver.observe(ctx.dom.subtitleRoot, {
      childList: true,
      subtree: true,
    });

    window.addEventListener(YOMITAN_POPUP_HIDDEN_EVENT, () => {
      clearNativeSubtitleSelection();
      if (!ctx.state.keyboardDrivenModeEnabled) {
        syncKeyboardTokenSelection();
        return;
      }
      restoreOverlayKeyboardFocus();
    });
    window.addEventListener(YOMITAN_POPUP_SHOWN_EVENT, () => {
      if (!ctx.state.keyboardDrivenModeEnabled) {
        return;
      }
      queueMicrotask(() => {
        scheduleOverlayFocusReclaim(8);
      });
    });

    document.addEventListener(
      'focusin',
      (e: FocusEvent) => {
        if (!ctx.state.keyboardDrivenModeEnabled) {
          return;
        }
        const target = e.target;
        if (
          target &&
          typeof target === 'object' &&
          'tagName' in target &&
          isYomitanPopupIframe(target as Element)
        ) {
          queueMicrotask(() => {
            scheduleOverlayFocusReclaim(8);
          });
        }
      },
      true,
    );

    document.addEventListener('keydown', (e: KeyboardEvent) => {
      if (isKeyboardDrivenModeToggle(e) && ctx.platform.isModalLayer) {
        e.preventDefault();
        handleKeyboardModeToggleRequested();
        return;
      }

      if (isLookupWindowToggle(e)) {
        e.preventDefault();
        handleLookupWindowToggleRequested();
        return;
      }

      if (ctx.state.playlistBrowserModalOpen) {
        if (options.handlePlaylistBrowserKeydown(e)) {
          return;
        }
      }

      if (handleKeyboardDrivenModeLookupControls(e)) {
        e.preventDefault();
        return;
      }

      if (ctx.state.runtimeOptionsModalOpen) {
        options.handleRuntimeOptionsKeydown(e);
        return;
      }
      if (ctx.state.subsyncModalOpen) {
        options.handleSubsyncKeydown(e);
        return;
      }
      if (ctx.state.kikuModalOpen) {
        options.handleKikuKeydown(e);
        return;
      }
      if (ctx.state.jimakuModalOpen) {
        options.handleJimakuKeydown(e);
        return;
      }
      if (ctx.state.youtubePickerModalOpen) {
        if (options.handleYoutubePickerKeydown(e)) {
          return;
        }
      }
      if (ctx.state.controllerSelectModalOpen) {
        options.handleControllerSelectKeydown(e);
        return;
      }
      if (ctx.state.controllerDebugModalOpen) {
        options.handleControllerDebugKeydown(e);
        return;
      }
      if (ctx.state.sessionHelpModalOpen) {
        options.handleSessionHelpKeydown(e);
        return;
      }

      if (isTextEntryTarget(e.target)) {
        return;
      }

      if (handlePendingNumericSelection(e)) {
        return;
      }

      if (isStatsOverlayToggle(e)) {
        e.preventDefault();
        window.electronAPI.toggleStatsOverlay();
        return;
      }

      if (isMarkWatchedKey(e)) {
        e.preventDefault();
        void handleMarkWatched();
        return;
      }

      if (isSubtitleSidebarToggle(e)) {
        e.preventDefault();
        options.toggleSubtitleSidebarModal?.();
        return;
      }

      if (
        (ctx.state.yomitanPopupVisible || isYomitanPopupVisible(document)) &&
        !isControllerModalShortcut(e)
      ) {
        if (handleYomitanPopupKeybind(e)) {
          e.preventDefault();
          return;
        }
      }

      if (ctx.state.keyboardDrivenModeEnabled && handleKeyboardDrivenModeNavigation(e)) {
        e.preventDefault();
        return;
      }

      if (ctx.state.chordPending) {
        const modifierKeys = [
          'ShiftLeft',
          'ShiftRight',
          'ControlLeft',
          'ControlRight',
          'AltLeft',
          'AltRight',
          'MetaLeft',
          'MetaRight',
        ];
        if (modifierKeys.includes(e.code)) {
          return;
        }

        e.preventDefault();
        const secondKey = keyEventToString(e);
        const action = CHORD_MAP.get(secondKey);
        resetChord();
        if (action) {
          if (action.type === 'mpv' && action.command) {
            window.electronAPI.sendMpvCommand(action.command);
          } else if (action.type === 'electron' && action.action) {
            action.action();
          }
        }
        return;
      }

      if (e.code === 'KeyY' && !e.ctrlKey && !e.altKey && !e.shiftKey && !e.metaKey && !e.repeat) {
        e.preventDefault();
        applySessionHelpChordBinding();
        ctx.state.chordPending = true;
        ctx.state.chordTimeout = setTimeout(() => {
          resetChord();
        }, CHORD_TIMEOUT_MS);
        return;
      }

      if ((e.ctrlKey || e.metaKey) && !e.altKey && !e.shiftKey && e.code === 'KeyA' && !e.repeat) {
        e.preventDefault();
        options.appendClipboardVideoToQueue();
        return;
      }

      if (isControllerModalShortcut(e)) {
        e.preventDefault();
        if (e.shiftKey) {
          options.openControllerDebugModal();
        } else {
          options.openControllerSelectModal();
        }
        return;
      }

      const keyString = keyEventToString(e);
      const binding = ctx.state.sessionBindingMap.get(keyString);
      if (binding) {
        e.preventDefault();
        dispatchSessionBinding(binding);
      }
    });

    document.addEventListener('mousedown', (e: MouseEvent) => {
      if (e.button === 2 && !isInteractiveTarget(e.target)) {
        e.preventDefault();
        window.electronAPI.sendMpvCommand(['cycle', 'pause']);
      }
    });

    document.addEventListener('contextmenu', (e: Event) => {
      if (!isInteractiveTarget(e.target)) {
        e.preventDefault();
      }
    });
  }

  return {
    beginSessionNumericSelection,
    setupMpvInputForwarding,
    refreshConfiguredShortcuts,
    updateSessionBindings,
    syncKeyboardTokenSelection,
    handleSubtitleContentUpdated,
    handleKeyboardModeToggleRequested,
    handleLookupWindowToggleRequested,
    closeLookupWindow,
    moveSelectionForController,
    forwardPopupKeydownForController,
    mineSelectedFromController,
    cyclePopupAudioSourceForController,
    playCurrentAudioForController,
    scrollPopupByController,
  };
}
