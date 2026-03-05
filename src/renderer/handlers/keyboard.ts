import type { Keybinding } from '../../types';
import type { RendererContext } from '../context';
import {
  YOMITAN_POPUP_HIDDEN_EVENT,
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
    handleSessionHelpKeydown: (e: KeyboardEvent) => boolean;
    openSessionHelpModal: (opening: {
      bindingKey: 'KeyH' | 'KeyK';
      fallbackUsed: boolean;
      fallbackUnavailable: boolean;
    }) => void;
    appendClipboardVideoToQueue: () => void;
  },
) {
  // Timeout for the modal chord capture window (e.g. Y followed by H/K).
  const CHORD_TIMEOUT_MS = 1000;
  const KEYBOARD_SELECTED_WORD_CLASS = 'keyboard-selected';

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

  function dispatchYomitanFrontendScanSelectedText() {
    window.dispatchEvent(
      new CustomEvent(YOMITAN_POPUP_COMMAND_EVENT, {
        detail: {
          type: 'scanSelectedText',
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

  function getSubtitleWordNodes(): HTMLElement[] {
    return Array.from(ctx.dom.subtitleRoot.querySelectorAll<HTMLElement>('.word[data-token-index]'));
  }

  function syncKeyboardTokenSelection(): void {
    const wordNodes = getSubtitleWordNodes();
    for (const wordNode of wordNodes) {
      wordNode.classList.remove(KEYBOARD_SELECTED_WORD_CLASS);
    }

    if (!ctx.state.keyboardDrivenModeEnabled || wordNodes.length === 0) {
      ctx.state.keyboardSelectedWordIndex = null;
      return;
    }

    const selectedIndex = Math.min(
      Math.max(ctx.state.keyboardSelectedWordIndex ?? 0, 0),
      wordNodes.length - 1,
    );
    ctx.state.keyboardSelectedWordIndex = selectedIndex;
    const selectedWordNode = wordNodes[selectedIndex];
    if (selectedWordNode) {
      selectedWordNode.classList.add(KEYBOARD_SELECTED_WORD_CLASS);
    }
  }

  function setKeyboardDrivenModeEnabled(enabled: boolean): void {
    ctx.state.keyboardDrivenModeEnabled = enabled;
    if (!enabled) {
      ctx.state.keyboardSelectedWordIndex = null;
    }
    syncKeyboardTokenSelection();
  }

  function toggleKeyboardDrivenMode(): void {
    setKeyboardDrivenModeEnabled(!ctx.state.keyboardDrivenModeEnabled);
  }

  function moveKeyboardSelection(delta: -1 | 1): boolean {
    const wordNodes = getSubtitleWordNodes();
    if (wordNodes.length === 0) {
      ctx.state.keyboardSelectedWordIndex = null;
      syncKeyboardTokenSelection();
      return true;
    }

    const currentIndex = ctx.state.keyboardSelectedWordIndex ?? 0;
    const nextIndex = Math.min(Math.max(currentIndex + delta, 0), wordNodes.length - 1);
    ctx.state.keyboardSelectedWordIndex = nextIndex;
    syncKeyboardTokenSelection();
    return true;
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
    if (typeof PointerEvent !== 'undefined') {
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

    syncKeyboardTokenSelection();
    selectWordNodeText(selectedWordNode);

    const rect = selectedWordNode.getBoundingClientRect();
    const clientX = rect.left + rect.width / 2;
    const clientY = rect.top + rect.height / 2;

    dispatchYomitanFrontendScanSelectedText();
    // Fallback only if the explicit scan path did not open popup quickly.
    setTimeout(() => {
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

  function handleLookupWindowToggleRequested(): void {
    if (ctx.state.yomitanPopupVisible) {
      dispatchYomitanPopupVisibility(false);
      if (ctx.state.keyboardDrivenModeEnabled) {
        queueMicrotask(() => {
          restoreOverlayKeyboardFocus();
        });
      }
      return;
    }
    triggerLookupForSelectedWord();
  }

  function restoreOverlayKeyboardFocus(): void {
    void window.electronAPI.focusMainWindow();
    window.focus();
    ctx.dom.overlay.focus({ preventScroll: true });
  }

  function handleKeyboardDrivenModeNavigation(e: KeyboardEvent): boolean {
    if (e.ctrlKey || e.metaKey || e.altKey || e.shiftKey) {
      return false;
    }

    const key = e.code;
    if (key === 'ArrowLeft' || key === 'ArrowUp' || key === 'KeyH' || key === 'KeyK') {
      return moveKeyboardSelection(-1);
    }
    if (key === 'ArrowRight' || key === 'ArrowDown' || key === 'KeyL' || key === 'KeyJ') {
      return moveKeyboardSelection(1);
    }
    return false;
  }

  function handleYomitanPopupKeybind(e: KeyboardEvent): boolean {
    if (e.repeat) return false;
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

    if (!e.ctrlKey && !e.metaKey && !e.altKey && e.code === 'KeyM') {
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
    if (!ctx.state.keybindingsMap.has('KeyH')) {
      return {
        bindingKey: firstChoice,
        fallbackUsed: false,
        fallbackUnavailable: false,
      };
    }

    if (ctx.state.keybindingsMap.has('KeyK')) {
      return {
        bindingKey: 'KeyK',
        fallbackUsed: true,
        fallbackUnavailable: true,
      };
    }

    return {
      bindingKey: 'KeyK',
      fallbackUsed: true,
      fallbackUnavailable: false,
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
    updateKeybindings(await window.electronAPI.getKeybindings());
    syncKeyboardTokenSelection();

    const subtitleMutationObserver = new MutationObserver(() => {
      syncKeyboardTokenSelection();
    });
    subtitleMutationObserver.observe(ctx.dom.subtitleRoot, {
      childList: true,
      subtree: true,
    });

    window.addEventListener(YOMITAN_POPUP_HIDDEN_EVENT, () => {
      if (!ctx.state.keyboardDrivenModeEnabled) {
        return;
      }
      restoreOverlayKeyboardFocus();
    });

    document.addEventListener('keydown', (e: KeyboardEvent) => {
      if (isKeyboardDrivenModeToggle(e)) {
        e.preventDefault();
        handleKeyboardModeToggleRequested();
        return;
      }

      if (isLookupWindowToggle(e)) {
        e.preventDefault();
        handleLookupWindowToggleRequested();
        return;
      }

      if (ctx.state.yomitanPopupVisible || isYomitanPopupVisible(document)) {
        if (handleYomitanPopupKeybind(e)) {
          e.preventDefault();
        }
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
      if (ctx.state.sessionHelpModalOpen) {
        options.handleSessionHelpKeydown(e);
        return;
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

      const keyString = keyEventToString(e);
      const command = ctx.state.keybindingsMap.get(keyString);

      if (command) {
        e.preventDefault();
        window.electronAPI.sendMpvCommand(command);
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

  function updateKeybindings(keybindings: Keybinding[]): void {
    ctx.state.keybindingsMap = new Map();
    for (const binding of keybindings) {
      if (binding.command) {
        ctx.state.keybindingsMap.set(binding.key, binding.command);
      }
    }
  }

  return {
    setupMpvInputForwarding,
    updateKeybindings,
    syncKeyboardTokenSelection,
    handleKeyboardModeToggleRequested,
    handleLookupWindowToggleRequested,
  };
}
