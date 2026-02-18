import type { Keybinding } from '../../types';
import type { RendererContext } from '../context';

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
    saveInvisiblePositionEdit: () => void;
    cancelInvisiblePositionEdit: () => void;
    setInvisiblePositionEditMode: (enabled: boolean) => void;
    applyInvisibleSubtitleOffsetPosition: () => void;
    updateInvisiblePositionEditHud: () => void;
    appendClipboardVideoToQueue: () => void;
  },
) {
  // Timeout for the modal chord capture window (e.g. Y followed by H/K).
  const CHORD_TIMEOUT_MS = 1000;

  const CHORD_MAP = new Map<
    string,
    { type: 'mpv' | 'electron'; command?: string[]; action?: () => void }
  >([
    ['KeyS', { type: 'mpv', command: ['script-message', 'subminer-start'] }],
    ['Shift+KeyS', { type: 'mpv', command: ['script-message', 'subminer-stop'] }],
    ['KeyT', { type: 'mpv', command: ['script-message', 'subminer-toggle'] }],
    ['KeyI', { type: 'mpv', command: ['script-message', 'subminer-toggle-invisible'] }],
    ['Shift+KeyI', { type: 'mpv', command: ['script-message', 'subminer-show-invisible'] }],
    ['KeyU', { type: 'mpv', command: ['script-message', 'subminer-hide-invisible'] }],
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
    if (target.tagName === 'IFRAME' && target.id?.startsWith('yomitan-popup')) {
      return true;
    }
    if (target.closest && target.closest('iframe[id^="yomitan-popup"]')) return true;
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

  function isInvisiblePositionToggleShortcut(e: KeyboardEvent): boolean {
    return (
      e.code === ctx.platform.invisiblePositionEditToggleCode &&
      !e.altKey &&
      e.shiftKey &&
      (e.ctrlKey || e.metaKey)
    );
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

  function handleInvisiblePositionEditKeydown(e: KeyboardEvent): boolean {
    if (!ctx.platform.isInvisibleLayer) return false;

    if (isInvisiblePositionToggleShortcut(e)) {
      e.preventDefault();
      if (ctx.state.invisiblePositionEditMode) {
        options.cancelInvisiblePositionEdit();
      } else {
        options.setInvisiblePositionEditMode(true);
      }
      return true;
    }

    if (!ctx.state.invisiblePositionEditMode) return false;

    const step = e.shiftKey
      ? ctx.platform.invisiblePositionStepFastPx
      : ctx.platform.invisiblePositionStepPx;

    if (e.key === 'Escape') {
      e.preventDefault();
      options.cancelInvisiblePositionEdit();
      return true;
    }

    if (e.key === 'Enter' || ((e.ctrlKey || e.metaKey) && e.code === 'KeyS')) {
      e.preventDefault();
      options.saveInvisiblePositionEdit();
      return true;
    }

    if (
      e.key === 'ArrowUp' ||
      e.key === 'ArrowDown' ||
      e.key === 'ArrowLeft' ||
      e.key === 'ArrowRight' ||
      e.key === 'h' ||
      e.key === 'j' ||
      e.key === 'k' ||
      e.key === 'l' ||
      e.key === 'H' ||
      e.key === 'J' ||
      e.key === 'K' ||
      e.key === 'L'
    ) {
      e.preventDefault();
      if (e.key === 'ArrowUp' || e.key === 'k' || e.key === 'K') {
        ctx.state.invisibleSubtitleOffsetYPx += step;
      } else if (e.key === 'ArrowDown' || e.key === 'j' || e.key === 'J') {
        ctx.state.invisibleSubtitleOffsetYPx -= step;
      } else if (e.key === 'ArrowLeft' || e.key === 'h' || e.key === 'H') {
        ctx.state.invisibleSubtitleOffsetXPx -= step;
      } else if (e.key === 'ArrowRight' || e.key === 'l' || e.key === 'L') {
        ctx.state.invisibleSubtitleOffsetXPx += step;
      }
      options.applyInvisibleSubtitleOffsetPosition();
      options.updateInvisiblePositionEditHud();
      return true;
    }

    return true;
  }

  function resetChord(): void {
    ctx.state.chordPending = false;
    if (ctx.state.chordTimeout !== null) {
      clearTimeout(ctx.state.chordTimeout);
      ctx.state.chordTimeout = null;
    }
  }

  async function setupMpvInputForwarding(): Promise<void> {
    const keybindings: Keybinding[] = await window.electronAPI.getKeybindings();
    ctx.state.keybindingsMap = new Map();
    for (const binding of keybindings) {
      if (binding.command) {
        ctx.state.keybindingsMap.set(binding.key, binding.command);
      }
    }

    document.addEventListener('keydown', (e: KeyboardEvent) => {
      const yomitanPopup = document.querySelector('iframe[id^="yomitan-popup"]');
      if (yomitanPopup) return;
      if (handleInvisiblePositionEditKeydown(e)) return;

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

      if (
        (e.ctrlKey || e.metaKey) &&
        !e.altKey &&
        !e.shiftKey &&
        e.code === 'KeyA' &&
        !e.repeat
      ) {
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

  return {
    setupMpvInputForwarding,
  };
}
