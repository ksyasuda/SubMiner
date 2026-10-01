import {
  MPV_MOUSE_BUTTON_BY_BUTTON,
  keyboardEventToMpvKey,
  normalizeMpvInputKey,
  wheelEventToMpvWheel,
} from '../../shared/mpv-input-bindings';
import type { MpvInputBindingsSnapshot } from '../../types/session-bindings';

type ModifierState = Pick<KeyboardEvent, 'ctrlKey' | 'altKey' | 'shiftKey' | 'metaKey'>;
type ForwardedKeyEvent = Parameters<typeof keyboardEventToMpvKey>[0] &
  Pick<KeyboardEvent, 'code' | 'repeat' | 'defaultPrevented' | 'preventDefault'>;
type ForwardedWheelEvent = Parameters<typeof wheelEventToMpvWheel>[0] &
  ModifierState &
  Pick<WheelEvent, 'defaultPrevented' | 'preventDefault'>;
type ForwardedMouseEvent = ModifierState &
  Pick<MouseEvent, 'button' | 'defaultPrevented' | 'preventDefault'>;

export function createMpvInputForwarding(deps: {
  load: () => Promise<MpvInputBindingsSnapshot>;
  send: (command: (string | number)[]) => void;
}) {
  let keys = new Set<string>();
  let blockedKeys: MpvInputBindingsSnapshot['blockedKeys'] = [];
  const heldKeys = new Map<string, string>();
  let generation = 0;
  let disposed = false;
  let pending: Promise<void> | null = null;

  function releaseAll(): void {
    for (const key of heldKeys.values()) deps.send(['keyup', key]);
    heldKeys.clear();
  }

  function refresh(): Promise<void> {
    if (disposed) return Promise.resolve();
    generation += 1;
    keys.clear();
    blockedKeys = [];
    releaseAll();
    if (pending) return pending;
    pending = (async () => {
      let requestedGeneration: number;
      do {
        requestedGeneration = generation;
        try {
          const snapshot = await deps.load();
          if (!disposed && requestedGeneration === generation) {
            keys = new Set(snapshot.keys);
            blockedKeys = snapshot.blockedKeys;
          }
        } catch {
          // Discovery is optional. Keep the existing overlay controls available.
        }
      } while (!disposed && requestedGeneration !== generation);
    })().finally(() => {
      pending = null;
    });
    return pending;
  }

  // Keys claimed by SubMiner's configured keybindings stay with SubMiner.
  function isBlocked(code: string, event: ModifierState): boolean {
    return blockedKeys.some(
      ({ code: blockedCode, modifiers }) =>
        blockedCode === code &&
        modifiers.includes('ctrl') === event.ctrlKey &&
        modifiers.includes('alt') === event.altKey &&
        modifiers.includes('shift') === event.shiftKey &&
        modifiers.includes('meta') === event.metaKey,
    );
  }

  function modifiedMpvKey(event: ModifierState, key: string): string | null {
    return normalizeMpvInputKey(
      [
        ...(event.ctrlKey ? ['ctrl'] : []),
        ...(event.altKey ? ['alt'] : []),
        ...(event.shiftKey ? ['shift'] : []),
        ...(event.metaKey ? ['meta'] : []),
        key,
      ].join('+'),
    );
  }

  function keydown(event: ForwardedKeyEvent): boolean {
    if (disposed || event.defaultPrevented) return false;
    if (heldKeys.has(event.code)) {
      event.preventDefault();
      return true;
    }
    if (event.repeat || event.code.startsWith('Numpad')) return false;
    if (isBlocked(event.code, event)) return false;
    const key = keyboardEventToMpvKey(event);
    if (!key || !keys.has(key)) return false;
    heldKeys.set(event.code, key);
    deps.send(['keydown', key]);
    event.preventDefault();
    return true;
  }

  function keyup(event: Pick<KeyboardEvent, 'code' | 'preventDefault'>): void {
    const key = heldKeys.get(event.code);
    if (!key) return;
    heldKeys.delete(event.code);
    deps.send(['keyup', key]);
    event.preventDefault();
  }

  // Wheel scrolls are single events, so they go through mpv's keypress with the notch
  // count as scale, matching how mpv handles precise scrolling natively.
  function wheel(event: ForwardedWheelEvent): boolean {
    if (disposed || event.defaultPrevented) return false;
    const scroll = wheelEventToMpvWheel(event);
    if (!scroll || isBlocked(scroll.key, event)) return false;
    const key = modifiedMpvKey(event, scroll.key);
    if (!key || !keys.has(key)) return false;
    deps.send(['keypress', key, scroll.notches]);
    event.preventDefault();
    return true;
  }

  // Buttons go through keydown/keyup so held-button bindings and mpv's own double-click
  // detection (MBTN_LEFT_DBL) work. A button is forwarded when mpv binds it or its
  // double-click.
  function mousedown(event: ForwardedMouseEvent): boolean {
    if (disposed || event.defaultPrevented) return false;
    const heldId = `mouse:${event.button}`;
    if (heldKeys.has(heldId)) {
      event.preventDefault();
      return true;
    }
    const button = MPV_MOUSE_BUTTON_BY_BUTTON[event.button];
    if (!button || isBlocked(button, event)) return false;
    const key = modifiedMpvKey(event, button);
    if (!key || (!keys.has(key) && !keys.has(`${key}_DBL`))) return false;
    heldKeys.set(heldId, key);
    deps.send(['keydown', key]);
    event.preventDefault();
    return true;
  }

  function mouseup(event: Pick<MouseEvent, 'button' | 'preventDefault'>): void {
    const heldId = `mouse:${event.button}`;
    const key = heldKeys.get(heldId);
    if (!key) return;
    heldKeys.delete(heldId);
    deps.send(['keyup', key]);
    event.preventDefault();
  }

  function dispose(): void {
    disposed = true;
    keys.clear();
    releaseAll();
  }

  return { refresh, keydown, keyup, wheel, mousedown, mouseup, releaseAll, dispose };
}
