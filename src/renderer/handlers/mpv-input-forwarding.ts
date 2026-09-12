import { keyboardEventToMpvKey } from '../../shared/mpv-input-bindings';
import type { MpvInputBindingsSnapshot } from '../../types/session-bindings';

type ForwardedKeyEvent = Parameters<typeof keyboardEventToMpvKey>[0] &
  Pick<KeyboardEvent, 'code' | 'repeat' | 'defaultPrevented' | 'preventDefault'>;

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

  function keydown(event: ForwardedKeyEvent): boolean {
    if (disposed || event.defaultPrevented) return false;
    if (heldKeys.has(event.code)) {
      event.preventDefault();
      return true;
    }
    if (event.repeat || event.code.startsWith('Numpad')) return false;
    if (
      blockedKeys.some(
        ({ code, modifiers }) =>
          code === event.code &&
          modifiers.includes('ctrl') === event.ctrlKey &&
          modifiers.includes('alt') === event.altKey &&
          modifiers.includes('shift') === event.shiftKey &&
          modifiers.includes('meta') === event.metaKey,
      )
    )
      return false;
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

  function dispose(): void {
    disposed = true;
    keys.clear();
    releaseAll();
  }

  return { refresh, keydown, keyup, releaseAll, dispose };
}
