import assert from 'node:assert/strict';
import test from 'node:test';
import {
  composeShortcutRuntimes,
  type ShortcutsRuntimeComposerOptions,
} from './shortcuts-runtime-composer';

type GlobalShortcutsOptions = ShortcutsRuntimeComposerOptions['globalShortcuts'];

function createComposerOptions(tryBeginOverlaySelection: (timeoutMs: number) => boolean) {
  const registeredAccelerators: string[] = [];
  const options: ShortcutsRuntimeComposerOptions = {
    // Global and overlay shortcut handlers are only constructed here, never invoked.
    globalShortcuts: {
      getConfiguredShortcutsMainDeps:
        {} as GlobalShortcutsOptions['getConfiguredShortcutsMainDeps'],
      buildRegisterGlobalShortcutsMainDeps: () =>
        ({}) as ReturnType<GlobalShortcutsOptions['buildRegisterGlobalShortcutsMainDeps']>,
      buildRefreshGlobalAndOverlayShortcutsMainDeps: () =>
        ({}) as ReturnType<GlobalShortcutsOptions['buildRefreshGlobalAndOverlayShortcutsMainDeps']>,
    },
    overlayShortcutsRuntimeMainDeps:
      {} as ShortcutsRuntimeComposerOptions['overlayShortcutsRuntimeMainDeps'],
    numericShortcutRuntimeMainDeps: {
      globalShortcut: {
        register: (accelerator) => {
          registeredAccelerators.push(accelerator);
          return true;
        },
        unregister: () => {},
      },
      showMpvOsd: () => {},
      setTimer: () => setTimeout(() => {}, 0),
      clearTimer: (timer) => clearTimeout(timer),
    },
    numericSessions: {
      onMultiCopyDigit: () => {},
      onMineSentenceDigit: () => {},
      tryBeginMultiCopyOverlaySelection: tryBeginOverlaySelection,
      tryBeginMineSentenceOverlaySelection: tryBeginOverlaySelection,
    },
  };
  return { options, registeredAccelerators };
}

test('numeric sessions delegate to the visible overlay when it accepts the selection', () => {
  const timeouts: number[] = [];
  const { options, registeredAccelerators } = createComposerOptions((timeoutMs) => {
    timeouts.push(timeoutMs);
    return true;
  });
  const runtimes = composeShortcutRuntimes(options);

  runtimes.startPendingMultiCopy(3000);
  runtimes.startPendingMineSentenceMultiple(4000);

  assert.deepEqual(timeouts, [3000, 4000]);
  assert.deepEqual(registeredAccelerators, []);
});

test('numeric sessions fall back to global digit shortcuts when the overlay declines', () => {
  const { options, registeredAccelerators } = createComposerOptions(() => false);
  const runtimes = composeShortcutRuntimes(options);

  runtimes.startPendingMultiCopy(3000);

  assert.ok(registeredAccelerators.length > 0);
  runtimes.cancelPendingMultiCopy();
});
