import assert from 'node:assert/strict';
import test from 'node:test';

import { createShortcutsRuntime } from './shortcuts-runtime';

test('shortcuts runtime bridges modal shortcut sync to unregister and sync', () => {
  const calls: string[] = [];

  const runtime = createShortcutsRuntime({
    globalShortcuts: {
      getConfiguredShortcutsMainDeps: {
        getResolvedConfig: () => ({}) as never,
        defaultConfig: {} as never,
        resolveConfiguredShortcuts: () => ({}) as never,
      },
      buildRegisterGlobalShortcutsMainDeps: () => ({
        getConfiguredShortcuts: () => ({}) as never,
        registerGlobalShortcutsCore: () => {
          calls.push('registerGlobalShortcutsCore');
        },
        toggleVisibleOverlay: () => {},
        openYomitanSettings: () => {},
        isDev: false,
        getMainWindow: () => null,
      }),
      buildRefreshGlobalAndOverlayShortcutsMainDeps: () => ({
        unregisterAllGlobalShortcuts: () => {
          calls.push('unregisterAllGlobalShortcuts');
        },
        registerGlobalShortcuts: () => {
          calls.push('registerGlobalShortcuts');
        },
        syncOverlayShortcuts: () => {
          calls.push('syncOverlayShortcuts');
        },
      }),
    },
    numericShortcutRuntimeMainDeps: {
      globalShortcut: {
        register: () => true,
        unregister: () => {},
      },
      showMpvOsd: () => {},
      setTimer: (handler, timeoutMs) => setTimeout(handler, timeoutMs),
      clearTimer: (timer) => clearTimeout(timer),
    },
    numericSessions: {
      onMultiCopyDigit: () => {},
      onMineSentenceDigit: () => {},
    },
    overlayShortcutsRuntimeMainDeps: {
      overlayShortcutsRuntime: {
        registerOverlayShortcuts: () => {
          calls.push('registerOverlayShortcuts');
        },
        unregisterOverlayShortcuts: () => {
          calls.push('unregisterOverlayShortcuts');
        },
        syncOverlayShortcuts: () => {
          calls.push('syncOverlayShortcutsRuntime');
        },
        refreshOverlayShortcuts: () => {
          calls.push('refreshOverlayShortcuts');
        },
      },
    },
  });

  assert.equal(typeof runtime.getConfiguredShortcuts, 'function');
  assert.equal(typeof runtime.registerGlobalShortcuts, 'function');
  assert.equal(typeof runtime.syncOverlayShortcutsForModal, 'function');

  runtime.syncOverlayShortcutsForModal(true);
  runtime.syncOverlayShortcutsForModal(false);

  assert.deepEqual(calls, ['unregisterOverlayShortcuts', 'syncOverlayShortcutsRuntime']);
});
