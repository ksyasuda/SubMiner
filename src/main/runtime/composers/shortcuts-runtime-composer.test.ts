import assert from 'node:assert/strict';
import test from 'node:test';
import { composeShortcutRuntimes } from './shortcuts-runtime-composer';

test('composeShortcutRuntimes returns callable shortcut runtime handlers', () => {
  const composed = composeShortcutRuntimes({
    globalShortcuts: {
      getConfiguredShortcutsMainDeps: {
        getResolvedConfig: () => ({}) as never,
        defaultConfig: {} as never,
        resolveConfiguredShortcuts: () => ({}) as never,
      },
      buildRegisterGlobalShortcutsMainDeps: () => ({
        getConfiguredShortcuts: () => ({}) as never,
        registerGlobalShortcutsCore: () => {},
        toggleVisibleOverlay: () => {},
        openYomitanSettings: () => {},
        isDev: false,
        getMainWindow: () => null,
      }),
      buildRefreshGlobalAndOverlayShortcutsMainDeps: () => ({
        unregisterAllGlobalShortcuts: () => {},
        registerGlobalShortcuts: () => {},
        syncOverlayShortcuts: () => {},
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
        registerOverlayShortcuts: () => {},
        unregisterOverlayShortcuts: () => {},
        syncOverlayShortcuts: () => {},
        refreshOverlayShortcuts: () => {},
      },
    },
  });

  assert.equal(typeof composed.getConfiguredShortcuts, 'function');
  assert.equal(typeof composed.registerGlobalShortcuts, 'function');
  assert.equal(typeof composed.refreshGlobalAndOverlayShortcuts, 'function');
  assert.equal(typeof composed.cancelPendingMultiCopy, 'function');
  assert.equal(typeof composed.startPendingMultiCopy, 'function');
  assert.equal(typeof composed.cancelPendingMineSentenceMultiple, 'function');
  assert.equal(typeof composed.startPendingMineSentenceMultiple, 'function');
  assert.equal(typeof composed.registerOverlayShortcuts, 'function');
  assert.equal(typeof composed.unregisterOverlayShortcuts, 'function');
  assert.equal(typeof composed.syncOverlayShortcuts, 'function');
  assert.equal(typeof composed.refreshOverlayShortcuts, 'function');
});
