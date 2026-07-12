import assert from 'node:assert/strict';
import test from 'node:test';
import type { ConfiguredShortcuts } from '../../core/utils/shortcut-config';
import { createGlobalShortcutsRuntimeHandlers } from './global-shortcuts-runtime-handlers';

function createShortcuts(): ConfiguredShortcuts {
  return {
    toggleVisibleOverlayGlobal: 'CommandOrControl+Shift+O',
    copySubtitle: 's',
    copySubtitleMultiple: 'CommandOrControl+s',
    updateLastCardFromClipboard: 'c',
    triggerFieldGrouping: null,
    triggerSubsync: null,
    mineSentence: 'q',
    mineSentenceMultiple: 'w',
    multiCopyTimeoutMs: 5000,
    toggleSecondarySub: null,
    markAudioCard: null,
    openCharacterDictionaryManager: null,
    openRuntimeOptions: null,
    openJimaku: null,
    openAnimetosho: null,
    openSessionHelp: null,
    openControllerSelect: null,
    openControllerDebug: null,
    toggleSubtitleSidebar: null,
    toggleNotificationHistory: null,
    appendClipboardVideoToQueue: null,
  };
}

test('global shortcuts runtime handlers compose get/register/refresh flow', () => {
  const calls: string[] = [];
  const shortcuts = createShortcuts();
  const runtime = createGlobalShortcutsRuntimeHandlers({
    getConfiguredShortcutsMainDeps: {
      getResolvedConfig: () => ({}) as never,
      defaultConfig: {} as never,
      resolveConfiguredShortcuts: () => shortcuts,
    },
    buildRegisterGlobalShortcutsMainDeps: (getConfiguredShortcuts) => ({
      getConfiguredShortcuts,
      registerGlobalShortcutsCore: (options) => {
        calls.push('register');
        assert.equal(options.shortcuts, shortcuts);
      },
      toggleVisibleOverlay: () => calls.push('toggle-visible'),
      openYomitanSettings: () => calls.push('open-yomitan'),
      isDev: false,
      getMainWindow: () => null,
    }),
    buildRefreshGlobalAndOverlayShortcutsMainDeps: (registerGlobalShortcuts) => ({
      unregisterAllGlobalShortcuts: () => calls.push('unregister'),
      registerGlobalShortcuts: () => registerGlobalShortcuts(),
      syncOverlayShortcuts: () => calls.push('sync-overlay'),
    }),
  });

  assert.equal(runtime.getConfiguredShortcuts(), shortcuts);
  runtime.registerGlobalShortcuts();
  runtime.refreshGlobalAndOverlayShortcuts();
  assert.deepEqual(calls, ['register', 'unregister', 'register', 'sync-overlay']);
});
