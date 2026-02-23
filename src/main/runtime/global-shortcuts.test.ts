import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createGetConfiguredShortcutsHandler,
  createRefreshGlobalAndOverlayShortcutsHandler,
  createRegisterGlobalShortcutsHandler,
} from './global-shortcuts';
import type { ConfiguredShortcuts } from '../../core/utils/shortcut-config';

function createShortcuts(): ConfiguredShortcuts {
  return {
    toggleVisibleOverlayGlobal: 'CommandOrControl+Shift+O',
    toggleInvisibleOverlayGlobal: 'CommandOrControl+Shift+I',
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
    openRuntimeOptions: null,
    openJimaku: null,
  };
}

test('configured shortcuts handler resolves from current config', () => {
  const calls: string[] = [];
  const config = {} as never;
  const defaultConfig = {} as never;
  const shortcuts = createShortcuts();
  const getConfiguredShortcuts = createGetConfiguredShortcutsHandler({
    getResolvedConfig: () => config,
    defaultConfig,
    resolveConfiguredShortcuts: (nextConfig, nextDefaultConfig) => {
      calls.push('resolve');
      assert.equal(nextConfig, config);
      assert.equal(nextDefaultConfig, defaultConfig);
      return shortcuts;
    },
  });

  assert.equal(getConfiguredShortcuts(), shortcuts);
  assert.deepEqual(calls, ['resolve']);
});

test('register global shortcuts handler passes through callbacks and shortcuts', () => {
  const calls: string[] = [];
  const shortcuts = createShortcuts();
  const mainWindow = {} as never;
  const registerGlobalShortcuts = createRegisterGlobalShortcutsHandler({
    getConfiguredShortcuts: () => shortcuts,
    registerGlobalShortcutsCore: (options) => {
      calls.push('register');
      assert.equal(options.shortcuts, shortcuts);
      assert.equal(options.isDev, true);
      assert.equal(options.getMainWindow(), mainWindow);
      options.onToggleVisibleOverlay();
      options.onToggleInvisibleOverlay();
      options.onOpenYomitanSettings();
    },
    onToggleVisibleOverlay: () => calls.push('toggle-visible'),
    onToggleInvisibleOverlay: () => calls.push('toggle-invisible'),
    onOpenYomitanSettings: () => calls.push('open-yomitan'),
    isDev: true,
    getMainWindow: () => mainWindow,
  });

  registerGlobalShortcuts();
  assert.deepEqual(calls, ['register', 'toggle-visible', 'toggle-invisible', 'open-yomitan']);
});

test('refresh global and overlay shortcuts unregisters then re-registers', () => {
  const calls: string[] = [];
  const refresh = createRefreshGlobalAndOverlayShortcutsHandler({
    unregisterAllGlobalShortcuts: () => calls.push('unregister'),
    registerGlobalShortcuts: () => calls.push('register'),
    syncOverlayShortcuts: () => calls.push('sync-overlay'),
  });

  refresh();
  assert.deepEqual(calls, ['unregister', 'register', 'sync-overlay']);
});
