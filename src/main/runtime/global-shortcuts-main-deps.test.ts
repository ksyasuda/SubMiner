import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildGetConfiguredShortcutsMainDepsHandler,
  createBuildRefreshGlobalAndOverlayShortcutsMainDepsHandler,
  createBuildRegisterGlobalShortcutsMainDepsHandler,
} from './global-shortcuts-main-deps';

test('get configured shortcuts main deps map config resolver inputs', () => {
  const config = { shortcuts: { copySubtitle: 's' } } as never;
  const defaults = { shortcuts: { copySubtitle: 'c' } } as never;
  const build = createBuildGetConfiguredShortcutsMainDepsHandler({
    getResolvedConfig: () => config,
    defaultConfig: defaults,
    resolveConfiguredShortcuts: (nextConfig, nextDefaults) => ({ nextConfig, nextDefaults }) as never,
  });

  const deps = build();
  assert.equal(deps.getResolvedConfig(), config);
  assert.equal(deps.defaultConfig, defaults);
  assert.deepEqual(deps.resolveConfiguredShortcuts(config, defaults), { nextConfig: config, nextDefaults: defaults });
});

test('register global shortcuts main deps map callbacks and flags', () => {
  const calls: string[] = [];
  const mainWindow = { id: 'main' };
  const build = createBuildRegisterGlobalShortcutsMainDepsHandler({
    getConfiguredShortcuts: () => ({ copySubtitle: 's' } as never),
    registerGlobalShortcutsCore: () => calls.push('register'),
    toggleVisibleOverlay: () => calls.push('toggle-visible'),
    toggleInvisibleOverlay: () => calls.push('toggle-invisible'),
    openYomitanSettings: () => calls.push('open-yomitan'),
    isDev: true,
    getMainWindow: () => mainWindow as never,
  });

  const deps = build();
  deps.registerGlobalShortcutsCore({
    shortcuts: deps.getConfiguredShortcuts(),
    onToggleVisibleOverlay: () => undefined,
    onToggleInvisibleOverlay: () => undefined,
    onOpenYomitanSettings: () => undefined,
    isDev: deps.isDev,
    getMainWindow: deps.getMainWindow,
  });
  deps.onToggleVisibleOverlay();
  deps.onToggleInvisibleOverlay();
  deps.onOpenYomitanSettings();
  assert.equal(deps.isDev, true);
  assert.deepEqual(deps.getMainWindow(), mainWindow);
  assert.deepEqual(calls, ['register', 'toggle-visible', 'toggle-invisible', 'open-yomitan']);
});

test('refresh global shortcuts main deps map passthrough handlers', () => {
  const calls: string[] = [];
  const deps = createBuildRefreshGlobalAndOverlayShortcutsMainDepsHandler({
    unregisterAllGlobalShortcuts: () => calls.push('unregister'),
    registerGlobalShortcuts: () => calls.push('register'),
    syncOverlayShortcuts: () => calls.push('sync'),
  })();

  deps.unregisterAllGlobalShortcuts();
  deps.registerGlobalShortcuts();
  deps.syncOverlayShortcuts();
  assert.deepEqual(calls, ['unregister', 'register', 'sync']);
});
