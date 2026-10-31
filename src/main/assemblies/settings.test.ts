import assert from 'node:assert/strict';
import test from 'node:test';
import { createSettingsAssembly, type SettingsAssemblyDeps } from './settings';

// Every callback throws, so a test fails if the assembly calls a dep it should only call later.
function makeDeps(): SettingsAssemblyDeps {
  const lazy = (name: string) => (): never => {
    throw new Error(`${name} called unexpectedly`);
  };
  return {
    appState: { configSettingsWindow: null },
    configService: {
      getConfig: lazy('getConfig'),
      getConfigPath: lazy('getConfigPath'),
      getRawConfig: lazy('getRawConfig'),
      getWarnings: lazy('getWarnings'),
      reloadConfigStrict: lazy('reloadConfigStrict'),
    },
    logger: { error: lazy('error') },
    getOverlayNotifications: lazy('getOverlayNotifications'),
    showMpvOsd: lazy('showMpvOsd'),
    showDesktopNotification: lazy('showDesktopNotification'),
    hotReloadApplied: {
      setKeybindings: lazy('setKeybindings'),
      setSessionBindings: lazy('setSessionBindings'),
      refreshGlobalAndOverlayShortcuts: lazy('refreshGlobalAndOverlayShortcuts'),
      setSecondarySubMode: lazy('setSecondarySubMode'),
      broadcastToOverlayWindows: lazy('broadcastToOverlayWindows'),
      applyAnkiRuntimeConfigPatch: lazy('applyAnkiRuntimeConfigPatch'),
    },
    getYomitanAnkiDeckName: lazy('getYomitanAnkiDeckName'),
    showErrorBox: lazy('showErrorBox'),
    createBrowserWindow: lazy('createBrowserWindow'),
    mainDirname: '/tmp/subminer-main',
    openPath: lazy('openPath'),
    ipcMain: { handle: lazy('ipcMain.handle') },
  };
}

// main.ts builds the assembly before most siblings exist, so construction must stay lazy.
test('createSettingsAssembly does not call its deps while constructing', () => {
  const settings = createSettingsAssembly(makeDeps());
  assert.equal(typeof settings.hotReload.start, 'function');
  assert.equal(typeof settings.window.registerHandlers, 'function');
});
