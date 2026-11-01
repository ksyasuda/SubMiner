import assert from 'node:assert/strict';
import test from 'node:test';
import { createJellyfinAssembly, type JellyfinAssemblyDeps } from './jellyfin';

// Every callback throws, so a test fails if the assembly calls a dep it should only call later.
function makeDeps(): JellyfinAssemblyDeps {
  const lazy = (name: string) => (): never => {
    throw new Error(`${name} called unexpectedly`);
  };
  return {
    appState: {
      mpvClient: null,
      mpvSocketPath: '/tmp/subminer-test.sock',
      immersionTracker: null,
      jellyfinRemoteSession: null,
      jellyfinSetupWindow: null,
    },
    configService: { getConfig: lazy('getConfig'), patchRawConfig: lazy('patchRawConfig') },
    jellyfinTokenStore: {
      loadSession: lazy('loadSession'),
      saveSession: lazy('saveSession'),
      clearSession: lazy('clearSession'),
    },
    logger: { info: () => {}, warn: () => {}, debug: () => {}, error: () => {} },
    defaultMpvLogPath: '/tmp/mpv.log',
    getAppDataPath: lazy('getAppDataPath'),
    getRuntimePluginEntrypoint: lazy('getRuntimePluginEntrypoint'),
    getMpvPluginRuntimeConfig: lazy('getMpvPluginRuntimeConfig'),
    createMpvClient: lazy('createMpvClient'),
    initSubtitlePrefetch: lazy('initSubtitlePrefetch'),
    showVisibleOverlay: lazy('showVisibleOverlay'),
    getOverlayNotifications: lazy('getOverlayNotifications'),
    updateCurrentMediaTitle: lazy('updateCurrentMediaTitle'),
    ensureImmersionTrackerStarted: lazy('ensureImmersionTrackerStarted'),
    refreshTrayMenu: lazy('refreshTrayMenu'),
    setupPreloadPath: '/tmp/preload-jellyfin-setup.js',
    createBrowserWindow: lazy('createBrowserWindow'),
    ipcMain: { handle: lazy('ipcMain.handle'), removeHandler: lazy('ipcMain.removeHandler') },
  };
}

// main.ts builds the assembly before several siblings exist, so construction must stay lazy.
test('createJellyfinAssembly does not call its deps while constructing', () => {
  const jellyfin = createJellyfinAssembly(makeDeps());
  assert.equal(jellyfin.getActivePlayback(), null);
  assert.equal(jellyfin.getQuitOnDisconnectArmed(), false);
});
