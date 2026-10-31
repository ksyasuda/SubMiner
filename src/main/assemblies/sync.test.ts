import assert from 'node:assert/strict';
import test from 'node:test';
import { createSyncAssembly, type SyncAssemblyDeps } from './sync';

// Every callback throws, so a test fails if the assembly calls a dep it should only call later.
function makeDeps(): SyncAssemblyDeps {
  const lazy = (name: string) => (): never => {
    throw new Error(`${name} called unexpectedly`);
  };
  return {
    appState: { syncUiWindow: null, initialArgs: null },
    configService: { getConfig: lazy('getConfig') },
    logger: { info: lazy('info'), error: lazy('error') },
    userDataPath: '/tmp/subminer-user-data',
    defaultImmersionDbPath: '/tmp/subminer-user-data/immersion.sqlite',
    getOverlayNotifications: lazy('getOverlayNotifications'),
    requestAppQuit: lazy('requestAppQuit'),
    createBrowserWindow: lazy('createBrowserWindow'),
    mainDirname: '/tmp/subminer-main',
    ipcMain: { handle: lazy('ipcMain.handle') },
    isPackaged: lazy('isPackaged'),
    getAppPath: lazy('getAppPath'),
    pickSnapshotFile: lazy('pickSnapshotFile'),
    revealPath: lazy('revealPath'),
  };
}

// main.ts builds the assembly before most siblings exist, so construction must stay lazy.
test('createSyncAssembly does not call its deps while constructing', () => {
  const sync = createSyncAssembly(makeDeps());
  assert.equal(sync.runtime.isRunning(), false);
});
