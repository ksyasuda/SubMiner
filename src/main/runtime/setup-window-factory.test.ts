import assert from 'node:assert/strict';
import test from 'node:test';
import { createCreateJellyfinSetupWindowHandler } from './setup-window-factory';

test('createCreateJellyfinSetupWindowHandler wires optional preload bridge', () => {
  const captured: { options?: Electron.BrowserWindowConstructorOptions } = {};
  const createSetupWindow = createCreateJellyfinSetupWindowHandler({
    createBrowserWindow: (nextOptions) => {
      captured.options = nextOptions;
      return { id: 'jellyfin' } as never;
    },
    preloadPath: 'C:\\SubMiner\\dist\\preload-jellyfin-setup.js',
  });

  assert.deepEqual(createSetupWindow(), { id: 'jellyfin' });
  const options = captured.options;
  assert.ok(options);
  assert.equal(options.webPreferences?.preload, 'C:\\SubMiner\\dist\\preload-jellyfin-setup.js');
  assert.equal(options.webPreferences?.sandbox, true);
});
