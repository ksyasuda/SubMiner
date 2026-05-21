import assert from 'node:assert/strict';
import test from 'node:test';
import { reloadOverlayWindowsForYomitanContentScripts } from './yomitan-extension-overlay-reload';

test('reloadOverlayWindowsForYomitanContentScripts reloads only live overlay windows', () => {
  const calls: string[] = [];
  const windows = [
    {
      isDestroyed: () => false,
      webContents: {
        isDestroyed: () => false,
        reload: () => calls.push('live'),
      },
    },
    {
      isDestroyed: () => true,
      webContents: {
        isDestroyed: () => false,
        reload: () => calls.push('destroyed-window'),
      },
    },
    {
      isDestroyed: () => false,
      webContents: {
        isDestroyed: () => true,
        reload: () => calls.push('destroyed-webcontents'),
      },
    },
  ];

  assert.equal(reloadOverlayWindowsForYomitanContentScripts(windows), 1);
  assert.deepEqual(calls, ['live']);
});
