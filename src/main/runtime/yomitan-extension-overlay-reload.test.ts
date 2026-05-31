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
    {
      isDestroyed: () => false,
      webContents: {
        isDestroyed: () => false,
        isLoading: () => true,
        reload: () => calls.push('loading-webcontents'),
      },
    },
  ];

  assert.equal(reloadOverlayWindowsForYomitanContentScripts(windows), 1);
  assert.deepEqual(calls, ['live']);
});

test('reloadOverlayWindowsForYomitanContentScripts retries loading webContents after load', () => {
  const calls: string[] = [];
  let finishLoad: (() => void) | null = null;
  const window = {
    isDestroyed: () => false,
    webContents: {
      isDestroyed: () => false,
      isLoading: () => true,
      once: (event: 'did-finish-load', listener: () => void) => {
        assert.equal(event, 'did-finish-load');
        finishLoad = listener;
      },
      reload: () => calls.push('reloaded-after-load'),
    },
  };

  assert.equal(reloadOverlayWindowsForYomitanContentScripts([window]), 0);
  assert.deepEqual(calls, []);

  assert.ok(finishLoad);
  const finish = finishLoad as () => void;
  finish();

  assert.deepEqual(calls, ['reloaded-after-load']);
});
