import test from 'node:test';
import assert from 'node:assert/strict';
import { sendToVisibleOverlayRuntime } from './overlay-bridge';

test('sendToVisibleOverlayRuntime restores visibility flag when opening hidden overlay modal', () => {
  const sent: unknown[][] = [];
  const restoreSet = new Set<'runtime-options' | 'subsync'>();
  let visibleOverlayVisible = false;

  const ok = sendToVisibleOverlayRuntime({
    mainWindow: {
      isDestroyed: () => false,
      webContents: {
        isLoading: () => false,
        send: (...args: unknown[]) => {
          sent.push(args);
        },
      },
    } as unknown as Electron.BrowserWindow,
    visibleOverlayVisible,
    setVisibleOverlayVisible: (visible: boolean) => {
      visibleOverlayVisible = visible;
    },
    channel: 'runtime-options:open',
    restoreOnModalClose: 'runtime-options',
    restoreVisibleOverlayOnModalClose: restoreSet,
  });

  assert.equal(ok, true);
  assert.equal(visibleOverlayVisible, true);
  assert.equal(restoreSet.has('runtime-options'), true);
  assert.deepEqual(sent, [['runtime-options:open']]);
});

test('sendToVisibleOverlayRuntime waits for overlay page before sending open command', () => {
  const sent: unknown[][] = [];
  const restoreSet = new Set<'runtime-options' | 'subsync'>();
  let loading = true;
  let currentURL = '';
  const finishCallbacks: Array<() => void> = [];

  const ok = sendToVisibleOverlayRuntime({
    mainWindow: {
      isDestroyed: () => false,
      webContents: {
        isLoading: () => loading,
        getURL: () => currentURL,
        send: (...args: unknown[]) => {
          sent.push(args);
        },
        once: (_event: string, callback: () => void) => {
          finishCallbacks.push(callback);
        },
      },
    } as unknown as Electron.BrowserWindow,
    visibleOverlayVisible: false,
    setVisibleOverlayVisible: () => {},
    channel: 'runtime-options:open',
    restoreOnModalClose: 'runtime-options',
    restoreVisibleOverlayOnModalClose: restoreSet,
  });

  assert.equal(ok, true);
  assert.deepEqual(sent, []);
  assert.equal(restoreSet.has('runtime-options'), true);

  loading = false;
  currentURL = 'file:///overlay/index.html?layer=visible';
  assert.ok(finishCallbacks[0]);
  finishCallbacks[0]!();

  assert.deepEqual(sent, [['runtime-options:open']]);
});
