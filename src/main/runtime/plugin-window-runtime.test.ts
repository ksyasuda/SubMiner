import assert from 'node:assert/strict';
import { EventEmitter } from 'node:events';
import test from 'node:test';

import type { BrowserWindowConstructorOptions } from 'electron';

import { createPluginWindowRuntime } from './plugin-window-runtime';

function createFakeWindow(loadResult: Promise<void>) {
  const events = new EventEmitter();
  const calls: string[] = [];
  const window = {
    calls,
    destroyed: false,
    isDestroyed: () => window.destroyed,
    show: () => calls.push('show'),
    focus: () => calls.push('focus'),
    loadFile: (filePath: string) => {
      calls.push(`load:${filePath}`);
      return loadResult;
    },
    destroy: () => {
      window.destroyed = true;
      events.emit('closed');
    },
    once: (event: 'closed', listener: () => void) => events.once(event, listener),
  };
  return window;
}

const request = {
  key: 'viewer:/plugins/viewer/index.html',
  htmlPath: '/plugins/viewer/index.html',
  preloadPath: '/plugins/viewer/preload.js',
  title: 'Viewer',
  width: 400,
  height: 300,
};

test('reopening a key focuses the open window and a closed window is recreated', () => {
  const created: Array<ReturnType<typeof createFakeWindow>> = [];
  const options: BrowserWindowConstructorOptions[] = [];
  const promoted: unknown[] = [];
  const runtime = createPluginWindowRuntime({
    createBrowserWindow: (windowOptions) => {
      options.push(windowOptions);
      const window = createFakeWindow(Promise.resolve());
      created.push(window);
      return window;
    },
    promoteAboveOverlay: (window) => promoted.push(window),
    logError: () => {},
  });

  const first = runtime.open(request);
  const again = runtime.open(request);
  assert.equal(first.created, true);
  assert.equal(again.created, false);
  assert.equal(again.window, first.window);
  assert.deepEqual(first.window.calls, [
    'load:/plugins/viewer/index.html',
    'show',
    'focus',
    'show',
    'focus',
  ]);
  assert.equal(promoted.length, 2);
  assert.equal(options[0]?.webPreferences?.preload, '/plugins/viewer/preload.js');
  assert.equal(options[0]?.webPreferences?.contextIsolation, true);

  first.window.destroy();
  const reopened = runtime.open(request);
  assert.equal(reopened.created, true);
  assert.equal(created.length, 2);
});

test('a window whose page fails to load is destroyed and logged', async () => {
  const errors: string[] = [];
  const window = createFakeWindow(Promise.reject(new Error('ENOENT')));
  const runtime = createPluginWindowRuntime({
    createBrowserWindow: () => window,
    promoteAboveOverlay: () => {},
    logError: (message) => errors.push(message),
  });

  runtime.open(request);
  await new Promise((resolve) => setImmediate(resolve));

  assert.equal(window.destroyed, true);
  assert.deepEqual(errors, ['Failed to load plugin window /plugins/viewer/index.html: ENOENT']);
});
