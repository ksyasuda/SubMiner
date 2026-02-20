import assert from 'node:assert/strict';
import test from 'node:test';
import type { Extension } from 'electron';
import { createYomitanExtensionRuntime } from './yomitan-extension-runtime';

test('yomitan extension runtime reuses in-flight ensure load and clears it after resolve', async () => {
  let extension: Extension | null = null;
  let inFlight: Promise<Extension | null> | null = null;
  let parserWindow: unknown = null;
  let readyPromise: Promise<void> | null = null;
  let initPromise: Promise<boolean> | null = null;
  let loadCalls = 0;
  const releaseLoadState: { releaseLoad: ((value: Extension | null) => void) | null } = {
    releaseLoad: null,
  };

  const runtime = createYomitanExtensionRuntime({
    loadYomitanExtensionCore: async (options) => {
      loadCalls += 1;
      options.setYomitanParserWindow(null);
      options.setYomitanParserReadyPromise(Promise.resolve());
      options.setYomitanParserInitPromise(Promise.resolve(true));
      return await new Promise<Extension | null>((resolve) => {
        releaseLoadState.releaseLoad = (value) => {
          options.setYomitanExtension(value);
          resolve(value);
        };
      });
    },
    userDataPath: '/tmp',
    getYomitanParserWindow: () => parserWindow as never,
    setYomitanParserWindow: (window) => {
      parserWindow = window;
    },
    setYomitanParserReadyPromise: (promise) => {
      readyPromise = promise as Promise<void> | null;
    },
    setYomitanParserInitPromise: (promise) => {
      initPromise = promise as Promise<boolean> | null;
    },
    setYomitanExtension: (next) => {
      extension = next;
    },
    getYomitanExtension: () => extension,
    getLoadInFlight: () => inFlight,
    setLoadInFlight: (promise) => {
      inFlight = promise;
    },
  });

  const first = runtime.ensureYomitanExtensionLoaded();
  const second = runtime.ensureYomitanExtensionLoaded();
  assert.equal(loadCalls, 1);
  assert.ok(inFlight);
  assert.equal(parserWindow, null);
  assert.ok(readyPromise);
  assert.ok(initPromise);

  const fakeExtension = { id: 'yomitan' } as Extension;
  const releaseLoad = releaseLoadState.releaseLoad;
  if (!releaseLoad) {
    throw new Error('expected in-flight yomitan load resolver');
  }
  releaseLoad(fakeExtension);
  assert.equal(await first, fakeExtension);
  assert.equal(await second, fakeExtension);
  assert.equal(extension, fakeExtension);
  assert.equal(inFlight, null);

  const third = await runtime.ensureYomitanExtensionLoaded();
  assert.equal(third, fakeExtension);
  assert.equal(loadCalls, 1);
});

test('yomitan extension runtime direct load delegates to core', async () => {
  let loadCalls = 0;

  const runtime = createYomitanExtensionRuntime({
    loadYomitanExtensionCore: async () => {
      loadCalls += 1;
      return null;
    },
    userDataPath: '/tmp',
    getYomitanParserWindow: () => null,
    setYomitanParserWindow: () => {},
    setYomitanParserReadyPromise: () => {},
    setYomitanParserInitPromise: () => {},
    setYomitanExtension: () => {},
    getYomitanExtension: () => null,
    getLoadInFlight: () => null,
    setLoadInFlight: () => {},
  });

  assert.equal(await runtime.loadYomitanExtension(), null);
  assert.equal(loadCalls, 1);
});
