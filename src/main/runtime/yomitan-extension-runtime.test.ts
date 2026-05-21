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
  let yomitanSession: unknown = null;
  let receivedExternalProfilePath = '';
  let loadCalls = 0;
  const releaseLoadState: { releaseLoad: ((value: Extension | null) => void) | null } = {
    releaseLoad: null,
  };

  const runtime = createYomitanExtensionRuntime({
    loadYomitanExtensionCore: async (options) => {
      loadCalls += 1;
      receivedExternalProfilePath = options.externalProfilePath ?? '';
      options.setYomitanParserWindow(null);
      options.setYomitanParserReadyPromise(Promise.resolve());
      options.setYomitanParserInitPromise(Promise.resolve(true));
      options.setYomitanSession({ id: 'session' } as never);
      return await new Promise<Extension | null>((resolve) => {
        releaseLoadState.releaseLoad = (value) => {
          options.setYomitanExtension(value);
          resolve(value);
        };
      });
    },
    userDataPath: '/tmp',
    externalProfilePath: '/tmp/gsm-profile',
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
    setYomitanSession: (next) => {
      yomitanSession = next;
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
  assert.deepEqual(yomitanSession, { id: 'session' });
  assert.equal(receivedExternalProfilePath, '/tmp/gsm-profile');

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
  let receivedExternalProfilePath = '';
  let yomitanSession: unknown = null;

  const runtime = createYomitanExtensionRuntime({
    loadYomitanExtensionCore: async (options) => {
      loadCalls += 1;
      receivedExternalProfilePath = options.externalProfilePath ?? '';
      options.setYomitanSession({ id: 'session' } as never);
      return null;
    },
    userDataPath: '/tmp',
    externalProfilePath: '/tmp/gsm-profile',
    getYomitanParserWindow: () => null,
    setYomitanParserWindow: () => {},
    setYomitanParserReadyPromise: () => {},
    setYomitanParserInitPromise: () => {},
    setYomitanExtension: () => {},
    setYomitanSession: (next) => {
      yomitanSession = next;
    },
    getYomitanExtension: () => null,
    getLoadInFlight: () => null,
    setLoadInFlight: () => {},
  });

  assert.equal(await runtime.loadYomitanExtension(), null);
  assert.equal(loadCalls, 1);
  assert.equal(receivedExternalProfilePath, '/tmp/gsm-profile');
  assert.deepEqual(yomitanSession, { id: 'session' });
});

test('yomitan extension runtime notifies once after concurrent ensure load resolves', async () => {
  let extension: Extension | null = null;
  let inFlight: Promise<Extension | null> | null = null;
  const notifications: Extension[] = [];
  const releaseLoadState: { releaseLoad: ((value: Extension | null) => void) | null } = {
    releaseLoad: null,
  };

  const runtime = createYomitanExtensionRuntime({
    loadYomitanExtensionCore: async (options) => {
      return await new Promise<Extension | null>((resolve) => {
        releaseLoadState.releaseLoad = (value) => {
          options.setYomitanExtension(value);
          resolve(value);
        };
      });
    },
    userDataPath: '/tmp',
    getYomitanParserWindow: () => null,
    setYomitanParserWindow: () => {},
    setYomitanParserReadyPromise: () => {},
    setYomitanParserInitPromise: () => {},
    setYomitanExtension: (next) => {
      extension = next;
    },
    setYomitanSession: () => {},
    getYomitanExtension: () => extension,
    getLoadInFlight: () => inFlight,
    setLoadInFlight: (promise) => {
      inFlight = promise;
    },
    onYomitanExtensionLoaded: (loadedExtension) => {
      notifications.push(loadedExtension);
    },
  });

  const first = runtime.ensureYomitanExtensionLoaded();
  const second = runtime.ensureYomitanExtensionLoaded();
  const fakeExtension = { id: 'yomitan' } as Extension;
  const releaseLoad = releaseLoadState.releaseLoad;
  if (!releaseLoad) {
    throw new Error('expected in-flight yomitan load resolver');
  }

  releaseLoad(fakeExtension);

  assert.equal(await first, fakeExtension);
  assert.equal(await second, fakeExtension);
  assert.deepEqual(notifications, [fakeExtension]);
});

test('yomitan extension runtime retries notification after callback failure', async () => {
  const fakeExtension = { id: 'yomitan' } as Extension;
  let calls = 0;

  const runtime = createYomitanExtensionRuntime({
    loadYomitanExtensionCore: async () => fakeExtension,
    userDataPath: '/tmp',
    getYomitanParserWindow: () => null,
    setYomitanParserWindow: () => {},
    setYomitanParserReadyPromise: () => {},
    setYomitanParserInitPromise: () => {},
    setYomitanExtension: () => {},
    setYomitanSession: () => {},
    getYomitanExtension: () => fakeExtension,
    getLoadInFlight: () => null,
    setLoadInFlight: () => {},
    onYomitanExtensionLoaded: () => {
      calls += 1;
      if (calls === 1) {
        throw new Error('overlay reload failed');
      }
    },
  });

  await assert.rejects(runtime.ensureYomitanExtensionLoaded(), /overlay reload failed/);
  assert.equal(await runtime.ensureYomitanExtensionLoaded(), fakeExtension);
  assert.equal(calls, 2);
});
