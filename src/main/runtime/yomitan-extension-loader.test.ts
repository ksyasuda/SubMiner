import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createEnsureYomitanExtensionLoadedHandler,
  createLoadYomitanExtensionHandler,
} from './yomitan-extension-loader';

test('load yomitan extension handler forwards parser state dependencies', async () => {
  const calls: string[] = [];
  const parserWindow = {} as never;
  const extension = { id: 'ext' } as never;
  const loadYomitanExtension = createLoadYomitanExtensionHandler({
    loadYomitanExtensionCore: async (options) => {
      calls.push(`path:${options.userDataPath}`);
      assert.equal(options.getYomitanParserWindow(), parserWindow);
      options.setYomitanParserWindow(null);
      options.setYomitanParserReadyPromise(null);
      options.setYomitanParserInitPromise(null);
      options.setYomitanExtension(extension);
      return extension;
    },
    userDataPath: '/tmp/subminer',
    getYomitanParserWindow: () => parserWindow,
    setYomitanParserWindow: () => calls.push('set-window'),
    setYomitanParserReadyPromise: () => calls.push('set-ready'),
    setYomitanParserInitPromise: () => calls.push('set-init'),
    setYomitanExtension: () => calls.push('set-ext'),
  });

  assert.equal(await loadYomitanExtension(), extension);
  assert.deepEqual(calls, ['path:/tmp/subminer', 'set-window', 'set-ready', 'set-init', 'set-ext']);
});

test('ensure yomitan loader returns existing extension when available', async () => {
  const extension = { id: 'ext' } as never;
  const ensureLoaded = createEnsureYomitanExtensionLoadedHandler({
    getYomitanExtension: () => extension,
    getLoadInFlight: () => null,
    setLoadInFlight: () => {
      throw new Error('unexpected');
    },
    loadYomitanExtension: async () => {
      throw new Error('unexpected');
    },
  });

  assert.equal(await ensureLoaded(), extension);
});

test('ensure yomitan loader reuses in-flight promise', async () => {
  const extension = { id: 'ext' } as never;
  const inflight = Promise.resolve(extension);
  const ensureLoaded = createEnsureYomitanExtensionLoadedHandler({
    getYomitanExtension: () => null,
    getLoadInFlight: () => inflight,
    setLoadInFlight: () => {
      throw new Error('unexpected');
    },
    loadYomitanExtension: async () => {
      throw new Error('unexpected');
    },
  });

  assert.equal(await ensureLoaded(), extension);
});

test('ensure yomitan loader starts load and clears in-flight when done', async () => {
  const calls: string[] = [];
  let inFlight: Promise<any> | null = null;
  const extension = { id: 'ext' } as never;
  const ensureLoaded = createEnsureYomitanExtensionLoadedHandler({
    getYomitanExtension: () => null,
    getLoadInFlight: () => inFlight,
    setLoadInFlight: (promise) => {
      inFlight = promise;
      calls.push(promise ? 'set:promise' : 'set:null');
    },
    loadYomitanExtension: async () => {
      calls.push('load');
      return extension;
    },
  });

  assert.equal(await ensureLoaded(), extension);
  assert.deepEqual(calls, ['load', 'set:promise', 'set:null']);
});
