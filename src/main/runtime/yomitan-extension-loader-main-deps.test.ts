import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildEnsureYomitanExtensionLoadedMainDepsHandler,
  createBuildLoadYomitanExtensionMainDepsHandler,
} from './yomitan-extension-loader-main-deps';

test('load yomitan extension main deps builder maps callbacks', async () => {
  const calls: string[] = [];
  const deps = createBuildLoadYomitanExtensionMainDepsHandler({
    loadYomitanExtensionCore: async () => {
      calls.push('load-core');
      return null;
    },
    userDataPath: '/tmp/subminer',
    getYomitanParserWindow: () => null,
    setYomitanParserWindow: () => calls.push('set-window'),
    setYomitanParserReadyPromise: () => calls.push('set-ready'),
    setYomitanParserInitPromise: () => calls.push('set-init'),
    setYomitanExtension: () => calls.push('set-ext'),
  })();

  assert.equal(deps.userDataPath, '/tmp/subminer');
  await deps.loadYomitanExtensionCore({} as never);
  deps.setYomitanParserWindow(null);
  deps.setYomitanParserReadyPromise(null);
  deps.setYomitanParserInitPromise(null);
  deps.setYomitanExtension(null);
  assert.deepEqual(calls, ['load-core', 'set-window', 'set-ready', 'set-init', 'set-ext']);
});

test('ensure yomitan extension loaded main deps builder maps callbacks', async () => {
  const calls: string[] = [];
  const deps = createBuildEnsureYomitanExtensionLoadedMainDepsHandler({
    getYomitanExtension: () => null,
    getLoadInFlight: () => null,
    setLoadInFlight: () => calls.push('set-inflight'),
    loadYomitanExtension: async () => {
      calls.push('load');
      return null;
    },
  })();

  assert.equal(deps.getYomitanExtension(), null);
  assert.equal(deps.getLoadInFlight(), null);
  deps.setLoadInFlight(null);
  await deps.loadYomitanExtension();
  assert.deepEqual(calls, ['set-inflight', 'load']);
});
