import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildDictionaryRootsMainHandler,
  createBuildFrequencyDictionaryRootsMainHandler,
  createBuildFrequencyDictionaryRuntimeMainDepsHandler,
  createBuildJlptDictionaryRuntimeMainDepsHandler,
} from './dictionary-runtime-main-deps';

test('dictionary roots main handler returns expected root list', () => {
  const roots = createBuildDictionaryRootsMainHandler({
    dirname: '/repo/dist/main',
    appPath: '/Applications/SubMiner.app/Contents/Resources/app.asar',
    resourcesPath: '/Applications/SubMiner.app/Contents/Resources',
    userDataPath: '/Users/a/.config/SubMiner',
    appUserDataPath: '/Users/a/Library/Application Support/SubMiner',
    homeDir: '/Users/a',
    cwd: '/repo',
    joinPath: (...parts) => parts.join('/'),
  })();

  assert.equal(roots.length, 11);
  assert.equal(roots[0], '/repo/dist/main/../../vendor/yomitan-jlpt-vocab');
  assert.equal(roots[10], '/repo');
});

test('jlpt dictionary runtime main deps builder maps search paths and log prefix', () => {
  const calls: string[] = [];
  const deps = createBuildJlptDictionaryRuntimeMainDepsHandler({
    isJlptEnabled: () => true,
    getDictionaryRoots: () => ['/root/a'],
    getJlptDictionarySearchPaths: ({ getDictionaryRoots }) => getDictionaryRoots().map((path) => `${path}/jlpt`),
    setJlptLevelLookup: () => calls.push('set-lookup'),
    logInfo: (message) => calls.push(`log:${message}`),
  })();

  assert.equal(deps.isJlptEnabled(), true);
  assert.deepEqual(deps.getSearchPaths(), ['/root/a/jlpt']);
  deps.setJlptLevelLookup(() => null);
  deps.log('loaded');
  assert.deepEqual(calls, ['set-lookup', 'log:[JLPT] loaded']);
});

test('frequency dictionary roots main handler returns expected root list', () => {
  const roots = createBuildFrequencyDictionaryRootsMainHandler({
    dirname: '/repo/dist/main',
    appPath: '/Applications/SubMiner.app/Contents/Resources/app.asar',
    resourcesPath: '/Applications/SubMiner.app/Contents/Resources',
    userDataPath: '/Users/a/.config/SubMiner',
    appUserDataPath: '/Users/a/Library/Application Support/SubMiner',
    homeDir: '/Users/a',
    cwd: '/repo',
    joinPath: (...parts) => parts.join('/'),
  })();

  assert.equal(roots.length, 15);
  assert.equal(roots[0], '/repo/dist/main/../../vendor/jiten_freq_global');
  assert.equal(roots[14], '/repo');
});

test('frequency dictionary runtime main deps builder maps search paths/source and log prefix', () => {
  const calls: string[] = [];
  const deps = createBuildFrequencyDictionaryRuntimeMainDepsHandler({
    isFrequencyDictionaryEnabled: () => true,
    getDictionaryRoots: () => ['/root/a', ''],
    getFrequencyDictionarySearchPaths: ({ getDictionaryRoots, getSourcePath }) => [
      ...getDictionaryRoots().map((path) => `${path}/freq`),
      getSourcePath() || '',
    ],
    getSourcePath: () => '/custom/freq.json',
    setFrequencyRankLookup: () => calls.push('set-rank'),
    logInfo: (message) => calls.push(`log:${message}`),
  })();

  assert.equal(deps.isFrequencyDictionaryEnabled(), true);
  assert.deepEqual(deps.getSearchPaths(), ['/root/a/freq', '/custom/freq.json']);
  deps.setFrequencyRankLookup(() => null);
  deps.log('loaded');
  assert.deepEqual(calls, ['set-rank', 'log:[Frequency] loaded']);
});
