import assert from 'node:assert/strict';
import test from 'node:test';
import { createDictionariesAssembly, type DictionariesAssemblyDeps } from './dictionaries';

// Every callback throws, so a test fails if the assembly calls a dep it should only call later.
function makeDeps(): DictionariesAssemblyDeps {
  const lazy = (name: string) => (): never => {
    throw new Error(`${name} called unexpectedly`);
  };
  return {
    roots: {
      platform: 'linux',
      dirname: '/tmp/subminer-main',
      appPath: '/tmp/subminer-app',
      resourcesPath: '/tmp/subminer-resources',
      userDataPath: '/tmp/subminer-user-data',
      appUserDataPath: '/tmp/subminer-app-user-data',
      homeDir: '/tmp/subminer-home',
      cwd: '/tmp/subminer-cwd',
      joinPath: lazy('joinPath'),
    },
    appState: { jlptLevelLookup: () => null, frequencyRankLookup: () => null },
    getConfig: lazy('getConfig'),
    logger: { info: lazy('info') },
  };
}

// main.ts builds the assembly before most siblings exist, so construction must stay lazy.
test('createDictionariesAssembly does not call its deps while constructing', () => {
  const dictionaries = createDictionariesAssembly(makeDeps());
  assert.equal(typeof dictionaries.jlpt.ensureJlptDictionaryLookup, 'function');
  assert.equal(typeof dictionaries.frequency.ensureFrequencyDictionaryLookup, 'function');
});
