import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildDictionaryRootsMainHandler,
  createBuildFrequencyDictionaryRootsMainHandler,
} from './dictionary-runtime-main-deps';

test('dictionary roots main handler returns expected root list', () => {
  const roots = createBuildDictionaryRootsMainHandler({
    platform: 'darwin',
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

test('frequency dictionary roots main handler returns expected root list', () => {
  const roots = createBuildFrequencyDictionaryRootsMainHandler({
    platform: 'darwin',
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
  assert.equal(roots[0], '/repo/dist/main/../../vendor/frequency-dictionary');
  assert.equal(roots[10], '/repo');
});

test('dictionary roots main handler uses APPDATA-style roots on windows', () => {
  const roots = createBuildDictionaryRootsMainHandler({
    platform: 'win32',
    dirname: 'C:\\repo\\dist\\main',
    appPath: 'C:\\Program Files\\SubMiner\\resources\\app.asar',
    resourcesPath: 'C:\\Program Files\\SubMiner\\resources',
    userDataPath: 'C:\\Users\\a\\AppData\\Roaming\\SubMiner',
    appUserDataPath: 'C:\\Users\\a\\AppData\\Roaming\\SubMiner',
    homeDir: 'C:\\Users\\a',
    appDataDir: 'C:\\Users\\a\\AppData\\Roaming',
    cwd: 'C:\\repo',
    joinPath: (...parts) => parts.join('\\'),
  })();

  assert.equal(roots.includes('C:\\Users\\a\\.config\\SubMiner'), false);
  assert.equal(roots.includes('C:\\Users\\a\\AppData\\Roaming\\SubMiner'), true);
});

test('frequency dictionary roots main handler uses APPDATA-style roots on windows', () => {
  const roots = createBuildFrequencyDictionaryRootsMainHandler({
    platform: 'win32',
    dirname: 'C:\\repo\\dist\\main',
    appPath: 'C:\\Program Files\\SubMiner\\resources\\app.asar',
    resourcesPath: 'C:\\Program Files\\SubMiner\\resources',
    userDataPath: 'C:\\Users\\a\\AppData\\Roaming\\SubMiner',
    appUserDataPath: 'C:\\Users\\a\\AppData\\Roaming\\SubMiner',
    homeDir: 'C:\\Users\\a',
    appDataDir: 'C:\\Users\\a\\AppData\\Roaming',
    cwd: 'C:\\repo',
    joinPath: (...parts) => parts.join('\\'),
  })();

  assert.equal(roots.includes('C:\\Users\\a\\.config\\SubMiner'), false);
  assert.equal(roots.includes('C:\\Users\\a\\AppData\\Roaming\\SubMiner'), true);
});
