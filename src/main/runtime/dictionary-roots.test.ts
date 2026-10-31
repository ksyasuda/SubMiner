import assert from 'node:assert/strict';
import test from 'node:test';
import { createBuildDictionaryRootsHandler } from './dictionary-roots';

const darwinDeps = {
  platform: 'darwin',
  dirname: '/repo/dist/main',
  appPath: '/Applications/SubMiner.app/Contents/Resources/app.asar',
  resourcesPath: '/Applications/SubMiner.app/Contents/Resources',
  userDataPath: '/Users/a/.config/SubMiner',
  appUserDataPath: '/Users/a/Library/Application Support/SubMiner',
  homeDir: '/Users/a',
  cwd: '/repo',
  joinPath: (...parts: string[]) => parts.join('/'),
} as const;

const win32Deps = {
  platform: 'win32',
  dirname: 'C:\\repo\\dist\\main',
  appPath: 'C:\\Program Files\\SubMiner\\resources\\app.asar',
  resourcesPath: 'C:\\Program Files\\SubMiner\\resources',
  userDataPath: 'C:\\Users\\a\\AppData\\Roaming\\SubMiner',
  appUserDataPath: 'C:\\Users\\a\\AppData\\Roaming\\SubMiner',
  homeDir: 'C:\\Users\\a',
  appDataDir: 'C:\\Users\\a\\AppData\\Roaming',
  cwd: 'C:\\repo',
  joinPath: (...parts: string[]) => parts.join('\\'),
} as const;

for (const vendorDirName of ['yomitan-jlpt-vocab', 'frequency-dictionary']) {
  test(`dictionary roots for ${vendorDirName} list vendor, user data, platform and cwd roots`, () => {
    const roots = createBuildDictionaryRootsHandler({ ...darwinDeps, vendorDirName })();

    assert.equal(roots.length, 11);
    assert.equal(roots[0], `/repo/dist/main/../../vendor/${vendorDirName}`);
    assert.equal(roots[2], `/Applications/SubMiner.app/Contents/Resources/${vendorDirName}`);
    assert.equal(roots[10], '/repo');
  });
}

test('dictionary roots use APPDATA-style roots on windows', () => {
  const roots = createBuildDictionaryRootsHandler({
    ...win32Deps,
    vendorDirName: 'yomitan-jlpt-vocab',
  })();

  assert.equal(roots.includes('C:\\Users\\a\\.config\\SubMiner'), false);
  assert.equal(roots.includes('C:\\Users\\a\\AppData\\Roaming\\SubMiner'), true);
});
