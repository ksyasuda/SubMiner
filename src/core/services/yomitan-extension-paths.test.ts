import assert from 'node:assert/strict';
import path from 'node:path';
import test from 'node:test';

import {
  getYomitanExtensionSearchPaths,
  resolveExistingYomitanExtensionPath,
} from './yomitan-extension-paths';

test('getYomitanExtensionSearchPaths prioritizes generated build output before packaged fallbacks', () => {
  const searchPaths = getYomitanExtensionSearchPaths({
    cwd: '/repo',
    moduleDir: '/repo/dist/core/services',
    resourcesPath: '/opt/SubMiner/resources',
    userDataPath: '/Users/kyle/.config/SubMiner',
  });

  assert.deepEqual(searchPaths, [
    path.join('/repo', 'build', 'yomitan'),
    path.join('/opt/SubMiner/resources', 'yomitan'),
    '/usr/share/SubMiner/yomitan',
    path.join('/Users/kyle/.config/SubMiner', 'yomitan'),
  ]);
});

test('resolveExistingYomitanExtensionPath returns first manifest-backed candidate', () => {
  const existing = new Set<string>([
    path.join('/repo', 'build', 'yomitan', 'manifest.json'),
    path.join('/repo', 'vendor', 'subminer-yomitan', 'ext', 'manifest.json'),
  ]);

  const resolved = resolveExistingYomitanExtensionPath(
    [
      path.join('/repo', 'build', 'yomitan'),
      path.join('/repo', 'vendor', 'subminer-yomitan', 'ext'),
    ],
    (candidate) => existing.has(candidate),
  );

  assert.equal(resolved, path.join('/repo', 'build', 'yomitan'));
});

test('resolveExistingYomitanExtensionPath ignores source tree without built manifest', () => {
  const resolved = resolveExistingYomitanExtensionPath(
    [path.join('/repo', 'vendor', 'subminer-yomitan', 'ext')],
    () => false,
  );

  assert.equal(resolved, null);
});
