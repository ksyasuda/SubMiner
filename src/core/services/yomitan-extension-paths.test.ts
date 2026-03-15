import assert from 'node:assert/strict';
import path from 'node:path';
import test from 'node:test';

import {
  getYomitanExtensionSearchPaths,
  resolveExternalYomitanExtensionPath,
  resolveExistingYomitanExtensionPath,
} from './yomitan-extension-paths';

test('getYomitanExtensionSearchPaths prioritizes generated build output before packaged fallbacks', () => {
  const repoRoot = path.resolve('repo');
  const resourcesPath = path.join(path.sep, 'opt', 'SubMiner', 'resources');
  const userDataPath = path.join(path.sep, 'Users', 'kyle', '.config', 'SubMiner');
  const searchPaths = getYomitanExtensionSearchPaths({
    cwd: repoRoot,
    moduleDir: path.join(repoRoot, 'dist', 'core', 'services'),
    resourcesPath,
    userDataPath,
  });

  assert.deepEqual(searchPaths, [
    path.join(repoRoot, 'build', 'yomitan'),
    path.join(resourcesPath, 'yomitan'),
    '/usr/share/SubMiner/yomitan',
    path.join(userDataPath, 'yomitan'),
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

test('resolveExternalYomitanExtensionPath returns external extension dir when manifest exists', () => {
  const profilePath = path.join('/Users', 'kyle', '.local', 'share', 'gsm-profile');
  const resolved = resolveExternalYomitanExtensionPath(
    profilePath,
    (candidate) => candidate === path.join(profilePath, 'extensions', 'yomitan', 'manifest.json'),
  );

  assert.equal(resolved, path.join(profilePath, 'extensions', 'yomitan'));
});

test('resolveExternalYomitanExtensionPath returns null when external profile has no extension', () => {
  const profilePath = path.join('/Users', 'kyle', '.local', 'share', 'gsm-profile');
  const resolved = resolveExternalYomitanExtensionPath(profilePath, () => false);

  assert.equal(resolved, null);
});
