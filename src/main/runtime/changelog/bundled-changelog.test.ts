import assert from 'node:assert/strict';
import test from 'node:test';

import { readBundledChangelog, resolveBundledChangelogPath } from './bundled-changelog';

test('bundled changelog path prefers the packaged resources copy', () => {
  const resolved = resolveBundledChangelogPath({
    resourcesPath: '/res',
    appPath: '/app',
    dirname: '/app/dist/main',
    joinPath: (...parts) => parts.join('/'),
    fileExists: (candidate) => candidate === '/res/CHANGELOG.md',
  });

  assert.equal(resolved, '/res/CHANGELOG.md');
});

test('bundled changelog path falls back to the repo root during development', () => {
  const resolved = resolveBundledChangelogPath({
    resourcesPath: '/res',
    appPath: '/app',
    dirname: '/repo/dist/main',
    joinPath: (...parts) => parts.join('/'),
    fileExists: (candidate) => candidate === '/repo/dist/main/../../CHANGELOG.md',
  });

  assert.equal(resolved, '/repo/dist/main/../../CHANGELOG.md');
});

test('bundled changelog returns null when no copy is installed', () => {
  const result = readBundledChangelog({
    resolvePath: () => null,
    readFile: () => {
      throw new Error('should not read');
    },
    logWarn: () => {},
  });

  assert.equal(result, null);
});

test('bundled changelog logs and returns null when the file cannot be read', () => {
  const warnings: string[] = [];
  const result = readBundledChangelog({
    resolvePath: () => '/res/CHANGELOG.md',
    readFile: () => {
      throw new Error('EACCES');
    },
    logWarn: (message) => warnings.push(message),
  });

  assert.equal(result, null);
  assert.equal(warnings.length, 1);
  assert.match(warnings[0] ?? '', /EACCES/);
});
