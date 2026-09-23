import assert from 'node:assert/strict';
import test from 'node:test';
import { findExtensionUpdates } from './extension-updates';

test('findExtensionUpdates returns only strictly newer builds for known installed versions', () => {
  const updates = findExtensionUpdates(
    [
      { pkg: 'old', versionCode: 10 },
      { pkg: 'current', versionCode: 12 },
      { pkg: 'newer-than-repo', versionCode: 20 },
      { pkg: 'unknown', versionCode: null },
    ],
    [
      { pkg: 'old', versionCode: 11, name: 'Old' },
      { pkg: 'current', versionCode: 12, name: 'Current' },
      { pkg: 'newer-than-repo', versionCode: 19, name: 'Newer' },
      { pkg: 'unknown', versionCode: 30, name: 'Unknown' },
      { pkg: 'not-installed', versionCode: 1, name: 'Not installed' },
    ],
  );

  assert.deepEqual(updates, [{ pkg: 'old', versionCode: 11, name: 'Old' }]);
});
