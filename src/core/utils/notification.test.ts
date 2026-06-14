import assert from 'node:assert/strict';
import test from 'node:test';
import { resolveDefaultNotificationIconPath } from './notification';

test('default notification icon resolves packaged SubMiner asset when no per-notification icon is provided', () => {
  const path = resolveDefaultNotificationIconPath({
    platform: 'linux',
    resourcesPath: '/opt/SubMiner/resources',
    appPath: '/opt/SubMiner/resources/app.asar',
    dirname: '/opt/SubMiner/resources/app.asar/dist/core/utils',
    joinPath: (...parts) => parts.join('/'),
    fileExists: (candidate) => candidate === '/opt/SubMiner/resources/assets/SubMiner.png',
  });

  assert.equal(path, '/opt/SubMiner/resources/assets/SubMiner.png');
});

test('default notification icon prefers the square app icon when bundled images are available', () => {
  const path = resolveDefaultNotificationIconPath({
    platform: 'linux',
    resourcesPath: '/opt/SubMiner/resources',
    appPath: '/opt/SubMiner/resources/app.asar',
    dirname: '/opt/SubMiner/resources/app.asar/dist/core/utils',
    joinPath: (...parts) => parts.join('/'),
    fileExists: (candidate) =>
      candidate === '/opt/SubMiner/resources/assets/SubMiner.png' ||
      candidate === '/opt/SubMiner/resources/assets/SubMiner-square.png',
  });

  assert.equal(path, '/opt/SubMiner/resources/assets/SubMiner-square.png');
});

test('default notification icon avoids macOS tray template assets', () => {
  const seen: string[] = [];
  const path = resolveDefaultNotificationIconPath({
    platform: 'darwin',
    resourcesPath: '/Applications/SubMiner.app/Contents/Resources',
    appPath: '/Applications/SubMiner.app/Contents/Resources/app.asar',
    dirname: '/Applications/SubMiner.app/Contents/Resources/app.asar/dist/core/utils',
    joinPath: (...parts) => parts.join('/'),
    fileExists: (candidate) => {
      seen.push(candidate);
      return candidate.endsWith('/assets/SubMiner-square.png');
    },
  });

  assert.equal(path, '/Applications/SubMiner.app/Contents/Resources/assets/SubMiner-square.png');
  assert.equal(
    seen.some((candidate) => candidate.includes('SubMinerTemplate')),
    false,
  );
});
