import test from 'node:test';
import assert from 'node:assert/strict';
import { buildTrayMenuTemplateRuntime, resolveTrayIconPathRuntime } from './tray-runtime';

test('resolve tray icon picks template icon first on darwin', () => {
  const path = resolveTrayIconPathRuntime({
    platform: 'darwin',
    resourcesPath: '/res',
    appPath: '/app',
    dirname: '/dist/main',
    joinPath: (...parts) => parts.join('/'),
    fileExists: (candidate) => candidate.endsWith('/res/assets/SubMinerTemplate.png'),
  });
  assert.equal(path, '/res/assets/SubMinerTemplate.png');
});

test('resolve tray icon returns null when no asset exists', () => {
  const path = resolveTrayIconPathRuntime({
    platform: 'linux',
    resourcesPath: '/res',
    appPath: '/app',
    dirname: '/dist/main',
    joinPath: (...parts) => parts.join('/'),
    fileExists: () => false,
  });
  assert.equal(path, null);
});

test('tray menu template contains expected entries and handlers', () => {
  const calls: string[] = [];
  const template = buildTrayMenuTemplateRuntime({
    openOverlay: () => calls.push('overlay'),
    openYomitanSettings: () => calls.push('yomitan'),
    openRuntimeOptions: () => calls.push('runtime'),
    openJellyfinSetup: () => calls.push('jellyfin'),
    openAnilistSetup: () => calls.push('anilist'),
    quitApp: () => calls.push('quit'),
  });

  assert.equal(template.length, 7);
  template[0].click?.();
  template[5].type === 'separator' ? calls.push('separator') : calls.push('bad');
  template[6].click?.();
  assert.deepEqual(calls, ['overlay', 'separator', 'quit']);
});
