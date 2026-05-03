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
    openSessionHelp: () => calls.push('help'),
    openTexthookerInBrowser: () => calls.push('texthooker'),
    openFirstRunSetup: () => calls.push('setup'),
    showFirstRunSetup: true,
    openWindowsMpvLauncherSetup: () => calls.push('windows-mpv'),
    showWindowsMpvLauncherSetup: true,
    openYomitanSettings: () => calls.push('yomitan'),
    openRuntimeOptions: () => calls.push('runtime'),
    openJellyfinSetup: () => calls.push('jellyfin'),
    openAnilistSetup: () => calls.push('anilist'),
    quitApp: () => calls.push('quit'),
  });

  assert.equal(template.length, 10);
  assert.equal(
    template.some((entry) => entry.label === 'Open Overlay'),
    false,
  );
  assert.equal(template[0]!.label, 'Open Help');
  template[0]!.click?.();
  assert.equal(template[1]!.label, 'Open Texthooker');
  template[1]!.click?.();
  template[8]!.type === 'separator' ? calls.push('separator') : calls.push('bad');
  template[9]!.click?.();
  assert.deepEqual(calls, ['help', 'texthooker', 'separator', 'quit']);
});

test('tray menu template omits first-run setup entry when setup is complete', () => {
  const labels = buildTrayMenuTemplateRuntime({
    openSessionHelp: () => undefined,
    openTexthookerInBrowser: () => undefined,
    openFirstRunSetup: () => undefined,
    showFirstRunSetup: false,
    openWindowsMpvLauncherSetup: () => undefined,
    showWindowsMpvLauncherSetup: false,
    openYomitanSettings: () => undefined,
    openRuntimeOptions: () => undefined,
    openJellyfinSetup: () => undefined,
    openAnilistSetup: () => undefined,
    quitApp: () => undefined,
  })
    .map((entry) => entry.label)
    .filter(Boolean);

  assert.equal(labels.includes('Complete Setup'), false);
  assert.equal(labels.includes('Manage Windows mpv launcher'), false);
});
