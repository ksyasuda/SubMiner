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
    showTexthookerPage: true,
    openFirstRunSetup: () => calls.push('setup'),
    showFirstRunSetup: true,
    openWindowsMpvLauncherSetup: () => calls.push('windows-mpv'),
    showWindowsMpvLauncherSetup: true,
    openYomitanSettings: () => calls.push('yomitan'),
    openConfigSettings: () => calls.push('configuration'),
    openJellyfinSetup: () => calls.push('jellyfin'),
    showJellyfinDiscovery: true,
    jellyfinDiscoveryActive: false,
    toggleJellyfinDiscovery: () => calls.push('jellyfin-discovery'),
    openAnilistSetup: () => calls.push('anilist'),
    checkForUpdates: () => calls.push('updates'),
    quitApp: () => calls.push('quit'),
  });

  assert.equal(template.length, 12);
  assert.equal(
    template.some((entry) => entry.label === 'Open Runtime Options'),
    false,
  );
  assert.equal(
    template.some((entry) => entry.label === 'Open Overlay'),
    false,
  );
  assert.equal(template[0]!.label, 'Open Help');
  const discovery = template.find((entry) => entry.label === 'Jellyfin Discovery');
  assert.equal(discovery?.type, 'checkbox');
  assert.equal(discovery?.checked, false);
  discovery?.click?.();
  template[0]!.click?.();
  assert.equal(template[1]!.label, 'Open Texthooker');
  template[1]!.click?.();
  assert.equal(template[5]!.label, 'Open SubMiner Settings');
  assert.equal(template[9]!.label, 'Check for Updates');
  template[9]!.click?.();
  template[10]!.type === 'separator' ? calls.push('separator') : calls.push('bad');
  template[11]!.click?.();
  assert.deepEqual(calls, [
    'jellyfin-discovery',
    'help',
    'texthooker',
    'updates',
    'separator',
    'quit',
  ]);
});

test('tray menu template omits first-run setup entry when setup is complete', () => {
  const labels = buildTrayMenuTemplateRuntime({
    openSessionHelp: () => undefined,
    openTexthookerInBrowser: () => undefined,
    showTexthookerPage: true,
    openFirstRunSetup: () => undefined,
    showFirstRunSetup: false,
    openWindowsMpvLauncherSetup: () => undefined,
    showWindowsMpvLauncherSetup: false,
    openYomitanSettings: () => undefined,
    openConfigSettings: () => undefined,
    openJellyfinSetup: () => undefined,
    showJellyfinDiscovery: false,
    jellyfinDiscoveryActive: false,
    toggleJellyfinDiscovery: () => undefined,
    openAnilistSetup: () => undefined,
    checkForUpdates: () => undefined,
    quitApp: () => undefined,
  })
    .map((entry) => entry.label)
    .filter(Boolean);

  assert.equal(labels.includes('Complete Setup'), false);
  assert.equal(labels.includes('Manage Windows mpv launcher'), false);
  assert.equal(labels.includes('Jellyfin Discovery'), false);
});

test('tray menu template omits texthooker entry when texthooker page is disabled', () => {
  const labels = buildTrayMenuTemplateRuntime({
    openSessionHelp: () => undefined,
    openTexthookerInBrowser: () => undefined,
    showTexthookerPage: false,
    openFirstRunSetup: () => undefined,
    showFirstRunSetup: false,
    openWindowsMpvLauncherSetup: () => undefined,
    showWindowsMpvLauncherSetup: false,
    openYomitanSettings: () => undefined,
    openConfigSettings: () => undefined,
    openJellyfinSetup: () => undefined,
    showJellyfinDiscovery: false,
    jellyfinDiscoveryActive: false,
    toggleJellyfinDiscovery: () => undefined,
    openAnilistSetup: () => undefined,
    checkForUpdates: () => undefined,
    quitApp: () => undefined,
  })
    .map((entry) => entry.label)
    .filter(Boolean);

  assert.equal(labels.includes('Open Texthooker'), false);
});

test('tray menu template renders active jellyfin discovery checkbox', () => {
  const template = buildTrayMenuTemplateRuntime({
    openSessionHelp: () => undefined,
    openTexthookerInBrowser: () => undefined,
    showTexthookerPage: true,
    openFirstRunSetup: () => undefined,
    showFirstRunSetup: false,
    openWindowsMpvLauncherSetup: () => undefined,
    showWindowsMpvLauncherSetup: false,
    openYomitanSettings: () => undefined,
    openConfigSettings: () => undefined,
    openJellyfinSetup: () => undefined,
    showJellyfinDiscovery: true,
    jellyfinDiscoveryActive: true,
    toggleJellyfinDiscovery: () => undefined,
    openAnilistSetup: () => undefined,
    checkForUpdates: () => undefined,
    quitApp: () => undefined,
  });

  const discovery = template.find((entry) => entry.label === 'Jellyfin Discovery');
  assert.equal(discovery?.type, 'checkbox');
  assert.equal(discovery?.checked, true);
});
