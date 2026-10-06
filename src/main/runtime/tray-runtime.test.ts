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

type TrayMenuHandlers = Parameters<typeof buildTrayMenuTemplateRuntime>[0];

function baseHandlers(overrides: Partial<TrayMenuHandlers> = {}): TrayMenuHandlers {
  return {
    openSessionHelp: () => undefined,
    openChangelog: () => undefined,
    openTexthookerInBrowser: () => undefined,
    showTexthookerPage: false,
    openFirstRunSetup: () => undefined,
    showFirstRunSetup: false,
    openWindowsMpvLauncherSetup: () => undefined,
    showWindowsMpvLauncherSetup: false,
    dictionaryBackend: 'yomitan',
    openHachidoriSettings: () => undefined,
    openYomitanSettings: () => undefined,
    openConfigSettings: () => undefined,
    openSyncUi: () => undefined,
    openYoutubeBrowser: () => undefined,
    exportLogs: () => undefined,
    openJellyfinSetup: () => undefined,
    showJellyfinDiscovery: false,
    jellyfinDiscoveryActive: false,
    toggleJellyfinDiscovery: () => undefined,
    openAnilistSetup: () => undefined,
    checkForUpdates: () => undefined,
    quitApp: () => undefined,
    ...overrides,
  };
}

function labelsFor(overrides: Partial<TrayMenuHandlers>): string[] {
  return buildTrayMenuTemplateRuntime(baseHandlers(overrides)).flatMap((entry) =>
    entry.label ? [entry.label] : [],
  );
}

const optionalEntries: Array<{ label: string; flag: Partial<TrayMenuHandlers> }> = [
  { label: 'Open Texthooker', flag: { showTexthookerPage: true } },
  { label: 'Complete Setup', flag: { showFirstRunSetup: true } },
  { label: 'Open SubMiner Setup', flag: { showWindowsMpvLauncherSetup: true } },
  { label: 'Jellyfin Discovery', flag: { showJellyfinDiscovery: true } },
];

for (const entry of optionalEntries) {
  test(`tray menu shows "${entry.label}" only when its flag is set`, () => {
    assert.equal(labelsFor({}).includes(entry.label), false);
    assert.equal(labelsFor(entry.flag).includes(entry.label), true);
  });
}

for (const dictionaryBackend of ['yomitan', 'hachidori'] as const) {
  test(`tray menu shows only ${dictionaryBackend} settings and dispatches its handler`, () => {
    const calls: string[] = [];
    const template = buildTrayMenuTemplateRuntime(
      baseHandlers({
        dictionaryBackend,
        openHachidoriSettings: () => calls.push('hachidori'),
        openYomitanSettings: () => calls.push('yomitan'),
      }),
    );
    const labels = template.flatMap((entry) => (entry.label ? [entry.label] : []));
    const [shown, hidden] =
      dictionaryBackend === 'hachidori'
        ? ['Open Hachidori Settings', 'Open Yomitan Settings']
        : ['Open Yomitan Settings', 'Open Hachidori Settings'];
    assert.equal(labels.includes(hidden), false);
    template.find((entry) => entry.label === shown)?.click?.();
    assert.deepEqual(calls, [dictionaryBackend]);
  });
}

test('tray menu ends with a separator and Quit', () => {
  const template = buildTrayMenuTemplateRuntime(baseHandlers());
  assert.deepEqual(
    template.slice(-2).map((entry) => entry.label ?? entry.type),
    ['separator', 'Quit'],
  );
});

test('jellyfin discovery checkbox reflects state and forwards toggles', () => {
  const toggles: boolean[] = [];
  const template = buildTrayMenuTemplateRuntime(
    baseHandlers({
      showJellyfinDiscovery: true,
      jellyfinDiscoveryActive: true,
      toggleJellyfinDiscovery: (checked) => toggles.push(checked),
    }),
  );

  const discovery = template.find((entry) => entry.label === 'Jellyfin Discovery');
  assert.equal(discovery?.type, 'checkbox');
  assert.equal(discovery?.checked, true);
  discovery?.click?.({ checked: true });
  // Without a menu item state the click inverts the current discovery state.
  discovery?.click?.();
  assert.deepEqual(toggles, [true, false]);
});

test('tray menu template renders a visible linux discovery check mark when active', () => {
  const template = buildTrayMenuTemplateRuntime(
    baseHandlers({
      platform: 'linux',
      showJellyfinDiscovery: true,
      jellyfinDiscoveryActive: true,
    }),
  );

  const discovery = template.find((entry) => entry.label === '✓ Jellyfin Discovery');
  assert.equal(discovery?.type, 'checkbox');
  assert.equal(discovery?.checked, true);
});
