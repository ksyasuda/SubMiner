import test from 'node:test';
import assert from 'node:assert/strict';
import { createDestroyTrayHandler, createEnsureTrayHandler } from './tray-lifecycle';

type EnsureTrayDeps = Parameters<typeof createEnsureTrayHandler>[0];
type TrayIcon = ReturnType<EnsureTrayDeps['createEmptyImage']>;

interface FakeImage extends TrayIcon {
  label: string;
  template: boolean;
}

function makeImage(label: string, empty = false): FakeImage {
  const image: FakeImage = {
    label,
    template: false,
    isEmpty: () => empty,
    resize: ({ width, height }) => makeImage(`${label}@${width}x${height}`),
    setTemplateImage: (enabled) => {
      image.template = enabled;
    },
  };
  return image;
}

function makeTray(icon: TrayIcon | null = null) {
  const tray = {
    icon,
    menus: [] as unknown[],
    tooltip: null as string | null,
    clickHandler: null as (() => void) | null,
    destroyed: false,
    setContextMenu: (menu: unknown) => {
      tray.menus.push(menu);
    },
    setToolTip: (tooltip: string) => {
      tray.tooltip = tooltip;
    },
    on: (_event: 'click', handler: () => void) => {
      tray.clickHandler = handler;
    },
    destroy: () => {
      tray.destroyed = true;
    },
  };
  return tray;
}

type FakeTray = ReturnType<typeof makeTray>;

function createHarness(overrides: Partial<EnsureTrayDeps> & { existingTray?: FakeTray } = {}) {
  const { existingTray, ...depOverrides } = overrides;
  const state = {
    tray: (existingTray ?? null) as FakeTray | null,
    createdTrays: [] as FakeTray[],
    warnings: [] as string[],
    overlayShows: 0,
  };
  const ensureTray = createEnsureTrayHandler({
    getTray: () => state.tray,
    setTray: (tray) => {
      state.tray = tray as FakeTray | null;
    },
    buildTrayMenu: () => ({ id: 'menu' }),
    resolveTrayIconPath: () => '/tmp/icon.png',
    createImageFromPath: (iconPath) => makeImage(iconPath),
    createEmptyImage: () => makeImage('empty', true),
    createTray: (icon) => {
      const tray = makeTray(icon);
      state.createdTrays.push(tray);
      return tray;
    },
    trayTooltip: 'SubMiner',
    platform: 'darwin',
    logWarn: (message) => state.warnings.push(message),
    ensureOverlayVisibleFromTrayClick: () => {
      state.overlayShows += 1;
    },
    ...depOverrides,
  });
  return { ensureTray, state };
}

test('ensure tray only refreshes the menu of an existing tray', () => {
  for (const platform of ['darwin', 'linux']) {
    const existingTray = makeTray();
    const { ensureTray, state } = createHarness({ platform, existingTray });

    ensureTray();

    assert.deepEqual(existingTray.menus, [{ id: 'menu' }], platform);
    assert.equal(existingTray.tooltip, null, platform);
    assert.equal(existingTray.clickHandler, null, platform);
    assert.equal(state.createdTrays.length, 0, platform);
    assert.equal(state.tray, existingTray, platform);
  }
});

test('ensure tray creates a darwin tray with an 18px template icon and click handler', () => {
  const { ensureTray, state } = createHarness({ platform: 'darwin' });

  ensureTray();

  const tray = state.tray;
  assert.ok(tray);
  const icon = tray.icon as FakeImage;
  assert.equal(icon.label, '/tmp/icon.png@18x18');
  assert.equal(icon.template, true);
  assert.equal(tray.tooltip, 'SubMiner');
  assert.deepEqual(tray.menus, [{ id: 'menu' }]);

  tray.clickHandler?.();
  assert.equal(state.overlayShows, 1);
});

test('ensure tray creates a linux tray with a 20px non-template icon', () => {
  const { ensureTray, state } = createHarness({ platform: 'linux' });

  ensureTray();

  const icon = state.tray?.icon as FakeImage;
  assert.equal(icon.label, '/tmp/icon.png@20x20');
  assert.equal(icon.template, false);
  assert.equal(state.tray?.tooltip, 'SubMiner');
});

test('ensure tray logs Linux tray registration failures without crashing startup', () => {
  const { ensureTray, state } = createHarness({
    platform: 'linux',
    createTray: () => {
      throw new Error('StatusNotifier watcher unavailable');
    },
  });

  ensureTray();

  assert.equal(state.tray, null);
  assert.deepEqual(state.warnings, [
    'Unable to create Linux tray icon. Ensure your desktop has a StatusNotifier/AppIndicator tray host. StatusNotifier watcher unavailable',
  ]);
});

test('destroy tray handler destroys active tray and clears ref', () => {
  let tray: FakeTray | null = makeTray();
  const activeTray = tray;
  const destroyTray = createDestroyTrayHandler({
    getTray: () => tray,
    setTray: (next) => {
      tray = next as FakeTray | null;
    },
  });

  destroyTray();
  assert.equal(activeTray.destroyed, true);
  assert.equal(tray, null);
});
