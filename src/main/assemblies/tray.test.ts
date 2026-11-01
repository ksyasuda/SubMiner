import assert from 'node:assert/strict';
import test from 'node:test';
import type { Menu, NativeImage, Tray } from 'electron';
import { createTrayAssembly, type TrayAssemblyDeps, type TrayMenuActions } from './tray';

const MENU_ACTIONS = [
  'openSessionHelpModal',
  'openChangelogModal',
  'openTexthookerInBrowser',
  'showTexthookerPage',
  'showFirstRunSetup',
  'openFirstRunSetupWindow',
  'getDictionaryBackend',
  'openHachidoriSettings',
  'openYomitanSettings',
  'openConfigSettingsWindow',
  'openSyncUiWindow',
  'openYoutubeBrowserWindow',
  'exportLogs',
  'openJellyfinSetupWindow',
  'isJellyfinConfigured',
  'isJellyfinDiscoveryActive',
  'toggleJellyfinDiscovery',
  'openAnilistSetupWindow',
  'checkForUpdates',
  'quitApp',
] satisfies Array<keyof TrayMenuActions>;

function menuActions(makeAction: (name: string) => () => unknown): TrayMenuActions {
  return Object.fromEntries(
    MENU_ACTIONS.map((name) => [name, makeAction(name)]),
  ) as unknown as TrayMenuActions;
}

// Every callback throws, so a test fails if the assembly calls a dep it should only call later.
function makeDeps(overrides: Partial<TrayAssemblyDeps> = {}): TrayAssemblyDeps {
  const lazy = (name: string) => (): never => {
    throw new Error(`${name} called unexpectedly`);
  };
  return {
    appState: { overlayRuntimeInitialized: false },
    logger: { warn: lazy('warn') },
    iconRoots: { resourcesPath: '/missing', appPath: '/missing', mainDirname: '/missing' },
    electron: {
      createTray: lazy('createTray'),
      createImageFromPath: lazy('createImageFromPath'),
      createEmptyImage: lazy('createEmptyImage'),
      buildMenuFromTemplate: lazy('buildMenuFromTemplate'),
    },
    menu: menuActions(lazy),
    initializeOverlayRuntime: lazy('initializeOverlayRuntime'),
    setVisibleOverlayVisible: lazy('setVisibleOverlayVisible'),
    ...overrides,
  };
}

// main.ts builds the assembly before most siblings exist, so construction must stay lazy.
test('createTrayAssembly does not call its deps while constructing', () => {
  const tray = createTrayAssembly(makeDeps());
  assert.equal(tray.hasTray(), false);
});

test('refreshMenu rebuilds the menu only once the tray exists', () => {
  let menusBuilt = 0;
  let menusSet = 0;
  const fakeTray = {
    setToolTip: () => {},
    setContextMenu: () => {
      menusSet += 1;
    },
    on: () => {},
  } as unknown as Tray;
  const tray = createTrayAssembly(
    makeDeps({
      logger: { warn: () => {} },
      electron: {
        createTray: () => fakeTray,
        createImageFromPath: () => ({ isEmpty: () => true }) as NativeImage,
        createEmptyImage: () => ({ isEmpty: () => true }) as NativeImage,
        buildMenuFromTemplate: () => {
          menusBuilt += 1;
          return {} as Menu;
        },
      },
      menu: menuActions(() => () => false),
    }),
  );

  tray.refreshMenu();
  assert.equal(menusBuilt, 0);

  tray.ensureTray();
  assert.equal(tray.hasTray(), true);
  const menusAfterCreate = menusSet;
  tray.refreshMenu();
  assert.equal(menusSet, menusAfterCreate + 1);
  assert.equal(menusBuilt, menusSet);
});
