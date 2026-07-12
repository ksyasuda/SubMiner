import assert from 'node:assert/strict';
import test from 'node:test';
import { createTrayRuntimeHandlers } from './tray-runtime-handlers';

test('tray runtime handlers compose resolve/menu/ensure/destroy handlers', () => {
  let tray: { destroyed: boolean } | null = null;
  let overlayInitialized = false;
  let visibleOverlay = false;
  const calls: string[] = [];

  const runtime = createTrayRuntimeHandlers({
    resolveTrayIconPathDeps: {
      resolveTrayIconPathRuntime: () => '/tmp/SubMiner.png',
      platform: 'darwin',
      resourcesPath: '/resources',
      appPath: '/app',
      dirname: '/dirname',
      joinPath: (...parts) => parts.join('/'),
      fileExists: () => true,
    },
    buildTrayMenuTemplateDeps: {
      buildTrayMenuTemplateRuntime: () => [{ label: 'Open Help' }],
      initializeOverlayRuntime: () => {
        overlayInitialized = true;
      },
      isOverlayRuntimeInitialized: () => overlayInitialized,
      openSessionHelpModal: () => {},
      openTexthookerInBrowser: () => {},
      showTexthookerPage: () => true,
      showFirstRunSetup: () => true,
      openFirstRunSetupWindow: () => {},
      showWindowsMpvLauncherSetup: () => true,
      openYomitanSettings: () => {},
      openConfigSettingsWindow: () => {},
      openSyncUiWindow: () => {},
      exportLogs: () => {},
      openJellyfinSetupWindow: () => {},
      isJellyfinConfigured: () => false,
      isJellyfinDiscoveryActive: () => false,
      toggleJellyfinDiscovery: () => {},
      openAnilistSetupWindow: () => {},
      checkForUpdates: () => {},
      quitApp: () => {},
    },
    ensureTrayDeps: {
      getTray: () => tray as never,
      setTray: (next) => {
        tray = next as { destroyed: boolean } | null;
      },
      createImageFromPath: () => ({
        isEmpty: () => false,
        resize: () => ({
          isEmpty: () => false,
          resize: () => ({}) as never,
          setTemplateImage: () => {},
        }),
        setTemplateImage: () => {},
      }),
      createEmptyImage: () => ({
        isEmpty: () => true,
        resize: () => ({}) as never,
        setTemplateImage: () => {},
      }),
      createTray: () => ({
        setContextMenu: () => calls.push('setContextMenu'),
        setToolTip: () => calls.push('setToolTip'),
        on: (_event: 'click', handler: () => void) => {
          calls.push('on-click');
          handler();
        },
        destroy: () => {
          calls.push('destroy');
          if (tray) tray.destroyed = true;
        },
      }),
      trayTooltip: 'SubMiner',
      platform: 'darwin',
      logWarn: (message) => calls.push(`warn:${message}`),
      initializeOverlayRuntime: () => {
        overlayInitialized = true;
      },
      isOverlayRuntimeInitialized: () => overlayInitialized,
      setVisibleOverlayVisible: (visible) => {
        visibleOverlay = visible;
      },
    },
    destroyTrayDeps: {
      getTray: () => tray as never,
      setTray: (next) => {
        tray = next as { destroyed: boolean } | null;
      },
    },
    buildMenuFromTemplate: (template) => ({ template }),
  });

  assert.equal(runtime.resolveTrayIconPath(), '/tmp/SubMiner.png');
  assert.deepEqual(runtime.buildTrayMenu(), { template: [{ label: 'Open Help' }] });
  runtime.ensureTray();
  assert.equal(overlayInitialized, true);
  assert.equal(visibleOverlay, true);
  assert.ok(tray);
  runtime.destroyTray();
  assert.equal(tray, null);
  assert.deepEqual(calls, ['setToolTip', 'setContextMenu', 'on-click', 'destroy']);
});
