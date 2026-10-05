import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createBuildTrayMenuTemplateHandler,
  shouldShowTexthookerTrayEntry,
} from './tray-main-actions';

type TrayTemplateDeps = Parameters<typeof createBuildTrayMenuTemplateHandler<string>>[0];
type TrayTemplateHandlers = Parameters<TrayTemplateDeps['buildTrayMenuTemplateRuntime']>[0];

// Builds the tray template and returns the handlers it passed to the runtime builder.
function buildHandlers(overrides: Partial<TrayTemplateDeps> = {}) {
  const calls: string[] = [];
  let handlers: TrayTemplateHandlers | null = null;
  const buildTemplate = createBuildTrayMenuTemplateHandler<string>({
    buildTrayMenuTemplateRuntime: (nextHandlers) => {
      handlers = nextHandlers;
      return ['ok'];
    },
    initializeOverlayRuntime: () => calls.push('init'),
    isOverlayRuntimeInitialized: () => false,
    openSessionHelpModal: () => calls.push('help'),
    openChangelogModal: () => calls.push('changelog'),
    openTexthookerInBrowser: () => {},
    showTexthookerPage: () => true,
    showFirstRunSetup: () => true,
    openFirstRunSetupWindow: (force?: boolean) => calls.push(force ? 'setup-forced' : 'setup'),
    showWindowsMpvLauncherSetup: () => false,
    openYomitanSettings: () => {},
    openConfigSettingsWindow: () => {},
    openSyncUiWindow: () => {},
    openYoutubeBrowserWindow: () => {},
    exportLogs: () => {},
    openJellyfinSetupWindow: () => {},
    isJellyfinConfigured: () => false,
    isJellyfinDiscoveryActive: () => false,
    toggleJellyfinDiscovery: () => {},
    openAnilistSetupWindow: () => {},
    checkForUpdates: () => {},
    quitApp: () => {},
    ...overrides,
  });

  assert.deepEqual(buildTemplate(), ['ok']);
  assert.ok(handlers);
  return { handlers: handlers as TrayTemplateHandlers, calls };
}

test('tray modal actions initialize the overlay runtime only once', () => {
  let initialized = false;
  const { handlers, calls } = buildHandlers({
    initializeOverlayRuntime: () => {
      initialized = true;
      calls.push('init');
    },
    isOverlayRuntimeInitialized: () => initialized,
  });

  handlers.openSessionHelp();
  handlers.openChangelog();
  handlers.openSessionHelp();

  assert.deepEqual(calls, ['init', 'help', 'changelog', 'help']);
});

test('first-run setup opens normally while the windows mpv launcher action forces it open', () => {
  const { handlers, calls } = buildHandlers({
    showFirstRunSetup: () => false,
    showWindowsMpvLauncherSetup: () => true,
  });

  assert.equal(handlers.showFirstRunSetup, false);
  assert.equal(handlers.showWindowsMpvLauncherSetup, true);
  handlers.openFirstRunSetup();
  handlers.openWindowsMpvLauncherSetup();
  assert.deepEqual(calls, ['setup', 'setup-forced']);
});

test('jellyfin discovery entry follows configuration and forwards toggles', async () => {
  const toggles: boolean[] = [];
  const { handlers } = buildHandlers({
    isJellyfinConfigured: () => true,
    isJellyfinDiscoveryActive: () => true,
    toggleJellyfinDiscovery: async (checked) => {
      toggles.push(checked);
    },
  });

  assert.equal(handlers.showJellyfinDiscovery, true);
  assert.equal(handlers.jellyfinDiscoveryActive, true);
  handlers.toggleJellyfinDiscovery(false);
  assert.deepEqual(toggles, [false]);

  assert.equal(buildHandlers().handlers.showJellyfinDiscovery, false);
});

test('texthooker tray visibility follows websocket server enabled state', () => {
  assert.equal(
    shouldShowTexthookerTrayEntry({
      websocket: { enabled: false },
      annotationWebsocket: { enabled: false },
    }),
    false,
  );
  assert.equal(
    shouldShowTexthookerTrayEntry({
      websocket: { enabled: true },
      annotationWebsocket: { enabled: false },
    }),
    true,
  );
  assert.equal(
    shouldShowTexthookerTrayEntry({
      websocket: { enabled: 'auto' },
      annotationWebsocket: { enabled: false },
    }),
    true,
  );
  assert.equal(
    shouldShowTexthookerTrayEntry({
      websocket: { enabled: false },
      annotationWebsocket: { enabled: true },
    }),
    true,
  );
});
