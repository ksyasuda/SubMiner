import assert from 'node:assert/strict';
import path from 'node:path';
import test from 'node:test';

import { createMainBootServicesBootstrap } from './main-boot-services-bootstrap';

test('main boot services bootstrap composes grouped inputs into boot services', () => {
  const calls: string[] = [];
  const modalWindow = {} as never;
  const overlayManager = {
    getModalWindow: () => modalWindow,
  };
  type AppStateStub = {
    kind: 'app-state';
    input: {
      mpvSocketPath: string;
      texthookerPort: number;
    };
  };
  type OverlayModalRuntimeStub = {
    kind: 'overlay-modal-runtime';
  };

  const overlayModalInputState = {
    kind: 'overlay-modal-input-state',
    handleModalInputStateChange: (isActive: boolean) => {
      calls.push(`modal-state:${String(isActive)}`);
    },
  };
  const overlayModalInputStateParams: {
    getModalWindow: () => unknown;
    syncOverlayShortcutsForModal: (isActive: boolean) => void;
    syncOverlayVisibilityForModal: () => void;
  }[] = [];
  const createOverlayModalInputState = (params: (typeof overlayModalInputStateParams)[number]) => {
    overlayModalInputStateParams.push(params);
    return overlayModalInputState as never;
  };

  const createOverlayModalRuntime = (params: {
    onModalStateChange: (isActive: boolean) => void;
  }) => {
    calls.push(`modal:${String(params.onModalStateChange(true))}`);
    return { kind: 'overlay-modal-runtime' } as OverlayModalRuntimeStub;
  };

  const boot = createMainBootServicesBootstrap({
    system: {
      platform: 'darwin',
      argv: ['node', 'main.js'],
      appDataDir: '/tmp/app-data',
      xdgConfigHome: '/tmp/xdg',
      homeDir: '/Users/test',
      defaultMpvLogFile: '/tmp/mpv.log',
      envMpvLog: '',
      defaultTexthookerPort: 5174,
      getDefaultSocketPath: () => '/tmp/mpv.sock',
      resolveConfigDir: () => '/tmp/config',
      existsSync: () => true,
      mkdirSync: () => undefined,
      joinPath: (...parts: string[]) => path.posix.join(...parts),
      app: {
        setPath: () => undefined,
        quit: () => undefined,
        on: () => undefined,
        whenReady: async () => undefined,
      },
    },
    singleInstance: {
      shouldBypassSingleInstanceLock: () => false,
      requestSingleInstanceLockEarly: () => true,
      registerSecondInstanceHandlerEarly: () => undefined,
      onConfigStartupParseError: () => undefined,
    },
    factories: {
      createConfigService: () => ({ kind: 'config-service' }) as never,
      createAnilistTokenStore: () => ({ kind: 'anilist-token-store' }) as never,
      createJellyfinTokenStore: () => ({ kind: 'jellyfin-token-store' }) as never,
      createAnilistUpdateQueue: () => ({ kind: 'anilist-update-queue' }) as never,
      createSubtitleWebSocket: () => ({ kind: 'subtitle-websocket' }) as never,
      createLogger: () =>
        ({
          warn: () => undefined,
          info: () => undefined,
          error: () => undefined,
        }) as never,
      createMainRuntimeRegistry: () => ({ kind: 'runtime-registry' }) as never,
      createOverlayManager: () => overlayManager as never,
      createOverlayModalInputState,
      createOverlayContentMeasurementStore: () => ({ kind: 'overlay-content-store' }) as never,
      getSyncOverlayShortcutsForModal: () => (isActive: boolean) => {
        calls.push(`shortcuts:${String(isActive)}`);
      },
      getSyncOverlayVisibilityForModal: () => () => {
        calls.push('visibility');
      },
      createOverlayModalRuntime,
      createAppState: (input) => ({ kind: 'app-state', input }) satisfies AppStateStub,
    },
  });

  assert.equal(boot.configDir, '/tmp/config');
  assert.equal(boot.userDataPath, '/tmp/config');
  assert.equal(boot.defaultImmersionDbPath, '/tmp/config/immersion.sqlite');
  assert.equal(boot.appState.input.mpvSocketPath, '/tmp/mpv.sock');
  assert.equal(boot.appState.input.texthookerPort, 5174);
  assert.equal(overlayModalInputStateParams.length, 1);
  assert.equal(overlayModalInputStateParams[0]?.getModalWindow(), modalWindow);
  overlayModalInputStateParams[0]?.syncOverlayShortcutsForModal(true);
  overlayModalInputStateParams[0]?.syncOverlayVisibilityForModal();
  assert.deepEqual(calls, ['modal-state:true', 'modal:undefined', 'shortcuts:true', 'visibility']);
  assert.equal(boot.overlayManager, overlayManager);
  assert.equal(boot.overlayModalRuntime.kind, 'overlay-modal-runtime');
});
