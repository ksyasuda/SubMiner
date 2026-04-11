import assert from 'node:assert/strict';
import test from 'node:test';
import { createMainBootServices } from './services';

test('createMainBootServices builds boot-phase service bundle', () => {
  type MockAppLifecycleApp = {
    requestSingleInstanceLock: () => boolean;
    quit: () => void;
    on: (event: string, listener: (...args: unknown[]) => void) => MockAppLifecycleApp;
    whenReady: () => Promise<void>;
  };

  const calls: string[] = [];
  let setPathValue: string | null = null;
  const appOnCalls: string[] = [];
  let secondInstanceHandlerRegistered = false;

  const services = createMainBootServices<
    { configDir: string },
    { targetPath: string },
    { targetPath: string },
    { targetPath: string },
    { kind: string },
    { scope: string; warn: () => void; info: () => void; error: () => void },
    { registry: boolean },
    { getMainWindow: () => null; getModalWindow: () => null },
    {
      inputState: boolean;
      getModalInputExclusive: () => boolean;
      handleModalInputStateChange: (isActive: boolean) => void;
    },
    { measurementStore: boolean },
    { modalRuntime: boolean },
    { mpvSocketPath: string; texthookerPort: number },
    MockAppLifecycleApp
  >({
    platform: 'linux',
    argv: ['node', 'main.ts'],
    appDataDir: undefined,
    xdgConfigHome: undefined,
    homeDir: '/home/tester',
    defaultMpvLogFile: '/tmp/default.log',
    envMpvLog: ' /tmp/custom.log ',
    defaultTexthookerPort: 5174,
    getDefaultSocketPath: () => '/tmp/subminer.sock',
    resolveConfigDir: () => '/tmp/subminer-config',
    existsSync: () => false,
    mkdirSync: (targetPath) => {
      calls.push(`mkdir:${targetPath}`);
    },
    joinPath: (...parts) => parts.join('/'),
    app: {
      setPath: (_name, value) => {
        setPathValue = value;
      },
      quit: () => {},
      on: (event: string) => {
        appOnCalls.push(event);
        return {};
      },
      whenReady: async () => {},
    },
    shouldBypassSingleInstanceLock: () => false,
    requestSingleInstanceLockEarly: () => true,
    registerSecondInstanceHandlerEarly: () => {
      secondInstanceHandlerRegistered = true;
    },
    onConfigStartupParseError: () => {
      throw new Error('unexpected parse failure');
    },
    createConfigService: (configDir) => ({ configDir }),
    createAnilistTokenStore: (targetPath) => ({ targetPath }),
    createJellyfinTokenStore: (targetPath) => ({ targetPath }),
    createAnilistUpdateQueue: (targetPath) => ({ targetPath }),
    createSubtitleWebSocket: () => ({ kind: 'ws' }),
    createLogger: (scope) =>
      ({
        scope,
        warn: () => {},
        info: () => {},
        error: () => {},
      }) as const,
    createMainRuntimeRegistry: () => ({ registry: true }),
    createOverlayManager: () => ({
      getMainWindow: () => null,
      getModalWindow: () => null,
    }),
    createOverlayModalInputState: () => ({
      inputState: true,
      getModalInputExclusive: () => false,
      handleModalInputStateChange: () => {},
    }),
    createOverlayContentMeasurementStore: () => ({ measurementStore: true }),
    getSyncOverlayShortcutsForModal: () => () => {},
    getSyncOverlayVisibilityForModal: () => () => {},
    createOverlayModalRuntime: () => ({ modalRuntime: true }),
    createAppState: (input) => ({ ...input }),
  });

  assert.equal(services.configDir, '/tmp/subminer-config');
  assert.equal(services.userDataPath, '/tmp/subminer-config');
  assert.equal(services.defaultMpvLogPath, '/tmp/custom.log');
  assert.equal(services.defaultImmersionDbPath, '/tmp/subminer-config/immersion.sqlite');
  assert.deepEqual(services.configService, { configDir: '/tmp/subminer-config' });
  assert.deepEqual(services.anilistTokenStore, {
    targetPath: '/tmp/subminer-config/anilist-token-store.json',
  });
  assert.deepEqual(services.jellyfinTokenStore, {
    targetPath: '/tmp/subminer-config/jellyfin-token-store.json',
  });
  assert.deepEqual(services.anilistUpdateQueue, {
    targetPath: '/tmp/subminer-config/anilist-retry-queue.json',
  });
  assert.deepEqual(services.appState, {
    mpvSocketPath: '/tmp/subminer.sock',
    texthookerPort: 5174,
  });
  assert.equal(
    services.appLifecycleApp.on('ready', () => {}),
    services.appLifecycleApp,
  );
  assert.equal(
    services.appLifecycleApp.on('second-instance', () => {}),
    services.appLifecycleApp,
  );
  assert.deepEqual(appOnCalls, ['ready']);
  assert.equal(secondInstanceHandlerRegistered, true);
  assert.deepEqual(calls, ['mkdir:/tmp/subminer-config']);
  assert.equal(setPathValue, '/tmp/subminer-config');
});
