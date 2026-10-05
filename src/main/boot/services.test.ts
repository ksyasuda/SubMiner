import assert from 'node:assert/strict';
import test from 'node:test';
import { ConfigStartupParseError } from '../../config';
import { createMainBootServices } from './services';

// Minimal params whose factories echo their inputs, so tests can assert what each service received.
function baseParams() {
  return {
    platform: 'linux' as NodeJS.Platform,
    argv: ['node', 'main.ts'],
    appDataDir: undefined,
    xdgConfigHome: undefined,
    homeDir: '/home/tester',
    defaultMpvLogFile: '/tmp/default.log',
    envMpvLog: undefined as string | undefined,
    defaultTexthookerPort: 5174,
    getDefaultSocketPath: () => '/tmp/subminer.sock',
    resolveConfigDir: () => '/tmp/subminer-config',
    existsSync: (_targetPath: string) => false,
    mkdirSync: (_targetPath: string) => {},
    joinPath: (...parts: string[]) => parts.join('/'),
    app: {
      setPath: (_name: string, _value: string) => {},
      quit: () => {},
      exit: (_code?: number) => {},
      on: (_event: string, _listener: unknown) => ({}),
      whenReady: async () => {},
    },
    shouldBypassSingleInstanceLock: () => false,
    requestSingleInstanceLockEarly: () => true,
    registerSecondInstanceHandlerEarly: (_listener: unknown) => {},
    onConfigStartupParseError: (_error: ConfigStartupParseError) => {},
    createConfigService: (configDir: string) => ({ configDir }),
    createAnilistTokenStore: (targetPath: string) => ({ targetPath }),
    createJellyfinTokenStore: (targetPath: string) => ({ targetPath }),
    createAnilistUpdateQueue: (targetPath: string) => ({ targetPath }),
    createSubtitleWebSocket: (payloadMode: 'plain' | 'annotated') => ({ payloadMode }),
    createLogger: () => ({ warn: () => {}, info: () => {}, error: () => {} }),
    createMainRuntimeRegistry: () => ({}),
    createOverlayManager: () => ({ getMainWindow: () => null, getModalWindow: () => null }),
    createOverlayModalInputState: () => ({
      getModalInputExclusive: () => false,
      handleModalInputStateChange: (_isActive: boolean) => {},
    }),
    createOverlayContentMeasurementStore: () => ({}),
    getSyncOverlayShortcutsForModal: () => (_isActive: boolean) => {},
    getSyncOverlayVisibilityForModal: () => () => {},
    createOverlayModalRuntime: () => ({}),
    createAppState: (input: { mpvSocketPath: string; texthookerPort: number }) => input,
  };
}

function createServices(
  overrides: Partial<ReturnType<typeof baseParams>> & { configDir?: string },
) {
  return createMainBootServices({ ...baseParams(), ...overrides });
}

test('createMainBootServices derives data paths from the resolved config dir', () => {
  const services = createServices({});

  assert.equal(services.configDir, '/tmp/subminer-config');
  assert.equal(services.userDataPath, '/tmp/subminer-config');
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
  assert.deepEqual(services.subtitleWsService, { payloadMode: 'plain' });
  assert.deepEqual(services.annotationSubtitleWsService, { payloadMode: 'annotated' });
  assert.deepEqual(services.appState, {
    mpvSocketPath: '/tmp/subminer.sock',
    texthookerPort: 5174,
  });
});

test('createMainBootServices creates the user data dir only when missing and registers it', () => {
  for (const exists of [false, true]) {
    const created: string[] = [];
    const setPaths: Array<[string, string]> = [];
    createServices({
      existsSync: () => exists,
      mkdirSync: (targetPath) => {
        created.push(targetPath);
      },
      app: {
        ...baseParams().app,
        setPath: (name, value) => {
          setPaths.push([name, value]);
        },
      },
    });

    assert.deepEqual(created, exists ? [] : ['/tmp/subminer-config']);
    assert.deepEqual(setPaths, [['userData', '/tmp/subminer-config']]);
  }
});

test('createMainBootServices uses a trimmed MPV log env path and falls back when blank', () => {
  const cases = [
    { envMpvLog: ' /tmp/custom.log ', expected: '/tmp/custom.log' },
    { envMpvLog: '   ', expected: '/tmp/default.log' },
    { envMpvLog: undefined, expected: '/tmp/default.log' },
  ];
  for (const { envMpvLog, expected } of cases) {
    assert.equal(createServices({ envMpvLog }).defaultMpvLogPath, expected);
  }
});

test('createMainBootServices honors the profile selected by the early entrypoint', () => {
  const services = createServices({
    configDir: '/tmp/SubMiner-dev',
    resolveConfigDir: () => {
      throw new Error('early profile should be authoritative');
    },
  });

  assert.equal(services.configDir, '/tmp/SubMiner-dev');
  assert.equal(services.userDataPath, '/tmp/SubMiner-dev');
  assert.deepEqual(services.configService, { configDir: '/tmp/SubMiner-dev' });
});

test('createMainBootServices routes second-instance listeners to the early handler', () => {
  const appEvents: string[] = [];
  const earlyListeners: unknown[] = [];
  const services = createServices({
    app: {
      ...baseParams().app,
      on: (event) => {
        appEvents.push(event);
        return {};
      },
    },
    registerSecondInstanceHandlerEarly: (listener) => {
      earlyListeners.push(listener);
    },
  });

  const secondInstanceListener = () => {};
  services.appLifecycleApp.on('ready', () => {});
  services.appLifecycleApp.on('second-instance', secondInstanceListener);

  assert.deepEqual(appEvents, ['ready']);
  assert.deepEqual(earlyListeners, [secondInstanceListener]);
});

test('createMainBootServices skips the single-instance lock when bypassed', () => {
  for (const bypass of [false, true]) {
    let lockRequests = 0;
    const services = createServices({
      shouldBypassSingleInstanceLock: () => bypass,
      requestSingleInstanceLockEarly: () => {
        lockRequests += 1;
        return false;
      },
    });

    assert.equal(services.appLifecycleApp.requestSingleInstanceLock(), bypass);
    assert.equal(lockRequests, bypass ? 0 : 1);
  }
});

test('createMainBootServices reports config parse errors and rethrows', () => {
  const parseError = new ConfigStartupParseError('/tmp/config.jsonc', 'unexpected token');
  const reported: ConfigStartupParseError[] = [];

  assert.throws(
    () =>
      createServices({
        createConfigService: () => {
          throw parseError;
        },
        onConfigStartupParseError: (error) => {
          reported.push(error);
        },
      }),
    (error) => error === parseError,
  );
  assert.deepEqual(reported, [parseError]);
});
