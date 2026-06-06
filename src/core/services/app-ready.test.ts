import test from 'node:test';
import assert from 'node:assert/strict';
import { AppReadyRuntimeDeps, runAppReadyRuntime } from './startup';

function waitTurn(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

function makeDeps(overrides: Partial<AppReadyRuntimeDeps> = {}) {
  const calls: string[] = [];
  const deps = {
    ensureDefaultConfigBootstrap: () => calls.push('ensureDefaultConfigBootstrap'),
    loadSubtitlePosition: () => calls.push('loadSubtitlePosition'),
    resolveKeybindings: () => calls.push('resolveKeybindings'),
    createMpvClient: () => calls.push('createMpvClient'),
    reloadConfig: () => calls.push('reloadConfig'),
    getResolvedConfig: () => ({
      websocket: { enabled: 'auto' },
      secondarySub: {},
    }),
    getConfigWarnings: () => [],
    logConfigWarning: () => calls.push('logConfigWarning'),
    setLogLevel: (level, source) => calls.push(`setLogLevel:${level}:${source}`),
    initRuntimeOptionsManager: () => calls.push('initRuntimeOptionsManager'),
    setSecondarySubMode: (mode) => calls.push(`setSecondarySubMode:${mode}`),
    defaultSecondarySubMode: 'hover',
    defaultWebsocketPort: 9001,
    defaultAnnotationWebsocketPort: 6678,
    defaultTexthookerPort: 5174,
    hasMpvWebsocketPlugin: () => true,
    startSubtitleWebsocket: (port) => calls.push(`startSubtitleWebsocket:${port}`),
    startAnnotationWebsocket: (port) => calls.push(`startAnnotationWebsocket:${port}`),
    startTexthooker: (port, websocketUrl) =>
      calls.push(`startTexthooker:${port}:${websocketUrl ?? ''}`),
    log: (message) => calls.push(`log:${message}`),
    createMecabTokenizerAndCheck: async () => {
      calls.push('createMecabTokenizerAndCheck');
    },
    createSubtitleTimingTracker: () => calls.push('createSubtitleTimingTracker'),
    createImmersionTracker: () => calls.push('createImmersionTracker'),
    startJellyfinRemoteSession: async () => {
      calls.push('startJellyfinRemoteSession');
    },
    loadYomitanExtension: async () => {
      calls.push('loadYomitanExtension');
    },
    handleFirstRunSetup: async () => {
      calls.push('handleFirstRunSetup');
    },
    prewarmSubtitleDictionaries: async () => {
      calls.push('prewarmSubtitleDictionaries');
    },
    startBackgroundWarmups: () => {
      calls.push('startBackgroundWarmups');
    },
    texthookerOnlyMode: false,
    shouldAutoInitializeOverlayRuntimeFromConfig: () => true,
    setVisibleOverlayVisible: (visible) => calls.push(`setVisibleOverlayVisible:${visible}`),
    initializeOverlayRuntime: () => calls.push('initializeOverlayRuntime'),
    handleInitialArgs: () => calls.push('handleInitialArgs'),
    logDebug: (message) => calls.push(`debug:${message}`),
    now: () => 1000,
    ...overrides,
  } as AppReadyRuntimeDeps;
  return { deps, calls };
}

test('runAppReadyRuntime starts websocket in auto mode when plugin missing', async () => {
  const { deps, calls } = makeDeps({
    hasMpvWebsocketPlugin: () => false,
  });
  await runAppReadyRuntime(deps);
  assert.ok(calls.includes('ensureDefaultConfigBootstrap'));
  assert.ok(calls.includes('startSubtitleWebsocket:9001'));
  assert.ok(calls.includes('startAnnotationWebsocket:6678'));
  assert.ok(calls.includes('setVisibleOverlayVisible:true'));
  assert.ok(calls.includes('initializeOverlayRuntime'));
  assert.ok(
    calls.indexOf('setVisibleOverlayVisible:true') < calls.indexOf('initializeOverlayRuntime'),
  );
  assert.ok(calls.includes('startBackgroundWarmups'));
  assert.ok(calls.includes('log:Runtime ready: immersion tracker startup requested.'));
});

test('runAppReadyRuntime starts texthooker on startup when enabled in config', async () => {
  const { deps, calls } = makeDeps({
    getResolvedConfig: () => ({
      websocket: { enabled: 'auto' },
      secondarySub: {},
      texthooker: { launchAtStartup: true },
    }),
  });

  await runAppReadyRuntime(deps);

  assert.ok(calls.includes('startTexthooker:5174:ws://127.0.0.1:6678'));
  assert.ok(calls.indexOf('handleFirstRunSetup') < calls.indexOf('handleInitialArgs'));
  assert.ok(
    calls.indexOf('createMpvClient') < calls.indexOf('startTexthooker:5174:ws://127.0.0.1:6678'),
  );
  assert.ok(
    calls.indexOf('startTexthooker:5174:ws://127.0.0.1:6678') < calls.indexOf('handleInitialArgs'),
  );
});

test('runAppReadyRuntime creates immersion tracker during heavy startup', async () => {
  const { deps, calls } = makeDeps({
    shouldAutoInitializeOverlayRuntimeFromConfig: () => false,
  });

  await runAppReadyRuntime(deps);

  assert.equal(calls.includes('createImmersionTracker'), false);
  assert.ok(calls.includes('log:Runtime ready: immersion tracker startup requested.'));
});

test('runAppReadyRuntime keeps annotation websocket enabled when regular websocket auto-skips', async () => {
  const { deps, calls } = makeDeps({
    getResolvedConfig: () => ({
      websocket: { enabled: 'auto' },
      annotationWebsocket: { enabled: true, port: 6678 },
      secondarySub: {},
      texthooker: { launchAtStartup: true },
    }),
    hasMpvWebsocketPlugin: () => true,
  });

  await runAppReadyRuntime(deps);

  assert.equal(calls.includes('startSubtitleWebsocket:9001'), false);
  assert.ok(calls.includes('startAnnotationWebsocket:6678'));
  assert.ok(calls.includes('startTexthooker:5174:ws://127.0.0.1:6678'));
  assert.ok(calls.includes('log:mpv_websocket detected, skipping built-in WebSocket server'));
});

test('runAppReadyRuntime skips heavy startup when shouldSkipHeavyStartup returns true', async () => {
  const { deps, calls } = makeDeps({
    shouldSkipHeavyStartup: () => true,
    reloadConfig: () => calls.push('reloadConfig'),
    getResolvedConfig: () => {
      calls.push('getResolvedConfig');
      return {
        websocket: { enabled: 'auto' },
        secondarySub: {},
      };
    },
    getConfigWarnings: () => {
      calls.push('getConfigWarnings');
      return [];
    },
    setLogLevel: (level, source) => calls.push(`setLogLevel:${level}:${source}`),
    initRuntimeOptionsManager: () => calls.push('initRuntimeOptionsManager'),
    startBackgroundWarmups: () => calls.push('startBackgroundWarmups'),
    loadSubtitlePosition: () => calls.push('loadSubtitlePosition'),
    resolveKeybindings: () => calls.push('resolveKeybindings'),
    createMpvClient: () => calls.push('createMpvClient'),
    logConfigWarning: () => calls.push('logConfigWarning'),
    startJellyfinRemoteSession: async () => {
      calls.push('startJellyfinRemoteSession');
    },
    createImmersionTracker: () => calls.push('createImmersionTracker'),
    handleInitialArgs: () => calls.push('handleInitialArgs'),
  });

  await runAppReadyRuntime(deps);

  assert.equal(calls.includes('ensureDefaultConfigBootstrap'), true);
  assert.equal(calls.includes('reloadConfig'), true);
  assert.equal(calls.includes('getResolvedConfig'), false);
  assert.equal(calls.includes('getConfigWarnings'), false);
  assert.equal(calls.includes('setLogLevel:warn:config'), false);
  assert.equal(calls.includes('startBackgroundWarmups'), false);
  assert.equal(calls.includes('loadSubtitlePosition'), false);
  assert.equal(calls.includes('resolveKeybindings'), false);
  assert.equal(calls.includes('createMpvClient'), false);
  assert.equal(calls.includes('initRuntimeOptionsManager'), false);
  assert.equal(calls.includes('createImmersionTracker'), false);
  assert.equal(calls.includes('startJellyfinRemoteSession'), false);
  assert.equal(calls.includes('logConfigWarning'), false);
  assert.equal(calls.includes('handleInitialArgs'), true);
  assert.equal(calls.includes('loadYomitanExtension'), true);
  assert.equal(calls.includes('handleFirstRunSetup'), true);
  assert.ok(calls.indexOf('loadYomitanExtension') < calls.indexOf('handleInitialArgs'));
  assert.ok(calls.indexOf('loadYomitanExtension') < calls.indexOf('reloadConfig'));
  assert.ok(calls.indexOf('reloadConfig') < calls.indexOf('handleFirstRunSetup'));
  assert.ok(calls.indexOf('loadYomitanExtension') < calls.indexOf('handleFirstRunSetup'));
  assert.ok(calls.indexOf('handleFirstRunSetup') < calls.indexOf('handleInitialArgs'));
});

test('runAppReadyRuntime keeps websocket startup in texthooker-only mode but skips overlay window', async () => {
  const { deps, calls } = makeDeps({
    texthookerOnlyMode: true,
    reloadConfig: () => calls.push('reloadConfig'),
    handleInitialArgs: () => calls.push('handleInitialArgs'),
  });

  await runAppReadyRuntime(deps);

  assert.ok(calls.includes('reloadConfig'));
  assert.ok(calls.includes('createMpvClient'));
  assert.ok(calls.includes('startAnnotationWebsocket:6678'));
  assert.ok(calls.includes('startTexthooker:5174:ws://127.0.0.1:6678'));
  assert.ok(calls.includes('createSubtitleTimingTracker'));
  assert.ok(calls.includes('handleFirstRunSetup'));
  assert.ok(calls.includes('handleInitialArgs'));
  assert.ok(calls.includes('log:Texthooker-only mode enabled; skipping overlay window.'));
  assert.equal(calls.includes('initializeOverlayRuntime'), false);
  assert.equal(calls.includes('setVisibleOverlayVisible:true'), false);
});

test('runAppReadyRuntime skips Jellyfin remote startup when dependency is not wired', async () => {
  const { deps, calls } = makeDeps({
    startJellyfinRemoteSession: undefined,
  });

  await runAppReadyRuntime(deps);

  assert.equal(calls.includes('startJellyfinRemoteSession'), false);
  assert.ok(calls.includes('createMpvClient'));
  assert.ok(calls.includes('createSubtitleTimingTracker'));
  assert.ok(calls.includes('handleInitialArgs'));
  assert.ok(calls.includes('startBackgroundWarmups'));
  assert.ok(
    calls.includes('initializeOverlayRuntime') ||
      calls.includes('log:Overlay runtime deferred: waiting for explicit overlay command.'),
  );
});

test('runAppReadyRuntime logs when createImmersionTracker dependency is missing', async () => {
  const { deps, calls } = makeDeps({
    createImmersionTracker: undefined,
  });
  await runAppReadyRuntime(deps);
  assert.ok(calls.includes('log:Runtime ready: immersion tracker dependency is missing.'));
});

test('runAppReadyRuntime logs defer message when overlay not auto-started', async () => {
  const { deps, calls } = makeDeps({
    shouldAutoInitializeOverlayRuntimeFromConfig: () => false,
  });
  await runAppReadyRuntime(deps);
  assert.ok(calls.includes('log:Overlay runtime deferred: waiting for explicit overlay command.'));
});

test('runAppReadyRuntime applies config logging level during app-ready', async () => {
  const { deps, calls } = makeDeps({
    getResolvedConfig: () => ({
      websocket: { enabled: 'auto' },
      secondarySub: {},
      logging: { level: 'warn' },
    }),
  });
  await runAppReadyRuntime(deps);
  assert.ok(calls.includes('setLogLevel:warn:config'));
});

test('runAppReadyRuntime does not await background warmups', async () => {
  const calls: string[] = [];
  let releaseWarmup: (() => void) | undefined;
  const warmupGate = new Promise<void>((resolve) => {
    releaseWarmup = resolve;
  });
  const { deps } = makeDeps({
    startBackgroundWarmups: () => {
      calls.push('startBackgroundWarmups');
      void warmupGate.then(() => {
        calls.push('warmupDone');
      });
    },
    handleInitialArgs: () => {
      calls.push('handleInitialArgs');
    },
  });

  await runAppReadyRuntime(deps);
  assert.ok(calls.includes('startBackgroundWarmups'));
  assert.ok(calls.includes('handleInitialArgs'));
  assert.ok(calls.indexOf('startBackgroundWarmups') < calls.indexOf('handleInitialArgs'));
  assert.equal(calls.includes('warmupDone'), false);
  assert.ok(releaseWarmup);
  releaseWarmup();
});

test('runAppReadyRuntime handles managed background initial args before deferred Yomitan wait', async () => {
  const calls: string[] = [];
  let releaseYomitan!: () => void;
  const yomitanGate = new Promise<void>((resolve) => {
    releaseYomitan = resolve;
  });
  const { deps } = makeDeps({
    shouldAutoInitializeOverlayRuntimeFromConfig: () => false,
    shouldHandleInitialArgsBeforeDeferredOverlayWarmup: () => true,
    loadYomitanExtension: async () => {
      calls.push('loadYomitanExtension:start');
      await yomitanGate;
      calls.push('loadYomitanExtension:done');
    },
    handleFirstRunSetup: async () => {
      calls.push('handleFirstRunSetup');
    },
    handleInitialArgs: () => {
      calls.push('handleInitialArgs');
    },
  } as Partial<AppReadyRuntimeDeps>);

  const readyPromise = runAppReadyRuntime(deps);
  await waitTurn();

  try {
    assert.ok(calls.includes('handleFirstRunSetup'));
    assert.ok(calls.includes('handleInitialArgs'));
    assert.equal(calls.includes('loadYomitanExtension:done'), false);
  } finally {
    releaseYomitan();
    await readyPromise;
  }
});

test('runAppReadyRuntime keeps non-managed deferred overlay startup behind Yomitan readiness', async () => {
  const calls: string[] = [];
  let releaseYomitan!: () => void;
  const yomitanGate = new Promise<void>((resolve) => {
    releaseYomitan = resolve;
  });
  const { deps } = makeDeps({
    shouldAutoInitializeOverlayRuntimeFromConfig: () => false,
    shouldHandleInitialArgsBeforeDeferredOverlayWarmup: () => false,
    loadYomitanExtension: async () => {
      calls.push('loadYomitanExtension:start');
      await yomitanGate;
      calls.push('loadYomitanExtension:done');
    },
    handleInitialArgs: () => {
      calls.push('handleInitialArgs');
    },
  } as Partial<AppReadyRuntimeDeps>);

  const readyPromise = runAppReadyRuntime(deps);
  await waitTurn();

  assert.equal(calls.includes('handleInitialArgs'), false);

  releaseYomitan();
  await readyPromise;

  assert.ok(calls.indexOf('loadYomitanExtension:done') < calls.indexOf('handleInitialArgs'));
});

test('runAppReadyRuntime starts background warmups before core runtime services', async () => {
  const calls: string[] = [];
  const { deps } = makeDeps({
    startBackgroundWarmups: () => {
      calls.push('startBackgroundWarmups');
    },
    loadSubtitlePosition: () => calls.push('loadSubtitlePosition'),
    createMpvClient: () => calls.push('createMpvClient'),
  });

  await runAppReadyRuntime(deps);

  assert.ok(calls.indexOf('startBackgroundWarmups') < calls.indexOf('loadSubtitlePosition'));
  assert.ok(calls.indexOf('startBackgroundWarmups') < calls.indexOf('createMpvClient'));
});

test('runAppReadyRuntime exits before service init when critical anki mappings are invalid', async () => {
  const capturedErrors: string[][] = [];
  const { deps, calls } = makeDeps({
    getResolvedConfig: () => ({
      websocket: { enabled: 'auto' },
      secondarySub: {},
      ankiConnect: {
        enabled: true,
        fields: {
          audio: 'ExpressionAudio',
          image: 'Picture',
          sentence: '   ',
          miscInfo: 'MiscInfo',
          translation: '',
        },
      },
    }),
    onCriticalConfigErrors: (errors) => {
      capturedErrors.push(errors);
    },
  });

  await runAppReadyRuntime(deps);

  assert.equal(capturedErrors.length, 1);
  assert.deepEqual(capturedErrors[0], [
    'ankiConnect.fields.sentence must be a non-empty string when ankiConnect is enabled.',
    'ankiConnect.fields.translation must be a non-empty string when ankiConnect is enabled.',
  ]);
  assert.ok(calls.includes('reloadConfig'));
  assert.equal(calls.includes('createMpvClient'), false);
  assert.equal(calls.includes('initRuntimeOptionsManager'), false);
  assert.equal(calls.includes('startBackgroundWarmups'), false);
});

test('runAppReadyRuntime aggregates multiple critical anki mapping errors', async () => {
  const capturedErrors: string[][] = [];
  const { deps, calls } = makeDeps({
    getResolvedConfig: () => ({
      websocket: { enabled: 'auto' },
      secondarySub: {},
      ankiConnect: {
        enabled: true,
        fields: {
          audio: ' ',
          image: '',
          sentence: '\t',
          miscInfo: '   ',
          translation: '',
        },
      },
    }),
    onCriticalConfigErrors: (errors) => {
      capturedErrors.push(errors);
    },
  });

  await runAppReadyRuntime(deps);

  const firstErrorSet = capturedErrors[0]!;
  assert.equal(capturedErrors.length, 1);
  assert.equal(firstErrorSet.length, 5);
  assert.ok(
    firstErrorSet.includes(
      'ankiConnect.fields.audio must be a non-empty string when ankiConnect is enabled.',
    ),
  );
  assert.ok(
    firstErrorSet.includes(
      'ankiConnect.fields.image must be a non-empty string when ankiConnect is enabled.',
    ),
  );
  assert.ok(
    firstErrorSet.includes(
      'ankiConnect.fields.sentence must be a non-empty string when ankiConnect is enabled.',
    ),
  );
  assert.ok(
    firstErrorSet.includes(
      'ankiConnect.fields.miscInfo must be a non-empty string when ankiConnect is enabled.',
    ),
  );
  assert.ok(
    firstErrorSet.includes(
      'ankiConnect.fields.translation must be a non-empty string when ankiConnect is enabled.',
    ),
  );
  assert.equal(calls.includes('loadSubtitlePosition'), false);
});
