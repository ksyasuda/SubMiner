import test from 'node:test';
import assert from 'node:assert/strict';
import { createHandleInitialArgsHandler } from './initial-args-handler';

test('initial args handler no-ops without initial args', () => {
  let handled = false;
  const handleInitialArgs = createHandleInitialArgsHandler({
    getInitialArgs: () => null,
    isBackgroundMode: () => false,
    shouldEnsureTrayOnStartup: () => false,
    shouldRunHeadlessInitialCommand: () => false,
    ensureTray: () => {},
    isTexthookerOnlyMode: () => false,
    hasImmersionTracker: () => false,
    getMpvClient: () => null,
    commandNeedsOverlayRuntime: () => false,
    ensureOverlayStartupPrereqs: () => {},
    isOverlayRuntimeInitialized: () => false,
    initializeOverlayRuntime: () => {},
    logInfo: () => {},
    handleCliCommand: () => {
      handled = true;
    },
  });

  handleInitialArgs();
  assert.equal(handled, false);
});

test('initial args handler ensures tray in background mode', () => {
  let ensuredTray = false;
  const handleInitialArgs = createHandleInitialArgsHandler({
    getInitialArgs: () => ({ start: true }) as never,
    isBackgroundMode: () => true,
    shouldEnsureTrayOnStartup: () => false,
    shouldRunHeadlessInitialCommand: () => false,
    ensureTray: () => {
      ensuredTray = true;
    },
    isTexthookerOnlyMode: () => true,
    hasImmersionTracker: () => false,
    getMpvClient: () => null,
    commandNeedsOverlayRuntime: () => false,
    ensureOverlayStartupPrereqs: () => {},
    isOverlayRuntimeInitialized: () => false,
    initializeOverlayRuntime: () => {},
    logInfo: () => {},
    handleCliCommand: () => {},
  });

  handleInitialArgs();
  assert.equal(ensuredTray, true);
});

test('initial args handler auto-connects mpv when needed', () => {
  let connectCalls = 0;
  let logged = false;
  const handleInitialArgs = createHandleInitialArgsHandler({
    getInitialArgs: () => ({ start: true }) as never,
    isBackgroundMode: () => false,
    shouldEnsureTrayOnStartup: () => false,
    shouldRunHeadlessInitialCommand: () => false,
    ensureTray: () => {},
    isTexthookerOnlyMode: () => false,
    hasImmersionTracker: () => true,
    getMpvClient: () => ({
      connected: false,
      connect: () => {
        connectCalls += 1;
      },
    }),
    commandNeedsOverlayRuntime: () => false,
    ensureOverlayStartupPrereqs: () => {},
    isOverlayRuntimeInitialized: () => false,
    initializeOverlayRuntime: () => {},
    logInfo: () => {
      logged = true;
    },
    handleCliCommand: () => {},
  });

  handleInitialArgs();
  assert.equal(connectCalls, 1);
  assert.equal(logged, true);
});

test('initial args handler forwards args to cli handler', () => {
  const seenSources: string[] = [];
  const handleInitialArgs = createHandleInitialArgsHandler({
    getInitialArgs: () => ({ start: true }) as never,
    isBackgroundMode: () => false,
    shouldEnsureTrayOnStartup: () => false,
    shouldRunHeadlessInitialCommand: () => false,
    ensureTray: () => {},
    isTexthookerOnlyMode: () => false,
    hasImmersionTracker: () => false,
    getMpvClient: () => null,
    commandNeedsOverlayRuntime: () => false,
    ensureOverlayStartupPrereqs: () => {
      seenSources.push('prereqs');
    },
    isOverlayRuntimeInitialized: () => false,
    initializeOverlayRuntime: () => {
      seenSources.push('init-overlay');
    },
    logInfo: () => {},
    handleCliCommand: (_args, source) => {
      seenSources.push(source);
    },
  });

  handleInitialArgs();
  assert.deepEqual(seenSources, ['initial']);
});

test('initial args handler bootstraps overlay before initial overlay-runtime commands', () => {
  const calls: string[] = [];
  const args = { youtubePlay: 'https://youtube.com/watch?v=abc' } as never;
  const handleInitialArgs = createHandleInitialArgsHandler({
    getInitialArgs: () => args,
    isBackgroundMode: () => false,
    shouldEnsureTrayOnStartup: () => false,
    shouldRunHeadlessInitialCommand: () => false,
    ensureTray: () => {},
    isTexthookerOnlyMode: () => false,
    hasImmersionTracker: () => false,
    getMpvClient: () => null,
    commandNeedsOverlayRuntime: (inputArgs) => inputArgs === args,
    ensureOverlayStartupPrereqs: () => {
      calls.push('prereqs');
    },
    isOverlayRuntimeInitialized: () => false,
    initializeOverlayRuntime: () => {
      calls.push('init-overlay');
    },
    logInfo: () => {},
    handleCliCommand: (_args, source) => {
      calls.push(`cli:${source}`);
    },
  });

  handleInitialArgs();

  assert.deepEqual(calls, ['prereqs', 'init-overlay', 'cli:initial']);
});

test('initial args handler can ensure tray outside background mode when requested', () => {
  let ensuredTray = false;
  const handleInitialArgs = createHandleInitialArgsHandler({
    getInitialArgs: () => ({ start: true }) as never,
    isBackgroundMode: () => false,
    shouldEnsureTrayOnStartup: () => true,
    shouldRunHeadlessInitialCommand: () => false,
    ensureTray: () => {
      ensuredTray = true;
    },
    isTexthookerOnlyMode: () => true,
    hasImmersionTracker: () => false,
    getMpvClient: () => null,
    commandNeedsOverlayRuntime: () => false,
    ensureOverlayStartupPrereqs: () => {},
    isOverlayRuntimeInitialized: () => false,
    initializeOverlayRuntime: () => {},
    logInfo: () => {},
    handleCliCommand: () => {},
  });

  handleInitialArgs();
  assert.equal(ensuredTray, true);
});

test('initial args handler skips tray and mpv auto-connect for headless refresh', () => {
  let ensuredTray = false;
  let connectCalls = 0;
  const handleInitialArgs = createHandleInitialArgsHandler({
    getInitialArgs: () => ({ refreshKnownWords: true }) as never,
    isBackgroundMode: () => true,
    shouldEnsureTrayOnStartup: () => true,
    shouldRunHeadlessInitialCommand: () => true,
    ensureTray: () => {
      ensuredTray = true;
    },
    isTexthookerOnlyMode: () => false,
    hasImmersionTracker: () => true,
    getMpvClient: () => ({
      connected: false,
      connect: () => {
        connectCalls += 1;
      },
    }),
    commandNeedsOverlayRuntime: () => true,
    ensureOverlayStartupPrereqs: () => {},
    isOverlayRuntimeInitialized: () => false,
    initializeOverlayRuntime: () => {},
    logInfo: () => {},
    handleCliCommand: () => {},
  });

  handleInitialArgs();
  assert.equal(ensuredTray, false);
  assert.equal(connectCalls, 0);
});
