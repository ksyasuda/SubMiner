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
    logInfo: () => {},
    handleCliCommand: (_args, source) => {
      seenSources.push(source);
    },
  });

  handleInitialArgs();
  assert.deepEqual(seenSources, ['initial']);
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
    logInfo: () => {},
    handleCliCommand: () => {},
  });

  handleInitialArgs();
  assert.equal(ensuredTray, false);
  assert.equal(connectCalls, 0);
});
