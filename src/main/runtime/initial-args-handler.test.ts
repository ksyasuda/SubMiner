import test from 'node:test';
import assert from 'node:assert/strict';
import { createHandleInitialArgsHandler } from './initial-args-handler';

test('initial args handler no-ops without initial args', () => {
  let handled = false;
  const handleInitialArgs = createHandleInitialArgsHandler({
    getInitialArgs: () => null,
    isBackgroundMode: () => false,
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
