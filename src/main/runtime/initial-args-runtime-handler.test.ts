import assert from 'node:assert/strict';
import test from 'node:test';
import { createInitialArgsRuntimeHandler } from './initial-args-runtime-handler';

test('initial args runtime handler composes main deps and runs initial command flow', () => {
  const calls: string[] = [];
  const handleInitialArgs = createInitialArgsRuntimeHandler({
    getInitialArgs: () => ({ start: true }) as never,
    isBackgroundMode: () => true,
    shouldEnsureTrayOnStartup: () => false,
    shouldRunHeadlessInitialCommand: () => false,
    ensureTray: () => calls.push('tray'),
    isTexthookerOnlyMode: () => false,
    hasImmersionTracker: () => true,
    getMpvClient: () => ({
      connected: false,
      connect: () => calls.push('connect'),
    }),
    commandNeedsOverlayStartupPrereqs: () => false,
    commandNeedsOverlayRuntime: () => false,
    ensureOverlayStartupPrereqs: () => calls.push('prereqs'),
    isOverlayRuntimeInitialized: () => false,
    initializeOverlayRuntime: () => calls.push('init-overlay'),
    logInfo: (message) => calls.push(`log:${message}`),
    handleCliCommand: (_args, source) => calls.push(`cli:${source}`),
  });

  handleInitialArgs();

  assert.deepEqual(calls, [
    'tray',
    'log:Auto-connecting MPV client for immersion tracking',
    'connect',
    'cli:initial',
  ]);
});

test('initial args runtime handler skips mpv auto-connect for stats mode', () => {
  const calls: string[] = [];
  const handleInitialArgs = createInitialArgsRuntimeHandler({
    getInitialArgs: () => ({ stats: true }) as never,
    isBackgroundMode: () => false,
    shouldEnsureTrayOnStartup: () => false,
    shouldRunHeadlessInitialCommand: () => false,
    ensureTray: () => calls.push('tray'),
    isTexthookerOnlyMode: () => false,
    hasImmersionTracker: () => true,
    getMpvClient: () => ({
      connected: false,
      connect: () => calls.push('connect'),
    }),
    commandNeedsOverlayStartupPrereqs: () => false,
    commandNeedsOverlayRuntime: () => false,
    ensureOverlayStartupPrereqs: () => calls.push('prereqs'),
    isOverlayRuntimeInitialized: () => false,
    initializeOverlayRuntime: () => calls.push('init-overlay'),
    logInfo: (message) => calls.push(`log:${message}`),
    handleCliCommand: (_args, source) => calls.push(`cli:${source}`),
  });

  handleInitialArgs();

  assert.deepEqual(calls, ['cli:initial']);
});

test('initial args runtime handler skips tray and mpv auto-connect for headless refresh', () => {
  const calls: string[] = [];
  const handleInitialArgs = createInitialArgsRuntimeHandler({
    getInitialArgs: () => ({ refreshKnownWords: true }) as never,
    isBackgroundMode: () => true,
    shouldEnsureTrayOnStartup: () => true,
    shouldRunHeadlessInitialCommand: () => true,
    ensureTray: () => calls.push('tray'),
    isTexthookerOnlyMode: () => false,
    hasImmersionTracker: () => true,
    getMpvClient: () => ({
      connected: false,
      connect: () => calls.push('connect'),
    }),
    commandNeedsOverlayStartupPrereqs: () => true,
    commandNeedsOverlayRuntime: () => true,
    ensureOverlayStartupPrereqs: () => calls.push('prereqs'),
    isOverlayRuntimeInitialized: () => false,
    initializeOverlayRuntime: () => calls.push('init-overlay'),
    logInfo: (message) => calls.push(`log:${message}`),
    handleCliCommand: (_args, source) => calls.push(`cli:${source}`),
  });

  handleInitialArgs();

  assert.deepEqual(calls, ['cli:initial']);
});
