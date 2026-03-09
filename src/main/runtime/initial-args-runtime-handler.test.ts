import assert from 'node:assert/strict';
import test from 'node:test';
import { createInitialArgsRuntimeHandler } from './initial-args-runtime-handler';

test('initial args runtime handler composes main deps and runs initial command flow', () => {
  const calls: string[] = [];
  const handleInitialArgs = createInitialArgsRuntimeHandler({
    getInitialArgs: () => ({ start: true }) as never,
    isBackgroundMode: () => true,
    shouldEnsureTrayOnStartup: () => false,
    ensureTray: () => calls.push('tray'),
    isTexthookerOnlyMode: () => false,
    hasImmersionTracker: () => true,
    getMpvClient: () => ({
      connected: false,
      connect: () => calls.push('connect'),
    }),
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
