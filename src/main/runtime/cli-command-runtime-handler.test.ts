import assert from 'node:assert/strict';
import test from 'node:test';
import { createCliCommandRuntimeHandler } from './cli-command-runtime-handler';

test('cli command runtime handler applies precheck and forwards command with context', () => {
  const calls: string[] = [];
  const handler = createCliCommandRuntimeHandler({
    handleTexthookerOnlyModeTransitionMainDeps: {
      isTexthookerOnlyMode: () => true,
      setTexthookerOnlyMode: () => calls.push('set-mode'),
      commandNeedsOverlayRuntime: () => true,
      ensureOverlayStartupPrereqs: () => calls.push('prereqs'),
      startBackgroundWarmups: () => calls.push('warmups'),
      logInfo: (message) => calls.push(`log:${message}`),
    },
    createCliCommandContext: () => {
      calls.push('context');
      return { id: 'ctx' };
    },
    handleCliCommandRuntimeServiceWithContext: (_args, source, context) => {
      calls.push(`cli:${source}:${context.id}`);
    },
  });

  handler({ start: true } as never);

  assert.deepEqual(calls, [
    'prereqs',
    'set-mode',
    'log:Disabling texthooker-only mode after overlay/start command.',
    'warmups',
    'context',
    'cli:initial:ctx',
  ]);
});

test('cli command runtime handler prepares overlay prerequisites before overlay runtime commands', () => {
  const calls: string[] = [];
  const handler = createCliCommandRuntimeHandler({
    handleTexthookerOnlyModeTransitionMainDeps: {
      isTexthookerOnlyMode: () => false,
      setTexthookerOnlyMode: () => calls.push('set-mode'),
      commandNeedsOverlayRuntime: () => true,
      ensureOverlayStartupPrereqs: () => calls.push('prereqs'),
      startBackgroundWarmups: () => calls.push('warmups'),
      logInfo: (message) => calls.push(`log:${message}`),
    },
    createCliCommandContext: () => {
      calls.push('context');
      return { id: 'ctx' };
    },
    handleCliCommandRuntimeServiceWithContext: (_args, source, context) => {
      calls.push(`cli:${source}:${context.id}`);
    },
  });

  handler({ youtubePlay: 'https://www.youtube.com/watch?v=test' } as never);

  assert.deepEqual(calls, ['prereqs', 'context', 'cli:initial:ctx']);
});
