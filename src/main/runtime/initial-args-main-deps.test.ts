import assert from 'node:assert/strict';
import test from 'node:test';
import { createBuildHandleInitialArgsMainDepsHandler } from './initial-args-main-deps';

test('initial args main deps builder maps runtime callbacks and state readers', () => {
  const calls: string[] = [];
  const args = { start: true } as never;
  const mpvClient = { connected: false, connect: () => calls.push('connect') };
  const deps = createBuildHandleInitialArgsMainDepsHandler({
    getInitialArgs: () => args,
    isBackgroundMode: () => true,
    ensureTray: () => calls.push('ensure-tray'),
    isTexthookerOnlyMode: () => false,
    hasImmersionTracker: () => true,
    getMpvClient: () => mpvClient,
    logInfo: (message) => calls.push(`info:${message}`),
    handleCliCommand: (_args, source) => calls.push(`cli:${source}`),
  })();

  assert.equal(deps.getInitialArgs(), args);
  assert.equal(deps.isBackgroundMode(), true);
  assert.equal(deps.isTexthookerOnlyMode(), false);
  assert.equal(deps.hasImmersionTracker(), true);
  assert.equal(deps.getMpvClient(), mpvClient);

  deps.ensureTray();
  deps.logInfo('x');
  deps.handleCliCommand(args, 'initial');
  assert.deepEqual(calls, ['ensure-tray', 'info:x', 'cli:initial']);
});
