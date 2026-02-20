import assert from 'node:assert/strict';
import test from 'node:test';
import { createBuildHandleTexthookerOnlyModeTransitionMainDepsHandler } from './cli-command-prechecks-main-deps';

test('cli prechecks main deps builder maps transition handlers', () => {
  const calls: string[] = [];
  const deps = createBuildHandleTexthookerOnlyModeTransitionMainDepsHandler({
    isTexthookerOnlyMode: () => true,
    setTexthookerOnlyMode: (enabled) => calls.push(`set:${enabled}`),
    commandNeedsOverlayRuntime: () => true,
    startBackgroundWarmups: () => calls.push('warmups'),
    logInfo: (message) => calls.push(`info:${message}`),
  })();

  assert.equal(deps.isTexthookerOnlyMode(), true);
  assert.equal(deps.commandNeedsOverlayRuntime({} as never), true);
  deps.setTexthookerOnlyMode(false);
  deps.startBackgroundWarmups();
  deps.logInfo('x');
  assert.deepEqual(calls, ['set:false', 'warmups', 'info:x']);
});
