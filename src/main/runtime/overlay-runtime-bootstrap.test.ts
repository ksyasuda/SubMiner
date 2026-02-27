import test from 'node:test';
import assert from 'node:assert/strict';
import { createInitializeOverlayRuntimeHandler } from './overlay-runtime-bootstrap';

test('overlay runtime bootstrap no-ops when already initialized', () => {
  let coreCalls = 0;
  const initialize = createInitializeOverlayRuntimeHandler({
    isOverlayRuntimeInitialized: () => true,
    initializeOverlayRuntimeCore: () => {
      coreCalls += 1;
    },
    buildOptions: () => ({} as never),
    setOverlayRuntimeInitialized: () => {},
    startBackgroundWarmups: () => {},
  });

  initialize();
  assert.equal(coreCalls, 0);
});

test('overlay runtime bootstrap runs core init and applies post-init state', () => {
  const calls: string[] = [];
  let initialized = false;
  const initialize = createInitializeOverlayRuntimeHandler({
    isOverlayRuntimeInitialized: () => initialized,
    initializeOverlayRuntimeCore: () => {
      calls.push('core');
    },
    buildOptions: () => {
      calls.push('options');
      return {} as never;
    },
    setOverlayRuntimeInitialized: (value) => {
      initialized = value;
      calls.push(`initialized:${value ? 'yes' : 'no'}`);
    },
    startBackgroundWarmups: () => {
      calls.push('warmups');
    },
  });

  initialize();
  assert.equal(initialized, true);
  assert.deepEqual(calls, ['options', 'initialized:yes', 'core', 'warmups']);
});
