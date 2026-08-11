import assert from 'node:assert/strict';
import test from 'node:test';
import { activateMacOSApp } from './macos-app-activation';

test('activateMacOSApp steals app focus on darwin', () => {
  const calls: string[] = [];

  const activated = activateMacOSApp({
    platform: 'darwin',
    stealAppFocus: () => calls.push('steal'),
  });

  assert.equal(activated, true);
  assert.deepEqual(calls, ['steal']);
});

test('activateMacOSApp is a no-op off darwin', () => {
  const calls: string[] = [];

  for (const platform of ['linux', 'win32'] as NodeJS.Platform[]) {
    assert.equal(activateMacOSApp({ platform, stealAppFocus: () => calls.push('steal') }), false);
  }

  assert.deepEqual(calls, []);
});

test('activateMacOSApp warns instead of throwing when activation fails', () => {
  const warnings: string[] = [];

  const activated = activateMacOSApp({
    platform: 'darwin',
    stealAppFocus: () => {
      throw new Error('no window server');
    },
    warn: (message) => warnings.push(message),
  });

  assert.equal(activated, false);
  assert.equal(warnings.length, 1);
});
