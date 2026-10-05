import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createGetRuntimeOptionsStateHandler,
  createRestorePreviousSecondarySubVisibilityHandler,
} from './overlay-runtime-main-actions';

test('runtime options state handler returns empty list without manager', () => {
  const getState = createGetRuntimeOptionsStateHandler({
    getRuntimeOptionsManager: () => null,
  });
  assert.deepEqual(getState(), []);
});

test('runtime options state handler returns list from manager', () => {
  const getState = createGetRuntimeOptionsStateHandler({
    getRuntimeOptionsManager: () =>
      ({
        listOptions: () => [
          {
            id: 'anki.autoUpdateNewCards',
            label: 'X',
            scope: 'ankiConnect',
            valueType: 'boolean',
            value: true,
            allowedValues: [true, false],
            requiresRestart: false,
          },
        ],
      }) as never,
  });
  assert.deepEqual(getState(), [
    {
      id: 'anki.autoUpdateNewCards',
      label: 'X',
      scope: 'ankiConnect',
      valueType: 'boolean',
      value: true,
      allowedValues: [true, false],
      requiresRestart: false,
    },
  ]);
});

test('restore previous secondary subtitle visibility no-ops without connected mpv client', () => {
  let restored = false;
  const restore = createRestorePreviousSecondarySubVisibilityHandler({
    getMpvClient: () => ({
      connected: false,
      restorePreviousSecondarySubVisibility: () => (restored = true),
    }),
  });
  restore();
  assert.equal(restored, false);
});

test('restore previous secondary subtitle visibility calls runtime when connected', () => {
  let restored = false;
  const restore = createRestorePreviousSecondarySubVisibilityHandler({
    getMpvClient: () => ({
      connected: true,
      restorePreviousSecondarySubVisibility: () => (restored = true),
    }),
  });
  restore();
  assert.equal(restored, true);
});
