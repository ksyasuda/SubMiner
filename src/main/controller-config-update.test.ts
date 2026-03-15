import assert from 'node:assert/strict';
import test from 'node:test';

import { applyControllerConfigUpdate } from './controller-config-update.js';

test('applyControllerConfigUpdate replaces binding descriptors instead of deep-merging them', () => {
  const next = applyControllerConfigUpdate(
    {
      preferredGamepadId: 'pad-1',
      bindings: {
        toggleLookup: { kind: 'axis', axisIndex: 4, direction: 'positive' },
        closeLookup: { kind: 'button', buttonIndex: 1 },
      },
    },
    {
      bindings: {
        toggleLookup: { kind: 'button', buttonIndex: 11 },
      },
    },
  );

  assert.deepEqual(next.bindings?.toggleLookup, { kind: 'button', buttonIndex: 11 });
  assert.deepEqual(next.bindings?.closeLookup, { kind: 'button', buttonIndex: 1 });
});

test('applyControllerConfigUpdate merges buttonIndices while replacing only updated binding leaves', () => {
  const next = applyControllerConfigUpdate(
    {
      buttonIndices: {
        select: 6,
        buttonSouth: 0,
      },
      bindings: {
        toggleLookup: { kind: 'button', buttonIndex: 0 },
        closeLookup: { kind: 'button', buttonIndex: 1 },
      },
    },
    {
      buttonIndices: {
        buttonSouth: 9,
      },
      bindings: {
        closeLookup: { kind: 'none' },
      },
    },
  );

  assert.deepEqual(next.buttonIndices, {
    select: 6,
    buttonSouth: 9,
  });
  assert.deepEqual(next.bindings?.toggleLookup, { kind: 'button', buttonIndex: 0 });
  assert.deepEqual(next.bindings?.closeLookup, { kind: 'none' });
});
