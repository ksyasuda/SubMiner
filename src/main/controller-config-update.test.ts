import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';

import { applyControllerConfigUpdate } from './controller-config-update.js';

test('SM-012 controller config update path does not use JSON serialize-clone helpers', () => {
  const source = fs.readFileSync(
    path.join(process.cwd(), 'src/main/controller-config-update.ts'),
    'utf-8',
  );
  assert.equal(source.includes('JSON.parse(JSON.stringify('), false);
});

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

test('applyControllerConfigUpdate detaches updated binding values from the patch object', () => {
  const update = {
    bindings: {
      toggleLookup: { kind: 'button' as const, buttonIndex: 7 },
    },
  };

  const next = applyControllerConfigUpdate(undefined, update);
  update.bindings.toggleLookup.buttonIndex = 99;

  assert.deepEqual(next.bindings?.toggleLookup, { kind: 'button', buttonIndex: 7 });
});

test('applyControllerConfigUpdate merges per-controller profile binding leaves', () => {
  const next = applyControllerConfigUpdate(
    {
      profiles: {
        'pad-1': {
          label: 'Pad 1',
          bindings: {
            toggleLookup: { kind: 'button', buttonIndex: 0 },
            closeLookup: { kind: 'button', buttonIndex: 1 },
          },
        },
      },
    },
    {
      profiles: {
        'pad-1': {
          bindings: {
            toggleLookup: { kind: 'button', buttonIndex: 11 },
          },
        },
        'pad-2': {
          label: 'Pad 2',
          bindings: {
            mineCard: { kind: 'button', buttonIndex: 8 },
          },
        },
      },
    },
  );

  assert.deepEqual(next.profiles?.['pad-1']?.bindings?.toggleLookup, {
    kind: 'button',
    buttonIndex: 11,
  });
  assert.deepEqual(next.profiles?.['pad-1']?.bindings?.closeLookup, {
    kind: 'button',
    buttonIndex: 1,
  });
  assert.deepEqual(next.profiles?.['pad-2']?.bindings?.mineCard, {
    kind: 'button',
    buttonIndex: 8,
  });
});
