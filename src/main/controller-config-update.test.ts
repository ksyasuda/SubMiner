import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';

import { ConfigService } from '../config/service';
import { parseControllerConfigUpdate } from '../shared/ipc/validators';
import {
  applyControllerConfigUpdate,
  saveControllerConfigUpdate,
} from './controller-config-update.js';

const L1 = { kind: 'button', buttonIndex: 4 } as const;
const L2 = { kind: 'button', buttonIndex: 6 } as const;
const DPAD_LEFT = { kind: 'button', buttonIndex: 14 } as const;
const DPAD_RIGHT = { kind: 'button', buttonIndex: 15 } as const;

/** One Alt+C save for a profile, through the same validator and merge as the real IPC handler. */
function createSaveHarness() {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-controller-save-'));
  const service = new ConfigService(dir);
  return {
    service,
    save: (action: string, binding: unknown) => {
      const update = parseControllerConfigUpdate({
        profiles: { 'pad-1': { bindings: { [action]: binding } } },
      });
      assert.ok(update, 'the validator accepts the update');
      saveControllerConfigUpdate(service, update);
    },
    savedBinding: (action: 'nextAudio' | 'previousAudio' | 'nextEntry') =>
      service.getConfig().controller.profiles['pad-1']?.bindings[action],
    cleanup: () => fs.rmSync(dir, { recursive: true, force: true }),
  };
}

test('combos sharing a hold button or a press button all survive later saves', () => {
  const harness = createSaveHarness();
  try {
    harness.save('nextAudio', { ...DPAD_LEFT, hold: L1 });
    harness.save('previousAudio', { ...DPAD_RIGHT, hold: L1 });
    harness.save('nextEntry', { ...DPAD_LEFT, hold: L2 });

    assert.deepEqual(harness.savedBinding('nextAudio'), { ...DPAD_LEFT, hold: L1 });
    assert.deepEqual(harness.savedBinding('previousAudio'), { ...DPAD_RIGHT, hold: L1 });
    assert.deepEqual(harness.savedBinding('nextEntry'), { ...DPAD_LEFT, hold: L2 });
  } finally {
    harness.cleanup();
  }
});

test('rebinding a combo to one button drops its hold', () => {
  const harness = createSaveHarness();
  try {
    harness.save('nextAudio', { ...DPAD_LEFT, hold: L1 });
    harness.save('nextAudio', DPAD_RIGHT);

    assert.deepEqual(harness.savedBinding('nextAudio'), DPAD_RIGHT);
    assert.deepEqual(
      harness.service.getRawConfig().controller?.profiles?.['pad-1']?.bindings?.nextAudio,
      DPAD_RIGHT,
    );
  } finally {
    harness.cleanup();
  }
});

test('the validator rejects an invalid hold instead of saving a plain press', () => {
  assert.equal(
    parseControllerConfigUpdate({
      bindings: { nextAudio: { ...DPAD_LEFT, hold: { kind: 'face' } } },
    }),
    null,
  );
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

test('applyControllerConfigUpdate ignores reserved profile ids', () => {
  const reservedProfiles = Object.create(null) as NonNullable<
    Parameters<typeof applyControllerConfigUpdate>[1]['profiles']
  >;
  reservedProfiles.__proto__ = { label: 'polluted' };
  reservedProfiles['constructor'] = { label: 'ctor' };
  reservedProfiles['prototype'] = { label: 'proto' };
  reservedProfiles['pad-1'] = { label: 'Pad 1' };

  const next = applyControllerConfigUpdate(undefined, {
    profiles: reservedProfiles,
  });

  assert.equal(Object.getPrototypeOf(next.profiles), Object.prototype);
  assert.equal(Object.hasOwn(next.profiles ?? {}, '__proto__'), false);
  assert.equal(Object.hasOwn(next.profiles ?? {}, 'constructor'), false);
  assert.equal(Object.hasOwn(next.profiles ?? {}, 'prototype'), false);
  assert.equal(next.profiles?.['pad-1']?.label, 'Pad 1');
});

test('applyControllerConfigUpdate merges review bindings per action inside a profile', () => {
  const next = applyControllerConfigUpdate(
    {
      profiles: {
        'pad-1': {
          bindings: { toggleLookup: { kind: 'button', buttonIndex: 0 } },
          reviewBindings: {
            confirm: { press: { kind: 'button', buttonIndex: 1 } },
            cancel: { kind: 'none' },
          },
        },
      },
    },
    {
      profiles: {
        'pad-1': {
          reviewBindings: {
            cancel: {
              kind: 'combo',
              press: { kind: 'button', buttonIndex: 0 },
              hold: { kind: 'button', buttonIndex: 4 },
            },
          },
        },
      },
    },
  );

  assert.deepEqual(next.profiles?.['pad-1'], {
    bindings: { toggleLookup: { kind: 'button', buttonIndex: 0 } },
    reviewBindings: {
      confirm: { press: { kind: 'button', buttonIndex: 1 } },
      cancel: {
        kind: 'combo',
        press: { kind: 'button', buttonIndex: 0 },
        hold: { kind: 'button', buttonIndex: 4 },
      },
    },
  });
});

test('applyControllerConfigUpdate merges field grouping bindings next to review bindings', () => {
  const next = applyControllerConfigUpdate(
    {
      reviewBindings: { confirm: { kind: 'none' } },
      fieldGroupingBindings: { back: { kind: 'none' } },
    },
    { fieldGroupingBindings: { cancelMerge: { press: { kind: 'face', letter: 'b' } } } },
  );

  assert.deepEqual(next.reviewBindings, { confirm: { kind: 'none' } });
  assert.deepEqual(next.fieldGroupingBindings, {
    back: { kind: 'none' },
    cancelMerge: { press: { kind: 'face', letter: 'b' } },
  });
});
