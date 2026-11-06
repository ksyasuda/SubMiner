import assert from 'node:assert/strict';
import test from 'node:test';

import { CORE_DEFAULT_CONFIG } from '../../config/definitions/defaults-core';
import type { ResolvedControllerConfig } from '../../types';
import { findControllerBindingIssues } from './controller-binding-issues.js';
import { CONTROLLER_BINDING_DEFINITIONS } from './controller-config-form.js';

type Bindings = ResolvedControllerConfig['bindings'];
const DEFAULTS = CORE_DEFAULT_CONFIG.controller.bindings as Bindings;

const issuesWith = (changes: Partial<Bindings>) =>
  findControllerBindingIssues(CONTROLLER_BINDING_DEFINITIONS, { ...DEFAULTS, ...changes });

test('the default bindings share no input', () => {
  assert.deepEqual(issuesWith({}), {});
});

test('two actions on the same button both warn', () => {
  // Next Audio is on R1 (button 5) by default.
  assert.deepEqual(issuesWith({ nextSubtitle: { kind: 'button', buttonIndex: 5 } }), {
    nextSubtitle: 'Same as Next Audio',
    nextAudio: 'Same as Next Subtitle',
  });
});

test('a button on the D-pad clashes with a D-pad row that is still on', () => {
  assert.deepEqual(issuesWith({ nextSubtitle: { kind: 'button', buttonIndex: 15 } }), {
    nextSubtitle: 'Same as Token Move (D-pad)',
    'leftStickHorizontal:dpad': 'Same as Next Subtitle',
  });
  // Clearing the D-pad row (as the docs advise) removes the warning.
  assert.deepEqual(
    issuesWith({
      nextSubtitle: { kind: 'button', buttonIndex: 15 },
      leftStickHorizontal: { kind: 'axis', axisIndex: 0, dpadFallback: 'none' },
    }),
    {},
  );
});

const L1_DPAD_RIGHT = {
  kind: 'button',
  buttonIndex: 15,
  hold: { kind: 'button', buttonIndex: 4 },
} as const;

test('a combo does not clash with its hold or press used elsewhere on their own', () => {
  // L1 is Play Current Audio by default; the D-pad is still on Token Move.
  assert.deepEqual(issuesWith({ nextAudio: L1_DPAD_RIGHT }), {});
});

test('two actions on the same combo both warn', () => {
  assert.deepEqual(issuesWith({ nextAudio: L1_DPAD_RIGHT, nextEntry: L1_DPAD_RIGHT }), {
    nextAudio: 'Same as Next Entry',
    nextEntry: 'Same as Next Audio',
  });
});

test('a combo holding and pressing the same button warns', () => {
  assert.deepEqual(
    issuesWith({
      nextEntry: { kind: 'button', buttonIndex: 7, hold: { kind: 'button', buttonIndex: 7 } },
    }),
    { nextEntry: 'Hold and press are the same button' },
  );
});

test('the secondary subtitle hold warns on a hold button, since it cannot wait for release', () => {
  assert.deepEqual(
    issuesWith({
      playCurrentAudio: { kind: 'none' },
      nextAudio: L1_DPAD_RIGHT,
      holdSecondarySub: { kind: 'button', buttonIndex: 4 },
    }),
    { holdSecondarySub: 'Also the hold button of Next Audio' },
  );
});

test('a D-pad reported as a hat axis also counts as the D-pad', () => {
  assert.deepEqual(
    issuesWith({ previousEntry: { kind: 'axis', axisIndex: 7, direction: 'negative' } }),
    {
      previousEntry: 'Same as Popup Scroll (D-pad)',
      'leftStickVertical:dpad': 'Same as Previous Entry',
    },
  );
});

test('a stick direction clashes with a stick action on that axis', () => {
  const bindings = {
    rightStickVertical: { kind: 'axis', axisIndex: 3, dpadFallback: 'none' },
    holdSecondarySub: { kind: 'axis', axisIndex: 3, direction: 'negative' },
  } as const;
  assert.deepEqual(issuesWith(bindings), {
    holdSecondarySub: 'Same as Popup Jump (Stick)',
    rightStickVertical: 'Same as Show Secondary Subtitle (Hold)',
  });
});

test('opposite directions of one stick are different inputs', () => {
  assert.deepEqual(
    issuesWith({
      previousSubtitle: { kind: 'axis', axisIndex: 2, direction: 'negative' },
      nextSubtitle: { kind: 'axis', axisIndex: 2, direction: 'positive' },
    }),
    {},
  );
});

test('two stick actions on the same stick both warn', () => {
  assert.deepEqual(
    issuesWith({ leftStickVertical: { kind: 'axis', axisIndex: 0, dpadFallback: 'vertical' } }),
    {
      leftStickHorizontal: 'Same as Popup Scroll (Stick)',
      leftStickVertical: 'Same as Token Move (Stick)',
    },
  );
});

test('a row shared by several actions lists them all', () => {
  assert.equal(
    issuesWith({
      mineCard: { kind: 'button', buttonIndex: 5 },
      nextSubtitle: { kind: 'button', buttonIndex: 5 },
    }).nextAudio,
    'Same as Mine Card, Next Subtitle',
  );
});
