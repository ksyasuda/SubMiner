import test from 'node:test';
import assert from 'node:assert/strict';
import { shouldUseLearnedMouseBinding } from './settings-keybinding-controls';

test('mouse key learning ignores primary left clicks in DOM-code mode', () => {
  const learnButton = {};
  const outsideTarget = {};

  assert.equal(
    shouldUseLearnedMouseBinding(
      'MBTN_LEFT',
      'dom-code',
      { button: 0, target: outsideTarget } as MouseEvent,
      learnButton as HTMLButtonElement,
    ),
    false,
  );
  assert.equal(
    shouldUseLearnedMouseBinding(
      'MBTN_BACK',
      'dom-code',
      { button: 3, target: outsideTarget } as MouseEvent,
      learnButton as HTMLButtonElement,
    ),
    true,
  );
});

test('mouse key learning still ignores primary learn-button activation', () => {
  const learnButton = {};

  assert.equal(
    shouldUseLearnedMouseBinding(
      'MBTN_LEFT',
      'dom-code',
      { button: 0, target: learnButton } as MouseEvent,
      learnButton as HTMLButtonElement,
    ),
    false,
  );
  assert.equal(
    shouldUseLearnedMouseBinding(
      'MBTN_BACK',
      'dom-code',
      { button: 3, target: learnButton } as MouseEvent,
      learnButton as HTMLButtonElement,
    ),
    true,
  );
});
