import assert from 'node:assert/strict';
import test from 'node:test';

import { CORE_DEFAULT_CONFIG } from '../config/definitions/defaults-core';
import {
  controllerReviewInputSourceKey,
  formatControllerReviewBinding,
  formatControllerReviewInput,
  getControllerFaceButtonName,
  isControllerFaceButtonLabels,
  parseControllerReviewBinding,
  resolveControllerReviewInput,
  type ControllerReviewInputContext,
} from './controller-review';
import { parseControllerConfigUpdate } from './ipc/validators';

const buttonIndices = CORE_DEFAULT_CONFIG.controller.buttonIndices;
const standardXbox: ControllerReviewInputContext = {
  mapping: 'standard',
  faceButtonLabels: 'xbox',
  buttonIndices,
};

test('review bindings parse none, single presses and combos; anything else is rejected', () => {
  assert.deepEqual(parseControllerReviewBinding({ kind: 'none' }), { kind: 'none' });
  assert.deepEqual(parseControllerReviewBinding({ press: { kind: 'face', letter: 'a' } }), {
    kind: 'combo',
    press: { kind: 'face', letter: 'a' },
    hold: null,
  });
  assert.deepEqual(
    parseControllerReviewBinding({
      kind: 'combo',
      press: { kind: 'dpad', direction: 'up' },
      hold: { kind: 'button', buttonIndex: 4 },
    }),
    {
      kind: 'combo',
      press: { kind: 'dpad', direction: 'up' },
      hold: { kind: 'button', buttonIndex: 4 },
    },
  );
  for (const bad of [
    null,
    {},
    { kind: 'button', buttonIndex: 1 },
    { press: { kind: 'face', letter: 'z' } },
    { press: { kind: 'button', buttonIndex: -1 } },
    { press: { kind: 'axis', axisIndex: 2 } },
    { press: { kind: 'dpad', direction: 'left' }, hold: { kind: 'dpad' } },
  ]) {
    assert.equal(parseControllerReviewBinding(bad), null, JSON.stringify(bad));
  }
});

test('the same physical input resolves to the same key across layouts', () => {
  const key = (input: Parameters<typeof resolveControllerReviewInput>[0], context = standardXbox) =>
    controllerReviewInputSourceKey(resolveControllerReviewInput(input, context));

  // Letters follow faceButtonLabels.
  assert.equal(key({ kind: 'face', letter: 'a' }), 'button:0');
  assert.equal(
    key({ kind: 'face', letter: 'a' }, { ...standardXbox, faceButtonLabels: 'nintendo' }),
    'button:1',
  );
  // A learned standard D-pad button and the default D-pad entry are one input.
  assert.equal(key({ kind: 'button', buttonIndex: 14 }), 'dpad:left');
  assert.equal(key({ kind: 'dpad', direction: 'left' }), 'dpad:left');
  // Non-standard pads report the D-pad on axes 6 and 7.
  const raw = { ...standardXbox, mapping: '' };
  assert.equal(key({ kind: 'axis', axisIndex: 7, direction: 'positive' }, raw), 'dpad:down');
  assert.equal(key({ kind: 'button', buttonIndex: 14 }, raw), 'button:14');
  assert.equal(key({ kind: 'axis', axisIndex: 2, direction: 'positive' }, raw), 'axis:2:positive');
});

test('review inputs are named as printed and by layout', () => {
  const name = (input: Parameters<typeof formatControllerReviewInput>[0], context = standardXbox) =>
    formatControllerReviewInput(input, context, { dpadPrefix: true });
  const nintendo = { ...standardXbox, faceButtonLabels: 'nintendo' as const };
  const raw = { ...standardXbox, mapping: '' };

  assert.equal(name({ kind: 'face', letter: 'x' }, nintendo), 'X');
  assert.equal(name({ kind: 'button', buttonIndex: 0 }, nintendo), 'B');
  assert.equal(name({ kind: 'button', buttonIndex: 6 }), 'L2');
  assert.equal(name({ kind: 'button', buttonIndex: 6 }, raw), 'Back');
  assert.equal(name({ kind: 'button', buttonIndex: 10 }), 'L3');
  assert.equal(name({ kind: 'button', buttonIndex: 15 }), 'D-pad ▶');
  assert.equal(name({ kind: 'axis', axisIndex: 1, direction: 'negative' }), 'LS ▲');
  assert.equal(name({ kind: 'axis', axisIndex: 5, direction: 'positive' }, raw), 'R2');
  assert.equal(name({ kind: 'axis', axisIndex: 9, direction: 'positive' }), 'Axis 9 +');
  assert.equal(
    formatControllerReviewBinding(
      {
        kind: 'combo',
        press: { kind: 'dpad', direction: 'up' },
        hold: { kind: 'button', buttonIndex: 5 },
      },
      standardXbox,
    ),
    'R1 + ▲',
  );
});

test('controller config updates from Alt+C accept review bindings and reject bad ones', () => {
  // The validator builds prototype-less objects; compare plain copies.
  const plain = (value: unknown) => JSON.parse(JSON.stringify(value));
  assert.deepEqual(
    plain(
      parseControllerConfigUpdate({
        profiles: {
          'pad-1': {
            reviewBindings: {
              confirm: { kind: 'combo', press: { kind: 'button', buttonIndex: 1 } },
            },
          },
        },
      }),
    ),
    {
      profiles: {
        'pad-1': {
          reviewBindings: {
            confirm: { kind: 'combo', press: { kind: 'button', buttonIndex: 1 }, hold: null },
          },
        },
      },
    },
  );
  assert.equal(parseControllerConfigUpdate({ reviewBindings: { nope: { kind: 'none' } } }), null);
  assert.equal(
    parseControllerConfigUpdate({ reviewBindings: { confirm: { press: { kind: 'face' } } } }),
    null,
  );
});

test('Alt+C updates accept field grouping bindings, checked against that dialog', () => {
  const plain = (value: unknown) => JSON.parse(JSON.stringify(value));
  assert.deepEqual(
    plain(
      parseControllerConfigUpdate({
        profiles: { 'pad-1': { fieldGroupingBindings: { back: { kind: 'none' } } } },
      }),
    ),
    { profiles: { 'pad-1': { fieldGroupingBindings: { back: { kind: 'none' } } } } },
  );
  // Each dialog only accepts its own actions.
  assert.equal(
    parseControllerConfigUpdate({ fieldGroupingBindings: { confirm: { kind: 'none' } } }),
    null,
  );
  assert.equal(parseControllerConfigUpdate({ reviewBindings: { back: { kind: 'none' } } }), null);
});

test('PlayStation pads print symbols at the Xbox positions', () => {
  const playstation = { ...standardXbox, faceButtonLabels: 'playstation' as const };
  const name = (input: Parameters<typeof formatControllerReviewInput>[0]) =>
    formatControllerReviewInput(input, playstation);

  // Same physical buttons as Xbox: A is the bottom button, shown as ╳.
  assert.equal(
    controllerReviewInputSourceKey(
      resolveControllerReviewInput({ kind: 'face', letter: 'a' }, playstation),
    ),
    'button:0',
  );
  assert.deepEqual(
    (['a', 'b', 'x', 'y'] as const).map((letter) => name({ kind: 'face', letter })),
    ['╳', '○', '□', '△'],
  );
  assert.equal(name({ kind: 'button', buttonIndex: 3 }), '△');
  assert.equal(
    formatControllerReviewBinding(
      {
        kind: 'combo',
        press: { kind: 'face', letter: 'a' },
        hold: { kind: 'button', buttonIndex: 4 },
      },
      playstation,
    ),
    'L1 + ╳',
  );
  assert.equal(getControllerFaceButtonName(1, 'nintendo'), 'A');
  assert.equal(getControllerFaceButtonName(4, 'playstation'), null);
  assert.equal(isControllerFaceButtonLabels('playstation'), true);
  assert.equal(isControllerFaceButtonLabels('sega'), false);
});

test('Alt+C updates accept a button style globally and per controller', () => {
  const plain = (value: unknown) => JSON.parse(JSON.stringify(value));
  assert.deepEqual(
    plain(
      parseControllerConfigUpdate({
        faceButtonLabels: 'xbox',
        profiles: { 'pad-1': { faceButtonLabels: 'playstation' } },
      }),
    ),
    { faceButtonLabels: 'xbox', profiles: { 'pad-1': { faceButtonLabels: 'playstation' } } },
  );
  assert.equal(parseControllerConfigUpdate({ faceButtonLabels: 'sega' }), null);
  assert.equal(
    parseControllerConfigUpdate({ profiles: { 'pad-1': { faceButtonLabels: 1 } } }),
    null,
  );
});
