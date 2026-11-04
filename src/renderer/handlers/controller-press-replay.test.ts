import assert from 'node:assert/strict';
import test from 'node:test';
import { createControllerPressReplay } from './controller-press-replay';
import type { GamepadLike } from './gamepad-controller';

function pad(pressed: number[] = [], mapping = 'standard'): GamepadLike {
  return {
    id: 'pad',
    index: 0,
    connected: true,
    mapping,
    axes: [0, 0, 0, 0],
    buttons: Array.from({ length: 17 }, (_, index) => ({
      value: pressed.includes(index) ? 1 : 0,
      pressed: pressed.includes(index),
    })),
  };
}

const pressedIndices = (gamepad: GamepadLike | null | undefined): number[] =>
  gamepad?.buttons.flatMap((button, index) => (button.pressed ? [index] : [])) ?? [];

test('a queued press reads as pressed for exactly one poll', () => {
  const replay = createControllerPressReplay();
  replay.queue(0, 1000);

  assert.deepEqual(pressedIndices(replay.apply([pad()], 1010)[0]), [0]);
  assert.deepEqual(pressedIndices(replay.apply([pad()], 1026)[0]), []);
});

test('a press the Gamepad API already reported is not replayed', () => {
  const replay = createControllerPressReplay();
  replay.apply([pad([1])], 990);
  replay.queue(1, 1000);

  assert.deepEqual(pressedIndices(replay.apply([pad()], 1010)[0]), []);
});

test('a press seen after the replay was queued is not replayed again', () => {
  const replay = createControllerPressReplay();
  replay.queue(2, 1000);

  assert.deepEqual(pressedIndices(replay.apply([pad([2])], 1010)[0]), [2]);
  assert.deepEqual(pressedIndices(replay.apply([pad()], 1026)[0]), []);
});

test('an older press of the same button does not block the replay', () => {
  const replay = createControllerPressReplay();
  replay.apply([pad([3])], 0);
  replay.queue(3, 5000);

  assert.deepEqual(pressedIndices(replay.apply([pad()], 5010)[0]), [3]);
});

test('nothing is replayed while several pads are connected', () => {
  const replay = createControllerPressReplay();
  replay.queue(0, 1000);

  const pads = replay.apply([pad(), { ...pad(), index: 1 }], 1010);
  assert.deepEqual(pads.map(pressedIndices), [[], []]);
  assert.deepEqual(pressedIndices(replay.apply([pad()], 1026)[0]), []);
});

test('pads without the standard mapping are left alone', () => {
  const replay = createControllerPressReplay();
  replay.queue(0, 1000);

  assert.deepEqual(pressedIndices(replay.apply([pad([], '')], 1010)[0]), []);
});
