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

test('a replayed native pad keeps its fields', () => {
  // Like the browser's Gamepad, these fields are prototype getters that object spread skips.
  class NativePad implements GamepadLike {
    get id() {
      return 'pad';
    }
    get index() {
      return 2;
    }
    get connected() {
      return true;
    }
    get mapping() {
      return 'standard';
    }
    get timestamp() {
      return 900;
    }
    get axes() {
      return [0.5, 0];
    }
    get buttons() {
      return pad().buttons;
    }
  }
  const replay = createControllerPressReplay();
  replay.queue(0, 1000);

  const replayed = replay.apply([new NativePad()], 1010)[0];
  assert.deepEqual(pressedIndices(replayed), [0]);
  assert.equal(replayed?.id, 'pad');
  assert.equal(replayed?.index, 2);
  assert.equal(replayed?.connected, true);
  assert.equal(replayed?.mapping, 'standard');
  assert.equal(replayed?.timestamp, 900);
  assert.deepEqual(replayed?.axes, [0.5, 0]);
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
