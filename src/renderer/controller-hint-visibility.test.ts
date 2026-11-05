import assert from 'node:assert/strict';
import test from 'node:test';

import { createControllerHintVisibility } from './controller-hint-visibility.js';

function createHarness(answers: Array<string | null | Error>) {
  const pending = [...answers];
  let requests = 0;
  const visibility = createControllerHintVisibility({
    getOverlayControllerId: async () => {
      requests += 1;
      const answer = pending.shift() ?? null;
      if (answer instanceof Error) throw answer;
      return answer;
    },
  });
  const frame = async (now: number, activeGamepadId: string | null = null) => {
    const visible = visibility.update({ enabled: true, dialogOpen: true, activeGamepadId, now });
    await new Promise((resolve) => setImmediate(resolve));
    return visible;
  };
  return { visibility, frame, getRequests: () => requests };
}

test('hints stay hidden while no window can see a controller', async () => {
  const { frame } = createHarness([null, null]);

  assert.equal(await frame(0), false);
  assert.equal(await frame(16), false);
  assert.equal(await frame(2_000), false);
});

test('a controller the dialog window can see shows hints without asking the overlay', async () => {
  const { frame, getRequests } = createHarness([]);

  assert.equal(await frame(0, 'pad-1'), true);
  assert.equal(getRequests(), 0);
});

test('a fresh dialog window shows hints for the controller the overlay reports', async () => {
  const { frame } = createHarness(['pad-1']);

  assert.equal(await frame(0), false);
  assert.equal(await frame(16), true);
});

test('hints hide on the next refresh after the overlay loses the controller', async () => {
  const { frame, getRequests } = createHarness(['pad-1', null]);

  await frame(0);
  assert.equal(await frame(500), true);
  assert.equal(getRequests(), 1);

  await frame(1_000);
  assert.equal(getRequests(), 2);
  assert.equal(await frame(1_016), false);
});

test('a failed request keeps the last answer and retries on the next refresh', async () => {
  const { frame, getRequests } = createHarness(['pad-1', new Error('ipc closed'), null]);

  await frame(0);
  await frame(1_000);
  assert.equal(await frame(1_016), true);

  await frame(2_000);
  assert.equal(getRequests(), 3);
  assert.equal(await frame(2_016), false);
});

test('nothing is requested while controller support is off or no dialog is open', () => {
  const { visibility, getRequests } = createHarness([]);

  assert.equal(
    visibility.update({ enabled: false, dialogOpen: true, activeGamepadId: 'pad-1', now: 0 }),
    false,
  );
  visibility.update({ enabled: true, dialogOpen: false, activeGamepadId: null, now: 0 });
  assert.equal(getRequests(), 0);
});
