import assert from 'node:assert/strict';
import test from 'node:test';
import { createMpvInputForwarding } from './mpv-input-forwarding';
import type { MpvInputBindingsSnapshot } from '../../types/session-bindings';

function keyEvent(
  overrides: Partial<Parameters<ReturnType<typeof createMpvInputForwarding>['keydown']>[0]> = {},
) {
  return {
    key: 'r',
    code: 'KeyR',
    ctrlKey: false,
    altKey: false,
    shiftKey: false,
    metaKey: false,
    repeat: false,
    defaultPrevented: false,
    isComposing: false,
    getModifierState: () => false,
    preventDefault: () => {},
    ...overrides,
  };
}

test('forwarded keys use mpv repeat and retain the pressed key through modifier changes', async () => {
  const commands: (string | number)[][] = [];
  const forwarding = createMpvInputForwarding({
    load: async () => ({ keys: ['r', 'ctrl+A'], blockedKeys: [] }),
    send: (command) => commands.push(command),
  });
  await forwarding.refresh();
  assert.equal(forwarding.keydown(keyEvent()), true);
  assert.equal(forwarding.keydown(keyEvent({ repeat: true })), true);
  forwarding.keyup(keyEvent());
  forwarding.keydown(keyEvent({ key: 'A', code: 'KeyA', ctrlKey: true, shiftKey: true }));
  forwarding.keyup(keyEvent({ key: 'a', code: 'KeyA' }));
  assert.deepEqual(commands, [
    ['keydown', 'r'],
    ['keyup', 'r'],
    ['keydown', 'ctrl+A'],
    ['keyup', 'ctrl+A'],
  ]);
});

test('configured and disabled keys, handled input, and unknown keys are not forwarded', async () => {
  const commands: (string | number)[][] = [];
  const forwarding = createMpvInputForwarding({
    load: async () => ({ keys: ['r', 't', '1'], blockedKeys: [{ code: 'KeyR', modifiers: [] }] }),
    send: (command) => commands.push(command),
  });
  await forwarding.refresh();
  assert.equal(forwarding.keydown(keyEvent()), false);
  assert.equal(
    forwarding.keydown(keyEvent({ key: 't', code: 'KeyT', defaultPrevented: true })),
    false,
  );
  assert.equal(forwarding.keydown(keyEvent({ key: 'z', code: 'KeyZ' })), false);
  assert.equal(forwarding.keydown(keyEvent({ key: '1', code: 'Numpad1' })), false);
  assert.deepEqual(commands, []);
});

test('refresh discards stale responses and coalesces concurrent requests', async () => {
  let resolveFirst: (snapshot: MpvInputBindingsSnapshot) => void = () => {};
  let requests = 0;
  const forwarding = createMpvInputForwarding({
    load: () => {
      requests += 1;
      if (requests === 1)
        return new Promise((resolve) => {
          resolveFirst = resolve;
        });
      return Promise.resolve({ keys: ['t'], blockedKeys: [] });
    },
    send: () => {},
  });
  const first = forwarding.refresh();
  const second = forwarding.refresh();
  forwarding.refresh();
  assert.equal(requests, 1);
  resolveFirst({ keys: ['r'], blockedKeys: [] });
  await Promise.all([first, second]);
  assert.equal(requests, 2);
  assert.equal(forwarding.keydown(keyEvent()), false);
  assert.equal(forwarding.keydown(keyEvent({ key: 't', code: 'KeyT' })), true);
});

test('focus loss releases held keys and failed refresh clears stale bindings', async () => {
  let fail = false;
  const commands: (string | number)[][] = [];
  const forwarding = createMpvInputForwarding({
    load: async () => {
      if (fail) throw new Error('disconnected');
      return { keys: ['r'], blockedKeys: [] };
    },
    send: (command) => commands.push(command),
  });
  await forwarding.refresh();
  forwarding.keydown(keyEvent());
  forwarding.releaseAll();
  forwarding.keyup(keyEvent());
  assert.deepEqual(commands, [
    ['keydown', 'r'],
    ['keyup', 'r'],
  ]);
  fail = true;
  await forwarding.refresh();
  assert.equal(forwarding.keydown(keyEvent()), false);
  forwarding.dispose();
  fail = false;
  await forwarding.refresh();
  assert.equal(forwarding.keydown(keyEvent()), false);
});
