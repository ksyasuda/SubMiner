import assert from 'node:assert/strict';
import test from 'node:test';
import type { AnimeBrowserIpcSender } from './anime-browser-ipc-handlers';
import { createAnimeBrowserSessionRegistry } from './anime-browser-sessions';

test('moving a live sender to a new anime browser session releases its stale session', () => {
  const destroyedListeners: Array<() => void> = [];
  const sender: AnimeBrowserIpcSender = {
    send: () => {},
    isDestroyed: () => false,
    once: (_event, listener) => destroyedListeners.push(listener),
  };
  const released: string[] = [];
  const sessions = createAnimeBrowserSessionRegistry((sessionId) => released.push(sessionId));

  sessions.register('old', sender);
  sessions.register('old', sender);
  assert.equal(destroyedListeners.length, 1, 'duplicate registration keeps its existing handler');

  sessions.register('new', sender);
  assert.deepEqual(released, ['old']);
  assert.equal(sessions.get('old'), undefined);
  assert.equal(sessions.get('new'), sender);

  for (const listener of destroyedListeners) listener();
  assert.deepEqual(released, ['old', 'new']);
  assert.equal(sessions.get('new'), undefined);
});
