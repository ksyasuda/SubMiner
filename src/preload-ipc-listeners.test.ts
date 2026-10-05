import assert from 'node:assert/strict';
import test from 'node:test';
import { createIpcListenerFactories } from './preload-ipc-listeners';

function createFakeIpc() {
  const handlers = new Map<string, (payload: unknown) => void>();
  return {
    factories: createIpcListenerFactories((channel, handler) => handlers.set(channel, handler)),
    emit: (channel: string, payload?: unknown) => handlers.get(channel)?.(payload),
  };
}

test('queued listener replays every event fired before registration, then delivers live', () => {
  const ipc = createFakeIpc();
  const onOpen = ipc.factories.createQueuedIpcListener('open');
  ipc.emit('open');
  ipc.emit('open');

  let calls = 0;
  onOpen(() => (calls += 1));
  assert.equal(calls, 2);

  ipc.emit('open');
  assert.equal(calls, 3);
});

test('queued payload listener replays early payloads in order through normalize', () => {
  const ipc = createFakeIpc();
  const onPayload = ipc.factories.createQueuedIpcListenerWithPayload('payload', (raw) =>
    String(raw).toUpperCase(),
  );
  ipc.emit('payload', 'a');
  ipc.emit('payload', 'b');

  const received: string[] = [];
  onPayload((value) => received.push(value));
  ipc.emit('payload', 'c');
  assert.deepEqual(received, ['A', 'B', 'C']);
});

test('latest-value listener replays only the newest early payload, once', () => {
  const ipc = createFakeIpc();
  const onSubtitle = ipc.factories.createLatestValueIpcListenerWithPayload(
    'subtitle',
    (raw) => raw,
  );
  ipc.emit('subtitle', 'stale');
  ipc.emit('subtitle', 'current');

  const first: unknown[] = [];
  onSubtitle((value) => first.push(value));
  assert.deepEqual(first, ['current']);

  const second: unknown[] = [];
  onSubtitle((value) => second.push(value));
  assert.deepEqual(second, [], 'a buffered value is consumed by the first listener');

  ipc.emit('subtitle', 'live');
  assert.deepEqual(first, ['current', 'live']);
  assert.deepEqual(second, ['live']);
});
