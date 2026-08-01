import test from 'node:test';
import assert from 'node:assert/strict';
import { watchPlaybackOutcome, type PlaybackEndFileEvent } from './playback-outcome';

type Harness = {
  emitEndFile: (event: PlaybackEndFileEvent) => void;
  listenerCount: () => number;
  setProperty: (name: string, value: unknown) => void;
  failProperty: (name: string) => void;
};

function createHarness(overrides?: { timeoutMs?: number; probeIntervalMs?: number }) {
  const listeners = new Set<(event: PlaybackEndFileEvent) => void>();
  const properties = new Map<string, unknown>();
  const failing = new Set<string>();

  const watch = watchPlaybackOutcome({
    onEndFile: (listener) => {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    readProperty: async (name) => {
      if (failing.has(name)) throw new Error(`Failed to read MPV property '${name}'`);
      return properties.get(name);
    },
    wait: async () => {},
    timeoutMs: overrides?.timeoutMs ?? 1000,
    probeIntervalMs: overrides?.probeIntervalMs ?? 100,
  });

  const harness: Harness = {
    emitEndFile: (event) => {
      for (const listener of listeners) listener(event);
    },
    listenerCount: () => listeners.size,
    setProperty: (name, value) => properties.set(name, value),
    failProperty: (name) => failing.add(name),
  };
  return { watch, harness };
}

test('resolves ok once mpv configures a video output', async () => {
  const { watch, harness } = createHarness();
  harness.setProperty('vo-configured', true);
  const outcome = await watch.wait();
  assert.deepEqual(outcome, { ok: true });
  watch.dispose();
});

test('resolves failure with the mpv error when the file ends in error', async () => {
  const { watch, harness } = createHarness();
  const pending = watch.wait();
  harness.emitEndFile({ reason: 'error', fileError: 'no audio or video data played' });
  const outcome = await pending;
  assert.equal(outcome.ok, false);
  assert.ok(!outcome.ok && outcome.error.includes('no audio or video data played'));
  watch.dispose();
});

test('ignores the end-file fired for the file being replaced', async () => {
  const { watch, harness } = createHarness();
  const pending = watch.wait();
  harness.emitEndFile({ reason: 'stop', fileError: null });
  harness.setProperty('vo-configured', true);
  const outcome = await pending;
  assert.deepEqual(outcome, { ok: true });
  watch.dispose();
});

test('times out with a failure when nothing ever starts', async () => {
  const { watch } = createHarness({ timeoutMs: 300, probeIntervalMs: 100 });
  const outcome = await watch.wait();
  assert.equal(outcome.ok, false);
  assert.ok(!outcome.ok && outcome.error.length > 0);
  watch.dispose();
});

test('keeps polling through property read failures', async () => {
  const { watch, harness } = createHarness();
  harness.failProperty('vo-configured');
  const pending = watch.wait();
  harness.emitEndFile({ reason: 'error', fileError: null });
  const outcome = await pending;
  assert.equal(outcome.ok, false);
  watch.dispose();
});

test('dispose removes the end-file subscription', () => {
  const { watch, harness } = createHarness();
  assert.equal(harness.listenerCount(), 1);
  watch.dispose();
  assert.equal(harness.listenerCount(), 0);
});
