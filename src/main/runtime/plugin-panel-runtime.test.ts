import assert from 'node:assert/strict';
import test from 'node:test';

import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import type { PluginPanelSnapshot } from '../../types/runtime';
import { createPluginPanelRuntime } from './plugin-panel-runtime';

function createHarness(openResult: Promise<boolean>) {
  const sent: Array<{ channel: string; payload: unknown }> = [];
  const opened: PluginPanelSnapshot[] = [];
  let openFailures = 0;
  const runtime = createPluginPanelRuntime({
    openModal: async (getSnapshot) => {
      opened.push(getSnapshot());
      return await openResult;
    },
    broadcast: (channel, payload) => sent.push({ channel, payload }),
    onOpenFailed: () => {
      openFailures += 1;
    },
  });
  return { runtime, sent, opened, failures: () => openFailures };
}

test('re-sends content updated before the renderer acknowledged the open', async () => {
  let acknowledge!: (opened: boolean) => void;
  const h = createHarness(new Promise((resolve) => (acknowledge = resolve)));

  const panel = h.runtime.open({ title: 'Explain', markdown: '_Asking..._' });
  panel.update({ markdown: 'done' });
  acknowledge(true);
  await new Promise((resolve) => setImmediate(resolve));

  assert.deepEqual(h.opened, [{ panelId: 1, title: 'Explain', markdown: '_Asking..._' }]);
  assert.deepEqual(h.sent.at(-1), {
    channel: IPC_CHANNELS.event.pluginPanelUpdate,
    payload: { panelId: 1, title: 'Explain', markdown: 'done' },
  });
});

test('closing from the overlay aborts the handle and ignores later updates', async () => {
  const h = createHarness(Promise.resolve(true));
  const panel = h.runtime.open({ title: 'Explain', markdown: '' });
  await new Promise((resolve) => setImmediate(resolve));
  const sentBefore = h.sent.length;

  assert.equal(h.runtime.handleModalClosed(1), true);
  panel.update({ markdown: 'too late' });

  assert.equal(panel.signal.aborted, true);
  assert.equal(panel.closed, true);
  assert.equal(h.sent.length, sentBefore);
});

test('a late close acknowledgement for a replaced panel leaves the new panel open', () => {
  const h = createHarness(Promise.resolve(true));
  const first = h.runtime.open({ title: 'One', markdown: '' });
  first.close();
  const second = h.runtime.open({ title: 'Two', markdown: '' });

  assert.equal(h.runtime.handleModalClosed(1), false, 'skip teardown of the newer panel');
  assert.equal(second.signal.aborted, false);
});

test('a close notice after the plugin closed its panel still tears the modal down', () => {
  const h = createHarness(Promise.resolve(true));
  h.runtime.open({ title: 'One', markdown: '' }).close();

  assert.equal(h.runtime.handleModalClosed(1), true);
});

test('opening a new panel aborts the previous one', () => {
  const h = createHarness(Promise.resolve(true));
  const first = h.runtime.open({ title: 'One', markdown: '' });
  const second = h.runtime.open({ title: 'Two', markdown: '' });

  assert.equal(first.signal.aborted, true);
  assert.equal(second.signal.aborted, false);
});

test('plugin close tells the renderer which panel to close', () => {
  const h = createHarness(Promise.resolve(true));
  const panel = h.runtime.open({ title: 'One', markdown: '' });

  panel.close();

  assert.equal(panel.signal.aborted, true);
  assert.deepEqual(h.sent.at(-1), { channel: IPC_CHANNELS.event.pluginPanelClose, payload: 1 });
});

test('a failed open aborts the handle and reports it once', async () => {
  const h = createHarness(Promise.resolve(false));
  const panel = h.runtime.open({ title: 'One', markdown: '' });
  await new Promise((resolve) => setImmediate(resolve));

  assert.equal(panel.signal.aborted, true);
  assert.equal(h.failures(), 1);
});
