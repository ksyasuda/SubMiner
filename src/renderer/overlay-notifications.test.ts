import assert from 'node:assert/strict';
import test from 'node:test';

import {
  createOverlayNotificationStore,
  handleOverlayNotificationEvent,
  overlayNotificationPositionClass,
} from './overlay-notifications';

test('overlay notification store caps transient notifications and keeps pinned jobs visible', () => {
  const store = createOverlayNotificationStore({ maxVisible: 3 });

  store.upsert({
    id: 'character-dictionary-auto-sync',
    title: 'Character dictionary',
    body: 'Generating character dictionary',
    persistent: true,
  });
  store.upsert({ id: 'one', title: 'One', body: 'First' });
  store.upsert({ id: 'two', title: 'Two', body: 'Second' });
  store.upsert({ id: 'three', title: 'Three', body: 'Third' });

  assert.deepEqual(
    store.visible().map((entry) => entry.id),
    ['character-dictionary-auto-sync', 'two', 'three'],
  );

  store.upsert({
    id: 'character-dictionary-auto-sync',
    title: 'Character dictionary',
    body: 'Ready',
    persistent: false,
  });

  assert.deepEqual(
    store.visible().map((entry) => `${entry.id}:${entry.body}`),
    ['two:Second', 'three:Third', 'character-dictionary-auto-sync:Ready'],
  );
});

test('overlay notification positions map to stack alignment classes', () => {
  assert.equal(overlayNotificationPositionClass(undefined), 'position-top-right');
  assert.equal(overlayNotificationPositionClass('top-left'), 'position-top-left');
  assert.equal(overlayNotificationPositionClass('top'), 'position-top');
  assert.equal(overlayNotificationPositionClass('top-right'), 'position-top-right');
});

test('overlay notification event handler dismisses notifications by id', () => {
  const calls: string[] = [];

  handleOverlayNotificationEvent(
    {
      show: (payload) => {
        calls.push(`show:${payload.id ?? ''}:${payload.title}`);
        return payload.id ?? '';
      },
      remove: (id) => {
        calls.push(`remove:${id}`);
      },
    },
    { id: 'overlay-loading-status', dismiss: true },
  );

  assert.deepEqual(calls, ['remove:overlay-loading-status']);
});
