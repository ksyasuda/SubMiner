import assert from 'node:assert/strict';
import test from 'node:test';

import type { OverlayNotificationEntry } from './overlay-notifications';
import {
  createOverlayNotificationHistoryStore,
  resolveHistorySideFromStack,
} from './overlay-notification-history';

function entry(
  overrides: Partial<OverlayNotificationEntry> & { id: string },
): OverlayNotificationEntry {
  return {
    title: overrides.title ?? overrides.id,
    persistent: false,
    createdAt: 0,
    ...overrides,
  };
}

test('history store lists newest entries first', () => {
  const store = createOverlayNotificationHistoryStore();
  store.record(entry({ id: 'a', title: 'A' }));
  store.record(entry({ id: 'b', title: 'B' }));
  store.record(entry({ id: 'c', title: 'C' }));

  assert.deepEqual(
    store.list().map((item) => item.id),
    ['c', 'b', 'a'],
  );
  assert.equal(store.size(), 3);
});

test('history store updates an entry in place without reordering or duplicating', () => {
  let clock = 100;
  const store = createOverlayNotificationHistoryStore({ now: () => clock });
  store.record(entry({ id: 'job', title: 'Working', body: 'Step 1', variant: 'progress' }));
  store.record(entry({ id: 'other', title: 'Other' }));
  clock = 200;
  store.record(entry({ id: 'job', title: 'Done', body: 'Step 2', variant: 'success' }));

  const list = store.list();
  assert.equal(store.size(), 2);
  // Newest-first ordering is by first-seen; the in-place update keeps 'other' on top.
  assert.deepEqual(
    list.map((item) => item.id),
    ['other', 'job'],
  );
  const job = list.find((item) => item.id === 'job');
  assert.equal(job?.title, 'Done');
  assert.equal(job?.body, 'Step 2');
  assert.equal(job?.variant, 'success');
  assert.equal(job?.createdAt, 100);
  assert.equal(job?.updatedAt, 200);
});

test('history store removes and clears entries', () => {
  const store = createOverlayNotificationHistoryStore();
  store.record(entry({ id: 'a' }));
  store.record(entry({ id: 'b' }));

  store.remove('a');
  assert.deepEqual(
    store.list().map((item) => item.id),
    ['b'],
  );

  store.clear();
  assert.equal(store.size(), 0);
  assert.deepEqual(store.list(), []);
});

test('history store caps to max and drops the oldest entries', () => {
  const store = createOverlayNotificationHistoryStore({ max: 2 });
  store.record(entry({ id: 'a' }));
  store.record(entry({ id: 'b' }));
  store.record(entry({ id: 'c' }));

  assert.equal(store.size(), 2);
  assert.deepEqual(
    store.list().map((item) => item.id),
    ['c', 'b'],
  );
});

test('history store defaults missing variant to info', () => {
  const store = createOverlayNotificationHistoryStore();
  store.record(entry({ id: 'a' }));
  assert.equal(store.list()[0]?.variant, 'info');
});

test('panel side mirrors the notification stack position', () => {
  const stackWith = (positionClass: string) =>
    ({ classList: { contains: (token: string) => token === positionClass } }) as unknown as Element;

  assert.equal(resolveHistorySideFromStack(stackWith('position-top-left')), 'left');
  assert.equal(resolveHistorySideFromStack(stackWith('position-top-right')), 'right');
  // Center notifications open the panel from the right.
  assert.equal(resolveHistorySideFromStack(stackWith('position-top')), 'right');
});
