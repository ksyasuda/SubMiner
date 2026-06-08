import assert from 'node:assert/strict';
import test from 'node:test';

import type { OverlayNotificationEntry } from './overlay-notifications';
import {
  createOverlayNotificationHistoryPanel,
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

test('history store keeps same live notification id when history ids differ', () => {
  const store = createOverlayNotificationHistoryStore();
  store.record(
    entry({
      id: 'character-dictionary-auto-sync',
      title: 'Character dictionary',
      body: 'Checking character dictionary...',
      variant: 'progress',
      historyId: 'character-dictionary-auto-sync-checking',
    }),
  );
  store.record(
    entry({
      id: 'character-dictionary-auto-sync',
      title: 'Character dictionary',
      body: 'Building character dictionary...',
      variant: 'progress',
      historyId: 'character-dictionary-auto-sync-building',
    }),
  );
  store.record(
    entry({
      id: 'character-dictionary-auto-sync',
      title: 'Character dictionary',
      body: 'Character dictionary ready',
      variant: 'success',
      historyId: 'character-dictionary-auto-sync-ready',
    }),
  );

  assert.deepEqual(
    store.list().map((item) => `${item.id}:${item.body}`),
    [
      'character-dictionary-auto-sync-ready:Character dictionary ready',
      'character-dictionary-auto-sync-building:Building character dictionary...',
      'character-dictionary-auto-sync-checking:Checking character dictionary...',
    ],
  );
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

function createClassList(initialTokens: string[] = []) {
  const tokens = new Set(initialTokens);
  return {
    add: (...entries: string[]) => {
      for (const entry of entries) tokens.add(entry);
    },
    remove: (...entries: string[]) => {
      for (const entry of entries) tokens.delete(entry);
    },
    contains: (entry: string) => tokens.has(entry),
    toggle: (entry: string, force?: boolean) => {
      if (force === true) tokens.add(entry);
      else if (force === false) tokens.delete(entry);
      else if (tokens.has(entry)) tokens.delete(entry);
      else tokens.add(entry);
    },
  };
}

function createPanelHarness(stackPositionClass: string) {
  const stack = {
    classList: createClassList([stackPositionClass]),
  };
  const clearButton = {
    disabled: false,
    addEventListener: () => undefined,
  };
  const closeButton = {
    addEventListener: () => undefined,
  };
  const list = {
    replaceChildren: () => undefined,
  };
  const empty = {
    classList: createClassList(),
  };
  const panel = {
    classList: createClassList(['notification-history', 'side-right']),
    setAttribute: () => undefined,
    addEventListener: () => undefined,
    querySelector: (selector: string) => {
      switch (selector) {
        case '.notification-history-list':
          return list;
        case '.notification-history-empty':
          return empty;
        case '.notification-history-clear':
          return clearButton;
        case '.notification-history-close':
          return closeButton;
        default:
          return null;
      }
    },
  };

  const controller = createOverlayNotificationHistoryPanel({
    dom: {
      overlayNotificationHistory: panel,
      overlayNotificationStack: stack,
    },
    state: {
      isOverNotificationHistory: false,
      notificationHistoryOpen: false,
    },
    platform: {
      shouldToggleMouseIgnore: false,
    },
  } as never);

  return { controller, panel, stack };
}

test('history panel applies the initial stack side while still closed', () => {
  const { panel } = createPanelHarness('position-top-left');

  assert.equal(panel.classList.contains('side-left'), true);
  assert.equal(panel.classList.contains('side-right'), false);
  assert.equal(panel.classList.contains('open'), false);
});

test('history panel resyncs the closed side before first open', () => {
  const { controller, panel, stack } = createPanelHarness('position-top-right');

  stack.classList.remove('position-top-right');
  stack.classList.add('position-top-left');

  const syncable = controller as unknown as { syncSide?: () => void };
  assert.equal(typeof syncable.syncSide, 'function');
  syncable.syncSide?.();

  assert.equal(panel.classList.contains('side-left'), true);
  assert.equal(panel.classList.contains('side-right'), false);
  assert.equal(panel.classList.contains('open'), false);
});
