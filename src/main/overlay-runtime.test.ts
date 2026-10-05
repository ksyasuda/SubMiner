import assert from 'node:assert/strict';
import test from 'node:test';
import type { HyprlandPlacementStatus } from '../core/services/hyprland-window-placement';
import {
  FakeOverlayWindow,
  LOADING_BLANK_WINDOW,
  MODAL_GEOMETRY,
  createHarness,
  type FakeOverlayWindowInit,
} from './overlay-runtime-test-harness';
import type { OverlayHostedModal } from '../shared/ipc/contracts';

const SUBSYNC_PAYLOAD = {
  ffsubsyncAvailable: true,
  videoReferenceAvailable: true,
  subtitleTracks: [],
  defaultReferenceTrackId: null,
  defaultTargetTrackId: null,
};

const PENDING_PLACEMENT: HyprlandPlacementStatus = {
  applicable: true,
  clientFound: false,
  dispatched: false,
};

type FakeTimer = { callback: () => void; active: boolean; unref: () => void };

/**
 * Swaps the global setTimeout/clearTimeout for a manual queue while `run` executes. The
 * placement reconcile ladder calls the globals directly, so there is no option to inject.
 */
function withFakeTimers(
  run: (timers: { runNext: () => boolean; runAll: () => void; activeCount: () => number }) => void,
): void {
  const timers: FakeTimer[] = [];
  const originalSetTimeout = globalThis.setTimeout;
  const originalClearTimeout = globalThis.clearTimeout;
  globalThis.setTimeout = ((callback: () => void) => {
    const timer: FakeTimer = { callback, active: true, unref() {} };
    timers.push(timer);
    return timer;
  }) as unknown as typeof globalThis.setTimeout;
  globalThis.clearTimeout = ((timer?: FakeTimer) => {
    if (timer) timer.active = false;
  }) as unknown as typeof globalThis.clearTimeout;

  const runNext = (): boolean => {
    const timer = timers.find((candidate) => candidate.active);
    if (!timer) return false;
    timer.active = false;
    timer.callback();
    return true;
  };
  try {
    run({
      runNext,
      runAll: () => {
        for (let ran = 0; runNext(); ran += 1) {
          assert.ok(ran < 50, 'timer queue did not drain');
        }
      },
      activeCount: () => timers.filter((timer) => timer.active).length,
    });
  } finally {
    globalThis.setTimeout = originalSetTimeout;
    globalThis.clearTimeout = originalClearTimeout;
  }
}

test('sendToActiveOverlayWindow targets modal window with full geometry and tracks close restore', () => {
  const modal = new FakeOverlayWindow();
  const h = createHarness({ modal });

  const sent = h.runtime.sendToActiveOverlayWindow('runtime-options:open', undefined, {
    restoreOnModalClose: 'runtime-options',
  });
  assert.equal(sent, true);
  assert.equal(h.runtime.getRestoreVisibleOverlayOnModalClose().has('runtime-options'), true);
  assert.deepEqual(h.boundsCalls, [MODAL_GEOMETRY]);
  assert.equal(modal.showCount, 0);

  h.runtime.notifyOverlayModalOpened('runtime-options');
  assert.equal(h.createCalls, 0);
  assert.equal(modal.showCount, 1);
  assert.equal(modal.isFocused(), true);
  assert.deepEqual(h.boundsCalls, [MODAL_GEOMETRY, MODAL_GEOMETRY]);
  assert.deepEqual(modal.alwaysOnTopCalls, ['top:true:screen-saver:3']);
  assert.deepEqual(modal.sent, [['runtime-options:open']]);
});

test('sendToActiveOverlayWindow creates modal window lazily when absent', () => {
  const modal = new FakeOverlayWindow();
  const h = createHarness({ createModal: () => modal });

  assert.equal(
    h.runtime.sendToActiveOverlayWindow('jimaku:open', undefined, {
      restoreOnModalClose: 'jimaku',
    }),
    true,
  );
  assert.equal(h.createCalls, 1);
  assert.equal(modal.showCount, 0);
  h.runtime.notifyOverlayModalOpened('jimaku');
  assert.equal(modal.showCount, 1);
  assert.deepEqual(modal.sent, [['jimaku:open']]);
});

for (const platform of ['darwin', 'win32'] as const) {
  test(`primeModalWindow creates and warms a hidden modal on ${platform}`, () => {
    const modal = new FakeOverlayWindow({ ...LOADING_BLANK_WINDOW, documentLoaded: false });
    const h = createHarness({ createModal: () => modal, platform });

    assert.equal(h.runtime.primeModalWindow(), true);
    assert.equal(h.createCalls, 1);
    assert.equal(modal.isVisible(), false);

    modal.finishLoad();
    modal.emitReadyToShow();

    assert.equal(
      h.runtime.sendToActiveOverlayWindow('session-help:open', undefined, {
        restoreOnModalClose: 'session-help',
        preferModalWindow: true,
      }),
      true,
    );
    assert.equal(h.createCalls, 1);
    assert.equal(modal.isVisible(), true);
    assert.deepEqual(modal.sent, [['session-help:open']]);
  });
}

test('sendToActiveOverlayWindow delivers to a modal window created for the send', () => {
  // ready-to-show can fire before the document loads, and Electron still reports
  // isLoading() inside did-finish-load; only did-stop-loading sees the settled state.
  const modal = new FakeOverlayWindow({ ...LOADING_BLANK_WINDOW, documentLoaded: false });
  const h = createHarness({ createModal: () => modal });

  assert.equal(
    h.runtime.sendToActiveOverlayWindow(
      'media-timing-review:open',
      { reviewId: 'r1' },
      { restoreOnModalClose: 'media-timing-review', preferModalWindow: true },
    ),
    true,
  );
  modal.emitReadyToShow();
  assert.deepEqual(modal.sent, []);
  modal.finishLoad();
  assert.deepEqual(modal.sent, [['media-timing-review:open', { reviewId: 'r1' }]]);
});

test('primeModalWindow leaves Linux modal creation lazy', () => {
  const h = createHarness({ createModal: () => new FakeOverlayWindow(), platform: 'linux' });

  assert.equal(h.runtime.primeModalWindow(), false);
  assert.equal(h.createCalls, 0);
});

test('sendToActiveOverlayWindow does not retain restore state when modal creation fails', () => {
  const h = createHarness({ createModal: () => null });

  assert.equal(
    h.runtime.sendToActiveOverlayWindow('runtime-options:open', undefined, {
      restoreOnModalClose: 'runtime-options',
    }),
    false,
  );
  assert.equal(h.runtime.getRestoreVisibleOverlayOnModalClose().has('runtime-options'), false);
});

const readinessCases: Array<{
  name: string;
  window: FakeOverlayWindowInit;
  opens: OverlayHostedModal[];
  sentBeforeLoad: unknown[][];
}> = [
  {
    name: 'delivers on first modal load without waiting for ready-to-show',
    window: LOADING_BLANK_WINDOW,
    opens: ['runtime-options'],
    sentBeforeLoad: [],
  },
  {
    name: 'delivers when the modal loaded before listeners were registered',
    window: { contentReady: false },
    opens: ['runtime-options'],
    sentBeforeLoad: [['runtime-options:open']],
  },
  {
    name: 'does not infer document readiness from a pending file URL',
    window: { contentReady: false, documentLoaded: false },
    opens: ['runtime-options'],
    sentBeforeLoad: [],
  },
  {
    name: 'rejects stale content readiness during document reload',
    window: { contentReady: true, documentLoaded: false },
    opens: ['session-help'],
    sentBeforeLoad: [],
  },
  {
    name: 'flushes every queued open in order once the modal loads',
    window: LOADING_BLANK_WINDOW,
    opens: ['runtime-options', 'session-help'],
    sentBeforeLoad: [],
  },
];

for (const c of readinessCases) {
  test(`sendToActiveOverlayWindow ${c.name}`, () => {
    const modal = new FakeOverlayWindow(c.window);
    const h = createHarness({ modal });

    for (const modalId of c.opens) {
      assert.equal(
        h.runtime.sendToActiveOverlayWindow(`${modalId}:open`, undefined, {
          restoreOnModalClose: modalId,
        }),
        true,
      );
    }
    assert.deepEqual(modal.sent, c.sentBeforeLoad);

    const expected = c.opens.map((modalId) => [`${modalId}:open`]);
    modal.finishLoad();
    assert.deepEqual(modal.sent, expected);

    // Later ready-to-show and the renderer ack must not resend, and the ack reveals once.
    modal.emitReadyToShow();
    h.runtime.notifyOverlayModalOpened(c.opens[0]!);
    assert.deepEqual(modal.sent, expected);
    assert.equal(modal.showCount, 1);
    assert.equal(h.createCalls, 0);
  });
}

test('handleOverlayModalClosed keeps the modal window warm after all pending modals close', () => {
  const modal = new FakeOverlayWindow();
  const h = createHarness({ modal, platform: 'darwin' });

  h.runtime.sendToActiveOverlayWindow('runtime-options:open', undefined, {
    restoreOnModalClose: 'runtime-options',
  });
  h.runtime.sendToActiveOverlayWindow('subsync:open-manual', SUBSYNC_PAYLOAD, {
    restoreOnModalClose: 'subsync',
  });

  h.runtime.handleOverlayModalClosed('runtime-options');
  assert.equal(modal.isDestroyed(), false);

  h.runtime.handleOverlayModalClosed('subsync');
  assert.equal(modal.isDestroyed(), false);
  assert.equal(modal.isVisible(), false);
  assert.equal(modal.ignoreMouseEvents, true);
  assert.equal(h.runtime.getRestoreVisibleOverlayOnModalClose().size, 0);
});

test('sendToActiveOverlayWindow prefers visible main overlay window for modal open', () => {
  const main = new FakeOverlayWindow({ visible: true });
  const h = createHarness({ main, createModal: () => new FakeOverlayWindow() });

  const sent = h.runtime.sendToActiveOverlayWindow('runtime-options:open', undefined, {
    restoreOnModalClose: 'runtime-options',
  });

  assert.equal(sent, true);
  assert.equal(h.createCalls, 0);
  assert.deepEqual(main.sent, [['runtime-options:open']]);
});

test('sendToActiveOverlayWindow can prefer modal window even when main overlay is visible', () => {
  const main = new FakeOverlayWindow({ visible: true });
  const modal = new FakeOverlayWindow();
  const h = createHarness({ main, modal });

  const sent = h.runtime.sendToActiveOverlayWindow(
    'youtube:picker-open',
    { sessionId: 'yt-1' },
    { restoreOnModalClose: 'youtube-track-picker', preferModalWindow: true },
  );

  assert.equal(sent, true);
  assert.deepEqual(main.sent, []);
  assert.deepEqual(modal.sent, [['youtube:picker-open', { sessionId: 'yt-1' }]]);
});

test('modal window path makes visible main overlay click-through until modal closes', () => {
  const main = new FakeOverlayWindow({ visible: true });
  const modal = new FakeOverlayWindow();
  const h = createHarness({ main, modal });

  const sent = h.runtime.sendToActiveOverlayWindow(
    'youtube:picker-open',
    { sessionId: 'yt-1' },
    { restoreOnModalClose: 'youtube-track-picker', preferModalWindow: true },
  );
  h.runtime.notifyOverlayModalOpened('youtube-track-picker');

  assert.equal(sent, true);
  assert.equal(main.ignoreMouseEvents, true);
  assert.equal(main.forwardedIgnoreMouseEvents, true);
  assert.equal(modal.ignoreMouseEvents, false);

  h.runtime.handleOverlayModalClosed('youtube-track-picker');

  assert.equal(main.ignoreMouseEvents, true);
});

test('modal window path restores visible main overlay before modal input deactivates', () => {
  const main = new FakeOverlayWindow({ visible: true });
  const events: string[] = [];
  const h = createHarness({
    main,
    modal: new FakeOverlayWindow(),
    options: {
      onModalStateChange: (active) => events.push(`state:${active}:visible:${main.isVisible()}`),
    },
  });

  h.runtime.sendToActiveOverlayWindow(
    'youtube:picker-open',
    { sessionId: 'yt-1' },
    { restoreOnModalClose: 'youtube-track-picker', preferModalWindow: true },
  );
  h.runtime.notifyOverlayModalOpened('youtube-track-picker');

  assert.equal(main.hideCount, 1);
  assert.equal(main.isVisible(), false);

  h.runtime.handleOverlayModalClosed('youtube-track-picker');

  assert.equal(main.showCount, 1);
  assert.equal(main.isVisible(), true);
  assert.deepEqual(events, ['state:true:visible:true', 'state:false:visible:true']);
});

test('macOS maps a new modal panel before focusing SubMiner and hiding the subtitle overlay', () => {
  const log: string[] = [];
  const main = new FakeOverlayWindow({ name: 'main', log, visible: true });
  const modal = new FakeOverlayWindow({ name: 'modal', log });
  const h = createHarness({
    main,
    modal,
    platform: 'darwin',
    options: { focusApplication: () => log.push('focus-application') },
  });

  h.runtime.sendToActiveOverlayWindow('runtime-options:open', undefined, {
    restoreOnModalClose: 'runtime-options',
    preferModalWindow: true,
  });
  h.runtime.notifyOverlayModalOpened('runtime-options');

  assert.deepEqual(log, ['modal:show-inactive', 'focus-application', 'main:hide']);
  assert.equal(modal.isVisible(), true);
  assert.equal(main.isVisible(), false);
});

test('modal window path runs final close handoff before modal input deactivates', () => {
  const main = new FakeOverlayWindow({ visible: true });
  const events: string[] = [];
  const h = createHarness({
    main,
    modal: new FakeOverlayWindow(),
    options: {
      onFinalModalClosed: () => events.push(`handoff:visible:${main.isVisible()}`),
      onModalStateChange: (active) => events.push(`state:${active}:visible:${main.isVisible()}`),
    },
  });

  h.runtime.sendToActiveOverlayWindow(
    'youtube:picker-open',
    { sessionId: 'yt-1' },
    { restoreOnModalClose: 'youtube-track-picker', preferModalWindow: true },
  );
  h.runtime.notifyOverlayModalOpened('youtube-track-picker');
  h.runtime.handleOverlayModalClosed('youtube-track-picker');

  assert.deepEqual(events, [
    'state:true:visible:true',
    'handoff:visible:true',
    'state:false:visible:true',
  ]);
});

test('modal runtime deactivates modal state when final close handoff throws', () => {
  const events: string[] = [];
  const h = createHarness({
    main: new FakeOverlayWindow({ visible: true }),
    modal: new FakeOverlayWindow(),
    options: {
      onFinalModalClosed: () => {
        events.push('handoff');
        throw new Error('handoff failed');
      },
      onModalStateChange: (active) => events.push(`state:${active}`),
    },
  });

  h.runtime.sendToActiveOverlayWindow(
    'youtube:picker-open',
    { sessionId: 'yt-1' },
    { restoreOnModalClose: 'youtube-track-picker', preferModalWindow: true },
  );
  h.runtime.notifyOverlayModalOpened('youtube-track-picker');

  assert.doesNotThrow(() => h.runtime.handleOverlayModalClosed('youtube-track-picker'));
  assert.deepEqual(events, ['state:true', 'handoff', 'state:false']);
});

test('modal runtime notifies callers when modal input state becomes active/inactive', () => {
  const state: boolean[] = [];
  const h = createHarness({
    modal: new FakeOverlayWindow(),
    options: { onModalStateChange: (active) => state.push(active) },
  });

  h.runtime.sendToActiveOverlayWindow('runtime-options:open', undefined, {
    restoreOnModalClose: 'runtime-options',
  });
  h.runtime.sendToActiveOverlayWindow('subsync:open-manual', SUBSYNC_PAYLOAD, {
    restoreOnModalClose: 'subsync',
  });
  assert.deepEqual(state, []);
  h.runtime.notifyOverlayModalOpened('runtime-options');
  assert.deepEqual(state, [true]);

  h.runtime.handleOverlayModalClosed('runtime-options');
  assert.deepEqual(state, [true]);

  h.runtime.handleOverlayModalClosed('subsync');
  assert.deepEqual(state, [true, false]);
});

test('notifyOverlayModalOpened enables input on visible main overlay window when no modal window exists', () => {
  const main = new FakeOverlayWindow({ visible: true, ignoreMouseEvents: true });
  const state: boolean[] = [];
  const h = createHarness({
    main,
    options: { onModalStateChange: (active) => state.push(active) },
  });

  const sent = h.runtime.sendToActiveOverlayWindow('runtime-options:open', undefined, {
    restoreOnModalClose: 'runtime-options',
  });
  h.runtime.notifyOverlayModalOpened('runtime-options');

  assert.equal(sent, true);
  assert.equal(h.createCalls, 0);
  assert.deepEqual(state, [true]);
  assert.equal(main.ignoreMouseEvents, false);
  assert.equal(main.isFocused(), true);
  assert.equal(main.webContentsFocused, true);
});

test('handleOverlayModalClosed is a no-op when no modal window can be targeted', () => {
  const state: boolean[] = [];
  const h = createHarness({
    createModal: () => null,
    options: { onModalStateChange: (active) => state.push(active) },
  });

  const sent = h.runtime.sendToActiveOverlayWindow('runtime-options:open', undefined, {
    restoreOnModalClose: 'runtime-options',
  });
  assert.equal(sent, false);
  h.runtime.notifyOverlayModalOpened('runtime-options');
  h.runtime.handleOverlayModalClosed('runtime-options');

  assert.deepEqual(state, []);
});

test('modal fallback reveal skips showing window when content is not ready', () => {
  const modal = new FakeOverlayWindow(LOADING_BLANK_WINDOW);
  let scheduledReveal: (() => void) | null = null;
  const h = createHarness({
    modal,
    platform: 'darwin',
    options: {
      scheduleRevealFallback: (callback) => {
        scheduledReveal = callback;
        return { scheduled: true } as never;
      },
      clearRevealFallback: () => {
        scheduledReveal = null;
      },
    },
  });

  const sent = h.runtime.sendToActiveOverlayWindow('jimaku:open', undefined, {
    restoreOnModalClose: 'jimaku',
  });

  assert.equal(sent, true);
  assert.ok(scheduledReveal, 'expected reveal callback');
  (scheduledReveal as () => void)();
  assert.equal(modal.showCount, 0);

  h.runtime.notifyOverlayModalOpened('jimaku');
  assert.equal(modal.showCount, 1);
  assert.equal(modal.ignoreMouseEvents, false);
});

test('modal reopen reuses the warm window and shows it immediately on macOS', () => {
  const modal = new FakeOverlayWindow();
  const state: boolean[] = [];
  const h = createHarness({
    modal,
    createModal: () => modal,
    platform: 'darwin',
    options: { onModalStateChange: (active) => state.push(active) },
  });

  h.runtime.sendToActiveOverlayWindow('runtime-options:open', undefined, {
    restoreOnModalClose: 'runtime-options',
  });
  h.runtime.notifyOverlayModalOpened('runtime-options');
  h.runtime.handleOverlayModalClosed('runtime-options');

  assert.equal(modal.isDestroyed(), false);
  assert.equal(modal.isVisible(), false);
  assert.deepEqual(state, [true, false]);

  const sent = h.runtime.sendToActiveOverlayWindow('runtime-options:open', undefined, {
    restoreOnModalClose: 'runtime-options',
  });

  assert.equal(sent, true);
  assert.equal(h.createCalls, 0);
  assert.equal(modal.isVisible(), true);
  assert.equal(modal.showCount, 2);

  h.runtime.notifyOverlayModalOpened('runtime-options');
  assert.deepEqual(state, [true, false, true]);
});

test('modal reopen on Windows uses a fresh prewarmed interactive window', () => {
  const firstWindow = new FakeOverlayWindow();
  const replacementWindow = new FakeOverlayWindow();
  const h = createHarness({
    modal: firstWindow,
    createModal: () => replacementWindow,
    platform: 'win32',
  });

  h.runtime.sendToActiveOverlayWindow('runtime-options:open', undefined, {
    restoreOnModalClose: 'runtime-options',
  });
  h.runtime.notifyOverlayModalOpened('runtime-options');
  h.runtime.handleOverlayModalClosed('runtime-options');

  assert.equal(firstWindow.isDestroyed(), true);
  assert.equal(h.modal, replacementWindow);
  assert.equal(replacementWindow.isVisible(), false);
  assert.equal(h.createCalls, 1);

  const sent = h.runtime.sendToActiveOverlayWindow('session-help:open', undefined, {
    restoreOnModalClose: 'session-help',
  });

  assert.equal(sent, true);
  assert.equal(h.createCalls, 1);
  assert.equal(replacementWindow.isVisible(), true);
  assert.equal(replacementWindow.ignoreMouseEvents, false);
  assert.deepEqual(replacementWindow.sent, [['session-help:open']]);
});

test('visible stale modal window is made interactive again before reopening', () => {
  const modal = new FakeOverlayWindow({
    visible: true,
    focused: true,
    webContentsFocused: false,
    ignoreMouseEvents: true,
  });
  const h = createHarness({ modal });

  const sent = h.runtime.sendToActiveOverlayWindow('runtime-options:open', undefined, {
    restoreOnModalClose: 'runtime-options',
  });

  assert.equal(sent, true);
  assert.equal(modal.ignoreMouseEvents, false);
  assert.equal(modal.isFocused(), true);
  assert.equal(modal.webContentsFocused, true);
  assert.deepEqual(modal.sent, [['runtime-options:open']]);
});

test('waitForModalOpen resolves true after modal acknowledgement', async () => {
  const h = createHarness({ modal: new FakeOverlayWindow() });

  h.runtime.sendToActiveOverlayWindow(
    'youtube:picker-open',
    { sessionId: 'yt-1' },
    { restoreOnModalClose: 'youtube-track-picker' },
  );
  const pending = h.runtime.waitForModalOpen('youtube-track-picker', 1000);
  h.runtime.notifyOverlayModalOpened('youtube-track-picker');

  assert.equal(await pending, true);
});

test('waitForModalOpen resolves true when modal acknowledgement arrives before waiter registration', async () => {
  const h = createHarness({ modal: new FakeOverlayWindow() });

  h.runtime.sendToActiveOverlayWindow(
    'kiku:field-grouping-request',
    {},
    { restoreOnModalClose: 'kiku' },
  );
  h.runtime.notifyOverlayModalOpened('kiku');

  assert.equal(await h.runtime.waitForModalOpen('kiku', 5), true);
});

test('waitForModalOpen resolves false on timeout', async () => {
  const h = createHarness();

  assert.equal(await h.runtime.waitForModalOpen('youtube-track-picker', 5), false);
});

function openKikuInModalWindow(h: ReturnType<typeof createHarness>): void {
  h.runtime.sendToActiveOverlayWindow(
    'kiku:field-grouping-open',
    { test: true },
    { restoreOnModalClose: 'kiku', preferModalWindow: true },
  );
}

test('modal placement reconcile retries until the Hyprland client is mapped', () => {
  withFakeTimers((timers) => {
    // The compositor never maps the window, so every reconcile reports pending.
    const h = createHarness({
      modal: new FakeOverlayWindow(),
      setModalWindowBounds: () => PENDING_PLACEMENT,
    });
    openKikuInModalWindow(h);
    h.runtime.notifyOverlayModalOpened('kiku');
    const reconcilesBeforeLadder = h.boundsCalls.length;

    timers.runAll();

    // One re-assert per post-show ladder delay.
    assert.equal(h.boundsCalls.length - reconcilesBeforeLadder, 6);
  });
});

test('modal placement reconcile stops retrying once the client is mapped', () => {
  withFakeTimers((timers) => {
    const h = createHarness({
      modal: new FakeOverlayWindow(),
      setModalWindowBounds: () => ({ applicable: true, clientFound: true, dispatched: true }),
    });
    openKikuInModalWindow(h);
    h.runtime.notifyOverlayModalOpened('kiku');
    const reconcilesBeforeLadder = h.boundsCalls.length;

    timers.runAll();

    assert.equal(h.boundsCalls.length - reconcilesBeforeLadder, 1);
  });
});

test('modal placement reconcile cancels stale retry ladder after a newer visible modal interaction', () => {
  withFakeTimers((timers) => {
    const h = createHarness({
      modal: new FakeOverlayWindow(),
      setModalWindowBounds: () => PENDING_PLACEMENT,
    });
    openKikuInModalWindow(h);
    h.runtime.notifyOverlayModalOpened('kiku');
    assert.equal(timers.activeCount(), 1);

    openKikuInModalWindow(h);
    assert.equal(timers.activeCount(), 2);

    timers.runNext();

    assert.equal(timers.activeCount(), 1, 'stale retry should not schedule a continuation');
  });
});

test('Linux keeps the dedicated modal window unmapped until the renderer opens the modal, then hides the overlay before revealing it', () => {
  const log: string[] = [];
  const main = new FakeOverlayWindow({ name: 'main', log, visible: true });
  const modal = new FakeOverlayWindow({ name: 'modal', log });
  let revealScheduled = false;
  const h = createHarness({
    main,
    modal,
    platform: 'linux',
    options: {
      scheduleRevealFallback: () => {
        revealScheduled = true;
        return { scheduled: true } as never;
      },
      clearRevealFallback: () => {},
    },
  });

  const open = () =>
    h.runtime.sendToActiveOverlayWindow(
      'media-timing-review:open',
      { reviewId: 'review' },
      { restoreOnModalClose: 'media-timing-review', preferModalWindow: true },
    );

  assert.equal(open(), true);
  assert.deepEqual(modal.sent, [['media-timing-review:open', { reviewId: 'review' }]]);
  assert.equal(revealScheduled, false);
  assert.equal(modal.showCount, 0);
  assert.equal(main.hideCount, 0);

  // The open retry must not map the window before the renderer answers either.
  assert.equal(open(), true);
  assert.equal(modal.showCount, 0);

  h.runtime.notifyOverlayModalOpened('media-timing-review');

  assert.deepEqual(log, ['main:hide', 'modal:show']);
  assert.equal(main.isVisible(), false);
  assert.equal(modal.isVisible(), true);
  assert.equal(modal.ignoreMouseEvents, false);
});
