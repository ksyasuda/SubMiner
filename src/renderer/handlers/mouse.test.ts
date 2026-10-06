import assert from 'node:assert/strict';
import test, { type TestContext } from 'node:test';

import { SUBTITLE_DEFAULT_CONFIG } from '../../config/definitions/defaults-subtitle';
import { createRendererState } from '../state';
import { createMouseHandlers } from './mouse.js';
import {
  HACHIDORI_HOST_SELECTOR,
  HACHIDORI_POPUP_HIDDEN_EVENT,
  HACHIDORI_POPUP_SELECTOR,
  HACHIDORI_POPUP_SHOWN_EVENT,
  YOMITAN_LOOKUP_EVENT,
  YOMITAN_POPUP_HIDDEN_EVENT,
  YOMITAN_POPUP_HOST_SELECTOR,
  YOMITAN_POPUP_MOUSE_ENTER_EVENT,
  YOMITAN_POPUP_MOUSE_LEAVE_EVENT,
  YOMITAN_POPUP_SHOWN_EVENT,
  YOMITAN_POPUP_VISIBLE_HOST_SELECTOR,
} from '../yomitan-popup.js';

function createClassList() {
  const classes = new Set<string>();
  return {
    add: (...tokens: string[]) => tokens.forEach((token) => classes.add(token)),
    remove: (...tokens: string[]) => tokens.forEach((token) => classes.delete(token)),
    toggle: (token: string, force?: boolean) => {
      const next = force ?? !classes.has(token);
      if (next) classes.add(token);
      else classes.delete(token);
      return next;
    },
    contains: (token: string) => classes.has(token),
  };
}

function createDeferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((nextResolve) => {
    resolve = nextResolve;
  });
  return { promise, resolve };
}

function waitForNextTick(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

type Listener = (event: unknown) => void;
type IgnoreCall = { ignore: boolean; forward?: boolean };
type HoverTarget = 'primary' | 'secondary' | null;

const MOCKED_GLOBALS = ['window', 'document', 'MutationObserver', 'Node'] as const;

/**
 * Installs window/document/MutationObserver/Node stubs (restored via `t.after`) and builds
 * mouse handlers with no-op defaults. `hoverPause`, `popupPause`, `hovered` and
 * `popupHostVisible` are mutable mid-test; `fire` dispatches to the registered listeners.
 * With `hachidori`, the popup host is a Hachidori shadow host: `popupHostVisible` is its
 * attention flag, `popupPaneOpen` its popup pane, and `popupHostAttached` its presence.
 */
function createMouseHarness(
  t: TestContext,
  options: {
    platform?: 'windows' | 'macos';
    hoverPause?: boolean;
    popupPause?: boolean;
    paused?: boolean;
    getPlaybackPaused?: () => Promise<boolean | null>;
    popupHostVisible?: boolean;
    hovered?: HoverTarget;
    hachidori?: boolean;
  } = {},
) {
  const overlayClassList = createClassList();
  const bodyClassList = createClassList();
  const focusCalls = { mainWindow: 0, window: 0, overlay: 0 };
  const ctx = {
    dom: {
      overlay: {
        classList: overlayClassList,
        focus: () => {
          focusCalls.overlay += 1;
        },
      },
      subtitleRoot: { classList: createClassList() },
      subtitleContainer: {
        classList: createClassList(),
        style: { cursor: '' },
        addEventListener: () => {},
      },
      secondarySubContainer: {
        classList: createClassList(),
        addEventListener: () => {},
      },
    },
    platform: {
      shouldToggleMouseIgnore: true,
      isLinuxPlatform: false,
      isMacOSPlatform: options.platform === 'macos',
    },
    state: { ...createRendererState(), primaryVisibleOnYomitanPopup: false },
  };

  const listeners = {
    window: new Map<string, Listener[]>(),
    document: new Map<string, Listener[]>(),
  };
  const addListener = (target: keyof typeof listeners) => (type: string, listener: Listener) => {
    listeners[target].set(type, [...(listeners[target].get(type) ?? []), listener]);
  };
  const ignoreCalls: IgnoreCall[] = [];
  const mpvCommands: Array<(string | number)[]> = [];
  const popupPane = {
    get hidden() {
      return !harness.popupPaneOpen;
    },
  };
  const popupHost = {
    tagName: options.hachidori ? 'HACHIDORI-HOST' : 'DIV',
    getAttribute: (name: string) =>
      name === 'data-subminer-yomitan-popup-visible' ? String(harness.popupHostVisible) : null,
    shadowRoot: {
      querySelectorAll: (selector: string) =>
        selector === HACHIDORI_POPUP_SELECTOR ? [popupPane] : [],
    },
  };
  const hostSelectors = options.hachidori
    ? [HACHIDORI_HOST_SELECTOR, YOMITAN_POPUP_HOST_SELECTOR]
    : [YOMITAN_POPUP_HOST_SELECTOR];

  const harness = {
    ctx,
    ignoreCalls,
    mpvCommands,
    focusCalls,
    bodyClassList,
    hoverPause: options.hoverPause ?? false,
    popupPause: options.popupPause ?? false,
    hovered: options.hovered ?? null,
    popupHostVisible: options.popupHostVisible ?? false,
    popupHostAttached: options.hachidori ?? false,
    popupPaneOpen: false,
    handlers: null as unknown as ReturnType<typeof createMouseHandlers>,
    fire: (target: keyof typeof listeners, type: string, event: unknown = {}) => {
      for (const listener of listeners[target].get(type) ?? []) listener(event);
    },
    /** Moves the tracked pointer onto `hovered` and fires a document mousemove. */
    move: (hovered: HoverTarget, clientX = 120, clientY = 240) => {
      harness.hovered = hovered;
      harness.fire('document', 'mousemove', { clientX, clientY });
    },
    isInteractive: () => overlayClassList.contains('interactive'),
    isSecondaryHoverActive: () =>
      ctx.dom.secondarySubContainer.classList.contains('secondary-sub-hover-active'),
  };

  const stubs: Record<(typeof MOCKED_GLOBALS)[number], unknown> = {
    window: {
      addEventListener: addListener('window'),
      electronAPI: {
        setIgnoreMouseEvents: (ignore: boolean, ignoreOptions?: { forward?: boolean }) => {
          ignoreCalls.push({ ignore, forward: ignoreOptions?.forward });
        },
        focusMainWindow: () => {
          focusCalls.mainWindow += 1;
        },
      },
      focus: () => {
        focusCalls.window += 1;
      },
      getComputedStyle: () => ({ visibility: 'visible', display: 'block', opacity: '1' }),
      innerHeight: 1000,
      getSelection: () => null,
      setTimeout,
      clearTimeout,
    },
    document: {
      visibilityState: 'visible',
      body: { classList: bodyClassList },
      addEventListener: addListener('document'),
      querySelector: () => null,
      querySelectorAll: (selector: string) => {
        if (selector === YOMITAN_POPUP_VISIBLE_HOST_SELECTOR) {
          return harness.popupHostVisible ? [popupHost] : [];
        }
        return hostSelectors.includes(selector) &&
          (harness.popupHostVisible || harness.popupHostAttached)
          ? [popupHost]
          : [];
      },
      elementFromPoint: () =>
        harness.hovered === 'primary'
          ? ctx.dom.subtitleContainer
          : harness.hovered === 'secondary'
            ? ctx.dom.secondarySubContainer
            : null,
    },
    MutationObserver: class {
      observe() {}
    },
    Node: class {
      static ELEMENT_NODE = 1;
    },
  };

  const previous = MOCKED_GLOBALS.map(
    (key) => [key, Object.getOwnPropertyDescriptor(globalThis, key)] as const,
  );
  for (const key of MOCKED_GLOBALS) {
    Object.defineProperty(globalThis, key, {
      configurable: true,
      writable: true,
      value: stubs[key],
    });
  }
  t.after(() => {
    for (const [key, descriptor] of previous) {
      if (descriptor) Object.defineProperty(globalThis, key, descriptor);
      else delete (globalThis as Record<string, unknown>)[key];
    }
  });

  harness.handlers = createMouseHandlers(ctx as never, {
    modalStateReader: {
      isAnySettingsModalOpen: () => false,
      isAnyModalOpen: () => false,
    },
    applyYPercent: () => {},
    getCurrentYPercent: () => 10,
    persistSubtitlePositionPatch: () => {},
    getSubtitleHoverAutoPauseEnabled: () => harness.hoverPause,
    getYomitanPopupAutoPauseEnabled: () => harness.popupPause,
    getPlaybackPaused: options.getPlaybackPaused ?? (async () => options.paused ?? false),
    sendMpvCommand: (command) => {
      mpvCommands.push(command);
    },
  });

  return harness;
}

const PAUSE = ['set_property', 'pause', 'yes'];
const RESUME = ['set_property', 'pause', 'no'];

// Subtitle hover auto-pause

for (const c of [
  {
    name: 'secondary hover pauses on enter, reveals secondary subtitle, and resumes on leave',
    target: 'secondary',
    revealsSecondary: true,
  },
  {
    name: 'primary hover pauses on enter without revealing secondary subtitle, and resumes on leave',
    target: 'primary',
    revealsSecondary: false,
  },
] as const) {
  test(c.name, async (t) => {
    const h = createMouseHarness(t, { hoverPause: true });
    const { handlers } = h;
    const [enter, leave] =
      c.target === 'secondary'
        ? [handlers.handleSecondaryMouseEnter, handlers.handleSecondaryMouseLeave]
        : [handlers.handlePrimaryMouseEnter, handlers.handlePrimaryMouseLeave];

    await enter();
    assert.equal(h.isSecondaryHoverActive(), c.revealsSecondary);
    await leave();
    assert.equal(h.isSecondaryHoverActive(), false);

    assert.deepEqual(h.mpvCommands, [PAUSE, RESUME]);
  });
}

test('moving between primary and secondary subtitle containers keeps the hover pause active', async (t) => {
  const h = createMouseHarness(t, { hoverPause: true });
  const { subtitleContainer, secondarySubContainer } = h.ctx.dom;

  await h.handlers.handleSecondaryMouseEnter();
  await h.handlers.handleSecondaryMouseLeave({
    relatedTarget: subtitleContainer,
  } as unknown as MouseEvent);
  assert.equal(h.ctx.state.isOverSubtitle, false);
  assert.equal(h.isSecondaryHoverActive(), false);

  await h.handlers.handlePrimaryMouseEnter({
    relatedTarget: secondarySubContainer,
  } as unknown as MouseEvent);

  assert.equal(h.ctx.state.isOverSubtitle, true);
  assert.deepEqual(h.mpvCommands, [PAUSE]);
});

for (const c of [
  { name: 'playback is already paused', hoverPause: true, paused: true },
  { name: 'disabled in config', hoverPause: false, paused: false },
]) {
  test(`auto-pause on subtitle hover is skipped when ${c.name}`, async (t) => {
    const h = createMouseHarness(t, { hoverPause: c.hoverPause, paused: c.paused });

    await h.handlers.handleMouseEnter();
    await h.handlers.handleMouseLeave();

    assert.deepEqual(h.mpvCommands, []);
  });
}

test('pending hover pause check is ignored when mouse leaves before pause state resolves', async (t) => {
  const deferred = createDeferred<boolean | null>();
  const h = createMouseHarness(t, { hoverPause: true, getPlaybackPaused: () => deferred.promise });

  const enterPromise = h.handlers.handleMouseEnter();
  await h.handlers.handleMouseLeave();
  deferred.resolve(false);
  await enterPromise;

  assert.deepEqual(h.mpvCommands, []);
});

test('subtitle leave restores passthrough while embedded sidebar is open but not hovered', async (t) => {
  const h = createMouseHarness(t);
  h.ctx.state.isOverSubtitle = true;
  h.ctx.state.subtitleSidebarModalOpen = true;
  h.ctx.state.subtitleSidebarConfig = {
    ...SUBTITLE_DEFAULT_CONFIG.subtitleSidebar,
    layout: 'embedded',
  };

  await h.handlers.handlePrimaryMouseLeave();

  assert.equal(h.isInteractive(), false);
  assert.deepEqual(h.ignoreCalls.at(-1), { ignore: true, forward: true });
});

// Yomitan popup interaction

test('hover pause resumes immediately on subtitle leave even when yomitan popup is visible', async (t) => {
  const h = createMouseHarness(t, { hoverPause: true });
  h.handlers.setupYomitanObserver();
  h.fire('window', YOMITAN_POPUP_SHOWN_EVENT);

  await h.handlers.handleMouseEnter();
  await h.handlers.handleMouseLeave();

  assert.deepEqual(h.mpvCommands, [PAUSE, RESUME]);
});

test('popup open pauses and popup close resumes when yomitan popup auto-pause is enabled', async (t) => {
  const h = createMouseHarness(t, { hoverPause: true, popupPause: true });
  h.handlers.setupYomitanObserver();

  h.fire('window', YOMITAN_POPUP_SHOWN_EVENT);
  await waitForNextTick();
  h.fire('window', YOMITAN_POPUP_HIDDEN_EVENT);

  assert.deepEqual(h.mpvCommands, [PAUSE, RESUME]);
});

for (const c of [
  {
    name: 'nested popup close reasserts interactive state and focus when another popup remains visible on Windows',
    platform: 'windows',
    event: YOMITAN_POPUP_HIDDEN_EVENT,
    reclaimsFocus: true,
  },
  {
    name: 'window blur reclaims overlay focus while a yomitan popup remains visible on Windows',
    platform: 'windows',
    event: 'blur',
    reclaimsFocus: true,
  },
  {
    name: 'window blur on macOS keeps yomitan popup interactive without stealing click-away focus',
    platform: 'macos',
    event: 'blur',
    reclaimsFocus: false,
  },
] as const) {
  test(c.name, async (t) => {
    const h = createMouseHarness(t, { platform: c.platform, popupHostVisible: true });
    h.handlers.setupYomitanObserver();
    assert.equal(h.ctx.state.yomitanPopupVisible, true);
    assert.equal(h.isInteractive(), true);
    h.ignoreCalls.length = 0;

    h.fire('window', c.event);
    // Blur reconciles in a microtask.
    await Promise.resolve();

    assert.equal(h.ctx.state.yomitanPopupVisible, true);
    assert.equal(h.isInteractive(), true);
    assert.deepEqual(h.ignoreCalls.at(-1), { ignore: false, forward: undefined });
    const expectedFocus = c.reclaimsFocus ? 1 : 0;
    assert.deepEqual(h.focusCalls, {
      mainWindow: expectedFocus,
      window: expectedFocus,
      overlay: expectedFocus,
    });
  });
}

// Hachidori publishes one attention signal for an open popup and for a press on
// subtitle text that may become a selection, so its popup pane is the source of
// truth for what is on screen.

test('Hachidori press on subtitle text does not pause before a popup opens', async (t) => {
  const h = createMouseHarness(t, { popupPause: true, hachidori: true });
  h.handlers.setupYomitanObserver();

  h.popupHostVisible = true;
  h.fire('window', HACHIDORI_POPUP_SHOWN_EVENT);
  await waitForNextTick();
  h.popupHostVisible = false;
  h.fire('window', HACHIDORI_POPUP_HIDDEN_EVENT);
  await waitForNextTick();

  assert.deepEqual(h.mpvCommands, []);
});

test('Hachidori press before any lookup does not pause while its host is unattached', async (t) => {
  const h = createMouseHarness(t, { popupPause: true, hachidori: true });
  h.handlers.setupYomitanObserver();

  // Hachidori attaches its host on the first lookup, so a press anywhere on
  // the overlay before that claims attention with nothing in the DOM.
  h.popupHostAttached = false;
  h.fire('window', HACHIDORI_POPUP_SHOWN_EVENT);
  await waitForNextTick();
  h.fire('window', HACHIDORI_POPUP_HIDDEN_EVENT);
  await waitForNextTick();

  assert.deepEqual(h.mpvCommands, []);
});

test('Hachidori selection lookup pauses once its popup opens and resumes on close', async (t) => {
  const h = createMouseHarness(t, { popupPause: true, hachidori: true });
  h.handlers.setupYomitanObserver();

  h.popupHostVisible = true;
  h.fire('window', HACHIDORI_POPUP_SHOWN_EVENT);
  await waitForNextTick();
  assert.deepEqual(h.mpvCommands, []);

  // The drag's attention carries over to its lookup, so no second shown event arrives.
  h.popupPaneOpen = true;
  h.fire('window', YOMITAN_LOOKUP_EVENT);
  await waitForNextTick();
  assert.deepEqual(h.mpvCommands, [PAUSE]);

  h.popupPaneOpen = false;
  h.popupHostVisible = false;
  h.fire('window', HACHIDORI_POPUP_HIDDEN_EVENT);
  assert.deepEqual(h.mpvCommands, [PAUSE, RESUME]);
});

test('Hachidori hover popup pauses when shown', async (t) => {
  const h = createMouseHarness(t, { popupPause: true, hachidori: true });
  h.handlers.setupYomitanObserver();

  h.popupPaneOpen = true;
  h.popupHostVisible = true;
  h.fire('window', HACHIDORI_POPUP_SHOWN_EVENT);
  await waitForNextTick();

  assert.deepEqual(h.mpvCommands, [PAUSE]);
});

test('popup shown reclaims overlay focus on macOS and captures click-away', (t) => {
  const h = createMouseHarness(t, { platform: 'macos' });
  h.handlers.setupYomitanObserver();

  h.fire('window', YOMITAN_POPUP_SHOWN_EVENT);

  assert.equal(h.ctx.state.yomitanPopupVisible, true);
  assert.equal(h.isInteractive(), true);
  assert.deepEqual(h.ignoreCalls.at(-1), { ignore: false, forward: undefined });
  assert.deepEqual(h.focusCalls, { mainWindow: 1, window: 1, overlay: 1 });
});

test('popup mouse enter and leave on macOS keep click-away captured while popup remains visible', (t) => {
  const h = createMouseHarness(t, { platform: 'macos', popupHostVisible: true });
  h.handlers.setupYomitanObserver();

  h.fire('window', YOMITAN_POPUP_MOUSE_ENTER_EVENT);
  assert.equal(h.ctx.state.yomitanPopupVisible, true);
  assert.equal(h.ctx.state.isOverYomitanPopup, true);
  assert.equal(h.isInteractive(), true);

  h.ignoreCalls.length = 0;
  h.fire('window', YOMITAN_POPUP_MOUSE_LEAVE_EVENT);

  assert.equal(h.ctx.state.yomitanPopupVisible, true);
  assert.equal(h.ctx.state.isOverYomitanPopup, false);
  assert.equal(h.isInteractive(), true);
  assert.deepEqual(h.ignoreCalls.at(-1), { ignore: false, forward: undefined });
});

test('popup hidden on macOS releases click-away capture back to mpv', (t) => {
  const h = createMouseHarness(t, { platform: 'macos', popupHostVisible: true });
  h.handlers.setupYomitanObserver();
  assert.equal(h.isInteractive(), true);

  h.popupHostVisible = false;
  h.fire('window', YOMITAN_POPUP_HIDDEN_EVENT);

  assert.equal(h.ctx.state.yomitanPopupVisible, false);
  assert.equal(h.isInteractive(), false);
  assert.deepEqual(h.ignoreCalls.at(-1), { ignore: true, forward: true });
});

test('yomitan popup visibility marks primary subtitle hover hold while enabled', (t) => {
  const h = createMouseHarness(t);
  h.ctx.state.primaryVisibleOnYomitanPopup = true;
  h.handlers.setupYomitanObserver();

  h.fire('window', YOMITAN_POPUP_SHOWN_EVENT);
  assert.equal(h.bodyClassList.contains('primary-sub-visible-on-yomitan-popup'), true);

  h.fire('window', YOMITAN_POPUP_HIDDEN_EVENT);
  assert.equal(h.bodyClassList.contains('primary-sub-visible-on-yomitan-popup'), false);
});

// Pointer tracking and interaction recovery

test('pointer tracking enables overlay interaction as soon as the cursor reaches subtitles', (t) => {
  const h = createMouseHarness(t);
  h.handlers.setupPointerTracking();

  h.move('primary');

  assert.equal(h.ctx.state.isOverSubtitle, true);
  assert.equal(h.isInteractive(), true);
  assert.deepEqual(h.ignoreCalls.at(-1), { ignore: false, forward: undefined });
});

test('pointer tracking restores click-through after the cursor leaves subtitles', (t) => {
  const h = createMouseHarness(t);
  h.handlers.setupPointerTracking();

  h.move('primary');
  h.move(null, 640, 360);

  assert.equal(h.ctx.state.isOverSubtitle, false);
  assert.equal(h.isInteractive(), false);
  assert.deepEqual(h.ignoreCalls.at(-1), { ignore: true, forward: true });
});

test('restorePointerInteractionState re-enables subtitle hover when pointer is already over subtitles', (t) => {
  const h = createMouseHarness(t);
  h.handlers.setupPointerTracking();
  h.move('primary');

  h.handlers.restorePointerInteractionState();

  assert.equal(h.ctx.state.isOverSubtitle, true);
  assert.equal(h.isInteractive(), true);
  assert.deepEqual(h.ignoreCalls.at(-1), { ignore: false, forward: undefined });
});

test('restorePointerInteractionState keeps overlay interactive until first real pointer move can resync hover', (t) => {
  const h = createMouseHarness(t);
  h.handlers.setupPointerTracking();

  h.handlers.restorePointerInteractionState();

  assert.equal(h.ctx.state.isOverSubtitle, false);
  assert.equal(h.isInteractive(), true);
  assert.deepEqual(h.ignoreCalls.at(-1), { ignore: false, forward: undefined });

  h.move(null, 24, 48);

  assert.equal(h.ctx.state.isOverSubtitle, false);
  assert.equal(h.isInteractive(), false);
  assert.deepEqual(h.ignoreCalls.at(-1), { ignore: true, forward: true });
});

test('restorePointerInteractionState reapplies the secondary hover class from pointer location', async (t) => {
  const h = createMouseHarness(t, { hovered: 'secondary' });
  h.handlers.setupPointerTracking();
  await h.handlers.handleSecondaryMouseEnter({ clientX: 10, clientY: 20 } as MouseEvent);

  h.handlers.restorePointerInteractionState();
  h.move('secondary', 10, 20);

  assert.equal(h.ctx.state.isOverSubtitle, true);
  assert.equal(h.isSecondaryHoverActive(), true);
});

test('visibility recovery re-enables subtitle hover without needing a fresh pointer move', (t) => {
  const h = createMouseHarness(t);
  h.handlers.setupPointerTracking();
  h.move('primary');
  h.ctx.state.isOverSubtitle = false;
  h.ctx.dom.overlay.classList.remove('interactive');

  h.fire('document', 'visibilitychange');

  assert.equal(h.ctx.state.isOverSubtitle, true);
  assert.equal(h.isInteractive(), true);
  assert.deepEqual(h.ignoreCalls.at(-1), { ignore: false, forward: undefined });
});

test('visibility recovery keeps overlay click-through when pointer is not over subtitles', (t) => {
  const h = createMouseHarness(t);
  h.handlers.setupPointerTracking();
  h.move(null, 320, 180);
  h.ctx.dom.overlay.classList.add('interactive');

  h.fire('document', 'visibilitychange');

  assert.equal(h.ctx.state.isOverSubtitle, false);
  assert.equal(h.isInteractive(), false);
  assert.deepEqual(h.ignoreCalls.at(-1), { ignore: true, forward: true });
});

for (const c of [
  { name: 'visibility recovery', target: 'document', event: 'visibilitychange' },
  { name: 'window resize', target: 'window', event: 'resize' },
] as const) {
  test(`${c.name} ignores synthetic subtitle enter until the pointer moves again`, async (t) => {
    const h = createMouseHarness(t);
    h.handlers.setupPointerTracking();
    h.handlers.setupResizeHandler();
    h.move('primary');
    await waitForNextTick();

    h.hoverPause = true;
    h.fire(c.target, c.event);
    await h.handlers.handlePrimaryMouseEnter();
    assert.deepEqual(h.mpvCommands, []);

    h.move(null, 32, 48);
    h.move('primary');
    await waitForNextTick();

    assert.deepEqual(h.mpvCommands, [PAUSE]);
  });
}

test('window resize allows primary hover pause from a real mouseenter over subtitles', async (t) => {
  const h = createMouseHarness(t, { hoverPause: true, hovered: 'primary' });
  h.handlers.setupResizeHandler();
  h.fire('window', 'resize');

  await h.handlers.handlePrimaryMouseEnter({ clientX: 120, clientY: 240 } as MouseEvent);

  assert.deepEqual(h.mpvCommands, [PAUSE]);
});
