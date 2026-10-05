import assert from 'node:assert/strict';
import test from 'node:test';

import type { WindowGeometry } from '../../types';
import { OVERLAY_WINDOW_CONTENT_READY_FLAG } from './overlay-window-flags';
import { updateVisibleOverlayVisibility } from './overlay-visibility';

type VisibilityArgs = Parameters<typeof updateVisibleOverlayVisibility>[0];
type Platform = 'linux' | 'macos' | 'windows';
type MouseMode = 'interactive' | 'passthrough' | 'passthrough-forward';

// Optional fields are omitted from the stub tracker entirely, matching trackers that
// cannot report target focus or minimized state.
type TrackerState = {
  tracking: boolean;
  geometry: WindowGeometry | null;
  targetFocused?: boolean;
  targetMinimized?: boolean;
};

// What a single updateVisibleOverlayVisibility call did. Reset at the start of every run.
type RunEffects = {
  shownWith: 'show' | 'showInactive' | null;
  hid: boolean;
  focusRequested: boolean;
  // Last window-level action: ensureOverlayWindowLevel ('raised') or the internal release.
  level: 'raised' | 'released' | null;
  orderEnforced: boolean;
  layerSynced: boolean;
  shortcutsSynced: boolean;
  zOrderSyncs: number;
  boundsUpdated: boolean;
  // Distinct mouse modes applied, in first-applied order.
  mouseModes: MouseMode[];
};

type Outcome = RunEffects & {
  visible: boolean;
  mouse: MouseMode | null;
  opacity: number;
  fullScreen: boolean | null;
  allWorkspaces: boolean | null;
  trackerWarning: boolean;
  osd: string[];
  dismissals: number;
};

const GEOMETRY: WindowGeometry = { x: 0, y: 0, width: 1280, height: 720 };
const LOADING_OSD = 'Overlay loading...';

const tracked = (extra: Partial<TrackerState> = {}): TrackerState => ({
  tracking: true,
  geometry: GEOMETRY,
  ...extra,
});
const lost = (extra: Partial<TrackerState> = {}): TrackerState => ({
  tracking: false,
  geometry: null,
  ...extra,
});

const emptyEffects = (): RunEffects => ({
  shownWith: null,
  hid: false,
  focusRequested: false,
  level: null,
  orderEnforced: false,
  layerSynced: false,
  shortcutsSynced: false,
  zOrderSyncs: 0,
  boundsUpdated: false,
  mouseModes: [],
});

/**
 * Drives updateVisibleOverlayVisibility against a fake overlay window and default deps.
 * `run(overrides)` returns an Outcome: the run's effects plus resulting window and OSD state.
 */
function createHarness(
  options: {
    platform?: Platform;
    tracker?: TrackerState | null;
    visible?: boolean;
    focused?: boolean;
    contentReady?: boolean;
    trackerWarning?: boolean;
    emitShowImmediately?: boolean;
  } = {},
) {
  const platform = options.platform ?? 'linux';
  const emitShowImmediately = options.emitShowImmediately ?? true;
  const listeners = new Map<string, Array<() => void>>();
  const win = {
    visible: options.visible ?? false,
    focused: options.focused ?? false,
    opacity: 1,
    mouse: null as MouseMode | null,
    fullScreen: null as boolean | null,
    allWorkspaces: null as boolean | null,
  };
  let effects = emptyEffects();
  const h = {
    tracker: options.tracker === undefined ? null : options.tracker,
    trackerWarning: options.trackerWarning ?? false,
    osd: [] as string[],
    dismissals: 0,
    // Every bounds update, with whether the window was already visible when it landed.
    bounds: [] as Array<{ width: number; visible: boolean }>,
  };

  const emitShow = (): void => {
    win.visible = true;
    const handlers = listeners.get('show') ?? [];
    listeners.delete('show');
    for (const handler of handlers) handler();
  };
  const show = (mode: 'show' | 'showInactive'): void => {
    effects.shownWith = mode;
    if (emitShowImmediately) emitShow();
  };
  const fakeWindow: Record<string, unknown> = {
    webContents: {},
    [OVERLAY_WINDOW_CONTENT_READY_FLAG]: options.contentReady ?? true,
    isDestroyed: () => false,
    isVisible: () => win.visible,
    isFocused: () => win.focused,
    once: (event: string, handler: () => void) => {
      listeners.set(event, [...(listeners.get(event) ?? []), handler]);
    },
    hide: () => {
      win.visible = false;
      win.focused = false;
      effects.hid = true;
    },
    show: () => show('show'),
    showInactive: () => show('showInactive'),
    focus: () => {
      win.focused = true;
      effects.focusRequested = true;
    },
    setAlwaysOnTop: (flag: boolean) => {
      if (!flag) effects.level = 'released';
    },
    setFullScreen: (flag: boolean) => {
      win.fullScreen = flag;
    },
    setVisibleOnAllWorkspaces: (flag: boolean) => {
      win.allWorkspaces = flag;
    },
    setIgnoreMouseEvents: (ignore: boolean, opts?: { forward?: boolean }) => {
      const mode: MouseMode = !ignore
        ? 'interactive'
        : opts?.forward
          ? 'passthrough-forward'
          : 'passthrough';
      win.mouse = mode;
      if (!effects.mouseModes.includes(mode)) effects.mouseModes.push(mode);
    },
    setOpacity: (opacity: number) => {
      win.opacity = opacity;
    },
  };

  const buildTracker = (state: TrackerState | null) =>
    state === null
      ? null
      : ({
          isTracking: () => state.tracking,
          getGeometry: () => state.geometry,
          ...(state.targetFocused === undefined
            ? {}
            : { isTargetWindowFocused: () => state.targetFocused }),
          ...(state.targetMinimized === undefined
            ? {}
            : { isTargetWindowMinimized: () => state.targetMinimized }),
        } as unknown as VisibilityArgs['windowTracker']);

  const outcome = (): Outcome => ({
    ...effects,
    mouseModes: [...effects.mouseModes],
    visible: win.visible,
    mouse: win.mouse,
    opacity: win.opacity,
    fullScreen: win.fullScreen,
    allWorkspaces: win.allWorkspaces,
    trackerWarning: h.trackerWarning,
    osd: [...h.osd],
    dismissals: h.dismissals,
  });

  const run = (overrides: Partial<VisibilityArgs> = {}): Outcome => {
    effects = emptyEffects();
    updateVisibleOverlayVisibility({
      visibleOverlayVisible: true,
      mainWindow: fakeWindow as unknown as VisibilityArgs['mainWindow'],
      windowTracker: buildTracker(h.tracker),
      trackerNotReadyWarningShown: h.trackerWarning,
      setTrackerNotReadyWarningShown: (shown) => {
        h.trackerWarning = shown;
      },
      updateVisibleOverlayBounds: (geometry) => {
        effects.boundsUpdated = true;
        h.bounds.push({ width: geometry.width, visible: win.visible });
      },
      ensureOverlayWindowLevel: () => {
        effects.level = 'raised';
      },
      syncWindowsOverlayToMpvZOrder: () => {
        effects.zOrderSyncs += 1;
      },
      syncPrimaryOverlayWindowLayer: () => {
        effects.layerSynced = true;
      },
      enforceOverlayLayerOrder: () => {
        effects.orderEnforced = true;
      },
      syncOverlayShortcuts: () => {
        effects.shortcutsSynced = true;
      },
      showOverlayLoadingOsd: (message) => {
        h.osd.push(message);
      },
      dismissOverlayLoadingOsd: () => {
        h.dismissals += 1;
      },
      isMacOSPlatform: platform === 'macos',
      isWindowsPlatform: platform === 'windows',
      ...overrides,
    });
    return outcome();
  };

  return {
    state: h,
    run,
    outcome,
    emitShow,
    setFocused: (focused: boolean) => {
      win.focused = focused;
    },
    setContentReady: (ready: boolean) => {
      fakeWindow[OVERLAY_WINDOW_CONTENT_READY_FLAG] = ready;
    },
  };
}

// Compares only the keys a scenario cares about.
function assertOutcome(actual: Outcome, expected: Partial<Outcome>): void {
  const picked: Partial<Outcome> = {};
  for (const key of Object.keys(expected) as Array<keyof Outcome>) {
    Object.assign(picked, { [key]: actual[key] });
  }
  assert.deepEqual(picked, expected);
}

// Shared "suppress repeat loading OSD for 5s" policy used by the OSD cooldown tests.
function loadingOsdCooldown(now: () => number) {
  let lastShownAtMs: number | null = null;
  return {
    shouldShowOverlayLoadingOsd: () => lastShownAtMs === null || now() - lastShownAtMs >= 5_000,
    markOverlayLoadingOsdShown: () => {
      lastShownAtMs = now();
    },
    resetOverlayLoadingOsdSuppression: () => {
      lastShownAtMs = null;
    },
  };
}

type Scenario = {
  name: string;
  platform: Platform;
  tracker: TrackerState | null;
  visible?: boolean;
  focused?: boolean;
  trackerWarning?: boolean;
  overrides?: Partial<VisibilityArgs>;
  expect: Partial<Outcome>;
};

function runScenarios(scenarios: Scenario[]): void {
  for (const c of scenarios) {
    test(c.name, () => {
      const h = createHarness({
        platform: c.platform,
        tracker: c.tracker,
        visible: c.visible,
        focused: c.focused,
        trackerWarning: c.trackerWarning,
      });
      assertOutcome(h.run(c.overrides), c.expect);
    });
  }
}

const hiddenNotShown = {
  visible: false,
  shownWith: null,
  boundsUpdated: false,
  focusRequested: false,
} satisfies Partial<Outcome>;

// Tracker not ready: the overlay stays hidden. Native platforms warn once with a loading OSD;
// untracked Linux has nothing to wait for, so it stays quiet.
for (const c of [
  { name: 'macOS tracker not ready', platform: 'macos', tracker: lost(), warns: true },
  { name: 'macOS tracker not initialized', platform: 'macos', tracker: null, warns: true },
  { name: 'Windows tracker not ready', platform: 'windows', tracker: lost(), warns: true },
  { name: 'Linux tracker not ready', platform: 'linux', tracker: lost(), warns: true },
  { name: 'Linux without a tracker', platform: 'linux', tracker: null, warns: false },
] satisfies Array<{
  name: string;
  platform: Platform;
  tracker: TrackerState | null;
  warns: boolean;
}>) {
  test(`${c.name}: overlay stays hidden${c.warns ? ' and emits one loading OSD' : ' without a loading OSD'}`, () => {
    const h = createHarness({ platform: c.platform, tracker: c.tracker });
    h.run();
    assertOutcome(h.run(), {
      ...hiddenNotShown,
      trackerWarning: c.warns,
      osd: c.warns ? [LOADING_OSD] : [],
    });
  });
}

test('untracked Linux overlay is released to click-through while hidden', () => {
  const h = createHarness({ platform: 'linux', tracker: null, visible: true });
  assertOutcome(h.run(), {
    hid: true,
    mouse: 'passthrough-forward',
    level: 'released',
    shortcutsSynced: true,
  });
});

test('macOS dismisses overlay loading OSD when tracker recovers', () => {
  const h = createHarness({ platform: 'macos', tracker: lost({ targetFocused: false }) });
  h.run();
  h.state.tracker = tracked({ targetFocused: true });

  assertOutcome(h.run(), {
    osd: [LOADING_OSD],
    dismissals: 1,
    trackerWarning: false,
    shownWith: 'showInactive',
  });
});

for (const platform of ['linux', 'windows'] satisfies Platform[]) {
  test(`${platform} tracked overlay shows loading OSD and waits for renderer content before first reveal`, () => {
    const h = createHarness({
      platform,
      tracker: tracked({ targetFocused: true }),
      contentReady: false,
    });
    h.run();
    assertOutcome(h.run(), {
      visible: false,
      shownWith: null,
      trackerWarning: true,
      osd: [LOADING_OSD],
      dismissals: 0,
    });

    h.setContentReady(true);
    assertOutcome(h.run(), {
      visible: true,
      shownWith: 'showInactive',
      trackerWarning: false,
      dismissals: 1,
    });
  });
}

runScenarios([
  {
    name: 'visible overlay stays hidden while a modal window is active',
    platform: 'macos',
    tracker: tracked(),
    overrides: { modalActive: true },
    expect: { hid: true, shownWith: null, boundsUpdated: false },
  },
  {
    name: 'suspended visible overlay hides without refreshing bounds or z-order',
    platform: 'linux',
    tracker: tracked(),
    visible: true,
    overrides: { suspendVisibleOverlay: true },
    expect: {
      visible: false,
      hid: true,
      mouse: 'passthrough-forward',
      level: 'released',
      shortcutsSynced: true,
      boundsUpdated: false,
      orderEnforced: false,
      shownWith: null,
      focusRequested: false,
    },
  },
]);

test('passive Linux overlay shows without focus and stays click-through on later updates', () => {
  const h = createHarness({ platform: 'linux', tracker: tracked() });
  assertOutcome(h.run(), {
    shownWith: 'showInactive',
    focusRequested: false,
    mouse: 'passthrough-forward',
  });
  assertOutcome(h.run(), { mouseModes: ['passthrough-forward'], shownWith: null });
});

runScenarios([
  {
    name: 'non-native shaped input region stays mouse-enabled without focusing the overlay',
    platform: 'linux',
    tracker: tracked({ targetFocused: true }),
    overrides: { nonNativeInputRegionActive: true },
    expect: { mouseModes: ['interactive'], shownWith: 'showInactive', focusRequested: false },
  },
  {
    name: 'passive Linux tracked overlay releases global topmost when mpv loses focus',
    platform: 'linux',
    tracker: tracked({ targetFocused: false }),
    visible: true,
    expect: {
      visible: true,
      hid: false,
      mouse: 'passthrough-forward',
      level: 'released',
      fullScreen: false,
      allWorkspaces: false,
      orderEnforced: false,
    },
  },
  {
    name: 'passive Linux fullscreen override overlay hides when mpv loses focus',
    platform: 'linux',
    tracker: tracked({ targetFocused: false }),
    visible: true,
    overrides: { hideNonNativeOverlayWhenTargetUnfocused: true },
    expect: { hid: true, mouse: 'passthrough-forward', level: 'released', orderEnforced: false },
  },
  {
    name: 'Linux active overlay interaction does not focus the overlay over fullscreen mpv',
    platform: 'linux',
    tracker: tracked({ targetFocused: true }),
    visible: true,
    overrides: { overlayInteractionActive: true },
    expect: { mouse: 'interactive', level: 'raised', orderEnforced: true, focusRequested: false },
  },
  {
    name: 'Linux active hover keeps global topmost when mpv loses focus and overlay is not focused',
    platform: 'linux',
    tracker: tracked({ targetFocused: false }),
    visible: true,
    overrides: { overlayInteractionActive: true },
    expect: { mouse: 'interactive', level: 'raised', orderEnforced: true },
  },
]);

test('tracked non-macOS overlay reapplies bounds after first show', () => {
  const h = createHarness({ platform: 'linux', tracker: tracked() });
  h.run();
  assert.deepEqual(h.state.bounds, [
    { width: 1280, visible: false },
    { width: 1280, visible: true },
  ]);
});

test('tracked non-macOS overlay queues only one first-show bounds refresh with the latest geometry', () => {
  const h = createHarness({ platform: 'linux', tracker: tracked(), emitShowImmediately: false });
  h.run();
  h.state.tracker = tracked({ geometry: { ...GEOMETRY, width: 1440 } });
  h.run();
  h.emitShow();

  assert.deepEqual(h.state.bounds, [
    { width: 1280, visible: false },
    { width: 1440, visible: false },
    { width: 1440, visible: true },
  ]);
});

// Windows binds the overlay to mpv via the owner window, so it never raises or enforces
// layer order itself; it stays plain click-through (no forwarding) unless the overlay is focused.
runScenarios([
  {
    name: 'Windows first show starts transparent and click-through, bound to mpv',
    platform: 'windows',
    tracker: tracked(),
    expect: {
      opacity: 0,
      shownWith: 'showInactive',
      mouse: 'passthrough',
      zOrderSyncs: 1,
      level: null,
      orderEnforced: false,
      focusRequested: false,
    },
  },
  {
    name: 'tracked Windows overlay refresh rebinds while already visible',
    platform: 'windows',
    tracker: tracked(),
    visible: true,
    expect: {
      shownWith: null,
      mouse: 'passthrough',
      zOrderSyncs: 1,
      level: null,
      shortcutsSynced: true,
    },
  },
  {
    name: 'forced passthrough still reapplies while visible on Windows',
    platform: 'windows',
    tracker: tracked(),
    visible: true,
    overrides: { forceMousePassthrough: true },
    expect: { mouse: 'passthrough', zOrderSyncs: 1, level: null, orderEnforced: false },
  },
  {
    name: 'forced passthrough still shows tracked overlay while bound to mpv on Windows',
    platform: 'windows',
    tracker: tracked(),
    overrides: { forceMousePassthrough: true },
    expect: { shownWith: 'showInactive', zOrderSyncs: 1, level: null },
  },
  {
    name: 'tracked Windows overlay rebinds without hiding when tracker focus changes',
    platform: 'windows',
    tracker: tracked({ targetFocused: false }),
    visible: true,
    expect: {
      visible: true,
      hid: false,
      shownWith: null,
      mouse: 'passthrough',
      zOrderSyncs: 1,
      level: null,
      orderEnforced: false,
    },
  },
  {
    name: 'tracked Windows overlay stays interactive while the overlay window itself is focused',
    platform: 'windows',
    tracker: tracked({ targetFocused: false }),
    visible: true,
    focused: true,
    expect: { mouse: 'interactive', zOrderSyncs: 1, level: null, orderEnforced: false },
  },
  {
    name: 'tracked Windows overlay reshows click-through even if focus state is stale after a modal closes',
    platform: 'windows',
    tracker: tracked({ targetFocused: false }),
    visible: false,
    focused: true,
    expect: { mouse: 'passthrough', shownWith: 'showInactive' },
  },
  {
    name: 'tracked Windows overlay binds above mpv even when tracker focus lags',
    platform: 'windows',
    tracker: tracked({ targetFocused: false }),
    expect: { mouse: 'passthrough', zOrderSyncs: 1, level: null },
  },
  {
    name: 'Windows preserves visible overlay and rebinds to mpv while tracker transiently loses a non-minimized window',
    platform: 'windows',
    tracker: tracked({ tracking: false, targetFocused: false, targetMinimized: false }),
    visible: true,
    expect: {
      hid: false,
      shownWith: null,
      mouse: 'passthrough',
      zOrderSyncs: 1,
      level: null,
      shortcutsSynced: true,
    },
  },
  {
    name: 'Windows hides the visible overlay when the tracked window is minimized',
    platform: 'windows',
    tracker: lost({ targetMinimized: true }),
    visible: true,
    expect: { visible: false, hid: true, zOrderSyncs: 0 },
  },
]);

test('Windows visible overlay restores opacity and rebinds after the deferred reveal delay', async () => {
  const h = createHarness({ platform: 'windows', tracker: tracked() });
  assertOutcome(h.run(), { opacity: 0, zOrderSyncs: 1 });

  await new Promise<void>((resolve) => setTimeout(resolve, 60));

  assertOutcome(h.outcome(), { opacity: 1, zOrderSyncs: 2 });
});

// macOS overlays stay click-through with mouse-move forwarding unless the renderer reports
// interaction, and release their window level (and hide) once mpv loses the foreground.
const macPassiveShown = {
  boundsUpdated: true,
  layerSynced: true,
  mouse: 'passthrough-forward',
  level: 'raised',
  orderEnforced: true,
  shortcutsSynced: true,
  hid: false,
} satisfies Partial<Outcome>;

const macReleasedHidden = {
  layerSynced: true,
  mouse: 'passthrough-forward',
  level: 'released',
  allWorkspaces: false,
  hid: true,
  shortcutsSynced: true,
  orderEnforced: false,
  focusRequested: false,
  shownWith: null,
} satisfies Partial<Outcome>;

runScenarios([
  {
    name: 'macOS tracked visible overlay starts click-through without passively stealing focus',
    platform: 'macos',
    tracker: tracked(),
    expect: { mouse: 'passthrough-forward', shownWith: 'showInactive', focusRequested: false },
  },
  {
    name: 'forced mouse passthrough keeps macOS tracked overlay passive while visible',
    platform: 'macos',
    tracker: tracked(),
    overrides: { forceMousePassthrough: true },
    expect: { mouse: 'passthrough-forward', shownWith: 'showInactive', focusRequested: false },
  },
  {
    name: 'macOS tracked visible overlay remains click-through even if the overlay had focus',
    platform: 'macos',
    tracker: tracked({ targetFocused: true }),
    focused: true,
    expect: { mouse: 'passthrough-forward', level: 'raised', focusRequested: false },
  },
  {
    name: 'macOS keeps active mpv overlay visible and click-through during tracker refresh',
    platform: 'macos',
    tracker: tracked({ targetFocused: true }),
    expect: { ...macPassiveShown, osd: [] },
  },
  {
    name: 'forced mouse passthrough keeps macOS tracked overlay above active mpv',
    platform: 'macos',
    tracker: tracked({ targetFocused: true }),
    overrides: { forceMousePassthrough: true },
    expect: { mouse: 'passthrough-forward', level: 'raised', orderEnforced: true },
  },
  {
    name: 'macOS tracked overlay hides when mpv loses foreground',
    platform: 'macos',
    tracker: tracked({ targetFocused: false }),
    visible: true,
    expect: { ...macReleasedHidden, boundsUpdated: true },
  },
  {
    name: 'forced mouse passthrough still hides macOS tracked overlay when mpv loses foreground',
    platform: 'macos',
    tracker: tracked({ targetFocused: false }),
    visible: true,
    overrides: { forceMousePassthrough: true },
    expect: macReleasedHidden,
  },
  {
    name: 'macOS keeps visible overlay stable while probing frontmost app after overlay blur',
    platform: 'macos',
    tracker: tracked({ targetFocused: false, targetMinimized: false }),
    visible: true,
    overrides: { macOSForegroundProbeActive: true },
    expect: macPassiveShown,
  },
  {
    // Interaction keeps the overlay up, lets it take mouse input, and focuses it so lookup
    // trigger keys reach the renderer.
    name: 'macOS active overlay stays visible, receives mouse input, and takes focus after mpv loses foreground',
    platform: 'macos',
    tracker: tracked({ targetFocused: false }),
    visible: true,
    overrides: { overlayInteractionActive: true },
    expect: {
      ...macPassiveShown,
      mouse: 'interactive',
      mouseModes: ['interactive'],
      focusRequested: true,
    },
  },
]);

test('macOS tracked overlay passively reappears when mpv regains foreground', () => {
  const h = createHarness({
    platform: 'macos',
    tracker: tracked({ targetFocused: false }),
    visible: true,
  });
  assertOutcome(h.run(), { hid: true, visible: false });

  h.state.tracker = tracked({ targetFocused: true });
  assertOutcome(h.run(), {
    mouse: 'passthrough-forward',
    level: 'raised',
    shownWith: 'showInactive',
    orderEnforced: true,
    focusRequested: false,
  });
});

// macOS tracker loss (isTracking false) on an already visible overlay: keep it up when there is
// retained geometry or an active focus signal, hide it once mpv is known to be in the background.
const macTrackerLossKept = {
  trackerWarning: false,
  osd: [],
  layerSynced: true,
  mouse: 'passthrough-forward',
  level: 'raised',
  orderEnforced: true,
  shortcutsSynced: true,
  hid: false,
} satisfies Partial<Outcome>;

runScenarios([
  {
    name: 'macOS preserves an already visible active mpv overlay while tracker is temporarily not ready',
    platform: 'macos',
    tracker: lost({ targetFocused: true }),
    visible: true,
    trackerWarning: true,
    expect: macTrackerLossKept,
  },
  {
    name: 'macOS preserves visible overlay during transient tracker loss with retained geometry',
    platform: 'macos',
    tracker: tracked({ tracking: false, targetFocused: true }),
    visible: true,
    trackerWarning: true,
    expect: { ...macTrackerLossKept, boundsUpdated: true, shownWith: null },
  },
  {
    name: 'macOS hides visible overlay during tracker loss after mpv loses foreground',
    platform: 'macos',
    tracker: lost({ targetFocused: false, targetMinimized: false }),
    visible: true,
    expect: { ...macReleasedHidden, osd: [] },
  },
  {
    name: 'macOS keeps a focused overlay visible during tracker loss',
    platform: 'macos',
    tracker: lost({ targetFocused: false, targetMinimized: false }),
    visible: true,
    focused: true,
    expect: macTrackerLossKept,
  },
  {
    name: 'macOS keeps an interactive overlay visible during tracker loss even when Electron focus drops',
    platform: 'macos',
    tracker: lost({ targetFocused: false, targetMinimized: false }),
    visible: true,
    overrides: { overlayInteractionActive: true },
    expect: { ...macTrackerLossKept, mouse: 'interactive' },
  },
]);

test('macOS suppresses immediate repeat loading OSD after tracker recovery until cooldown expires', () => {
  let nowMs = 1_000;
  const cooldown = loadingOsdCooldown(() => nowMs);
  const h = createHarness({ platform: 'macos' });
  const runWith = (tracker: TrackerState) => {
    h.state.tracker = tracker;
    h.run(cooldown);
  };

  runWith(lost());
  runWith(tracked());
  nowMs = 2_000;
  runWith(lost());
  runWith(tracked());
  nowMs = 6_500;
  runWith(lost());

  assert.deepEqual(h.state.osd, [LOADING_OSD, LOADING_OSD]);
});

test('macOS explicit hide resets loading OSD suppression before retry', () => {
  let nowMs = 1_000;
  const cooldown = loadingOsdCooldown(() => nowMs);
  const h = createHarness({ platform: 'macos', tracker: null });

  h.run(cooldown);
  nowMs = 1_500;
  assertOutcome(h.run({ ...cooldown, visibleOverlayVisible: false }), {
    trackerWarning: false,
    dismissals: 1,
    hid: true,
  });
  h.run(cooldown);

  assert.deepEqual(h.state.osd, [LOADING_OSD, LOADING_OSD]);
});
