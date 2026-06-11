import assert from 'node:assert/strict';
import test from 'node:test';
import {
  applyLinuxOverlayInputShape,
  applyLinuxOverlayPointerInteractionMousePassthrough,
  type LinuxOverlayPointerInteractionDeps,
  isCursorOverSubtitle,
  type ForegroundSuppressionGraceState,
  mapOverlayMeasurementForPointerInteraction,
  resolveDesiredOverlayInteractive,
  resolveForegroundSuppressionWithGrace,
  shouldSuppressPointerInteractionForForegroundWindow,
  shouldPrimeLinuxOverlayInteractionFromMeasurement,
  tickLinuxOverlayPointerInteraction,
} from './linux-overlay-pointer-interaction';

const BOUNDS = { x: 100, y: 100, width: 1920, height: 1080 };
const MEASUREMENT = {
  viewport: { width: 1920, height: 1080 },
  contentRect: { x: 800, y: 900, width: 320, height: 80 },
};

test('isCursorOverSubtitle hit-tests the subtitle rect in screen coords (1:1 scale)', () => {
  // Subtitle rect maps to screen [900..1220] x [1000..1080] (+100 window origin).
  assert.equal(isCursorOverSubtitle({ x: 1000, y: 1040 }, BOUNDS, MEASUREMENT), true);
  assert.equal(isCursorOverSubtitle({ x: 500, y: 1040 }, BOUNDS, MEASUREMENT), false);
  assert.equal(isCursorOverSubtitle({ x: 1000, y: 500 }, BOUNDS, MEASUREMENT), false);
});

test('isCursorOverSubtitle scales viewport px to window px', () => {
  // Window is 2x the reported viewport → rect doubles.
  const scaled = { ...BOUNDS, width: 3840, height: 2160 };
  // contentRect.x*2=1600 +100 origin → left ~1700; a point at 1700,1900 is inside.
  assert.equal(isCursorOverSubtitle({ x: 1700, y: 1900 }, scaled, MEASUREMENT), true);
});

test('isCursorOverSubtitle returns false without a content rect', () => {
  assert.equal(
    isCursorOverSubtitle({ x: 1000, y: 1040 }, BOUNDS, {
      viewport: MEASUREMENT.viewport,
      contentRect: null,
    }),
    false,
  );
  assert.equal(isCursorOverSubtitle({ x: 1000, y: 1040 }, BOUNDS, null), false);
});

test('isCursorOverSubtitle falls back to content rect when interactive rects are empty', () => {
  assert.equal(
    isCursorOverSubtitle({ x: 1000, y: 1040 }, BOUNDS, {
      ...MEASUREMENT,
      interactiveRects: [],
    }),
    true,
  );
});

function makeDeps(overrides: Partial<LinuxOverlayPointerInteractionDeps>): {
  deps: LinuxOverlayPointerInteractionDeps;
  state: { active: boolean };
} {
  const state = { active: false };
  const deps: LinuxOverlayPointerInteractionDeps = {
    getVisibleOverlayVisible: () => true,
    getMainWindow: () => ({
      isDestroyed: () => false,
      isVisible: () => true,
      getBounds: () => BOUNDS,
    }),
    getCursorScreenPoint: () => ({ x: 1000, y: 1040 }),
    getSubtitleMeasurement: () => MEASUREMENT,
    getRendererInteractiveHint: () => false,
    shouldSuspend: () => false,
    getInteractionActive: () => state.active,
    setInteractionActive: (active) => {
      state.active = active;
    },
    ...overrides,
  };
  return { deps, state };
}

test('resolveDesiredOverlayInteractive: interactive over subtitle, passthrough off it', () => {
  assert.equal(resolveDesiredOverlayInteractive(makeDeps({}).deps), true);
  assert.equal(
    resolveDesiredOverlayInteractive(
      makeDeps({ getCursorScreenPoint: () => ({ x: 200, y: 200 }) }).deps,
    ),
    false,
  );
});

test('resolveDesiredOverlayInteractive: renderer hint keeps it interactive off the rect', () => {
  const { deps } = makeDeps({
    getCursorScreenPoint: () => ({ x: 200, y: 200 }),
    getRendererInteractiveHint: () => true,
  });
  assert.equal(resolveDesiredOverlayInteractive(deps), true);
});

test('resolveDesiredOverlayInteractive: hit-tests separate subtitle bars without blocking between them', () => {
  const measurement = {
    viewport: { width: 1920, height: 1080 },
    contentRect: { x: 700, y: 40, width: 520, height: 940 },
    interactiveRects: [
      { x: 700, y: 40, width: 520, height: 80 },
      { x: 760, y: 900, width: 400, height: 80 },
    ],
  } as unknown as ReturnType<LinuxOverlayPointerInteractionDeps['getSubtitleMeasurement']>;

  assert.equal(
    resolveDesiredOverlayInteractive(
      makeDeps({
        getCursorScreenPoint: () => ({ x: 900, y: 300 }),
        getSubtitleMeasurement: () => measurement,
      }).deps,
    ),
    false,
  );
  assert.equal(
    resolveDesiredOverlayInteractive(
      makeDeps({
        getCursorScreenPoint: () => ({ x: 900, y: 1060 }),
        getSubtitleMeasurement: () => measurement,
      }).deps,
    ),
    true,
  );
  assert.equal(
    resolveDesiredOverlayInteractive(
      makeDeps({
        getCursorScreenPoint: () => ({ x: 900, y: 180 }),
        getSubtitleMeasurement: () => measurement,
      }).deps,
    ),
    true,
  );
});

test('shouldPrimeLinuxOverlayInteractionFromMeasurement primes input from first measured rect', () => {
  assert.equal(
    shouldPrimeLinuxOverlayInteractionFromMeasurement({
      getVisibleOverlayVisible: () => true,
      getMainWindow: () => ({
        isDestroyed: () => false,
        isVisible: () => true,
        getBounds: () => BOUNDS,
      }),
      getSubtitleMeasurement: () => ({
        ...MEASUREMENT,
        interactiveRects: [{ x: 900, y: 900, width: 320, height: 80 }],
      }),
      shouldSuspend: () => false,
      shouldSuppressInteraction: () => false,
    }),
    true,
  );
});

test('shouldPrimeLinuxOverlayInteractionFromMeasurement skips hidden or empty startup surfaces', () => {
  assert.equal(
    shouldPrimeLinuxOverlayInteractionFromMeasurement({
      getVisibleOverlayVisible: () => true,
      getMainWindow: () => ({
        isDestroyed: () => false,
        isVisible: () => false,
        getBounds: () => BOUNDS,
      }),
      getSubtitleMeasurement: () => MEASUREMENT,
      shouldSuspend: () => false,
    }),
    false,
  );
  assert.equal(
    shouldPrimeLinuxOverlayInteractionFromMeasurement({
      getVisibleOverlayVisible: () => true,
      getMainWindow: () => ({
        isDestroyed: () => false,
        isVisible: () => true,
        getBounds: () => BOUNDS,
      }),
      getSubtitleMeasurement: () => ({
        viewport: MEASUREMENT.viewport,
        contentRect: null,
        interactiveRects: [],
      }),
      shouldSuspend: () => false,
    }),
    false,
  );
});

test('mapOverlayMeasurementForPointerInteraction preserves renderer interactive rects', () => {
  const mapped = mapOverlayMeasurementForPointerInteraction({
    layer: 'visible',
    measuredAtMs: 1,
    viewport: { width: 1920, height: 1080 },
    contentRect: { x: 700, y: 40, width: 520, height: 940 },
    interactiveRects: [
      { x: 700, y: 40, width: 520, height: 80 },
      { x: 760, y: 900, width: 400, height: 80 },
    ],
  });

  assert.deepEqual(mapped, {
    viewport: { width: 1920, height: 1080 },
    contentRect: { x: 700, y: 40, width: 520, height: 940 },
    interactiveRects: [
      { x: 700, y: 40, width: 520, height: 80 },
      { x: 760, y: 900, width: 400, height: 80 },
    ],
  });
});

test('shouldSuppressPointerInteractionForForegroundWindow suppresses hover when another app is foreground', () => {
  assert.equal(
    shouldSuppressPointerInteractionForForegroundWindow({
      hasForegroundSeparateWindow: false,
      isTrackingMpvWindow: true,
      isMpvWindowFocused: false,
      isOverlayWindowFocused: false,
    }),
    true,
  );
  assert.equal(
    shouldSuppressPointerInteractionForForegroundWindow({
      hasForegroundSeparateWindow: false,
      isTrackingMpvWindow: true,
      isMpvWindowFocused: true,
      isOverlayWindowFocused: false,
    }),
    false,
  );
  assert.equal(
    shouldSuppressPointerInteractionForForegroundWindow({
      hasForegroundSeparateWindow: false,
      isTrackingMpvWindow: true,
      isMpvWindowFocused: false,
      isOverlayWindowFocused: true,
    }),
    false,
  );
});

test('resolveForegroundSuppressionWithGrace ignores a transient startup focus blip', () => {
  // Regression: right after playback starts the overlay can briefly become the X11 active
  // window, so the tracker reports mpv unfocused. Suppressing immediately leaves subtitles
  // inert for ~1s. The grace must hold interaction available until the loss is *stable*.
  const state: ForegroundSuppressionGraceState = { lossSinceMs: null };
  const base = {
    hasForegroundSeparateWindow: false,
    isTrackingMpvWindow: true,
    isMpvWindowFocused: false,
    isOverlayWindowFocused: false,
    graceMs: 500,
    state,
  };

  // Blip starts: not yet suppressed.
  assert.equal(resolveForegroundSuppressionWithGrace({ ...base, nowMs: 1_000 }), false);
  // Still within grace.
  assert.equal(resolveForegroundSuppressionWithGrace({ ...base, nowMs: 1_400 }), false);
  // mpv regains focus before the grace elapses → reset, never suppressed.
  assert.equal(
    resolveForegroundSuppressionWithGrace({ ...base, isMpvWindowFocused: true, nowMs: 1_450 }),
    false,
  );
  assert.equal(state.lossSinceMs, null);
});

test('resolveForegroundSuppressionWithGrace suppresses once foreground loss is stable', () => {
  const state: ForegroundSuppressionGraceState = { lossSinceMs: null };
  const base = {
    hasForegroundSeparateWindow: false,
    isTrackingMpvWindow: true,
    isMpvWindowFocused: false,
    isOverlayWindowFocused: false,
    graceMs: 500,
    state,
  };

  assert.equal(resolveForegroundSuppressionWithGrace({ ...base, nowMs: 1_000 }), false);
  // A real app stays foreground past the grace → suppress.
  assert.equal(resolveForegroundSuppressionWithGrace({ ...base, nowMs: 1_500 }), true);
});

test('resolveForegroundSuppressionWithGrace defers to a separate window immediately', () => {
  const state: ForegroundSuppressionGraceState = { lossSinceMs: 1_000 };
  assert.equal(
    resolveForegroundSuppressionWithGrace({
      hasForegroundSeparateWindow: true,
      isTrackingMpvWindow: true,
      isMpvWindowFocused: true,
      isOverlayWindowFocused: false,
      nowMs: 2_000,
      graceMs: 500,
      state,
    }),
    true,
  );
  assert.equal(state.lossSinceMs, null);
});

test('shouldSuppressPointerInteractionForForegroundWindow suppresses hover for separate app windows', () => {
  assert.equal(
    shouldSuppressPointerInteractionForForegroundWindow({
      hasForegroundSeparateWindow: true,
      isTrackingMpvWindow: true,
      isMpvWindowFocused: true,
      isOverlayWindowFocused: false,
    }),
    true,
  );
});

test('resolveDesiredOverlayInteractive: false when overlay hidden, null when suspended/no window', () => {
  assert.equal(
    resolveDesiredOverlayInteractive(makeDeps({ getVisibleOverlayVisible: () => false }).deps),
    false,
  );
  assert.equal(
    resolveDesiredOverlayInteractive(makeDeps({ shouldSuspend: () => true }).deps),
    null,
  );
  assert.equal(
    resolveDesiredOverlayInteractive(makeDeps({ getMainWindow: () => null }).deps),
    null,
  );
});

test('tick only writes interaction state on change', () => {
  const calls: boolean[] = [];
  const { deps, state } = makeDeps({
    setInteractionActive: (active) => {
      calls.push(active);
      state.active = active;
    },
  });
  tickLinuxOverlayPointerInteraction(deps); // off→on
  tickLinuxOverlayPointerInteraction(deps); // no change
  assert.deepEqual(calls, [true]);
});

test('tick does not flip state when suspended (returns null)', () => {
  const calls: boolean[] = [];
  const { deps } = makeDeps({
    getInteractionActive: () => true,
    shouldSuspend: () => true,
    setInteractionActive: (active) => calls.push(active),
  });
  tickLinuxOverlayPointerInteraction(deps);
  assert.deepEqual(calls, []);
});

test('tick clears active hover while a separate SubMiner window suppresses overlay interaction', () => {
  const calls: boolean[] = [];
  const { deps, state } = makeDeps({
    getInteractionActive: () => true,
    shouldSuppressInteraction: () => true,
    setInteractionActive: (active) => {
      calls.push(active);
      state.active = active;
    },
  });

  state.active = true;
  tickLinuxOverlayPointerInteraction(deps);
  assert.deepEqual(calls, [false]);
});

test('tick skips cursor-driven mouse-ignore toggles when Linux input shape owns hit rects', () => {
  const calls: boolean[] = [];
  const { deps } = makeDeps({
    getInteractionActive: () => false,
    shouldUseInputShape: () => true,
    setInteractionActive: (active) => calls.push(active),
  });

  tickLinuxOverlayPointerInteraction(deps);
  assert.deepEqual(calls, []);
});

test('applyLinuxOverlayInputShape shapes measured subtitle rects and enables mouse input', () => {
  const calls: string[] = [];
  const window = {
    isDestroyed: () => false,
    isVisible: () => true,
    getBounds: () => ({ ...BOUNDS, width: 3840, height: 2160 }),
    setShape: (rects: Array<{ x: number; y: number; width: number; height: number }>) => {
      calls.push(`shape:${JSON.stringify(rects)}`);
    },
    setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => {
      calls.push(`ignore:${ignore}:${options?.forward === true ? 'forward' : 'plain'}`);
    },
  };

  assert.deepEqual(
    applyLinuxOverlayInputShape({
      getVisibleOverlayVisible: () => true,
      getMainWindow: () => window,
      getSubtitleMeasurement: () => MEASUREMENT,
      getRendererInteractiveHint: () => false,
      shouldSuspend: () => false,
      shouldSuppressInteraction: () => false,
    }),
    { handled: true, active: true },
  );
  assert.deepEqual(calls, [
    'shape:[{"x":1594,"y":1794,"width":652,"height":172}]',
    'ignore:false:plain',
  ]);
});

test('applyLinuxOverlayInputShape uses the full window while renderer reports off-rect interaction', () => {
  const calls: string[] = [];

  assert.deepEqual(
    applyLinuxOverlayInputShape({
      getVisibleOverlayVisible: () => true,
      getMainWindow: () => ({
        isDestroyed: () => false,
        isVisible: () => true,
        getBounds: () => BOUNDS,
        setShape: (rects: Array<{ x: number; y: number; width: number; height: number }>) => {
          calls.push(`shape:${JSON.stringify(rects)}`);
        },
        setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => {
          calls.push(`ignore:${ignore}:${options?.forward === true ? 'forward' : 'plain'}`);
        },
      }),
      getSubtitleMeasurement: () => null,
      getRendererInteractiveHint: () => true,
      shouldSuspend: () => false,
    }),
    { handled: true, active: true },
  );
  assert.deepEqual(calls, [
    'shape:[{"x":0,"y":0,"width":1920,"height":1080}]',
    'ignore:false:plain',
  ]);
});

test('applyLinuxOverlayInputShape falls back when setShape is unavailable', () => {
  const calls: string[] = [];

  assert.deepEqual(
    applyLinuxOverlayInputShape({
      getVisibleOverlayVisible: () => true,
      getMainWindow: () => ({
        isDestroyed: () => false,
        isVisible: () => true,
        getBounds: () => BOUNDS,
        setIgnoreMouseEvents: () => {
          calls.push('ignore');
        },
      }),
      getSubtitleMeasurement: () => MEASUREMENT,
      getRendererInteractiveHint: () => false,
      shouldSuspend: () => false,
    }),
    { handled: false, active: false },
  );
  assert.deepEqual(calls, []);
});

test('applyLinuxOverlayPointerInteractionMousePassthrough toggles mouse input without full visibility refresh', () => {
  const calls: string[] = [];
  const window = {
    isDestroyed: () => false,
    isVisible: () => true,
    setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => {
      calls.push(`ignore:${ignore}:${options?.forward === true ? 'forward' : 'plain'}`);
    },
  };

  assert.equal(
    applyLinuxOverlayPointerInteractionMousePassthrough({
      active: true,
      getVisibleOverlayVisible: () => true,
      getMainWindow: () => window,
      shouldSuspend: () => false,
      shouldSuppressInteraction: () => false,
      updateVisibleOverlayVisibility: () => {
        calls.push('full-refresh');
      },
    }),
    true,
  );
  assert.deepEqual(calls, ['ignore:false:plain']);
});

test('applyLinuxOverlayPointerInteractionMousePassthrough falls back when pointer interaction is suppressed', () => {
  const calls: string[] = [];

  assert.equal(
    applyLinuxOverlayPointerInteractionMousePassthrough({
      active: false,
      getVisibleOverlayVisible: () => true,
      getMainWindow: () => ({
        isDestroyed: () => false,
        isVisible: () => true,
        setIgnoreMouseEvents: () => {
          calls.push('mouse-ignore');
        },
      }),
      shouldSuspend: () => false,
      shouldSuppressInteraction: () => true,
      updateVisibleOverlayVisibility: () => {
        calls.push('full-refresh');
      },
    }),
    false,
  );
  assert.deepEqual(calls, ['full-refresh']);
});
