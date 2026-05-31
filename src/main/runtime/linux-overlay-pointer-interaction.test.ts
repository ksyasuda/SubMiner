import assert from 'node:assert/strict';
import test from 'node:test';
import {
  type LinuxOverlayPointerInteractionDeps,
  isCursorOverSubtitle,
  mapOverlayMeasurementForPointerInteraction,
  resolveDesiredOverlayInteractive,
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
