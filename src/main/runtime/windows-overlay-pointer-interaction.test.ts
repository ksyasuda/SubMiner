import assert from 'node:assert/strict';
import test from 'node:test';
import {
  isCursorOverWindowsOverlayInteractiveRect,
  resolveDesiredWindowsOverlayInteractive,
  tickWindowsOverlayPointerInteraction,
  type WindowsOverlayPointerInteractionDeps,
} from './windows-overlay-pointer-interaction';
import type { OverlayContentMeasurement } from '../../types';

const BOUNDS = { x: 100, y: 100, width: 1920, height: 1080 };
const MEASUREMENT: OverlayContentMeasurement = {
  layer: 'visible',
  measuredAtMs: 1,
  viewport: { width: 1920, height: 1080 },
  contentRect: { x: 800, y: 900, width: 320, height: 80 },
};

function makeDeps(overrides: Partial<WindowsOverlayPointerInteractionDeps>): {
  deps: WindowsOverlayPointerInteractionDeps;
  state: { active: boolean };
} {
  const state = { active: false };
  const deps: WindowsOverlayPointerInteractionDeps = {
    getVisibleOverlayVisible: () => true,
    getMainWindow: () => ({
      isDestroyed: () => false,
      isVisible: () => true,
      getBounds: () => BOUNDS,
    }),
    getCursorScreenPoint: () => ({ x: 1000, y: 1040 }),
    getSubtitleMeasurement: () => MEASUREMENT,
    shouldSuspend: () => false,
    getInteractionActive: () => state.active,
    setInteractionActive: (active) => {
      state.active = active;
    },
    ...overrides,
  };
  return { deps, state };
}

test('isCursorOverWindowsOverlayInteractiveRect hit-tests measured overlay rects', () => {
  assert.equal(
    isCursorOverWindowsOverlayInteractiveRect({ x: 1000, y: 1040 }, BOUNDS, MEASUREMENT),
    true,
  );
  assert.equal(
    isCursorOverWindowsOverlayInteractiveRect({ x: 500, y: 1040 }, BOUNDS, MEASUREMENT),
    false,
  );
});

test('isCursorOverWindowsOverlayInteractiveRect scales viewport px to window px', () => {
  const scaled = { ...BOUNDS, width: 3840, height: 2160 };
  assert.equal(
    isCursorOverWindowsOverlayInteractiveRect({ x: 1700, y: 1900 }, scaled, MEASUREMENT),
    true,
  );
});

test('isCursorOverWindowsOverlayInteractiveRect uses separate interactive rects', () => {
  const measurement: OverlayContentMeasurement = {
    layer: 'visible',
    measuredAtMs: 1,
    viewport: { width: 1920, height: 1080 },
    contentRect: { x: 700, y: 40, width: 520, height: 940 },
    interactiveRects: [
      { x: 700, y: 40, width: 520, height: 80 },
      { x: 760, y: 900, width: 400, height: 80 },
    ],
  };

  assert.equal(
    isCursorOverWindowsOverlayInteractiveRect({ x: 900, y: 300 }, BOUNDS, measurement),
    false,
  );
  assert.equal(
    isCursorOverWindowsOverlayInteractiveRect({ x: 900, y: 180 }, BOUNDS, measurement),
    true,
  );
  assert.equal(
    isCursorOverWindowsOverlayInteractiveRect({ x: 900, y: 1060 }, BOUNDS, measurement),
    true,
  );
});

test('resolveDesiredWindowsOverlayInteractive: interactive over subtitle, passthrough off it', () => {
  assert.equal(resolveDesiredWindowsOverlayInteractive(makeDeps({}).deps), true);
  assert.equal(
    resolveDesiredWindowsOverlayInteractive(
      makeDeps({ getCursorScreenPoint: () => ({ x: 200, y: 200 }) }).deps,
    ),
    false,
  );
});

test('resolveDesiredWindowsOverlayInteractive returns null while another surface owns input', () => {
  assert.equal(
    resolveDesiredWindowsOverlayInteractive(makeDeps({ shouldSuspend: () => true }).deps),
    null,
  );
  assert.equal(
    resolveDesiredWindowsOverlayInteractive(makeDeps({ getMainWindow: () => null }).deps),
    null,
  );
});

test('tickWindowsOverlayPointerInteraction toggles only the fallback-owned state', () => {
  const calls: boolean[] = [];
  const { deps, state } = makeDeps({
    setInteractionActive: (active) => {
      calls.push(active);
      state.active = active;
    },
  });

  tickWindowsOverlayPointerInteraction(deps);
  tickWindowsOverlayPointerInteraction(deps);
  assert.deepEqual(calls, [true]);

  deps.getCursorScreenPoint = () => ({ x: 200, y: 200 });
  tickWindowsOverlayPointerInteraction(deps);
  assert.deepEqual(calls, [true, false]);
});

test('tickWindowsOverlayPointerInteraction leaves renderer-owned state alone while suspended', () => {
  const calls: boolean[] = [];
  const { deps } = makeDeps({
    getInteractionActive: () => true,
    shouldSuspend: () => true,
    setInteractionActive: (active) => calls.push(active),
  });

  tickWindowsOverlayPointerInteraction(deps);
  assert.deepEqual(calls, []);
});
