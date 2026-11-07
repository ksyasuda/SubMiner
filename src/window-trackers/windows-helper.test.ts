import test from 'node:test';
import assert from 'node:assert/strict';
import { createWindowsOverlayOperations, findWindowsMpvTargetWindowHandle } from './windows-helper';
import type { MpvPollResult } from './win32';

function overlayOperationsHarness() {
  let nowMs = 0;
  let failing = true;
  const warnings: Array<{ message: string; operation: unknown }> = [];
  const nativeCall = () => {
    if (failing) throw new Error('native failure');
  };
  const operations = createWindowsOverlayOperations({
    loadWin32: () => ({
      setOverlayOwner: nativeCall,
      ensureOverlayTransparency: nativeCall,
      bindOverlayAboveMpv: nativeCall,
      clearOverlayOwner: nativeCall,
    }),
    warn: (message, details) => warnings.push({ message, operation: details.operation }),
    now: () => nowMs,
  });
  return {
    operations,
    warnings,
    setNow: (value: number) => {
      nowMs = value;
    },
    setFailing: (value: boolean) => {
      failing = value;
    },
  };
}

for (const [operation, call] of [
  ['set-owner', (ops) => ops.setOwner(1, 2)],
  ['transparency', (ops) => ops.ensureTransparency(1)],
  ['bind-above-mpv', (ops) => ops.bindAboveMpv(1, 2)],
  ['clear-owner', (ops) => ops.clearOwner(1)],
] as const satisfies ReadonlyArray<
  readonly [string, (ops: ReturnType<typeof createWindowsOverlayOperations>) => boolean]
>) {
  test(`${operation} reports failure, rate limits warnings and records recovery`, () => {
    const h = overlayOperationsHarness();
    const failed = { message: 'Windows native overlay operation failed', operation };

    assert.equal(call(h.operations), false);
    assert.deepEqual(h.warnings, [failed]);

    h.setNow(29_999);
    assert.equal(call(h.operations), false);
    assert.deepEqual(h.warnings, [failed]);

    h.setNow(30_000);
    assert.equal(call(h.operations), false);
    assert.deepEqual(h.warnings, [failed, failed]);

    h.setFailing(false);
    const recovered = { message: 'Windows native overlay operation recovered', operation };
    assert.equal(call(h.operations), true);
    assert.deepEqual(h.warnings, [failed, failed, recovered]);

    assert.equal(call(h.operations), true);
    assert.equal(h.warnings.length, 3);
  });
}

test('native overlay operations rate limit independently', () => {
  const h = overlayOperationsHarness();
  h.operations.setOwner(1, 2);
  h.operations.clearOwner(1);
  assert.deepEqual(
    h.warnings.map((warning) => warning.operation),
    ['set-owner', 'clear-owner'],
  );
});

test('findWindowsMpvTargetWindowHandle prefers the focused mpv window', () => {
  const result: MpvPollResult = {
    matches: [
      {
        hwnd: 111,
        bounds: { x: 0, y: 0, width: 1280, height: 720 },
        area: 1280 * 720,
        isForeground: false,
      },
      {
        hwnd: 222,
        bounds: { x: 10, y: 10, width: 800, height: 600 },
        area: 800 * 600,
        isForeground: true,
      },
    ],
    focusState: true,
    windowState: 'visible',
  };

  assert.equal(findWindowsMpvTargetWindowHandle(result), 222);
});

test('findWindowsMpvTargetWindowHandle falls back to the largest visible mpv window', () => {
  const result: MpvPollResult = {
    matches: [
      {
        hwnd: 111,
        bounds: { x: 0, y: 0, width: 640, height: 360 },
        area: 640 * 360,
        isForeground: false,
      },
      {
        hwnd: 222,
        bounds: { x: 10, y: 10, width: 1920, height: 1080 },
        area: 1920 * 1080,
        isForeground: false,
      },
    ],
    focusState: false,
    windowState: 'visible',
  };

  assert.equal(findWindowsMpvTargetWindowHandle(result), 222);
});

test('findWindowsMpvTargetWindowHandle returns null when no mpv windows are visible', () => {
  const result: MpvPollResult = {
    matches: [],
    focusState: false,
    windowState: 'not-found',
  };

  assert.equal(findWindowsMpvTargetWindowHandle(result), null);
});
