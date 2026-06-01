import assert from 'node:assert/strict';
import test from 'node:test';
import {
  ensureLinuxOverlayZOrderKeepAliveLoop,
  type LinuxOverlayZOrderKeepAliveDeps,
  shouldRunLinuxOverlayZOrderKeepAlive,
  stopLinuxOverlayZOrderKeepAliveLoop,
  tickLinuxOverlayZOrderKeepAlive,
} from './linux-overlay-zorder-keepalive';

function withPlatform(platform: NodeJS.Platform, run: () => void): void {
  const original = Object.getOwnPropertyDescriptor(process, 'platform');
  Object.defineProperty(process, 'platform', { configurable: true, value: platform });
  try {
    run();
  } finally {
    if (original) Object.defineProperty(process, 'platform', original);
  }
}

function makeDeps(
  overrides: Partial<LinuxOverlayZOrderKeepAliveDeps>,
  calls: string[],
): LinuxOverlayZOrderKeepAliveDeps {
  return {
    getVisibleOverlayVisible: () => true,
    getMainWindow: () => ({ isDestroyed: () => false, isVisible: () => true }),
    isTrackingMpvWindow: () => true,
    isMpvWindowFocused: () => true,
    isOverlayWindowFocused: () => false,
    shouldSuppressReassert: () => false,
    raiseMpvWindow: async () => {
      calls.push('raise-mpv');
      return true;
    },
    releaseOverlayLayerOrder: () => calls.push('release'),
    enforceOverlayLayerOrder: () => calls.push('enforce'),
    focusOverlayWindow: () => calls.push('focus-overlay'),
    ...overrides,
  };
}

test('shouldRunLinuxOverlayZOrderKeepAlive runs on Linux except Hyprland/Sway', () => {
  withPlatform('linux', () => {
    assert.equal(shouldRunLinuxOverlayZOrderKeepAlive({ XDG_CURRENT_DESKTOP: 'KDE' }), true);
    assert.equal(shouldRunLinuxOverlayZOrderKeepAlive({ HYPRLAND_INSTANCE_SIGNATURE: 'h' }), false);
    assert.equal(shouldRunLinuxOverlayZOrderKeepAlive({ SWAYSOCK: '/tmp/s' }), false);
  });
  withPlatform('win32', () => {
    assert.equal(shouldRunLinuxOverlayZOrderKeepAlive({}), false);
  });
});

test('tick re-asserts overlay level when the overlay is shown and unobstructed', async () => {
  const calls: string[] = [];
  await tickLinuxOverlayZOrderKeepAlive(makeDeps({}, calls));
  assert.deepEqual(calls, ['enforce']);
});

test('tick raises mpv behind a focused overlay when mpv is behind another app', async () => {
  const calls: string[] = [];
  await tickLinuxOverlayZOrderKeepAlive(
    makeDeps(
      {
        isMpvWindowFocused: () => false,
        isOverlayWindowFocused: () => true,
      },
      calls,
    ),
  );
  assert.deepEqual(calls, ['raise-mpv', 'enforce', 'focus-overlay']);
});

test('tick releases stale overlay topmost when another app is focused', async () => {
  const calls: string[] = [];
  await tickLinuxOverlayZOrderKeepAlive(
    makeDeps(
      {
        isMpvWindowFocused: () => false,
        isOverlayWindowFocused: () => false,
      },
      calls,
    ),
  );
  assert.deepEqual(calls, ['release']);
});

test('tick skips when overlay hidden, mpv untracked, suppressed, or window gone', async () => {
  for (const override of [
    { getVisibleOverlayVisible: () => false },
    { isTrackingMpvWindow: () => false },
    { shouldSuppressReassert: () => true },
    { getMainWindow: () => null },
    { getMainWindow: () => ({ isDestroyed: () => true, isVisible: () => true }) },
    { getMainWindow: () => ({ isDestroyed: () => false, isVisible: () => false }) },
  ] satisfies Array<Partial<LinuxOverlayZOrderKeepAliveDeps>>) {
    const calls: string[] = [];
    await tickLinuxOverlayZOrderKeepAlive(makeDeps(override, calls));
    assert.deepEqual(calls, []);
  }
});

test('keep-alive loop skips overlapping ticks and resets after async completion', async () => {
  const originalSetInterval = globalThis.setInterval;
  const originalClearInterval = globalThis.clearInterval;
  let intervalCallback: (() => void) | null = null;
  let resolveRaise: (() => void) | null = null;
  let raiseCalls = 0;

  globalThis.setInterval = ((callback: () => void) => {
    intervalCallback = callback;
    return { unref: () => {} } as ReturnType<typeof setInterval>;
  }) as typeof setInterval;
  globalThis.clearInterval = (() => {}) as typeof clearInterval;

  try {
    withPlatform('linux', () => {
      ensureLinuxOverlayZOrderKeepAliveLoop(
        makeDeps(
          {
            isMpvWindowFocused: () => false,
            isOverlayWindowFocused: () => true,
            raiseMpvWindow: async () => {
              raiseCalls += 1;
              await new Promise<void>((resolve) => {
                resolveRaise = resolve;
              });
              return true;
            },
          },
          [],
        ),
        {},
      );
    });

    assert.ok(intervalCallback);
    const tick = intervalCallback as () => void;
    tick();
    tick();
    assert.equal(raiseCalls, 1);

    assert.ok(resolveRaise);
    const finishRaise = resolveRaise as () => void;
    finishRaise();
    await new Promise((resolve) => setTimeout(resolve, 0));

    tick();
    assert.equal(raiseCalls, 2);
  } finally {
    stopLinuxOverlayZOrderKeepAliveLoop();
    globalThis.setInterval = originalSetInterval;
    globalThis.clearInterval = originalClearInterval;
  }
});
