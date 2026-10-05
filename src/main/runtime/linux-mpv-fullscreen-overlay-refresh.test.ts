import assert from 'node:assert/strict';
import test, { beforeEach, afterEach } from 'node:test';
import {
  clearLinuxMpvFullscreenOverlayRefreshTimeouts,
  updateLinuxMpvFullscreenOverlayRefreshBurst,
  scheduleLinuxVisibleOverlayFullscreenRefreshBurst,
  type LinuxMpvFullscreenOverlayRefreshDeps,
} from './linux-mpv-fullscreen-overlay-refresh';

const compositorEnvKeys = [
  'HYPRLAND_INSTANCE_SIGNATURE',
  'SWAYSOCK',
  'XDG_CURRENT_DESKTOP',
  'XDG_SESSION_DESKTOP',
] as const;
const originalCompositorEnv = compositorEnvKeys.map((key) => [key, process.env[key]] as const);
beforeEach(() => {
  for (const key of compositorEnvKeys) delete process.env[key];
});
afterEach(() => {
  for (const [key, value] of originalCompositorEnv) {
    if (value === undefined) delete process.env[key];
    else process.env[key] = value;
  }
});

// The subject schedules with the global setTimeout (0/50/150/300/600 ms), so these tests run
// on real timers and poll for the expected calls instead of sleeping a fixed time.
async function waitFor(predicate: () => boolean, timeoutMs = 1_500): Promise<void> {
  const deadline = Date.now() + timeoutMs;
  while (!predicate() && Date.now() < deadline) {
    await new Promise((resolve) => setTimeout(resolve, 5));
  }
  assert.ok(predicate(), 'timed out waiting for overlay refresh calls');
}

/** The subject only refreshes on linux; clears pending burst timers afterwards. */
async function withLinuxPlatform(run: () => Promise<void>): Promise<void> {
  const originalPlatformDescriptor = Object.getOwnPropertyDescriptor(process, 'platform');
  Object.defineProperty(process, 'platform', { configurable: true, value: 'linux' });
  try {
    await run();
  } finally {
    clearLinuxMpvFullscreenOverlayRefreshTimeouts();
    if (originalPlatformDescriptor) {
      Object.defineProperty(process, 'platform', originalPlatformDescriptor);
    }
  }
}

function makeDeps(
  calls: string[],
  overrides: Partial<LinuxMpvFullscreenOverlayRefreshDeps> = {},
  options: { windowVisible?: boolean } = {},
): LinuxMpvFullscreenOverlayRefreshDeps {
  const window = {
    hide: () => calls.push('hide'),
    showInactive: () => calls.push('showInactive'),
    isDestroyed: () => false,
    isVisible: () => options.windowVisible ?? true,
    setIgnoreMouseEvents: (ignore: boolean, mouseOptions?: { forward?: boolean }) =>
      calls.push(`mouse:${ignore}:${mouseOptions?.forward === true ? 'forward' : 'plain'}`),
  };
  return {
    overlayManager: {
      getMainWindow: () => window,
      getVisibleOverlayVisible: () => true,
    },
    overlayVisibilityRuntime: {
      updateVisibleOverlayVisibility: () => calls.push('visibility'),
    },
    syncVisibleOverlayMpvFullscreenMode: (fullscreen) => calls.push(`mode:${fullscreen}`),
    ensureOverlayWindowLevel: () => calls.push('restack'),
    ...overrides,
  };
}

const count = (calls: string[], entry: string): number =>
  calls.filter((call) => call === entry).length;

const REFRESH_TICKS = 5;

for (const { compositorKey, expectedTickCalls } of [
  {
    // Hyprland restacks in place; hide/show would let it cancel the fullscreen transition.
    compositorKey: 'HYPRLAND_INSTANCE_SIGNATURE',
    expectedTickCalls: ['mode:true', 'visibility', 'mouse:true:forward', 'restack'],
  },
  {
    compositorKey: 'SWAYSOCK',
    expectedTickCalls: [
      'mode:true',
      'visibility',
      'hide',
      'showInactive',
      'mouse:true:forward',
      'restack',
    ],
  },
]) {
  test(`${compositorKey} fullscreen refresh uses compositor-specific restacking`, async () => {
    await withLinuxPlatform(async () => {
      process.env[compositorKey] = 'fullscreen-refresh-test';
      const calls: string[] = [];
      scheduleLinuxVisibleOverlayFullscreenRefreshBurst(true, makeDeps(calls));

      await waitFor(() => count(calls, 'restack') === REFRESH_TICKS);

      // Click-through is restored after the remap and before the window level is reasserted.
      assert.deepEqual(
        calls,
        Array.from({ length: REFRESH_TICKS }, () => expectedTickCalls).flat(),
      );
    });
  });
}

test('linux mpv fullscreen overlay refresh remembers mode even when overlay is hidden', async () => {
  await withLinuxPlatform(async () => {
    const calls: string[] = [];
    scheduleLinuxVisibleOverlayFullscreenRefreshBurst(
      true,
      makeDeps(calls, {
        overlayManager: {
          getMainWindow: () => null,
          getVisibleOverlayVisible: () => false,
        },
      }),
    );

    await waitFor(() => count(calls, 'mode:true') >= 2);

    assert.ok(calls.every((call) => call === 'mode:true'));
  });
});

test('linux mpv fullscreen overlay refresh updates mode without hide/show when fullscreen exits', async () => {
  await withLinuxPlatform(async () => {
    const calls: string[] = [];
    const deps = makeDeps(calls);

    // Leaving fullscreen cancels the pending fullscreen-enter burst.
    const cancel = updateLinuxMpvFullscreenOverlayRefreshBurst(true, deps, null);
    updateLinuxMpvFullscreenOverlayRefreshBurst(false, deps, cancel);

    await waitFor(() => count(calls, 'mode:false') >= 2);

    assert.ok(calls.every((call) => call === 'mode:false' || call === 'visibility'));
  });
});

test('linux mpv fullscreen overlay refresh preserves active subtitle interaction after restacking', async () => {
  await withLinuxPlatform(async () => {
    const calls: string[] = [];
    scheduleLinuxVisibleOverlayFullscreenRefreshBurst(
      true,
      makeDeps(calls, { getOverlayInteractionActive: () => true }),
    );

    await waitFor(() => count(calls, 'restack') >= 1);

    assert.ok(calls.indexOf('mouse:false:plain') > calls.indexOf('showInactive'));
    assert.equal(calls.includes('mouse:true:forward'), false);
  });
});
