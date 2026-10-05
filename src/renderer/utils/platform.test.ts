import test from 'node:test';
import assert from 'node:assert/strict';

import { resolvePlatformInfo, type PlatformInfo } from './platform.js';

// Installs fake `window` and `navigator` globals for the duration of `run`, then restores them.
function withGlobals(globals: { window: unknown; navigator: unknown }, run: () => void): void {
  const previousWindow = (globalThis as { window?: unknown }).window;
  const previousNavigator = (globalThis as { navigator?: unknown }).navigator;
  Object.defineProperty(globalThis, 'window', { configurable: true, value: globals.window });
  Object.defineProperty(globalThis, 'navigator', { configurable: true, value: globals.navigator });
  try {
    run();
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'navigator', {
      configurable: true,
      value: previousNavigator,
    });
  }
}

const MAC = { platform: 'MacIntel', userAgent: 'Mozilla/5.0 (Macintosh)' };
const WINDOWS = { platform: 'Win32', userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' };

const MAC_VISIBLE: PlatformInfo = {
  overlayLayer: 'visible',
  isModalLayer: false,
  isLinuxPlatform: false,
  isMacOSPlatform: true,
  isWindowsPlatform: false,
  shouldToggleMouseIgnore: true,
};

const cases: Array<{
  name: string;
  search: string;
  preloadLayer: string;
  navigator: { platform: string; userAgent: string };
  expected: PlatformInfo;
}> = [
  {
    name: 'prefers query layer over preload layer',
    search: '?layer=visible',
    preloadLayer: 'modal',
    navigator: MAC,
    expected: MAC_VISIBLE,
  },
  {
    name: 'ignores legacy secondary layer and falls back to visible',
    search: '',
    preloadLayer: 'secondary',
    navigator: MAC,
    expected: MAC_VISIBLE,
  },
  {
    name: 'supports modal layer and disables mouse-ignore toggles',
    search: '',
    preloadLayer: 'modal',
    navigator: MAC,
    expected: {
      ...MAC_VISIBLE,
      overlayLayer: 'modal',
      isModalLayer: true,
      shouldToggleMouseIgnore: false,
    },
  },
  {
    name: 'flags Windows platforms',
    search: '',
    preloadLayer: 'visible',
    navigator: WINDOWS,
    expected: { ...MAC_VISIBLE, isMacOSPlatform: false, isWindowsPlatform: true },
  },
];

for (const c of cases) {
  test(`resolvePlatformInfo ${c.name}`, () => {
    withGlobals(
      {
        window: {
          electronAPI: { getOverlayLayer: () => c.preloadLayer },
          location: { search: c.search },
        },
        navigator: c.navigator,
      },
      () => {
        assert.deepEqual(resolvePlatformInfo(), c.expected);
      },
    );
  });
}
