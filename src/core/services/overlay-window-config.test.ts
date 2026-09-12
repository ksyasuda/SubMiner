import test from 'node:test';
import assert from 'node:assert/strict';
import { buildOverlayWindowOptions } from './overlay-window-options';

test('overlay window config explicitly disables renderer sandbox for preload compatibility', () => {
  const options = buildOverlayWindowOptions('visible', {
    isDev: false,
    yomitanSession: null,
  });

  assert.equal(options.title, 'SubMiner Overlay');
  assert.equal(options.backgroundColor, '#00000000');
  assert.equal(options.paintWhenInitiallyHidden, true);
  assert.equal(options.webPreferences?.sandbox, false);
  assert.equal(options.webPreferences?.backgroundThrottling, false);
});

test('macOS modal overlay uses a fullscreen auxiliary panel without changing the passive overlay', () => {
  const visibleOptions = buildOverlayWindowOptions('visible', {
    isDev: false,
    platform: 'darwin',
    yomitanSession: null,
  });
  const modalOptions = buildOverlayWindowOptions('modal', {
    isDev: false,
    platform: 'darwin',
    yomitanSession: null,
  });

  assert.equal(visibleOptions.type, undefined);
  assert.equal(modalOptions.type, 'panel');
});

test('non-macOS modal overlay remains a regular window', () => {
  const options = buildOverlayWindowOptions('modal', {
    isDev: false,
    platform: 'linux',
    yomitanSession: null,
  });

  assert.equal(options.type, undefined);
  assert.equal(options.roundedCorners, false);
});

test('Linux visible overlay window allows compositor resize for mpv-sized placement', () => {
  const originalPlatformDescriptor = Object.getOwnPropertyDescriptor(process, 'platform');

  Object.defineProperty(process, 'platform', {
    configurable: true,
    value: 'linux',
  });

  try {
    const visibleOptions = buildOverlayWindowOptions('visible', {
      isDev: false,
      yomitanSession: null,
    });
    const modalOptions = buildOverlayWindowOptions('modal', {
      isDev: false,
      yomitanSession: null,
    });

    assert.equal(visibleOptions.resizable, true);
    assert.equal(modalOptions.resizable, false);
  } finally {
    if (originalPlatformDescriptor) {
      Object.defineProperty(process, 'platform', originalPlatformDescriptor);
    }
  }
});

test('Linux visible overlay window stays managed so native apps can cover it', () => {
  const originalPlatformDescriptor = Object.getOwnPropertyDescriptor(process, 'platform');

  Object.defineProperty(process, 'platform', {
    configurable: true,
    value: 'linux',
  });

  try {
    const visibleOptions = buildOverlayWindowOptions('visible', {
      isDev: false,
      yomitanSession: null,
    });
    const modalOptions = buildOverlayWindowOptions('modal', {
      isDev: false,
      yomitanSession: null,
    });

    assert.equal(visibleOptions.alwaysOnTop, false);
    assert.equal(visibleOptions.focusable, true);
    assert.equal(modalOptions.focusable, true);
  } finally {
    if (originalPlatformDescriptor) {
      Object.defineProperty(process, 'platform', originalPlatformDescriptor);
    }
  }
});

test('Linux fullscreen visible overlay window uses X11 override-redirect-friendly options', () => {
  const originalPlatformDescriptor = Object.getOwnPropertyDescriptor(process, 'platform');

  Object.defineProperty(process, 'platform', {
    configurable: true,
    value: 'linux',
  });

  try {
    const visibleOptions = buildOverlayWindowOptions('visible', {
      isDev: false,
      linuxX11FullscreenOverlay: true,
      yomitanSession: null,
    });

    assert.equal(visibleOptions.alwaysOnTop, true);
    assert.equal(visibleOptions.focusable, false);
    assert.equal(visibleOptions.resizable, false);
  } finally {
    if (originalPlatformDescriptor) {
      Object.defineProperty(process, 'platform', originalPlatformDescriptor);
    }
  }
});

test('Windows visible overlay window config does not start as always-on-top', () => {
  const originalPlatformDescriptor = Object.getOwnPropertyDescriptor(process, 'platform');

  Object.defineProperty(process, 'platform', {
    configurable: true,
    value: 'win32',
  });

  try {
    const options = buildOverlayWindowOptions('visible', {
      isDev: false,
      yomitanSession: null,
    });

    assert.equal(options.alwaysOnTop, false);
  } finally {
    if (originalPlatformDescriptor) {
      Object.defineProperty(process, 'platform', originalPlatformDescriptor);
    }
  }
});

test('overlay window config uses the provided Yomitan session when available', () => {
  const yomitanSession = { id: 'session' } as never;
  const withSession = buildOverlayWindowOptions('visible', {
    isDev: false,
    yomitanSession,
  });
  const withoutSession = buildOverlayWindowOptions('visible', {
    isDev: false,
    yomitanSession: null,
  });

  assert.equal(withSession.webPreferences?.session, yomitanSession);
  assert.equal(withoutSession.webPreferences?.session, undefined);
});
