import assert from 'node:assert/strict';
import test from 'node:test';
import { resolveX11ElectronRelaunchArgs, shouldForceX11ElectronBackend } from './electron-backend';

function withPlatform(platform: NodeJS.Platform, run: () => void): void {
  const original = Object.getOwnPropertyDescriptor(process, 'platform');
  Object.defineProperty(process, 'platform', { configurable: true, value: platform });
  try {
    run();
  } finally {
    if (original) Object.defineProperty(process, 'platform', original);
  }
}

test('shouldForceX11ElectronBackend forces X11 on Linux except Hyprland/Sway', () => {
  withPlatform('linux', () => {
    assert.equal(shouldForceX11ElectronBackend({ XDG_CURRENT_DESKTOP: 'KDE' }), true);
    assert.equal(shouldForceX11ElectronBackend({ WAYLAND_DISPLAY: 'wayland-0' }), true);
    // Even an explicit Wayland hint is overridden to x11 on unsupported compositors.
    assert.equal(shouldForceX11ElectronBackend({ ELECTRON_OZONE_PLATFORM_HINT: 'wayland' }), true);
    // Hyprland/Sway keep native Wayland (guard reports explicit wayland hints elsewhere).
    assert.equal(shouldForceX11ElectronBackend({ HYPRLAND_INSTANCE_SIGNATURE: 'hypr' }), false);
    assert.equal(shouldForceX11ElectronBackend({ SWAYSOCK: '/tmp/sway.sock' }), false);
  });
});

test('shouldForceX11ElectronBackend is false off Linux', () => {
  withPlatform('darwin', () => {
    assert.equal(shouldForceX11ElectronBackend({ XDG_CURRENT_DESKTOP: 'KDE' }), false);
  });
  withPlatform('win32', () => {
    assert.equal(shouldForceX11ElectronBackend({}), false);
  });
});

test('resolveX11ElectronRelaunchArgs adds the raw X11 Ozone argument on unsupported Linux', () => {
  assert.deepEqual(
    resolveX11ElectronRelaunchArgs(
      ['--start'],
      {
        DISPLAY: ':1',
        WAYLAND_DISPLAY: 'wayland-0',
        XDG_CURRENT_DESKTOP: 'KDE',
      },
      'linux',
    ),
    ['--start', '--ozone-platform=x11'],
  );
});

test('resolveX11ElectronRelaunchArgs avoids loops and preserves native Wayland backends', () => {
  const kdeWayland = {
    DISPLAY: ':1',
    WAYLAND_DISPLAY: 'wayland-0',
    XDG_CURRENT_DESKTOP: 'KDE',
  };
  assert.equal(
    resolveX11ElectronRelaunchArgs(['--start', '--ozone-platform=x11'], kdeWayland, 'linux'),
    null,
  );
  assert.equal(
    resolveX11ElectronRelaunchArgs(
      ['--start'],
      { ...kdeWayland, HYPRLAND_INSTANCE_SIGNATURE: 'hypr' },
      'linux',
    ),
    null,
  );
  assert.equal(resolveX11ElectronRelaunchArgs(['--start'], kdeWayland, 'darwin'), null);
  assert.equal(
    resolveX11ElectronRelaunchArgs(
      [],
      {
        ...kdeWayland,
        SUBMINER_APP_ARGC: '1',
        SUBMINER_APP_ARG_0: '--start',
      },
      'linux',
    )?.at(-1),
    '--ozone-platform=x11',
  );
  assert.equal(
    resolveX11ElectronRelaunchArgs([], { ...kdeWayland, SUBMINER_X11_BOOTSTRAPPED: '1' }, 'linux'),
    null,
  );
});

test('resolveX11ElectronRelaunchArgs replaces an explicit unsupported Wayland argument', () => {
  assert.deepEqual(
    resolveX11ElectronRelaunchArgs(
      ['--start', '--ozone-platform', 'wayland'],
      {
        DISPLAY: ':1',
        WAYLAND_DISPLAY: 'wayland-0',
        XDG_CURRENT_DESKTOP: 'KDE',
      },
      'linux',
    ),
    ['--start', '--ozone-platform=x11'],
  );
});
