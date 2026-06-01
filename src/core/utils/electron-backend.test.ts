import assert from 'node:assert/strict';
import test from 'node:test';
import { shouldForceX11ElectronBackend } from './electron-backend';

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
