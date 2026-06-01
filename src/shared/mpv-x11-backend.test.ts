import assert from 'node:assert/strict';
import test from 'node:test';
import {
  MPV_X11_BACKEND_ARGS,
  applyX11EnvOverrides,
  isSupportedWaylandCompositor,
  shouldForceX11MpvBackend,
  shouldForceX11WaylandSession,
} from './mpv-x11-backend';

function withPlatform(platform: NodeJS.Platform, run: () => void): void {
  const original = Object.getOwnPropertyDescriptor(process, 'platform');
  Object.defineProperty(process, 'platform', { configurable: true, value: platform });
  try {
    run();
  } finally {
    if (original) Object.defineProperty(process, 'platform', original);
  }
}

const KDE_WAYLAND = {
  DISPLAY: ':1',
  WAYLAND_DISPLAY: 'wayland-0',
  XDG_SESSION_TYPE: 'wayland',
  XDG_CURRENT_DESKTOP: 'KDE',
  XDG_SESSION_DESKTOP: 'plasma',
};

test('isSupportedWaylandCompositor detects Hyprland and Sway via env or xdg desktop', () => {
  assert.equal(isSupportedWaylandCompositor({ HYPRLAND_INSTANCE_SIGNATURE: 'hypr' }), true);
  assert.equal(isSupportedWaylandCompositor({ SWAYSOCK: '/tmp/sway.sock' }), true);
  assert.equal(isSupportedWaylandCompositor({ XDG_CURRENT_DESKTOP: 'Hyprland' }), true);
  assert.equal(isSupportedWaylandCompositor({ XDG_SESSION_DESKTOP: 'sway' }), true);
  assert.equal(isSupportedWaylandCompositor(KDE_WAYLAND), false);
});

test('shouldForceX11WaylandSession forces X11 for unsupported Wayland sessions only', () => {
  withPlatform('linux', () => {
    assert.equal(shouldForceX11WaylandSession(KDE_WAYLAND), true);
    // GNOME Wayland (also unsupported) → forced.
    assert.equal(
      shouldForceX11WaylandSession({
        DISPLAY: ':0',
        WAYLAND_DISPLAY: 'wayland-0',
        XDG_CURRENT_DESKTOP: 'GNOME',
      }),
      true,
    );
    // Hyprland keeps native Wayland.
    assert.equal(
      shouldForceX11WaylandSession({ ...KDE_WAYLAND, HYPRLAND_INSTANCE_SIGNATURE: 'hypr' }),
      false,
    );
    // No X11 display to fall back to.
    assert.equal(shouldForceX11WaylandSession({ WAYLAND_DISPLAY: 'wayland-0' }), false);
    // Pure X11 session (no Wayland) → nothing to force.
    assert.equal(shouldForceX11WaylandSession({ DISPLAY: ':0', XDG_SESSION_TYPE: 'x11' }), false);
  });
});

test('shouldForceX11WaylandSession is false off Linux', () => {
  withPlatform('darwin', () => {
    assert.equal(shouldForceX11WaylandSession(KDE_WAYLAND), false);
  });
  withPlatform('win32', () => {
    assert.equal(shouldForceX11WaylandSession(KDE_WAYLAND), false);
  });
});

test('shouldForceX11MpvBackend honors explicit x11 and auto modes', () => {
  withPlatform('linux', () => {
    // Explicit x11 forces even without Wayland.
    assert.equal(shouldForceX11MpvBackend('x11', { DISPLAY: ':0' }), true);
    // Auto defers to the session check.
    assert.equal(shouldForceX11MpvBackend('auto', KDE_WAYLAND), true);
    assert.equal(
      shouldForceX11MpvBackend('auto', { ...KDE_WAYLAND, SWAYSOCK: '/tmp/sway.sock' }),
      false,
    );
    // No display at all.
    assert.equal(shouldForceX11MpvBackend('x11', {}), false);
  });
});

test('applyX11EnvOverrides strips Wayland hints and pins session type to x11', () => {
  const env = {
    DISPLAY: ':1',
    WAYLAND_DISPLAY: 'wayland-0',
    XDG_SESSION_TYPE: 'wayland',
    HYPRLAND_INSTANCE_SIGNATURE: 'hypr',
    SWAYSOCK: '/tmp/sway.sock',
  };
  const result = applyX11EnvOverrides(env);
  assert.equal(result, env); // mutates in place
  assert.equal(result.DISPLAY, ':1');
  assert.equal(result.WAYLAND_DISPLAY, undefined);
  assert.equal(result.HYPRLAND_INSTANCE_SIGNATURE, undefined);
  assert.equal(result.SWAYSOCK, undefined);
  assert.equal(result.XDG_SESSION_TYPE, 'x11');
});

test('MPV_X11_BACKEND_ARGS pins the GPU stack to X11', () => {
  assert.deepEqual(
    [...MPV_X11_BACKEND_ARGS],
    ['--vo=gpu', '--gpu-api=opengl', '--gpu-context=x11egl,x11'],
  );
});
