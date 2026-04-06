import test from 'node:test';
import assert from 'node:assert/strict';
import { detectSessionBackend } from './backend-detection';

test('detectSessionBackend returns kwin for KDE Plasma Wayland sessions', () => {
  assert.equal(
    detectSessionBackend('linux', {
      WAYLAND_DISPLAY: 'wayland-0',
      XDG_CURRENT_DESKTOP: 'KDE',
      XDG_SESSION_DESKTOP: 'KDE',
      XDG_SESSION_TYPE: 'wayland',
    } as NodeJS.ProcessEnv),
    'kwin',
  );
});

test('detectSessionBackend returns x11 only when DISPLAY is present', () => {
  assert.equal(
    detectSessionBackend('linux', {
      DISPLAY: ':0',
    } as NodeJS.ProcessEnv),
    'x11',
  );
});

test('detectSessionBackend does not auto-detect x11 for Wayland-only unsupported sessions', () => {
  assert.equal(
    detectSessionBackend('linux', {
      WAYLAND_DISPLAY: 'wayland-0',
      XDG_SESSION_TYPE: 'wayland',
    } as NodeJS.ProcessEnv),
    null,
  );
});
