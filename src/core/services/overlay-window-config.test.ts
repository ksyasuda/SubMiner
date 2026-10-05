import test from 'node:test';
import assert from 'node:assert/strict';
import { buildOverlayWindowOptions } from './overlay-window-options';

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

test('Linux visible overlay window allows compositor resize for mpv-sized placement', () => {
  const visibleOptions = buildOverlayWindowOptions('visible', {
    isDev: false,
    platform: 'linux',
    yomitanSession: null,
  });
  const modalOptions = buildOverlayWindowOptions('modal', {
    isDev: false,
    platform: 'linux',
    yomitanSession: null,
  });

  assert.equal(visibleOptions.resizable, true);
  assert.equal(modalOptions.resizable, false);
});

test('Linux visible overlay window stays managed so native apps can cover it', () => {
  const visibleOptions = buildOverlayWindowOptions('visible', {
    isDev: false,
    platform: 'linux',
    yomitanSession: null,
  });
  const modalOptions = buildOverlayWindowOptions('modal', {
    isDev: false,
    platform: 'linux',
    yomitanSession: null,
  });

  assert.equal(visibleOptions.alwaysOnTop, false);
  assert.equal(visibleOptions.focusable, true);
  assert.equal(modalOptions.focusable, true);
});

test('Linux fullscreen visible overlay window uses X11 override-redirect-friendly options', () => {
  const visibleOptions = buildOverlayWindowOptions('visible', {
    isDev: false,
    platform: 'linux',
    linuxX11FullscreenOverlay: true,
    yomitanSession: null,
  });

  assert.equal(visibleOptions.alwaysOnTop, true);
  assert.equal(visibleOptions.focusable, false);
  assert.equal(visibleOptions.resizable, false);
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
