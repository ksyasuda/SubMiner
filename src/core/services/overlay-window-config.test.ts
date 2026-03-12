import test from 'node:test';
import assert from 'node:assert/strict';
import { buildOverlayWindowOptions } from './overlay-window-options';

test('overlay window config explicitly disables renderer sandbox for preload compatibility', () => {
  const options = buildOverlayWindowOptions('visible', {
    isDev: false,
    yomitanSession: null,
  });

  assert.equal(options.webPreferences?.sandbox, false);
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
