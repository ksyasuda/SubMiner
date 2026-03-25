import assert from 'node:assert/strict';
import test from 'node:test';
import { shouldEnsureTrayOnStartupForInitialArgs } from './startup-tray-policy';

test('startup tray policy enables tray on Windows by default', () => {
  assert.equal(shouldEnsureTrayOnStartupForInitialArgs('win32', null), true);
});

test('startup tray policy skips tray for direct youtube playback on Windows', () => {
  assert.equal(
    shouldEnsureTrayOnStartupForInitialArgs('win32', {
      youtubePlay: 'https://www.youtube.com/watch?v=abc',
    } as never),
    false,
  );
});

test('startup tray policy skips tray outside Windows', () => {
  assert.equal(shouldEnsureTrayOnStartupForInitialArgs('linux', null), false);
});
