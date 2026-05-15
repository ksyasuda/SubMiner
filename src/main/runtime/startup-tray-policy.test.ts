import assert from 'node:assert/strict';
import test from 'node:test';
import {
  shouldEnsureTrayOnStartupForInitialArgs,
  shouldQuitOnWindowAllClosedForTrayState,
} from './startup-tray-policy';

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

test('window-all-closed keeps tray-resident app alive', () => {
  assert.equal(
    shouldQuitOnWindowAllClosedForTrayState({ backgroundMode: false, hasTray: true }),
    false,
  );
});

test('window-all-closed quits non-background app without tray', () => {
  assert.equal(
    shouldQuitOnWindowAllClosedForTrayState({ backgroundMode: false, hasTray: false }),
    true,
  );
});

test('window-all-closed keeps background app alive without tray', () => {
  assert.equal(
    shouldQuitOnWindowAllClosedForTrayState({ backgroundMode: true, hasTray: false }),
    false,
  );
});
