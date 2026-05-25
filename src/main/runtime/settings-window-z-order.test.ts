import assert from 'node:assert/strict';
import test from 'node:test';
import { shouldSuppressVisibleOverlayRaiseForSeparateWindow } from './settings-window-z-order';

test('separate settings windows suppress visible overlay restacking', () => {
  const mainWindow = { id: 'overlay', isDestroyed: () => false };
  const settingsWindow = { id: 'settings', isDestroyed: () => false };

  assert.equal(
    shouldSuppressVisibleOverlayRaiseForSeparateWindow({
      window: mainWindow,
      mainWindow,
      separateWindows: [settingsWindow],
    }),
    true,
  );
});

test('separate settings windows do not suppress unrelated or closed overlay work', () => {
  const mainWindow = { id: 'overlay', isDestroyed: () => false };
  const modalWindow = { id: 'modal', isDestroyed: () => false };
  const closedSettingsWindow = { id: 'settings', isDestroyed: () => true };

  assert.equal(
    shouldSuppressVisibleOverlayRaiseForSeparateWindow({
      window: modalWindow,
      mainWindow,
      separateWindows: [{ isDestroyed: () => false }],
    }),
    false,
  );
  assert.equal(
    shouldSuppressVisibleOverlayRaiseForSeparateWindow({
      window: mainWindow,
      mainWindow,
      separateWindows: [closedSettingsWindow, null],
    }),
    false,
  );
});
