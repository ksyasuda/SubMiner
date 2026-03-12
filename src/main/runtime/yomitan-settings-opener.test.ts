import test from 'node:test';
import assert from 'node:assert/strict';
import { createOpenYomitanSettingsHandler } from './yomitan-settings-opener';

test('yomitan opener warns when extension cannot be loaded', async () => {
  const logs: string[] = [];
  const openSettings = createOpenYomitanSettingsHandler({
    ensureYomitanExtensionLoaded: async () => null,
    openYomitanSettingsWindow: () => {
      throw new Error('should not open');
    },
    getExistingWindow: () => null,
    setWindow: () => {},
    logWarn: (message) => logs.push(message),
    logError: () => logs.push('error'),
  });

  openSettings();
  await Promise.resolve();
  await Promise.resolve();
  assert.ok(logs.includes('Unable to open Yomitan settings: extension failed to load.'));
});

test('yomitan opener opens settings window when extension is available', async () => {
  let opened = false;
  const yomitanSession = { id: 'session' };
  const openSettings = createOpenYomitanSettingsHandler({
    ensureYomitanExtensionLoaded: async () => ({ id: 'ext' }),
    openYomitanSettingsWindow: ({ yomitanSession: forwardedSession }) => {
      opened = (forwardedSession as { id: string } | null)?.id === 'session';
    },
    getExistingWindow: () => null,
    setWindow: () => {},
    getYomitanSession: () => yomitanSession,
    logWarn: () => {},
    logError: () => {},
  });

  openSettings();
  await Promise.resolve();
  await Promise.resolve();
  assert.equal(opened, true);
});
