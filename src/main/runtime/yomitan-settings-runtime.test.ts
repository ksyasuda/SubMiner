import assert from 'node:assert/strict';
import test from 'node:test';
import { createYomitanSettingsRuntime } from './yomitan-settings-runtime';

test('yomitan settings runtime composes opener with built deps', async () => {
  let existingWindow: { id: string } | null = null;
  const calls: string[] = [];
  const yomitanSession = { id: 'session' };

  const runtime = createYomitanSettingsRuntime({
    ensureYomitanExtensionLoaded: async () => ({ id: 'ext' }),
    openYomitanSettingsWindow: ({ getExistingWindow, setWindow, yomitanSession: forwardedSession }) => {
      calls.push(`open-window:${(forwardedSession as { id: string } | null)?.id ?? 'null'}`);
      const current = getExistingWindow();
      if (!current) {
        setWindow({ id: 'settings' });
      }
    },
    getExistingWindow: () => existingWindow as never,
    setWindow: (window) => {
      existingWindow = window as { id: string } | null;
    },
    getYomitanSession: () => yomitanSession,
    logWarn: (message) => calls.push(`warn:${message}`),
    logError: (message) => calls.push(`error:${message}`),
  });

  runtime.openYomitanSettings();
  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.deepEqual(existingWindow, { id: 'settings' });
  assert.deepEqual(calls, ['open-window:session']);
});

test('yomitan settings runtime warns and does not open when no yomitan session is available', async () => {
  let existingWindow: { id: string } | null = null;
  const calls: string[] = [];

  const runtime = createYomitanSettingsRuntime({
    ensureYomitanExtensionLoaded: async () => ({ id: 'ext' }),
    openYomitanSettingsWindow: () => {
      calls.push('open-window');
    },
    getExistingWindow: () => existingWindow as never,
    setWindow: (window) => {
      existingWindow = window as { id: string } | null;
    },
    getYomitanSession: () => null,
    logWarn: (message) => calls.push(`warn:${message}`),
    logError: (message) => calls.push(`error:${message}`),
  });

  runtime.openYomitanSettings();
  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.equal(existingWindow, null);
  assert.deepEqual(calls, ['warn:Unable to open Yomitan settings: Yomitan session is unavailable.']);
});
