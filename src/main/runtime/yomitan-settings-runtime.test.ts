import assert from 'node:assert/strict';
import test from 'node:test';
import { createYomitanSettingsRuntime } from './yomitan-settings-runtime';

test('yomitan settings runtime composes opener with built deps', async () => {
  let existingWindow: { id: string } | null = null;
  const calls: string[] = [];

  const runtime = createYomitanSettingsRuntime({
    ensureYomitanExtensionLoaded: async () => ({ id: 'ext' }),
    openYomitanSettingsWindow: ({ getExistingWindow, setWindow }) => {
      calls.push('open-window');
      const current = getExistingWindow();
      if (!current) {
        setWindow({ id: 'settings' });
      }
    },
    getExistingWindow: () => existingWindow as never,
    setWindow: (window) => {
      existingWindow = window as { id: string } | null;
    },
    logWarn: (message) => calls.push(`warn:${message}`),
    logError: (message) => calls.push(`error:${message}`),
  });

  runtime.openYomitanSettings();
  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.deepEqual(existingWindow, { id: 'settings' });
  assert.deepEqual(calls, ['open-window']);
});
