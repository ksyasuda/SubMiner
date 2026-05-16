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
  let forwardedSession: { id: string } | null | undefined;
  const yomitanSession = { id: 'session' };
  const openSettings = createOpenYomitanSettingsHandler({
    ensureYomitanExtensionLoaded: async () => ({ id: 'ext' }),
    openYomitanSettingsWindow: ({ yomitanSession: nextSession }) => {
      forwardedSession = nextSession as { id: string } | null;
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
  assert.equal(forwardedSession, yomitanSession);
});

test('yomitan opener opens settings window without a dedicated session', async () => {
  let forwardedSession: unknown = 'unset';
  const logs: string[] = [];
  const openSettings = createOpenYomitanSettingsHandler({
    ensureYomitanExtensionLoaded: async () => ({ id: 'ext' }),
    openYomitanSettingsWindow: ({ yomitanSession: nextSession }) => {
      forwardedSession = nextSession;
    },
    getExistingWindow: () => null,
    setWindow: () => {},
    getYomitanSession: () => null,
    logWarn: (message) => logs.push(message),
    logError: () => logs.push('error'),
  });

  openSettings();
  await Promise.resolve();
  await Promise.resolve();

  assert.equal(forwardedSession, null);
  assert.deepEqual(logs, []);
});

test('yomitan opener does not start settings-triggered extension load while startup load is in flight', async () => {
  let ensureCalled = false;
  const logs: string[] = [];
  const startupLoad = new Promise<unknown>(() => {});
  const openSettings = createOpenYomitanSettingsHandler({
    ensureYomitanExtensionLoaded: async () => {
      ensureCalled = true;
      return { id: 'ext' };
    },
    getYomitanExtension: () => null,
    getYomitanExtensionLoadInFlight: () => startupLoad,
    openYomitanSettingsWindow: () => {
      throw new Error('should not open while startup load is in flight');
    },
    getExistingWindow: () => null,
    setWindow: () => {},
    logWarn: (message) => logs.push(message),
    logError: () => logs.push('error'),
  });

  openSettings();
  await Promise.resolve();
  await Promise.resolve();

  assert.equal(ensureCalled, false);
  assert.deepEqual(logs, [
    'Yomitan settings requested while Yomitan is still loading. Try again in a few seconds.',
  ]);
});

test('yomitan opener uses loaded extension from app state without calling loader', async () => {
  let forwardedExtension: { id: string } | null = null;
  const appStateExtension = { id: 'loaded-ext' };
  const openSettings = createOpenYomitanSettingsHandler({
    ensureYomitanExtensionLoaded: async () => {
      throw new Error('should not load extension from settings click');
    },
    getYomitanExtension: () => appStateExtension,
    getYomitanExtensionLoadInFlight: () => null,
    openYomitanSettingsWindow: ({ yomitanExt }) => {
      forwardedExtension = yomitanExt as { id: string };
    },
    getExistingWindow: () => null,
    setWindow: () => {},
    logWarn: () => {},
    logError: () => {},
  });

  openSettings();
  await Promise.resolve();

  assert.equal(forwardedExtension, appStateExtension);
});

test('yomitan opener lazy-loads extension when app state is empty and no load is in flight', async () => {
  let ensureCalled = false;
  let forwardedExtension: { id: string } | null = null;
  const openSettings = createOpenYomitanSettingsHandler({
    ensureYomitanExtensionLoaded: async () => {
      ensureCalled = true;
      return { id: 'ext' };
    },
    getYomitanExtension: () => null,
    getYomitanExtensionLoadInFlight: () => null,
    openYomitanSettingsWindow: ({ yomitanExt }) => {
      forwardedExtension = yomitanExt as { id: string };
    },
    getExistingWindow: () => null,
    setWindow: () => {},
    logWarn: () => {},
    logError: () => {},
  });

  openSettings();
  await Promise.resolve();
  await Promise.resolve();

  assert.equal(ensureCalled, true);
  assert.deepEqual(forwardedExtension, { id: 'ext' });
});
