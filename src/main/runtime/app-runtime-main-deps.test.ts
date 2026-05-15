import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildDestroyTrayMainDepsHandler,
  createBuildEnsureTrayMainDepsHandler,
  createBuildInitializeOverlayRuntimeBootstrapMainDepsHandler,
  createBuildOpenYomitanSettingsMainDepsHandler,
} from './app-runtime-main-deps';

test('ensure tray main deps trigger overlay bootstrap on tray click when runtime not initialized', () => {
  const calls: string[] = [];
  const deps = createBuildEnsureTrayMainDepsHandler({
    getTray: () => null,
    setTray: () => calls.push('set-tray'),
    buildTrayMenu: () => ({}),
    resolveTrayIconPath: () => null,
    createImageFromPath: () => ({}),
    createEmptyImage: () => ({}),
    createTray: () => ({}),
    trayTooltip: 'SubMiner',
    platform: 'darwin',
    logWarn: (message) => calls.push(`warn:${message}`),
    initializeOverlayRuntime: () => calls.push('init-overlay'),
    isOverlayRuntimeInitialized: () => false,
    setVisibleOverlayVisible: (visible) => calls.push(`set-visible:${visible}`),
  })();

  deps.ensureOverlayVisibleFromTrayClick();
  assert.deepEqual(calls, ['init-overlay', 'set-visible:true']);
});

test('destroy tray main deps map passthrough getters/setters', () => {
  let tray: unknown = { id: 'tray' };
  const deps = createBuildDestroyTrayMainDepsHandler({
    getTray: () => tray,
    setTray: (next) => {
      tray = next;
    },
  })();

  assert.deepEqual(deps.getTray(), { id: 'tray' });
  deps.setTray(null);
  assert.equal(tray, null);
});

test('initialize overlay runtime main deps map build options and callbacks', () => {
  const calls: string[] = [];
  const options = { id: 'opts' };
  const deps = createBuildInitializeOverlayRuntimeBootstrapMainDepsHandler({
    isOverlayRuntimeInitialized: () => false,
    initializeOverlayRuntimeCore: (value) => {
      calls.push(`core:${JSON.stringify(value)}`);
    },
    buildOptions: () => options,
    setOverlayRuntimeInitialized: (initialized) => calls.push(`set-initialized:${initialized}`),
    startBackgroundWarmups: () => calls.push('warmups'),
  })();

  assert.equal(deps.isOverlayRuntimeInitialized(), false);
  assert.equal(deps.buildOptions(), options);
  assert.equal(deps.initializeOverlayRuntimeCore(options), undefined);
  deps.setOverlayRuntimeInitialized(true);
  deps.startBackgroundWarmups();
  assert.deepEqual(calls, ['core:{"id":"opts"}', 'set-initialized:true', 'warmups']);
});

test('open yomitan settings main deps map async open callbacks', async () => {
  const calls: string[] = [];
  let currentWindow: unknown = null;
  const extension = { id: 'ext' };
  const startupLoad = Promise.resolve(extension);
  const yomitanSession = { id: 'session' };
  const deps = createBuildOpenYomitanSettingsMainDepsHandler({
    ensureYomitanExtensionLoaded: async () => extension,
    getYomitanExtension: () => extension,
    getYomitanExtensionLoadInFlight: () => startupLoad,
    openYomitanSettingsWindow: ({ yomitanExt, yomitanSession: forwardedSession }) =>
      calls.push(
        `open:${(yomitanExt as { id: string }).id}:${(forwardedSession as { id: string } | null)?.id ?? 'null'}`,
      ),
    getExistingWindow: () => currentWindow,
    setWindow: (window) => {
      currentWindow = window;
      calls.push('set-window');
    },
    getYomitanSession: () => yomitanSession,
    logWarn: (message) => calls.push(`warn:${message}`),
    logError: (message) => calls.push(`error:${message}`),
  })();

  assert.equal(await deps.ensureYomitanExtensionLoaded(), extension);
  assert.equal(deps.getYomitanExtension?.(), extension);
  assert.equal(deps.getYomitanExtensionLoadInFlight?.(), startupLoad);
  assert.equal(deps.getExistingWindow(), null);
  deps.setWindow({ id: 'win' });
  deps.openYomitanSettingsWindow({
    yomitanExt: extension,
    getExistingWindow: () => deps.getExistingWindow(),
    setWindow: (window) => deps.setWindow(window),
    yomitanSession: deps.getYomitanSession(),
  });
  deps.logWarn('warn');
  deps.logError('error', new Error('boom'));
  assert.deepEqual(calls, ['set-window', 'open:ext:session', 'warn:warn', 'error:error']);
  assert.deepEqual(currentWindow, { id: 'win' });
});
