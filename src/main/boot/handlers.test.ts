import assert from 'node:assert/strict';
import test from 'node:test';
import { createMainBootHandlers } from './handlers';

test('createMainBootHandlers returns grouped handler bundles', () => {
  const handlers = createMainBootHandlers<any, any, any, any>({
    startupLifecycleDeps: {
      registerProtocolUrlHandlersMainDeps: {} as never,
      onWillQuitCleanupMainDeps: {} as never,
      shouldRestoreWindowsOnActivateMainDeps: {} as never,
      restoreWindowsOnActivateMainDeps: {} as never,
    },
    ipcRuntimeDeps: {
      mpvCommandMainDeps: {} as never,
      handleMpvCommandFromIpcRuntime: () => ({ ok: true }) as never,
      runSubsyncManualFromIpc: () => Promise.resolve({ ok: true }) as never,
      registration: {
        runtimeOptions: {} as never,
        mainDeps: {} as never,
        ankiJimakuDeps: {} as never,
        registerIpcRuntimeServices: () => {},
      },
    },
    cliStartupDeps: {
      cliCommandContextMainDeps: {} as never,
      cliCommandRuntimeHandlerMainDeps: {} as never,
      initialArgsRuntimeHandlerMainDeps: {} as never,
    },
    headlessStartupDeps: {
      startupRuntimeHandlersDeps: {
        appLifecycleRuntimeRunnerMainDeps: {
          app: { on: () => {} } as never,
          platform: 'darwin',
          shouldStartApp: () => true,
          parseArgs: () => ({}) as never,
          handleCliCommand: () => {},
          printHelp: () => {},
          logNoRunningInstance: () => {},
          onReady: async () => {},
          onWillQuitCleanup: () => {},
          shouldRestoreWindowsOnActivate: () => false,
          restoreWindowsOnActivate: () => {},
          shouldQuitOnWindowAllClosed: () => false,
        },
        createAppLifecycleRuntimeRunner: () => () => {},
        buildStartupBootstrapMainDeps: (startAppLifecycle) => ({
          argv: ['node', 'main.js'],
          parseArgs: () => ({ command: 'start' }) as never,
          setLogLevel: () => {},
          forceX11Backend: () => {},
          enforceUnsupportedWaylandMode: () => {},
          shouldStartApp: () => true,
          getDefaultSocketPath: () => '/tmp/mpv.sock',
          defaultTexthookerPort: 5174,
          configDir: '/tmp/config',
          defaultConfig: {} as never,
          generateConfigTemplate: () => 'template',
          generateDefaultConfigFile: async () => 0,
          setExitCode: () => {},
          quitApp: () => {},
          logGenerateConfigError: () => {},
          startAppLifecycle: (args) => startAppLifecycle(args as never),
        }),
        createStartupBootstrapRuntimeDeps: (deps) => ({
          startAppLifecycle: deps.startAppLifecycle,
        }),
        runStartupBootstrapRuntime: () => ({ mode: 'started' } as never),
        applyStartupState: () => {},
      },
    },
    overlayWindowDeps: {
      createOverlayWindowDeps: {
        createOverlayWindowCore: (kind) => ({ kind }),
        isDev: false,
        ensureOverlayWindowLevel: () => {},
        onRuntimeOptionsChanged: () => {},
        setOverlayDebugVisualizationEnabled: () => {},
        isOverlayVisible: () => false,
        tryHandleOverlayShortcutLocalFallback: () => false,
        forwardTabToMpv: () => {},
        onWindowClosed: () => {},
        getYomitanSession: () => null,
      },
      setMainWindow: () => {},
      setModalWindow: () => {},
    },
  });

  assert.equal(typeof handlers.startupLifecycle.registerProtocolUrlHandlers, 'function');
  assert.equal(typeof handlers.ipcRuntime.registerIpcRuntimeHandlers, 'function');
  assert.equal(typeof handlers.cliStartup.handleCliCommand, 'function');
  assert.equal(typeof handlers.headlessStartup.runAndApplyStartupState, 'function');
  assert.equal(typeof handlers.overlayWindow.createMainWindow, 'function');
});
