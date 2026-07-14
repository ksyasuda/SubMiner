import assert from 'node:assert/strict';
import test from 'node:test';
import type { CliArgs } from '../../cli/args';
import { createStartupRuntimeHandlers } from './startup-runtime-handlers';

test('startup runtime handlers compose lifecycle runner and startup bootstrap apply flow', () => {
  const calls: string[] = [];
  const appliedStates: Array<{ mode: string }> = [];

  const runtime = createStartupRuntimeHandlers<
    { command: string },
    { mode: string },
    { startAppLifecycle: (args: CliArgs) => void }
  >({
    appLifecycleRuntimeRunnerMainDeps: {
      app: { on: () => {} } as never,
      platform: 'darwin',
      shouldStartApp: () => true,
      parseArgs: () => ({}) as never,
      handleCliCommand: () => calls.push('handle-cli'),
      printHelp: () => calls.push('help'),
      logNoRunningInstance: () => calls.push('no-instance'),
      onReady: async () => {
        calls.push('ready');
      },
      onWillQuitCleanup: () => {
        calls.push('cleanup');
      },
      shouldRestoreWindowsOnActivate: () => true,
      restoreWindowsOnActivate: () => calls.push('restore'),
      shouldQuitOnWindowAllClosed: () => false,
    },
    createAppLifecycleRuntimeRunner: () => (args) => {
      calls.push(`lifecycle:${args.command}`);
    },
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
      startAppLifecycle: (args) => {
        const typedArgs = args as unknown as { command: string };
        calls.push(`start-lifecycle:${typedArgs.command}`);
        startAppLifecycle(typedArgs);
      },
    }),
    createStartupBootstrapRuntimeDeps: (deps) => ({
      startAppLifecycle: deps.startAppLifecycle,
    }),
    runStartupBootstrapRuntime: (deps) => {
      deps.startAppLifecycle({ command: 'start' } as never);
      return { mode: 'started' };
    },
    applyStartupState: (state) => {
      appliedStates.push(state);
    },
  });

  const state = runtime.runAndApplyStartupState();
  assert.deepEqual(state, { mode: 'started' });
  assert.deepEqual(appliedStates, [{ mode: 'started' }]);
  assert.deepEqual(calls, ['start-lifecycle:start', 'lifecycle:start']);
});
