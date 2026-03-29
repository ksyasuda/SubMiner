import assert from 'node:assert/strict';
import test from 'node:test';
import type { CliArgs } from '../../../cli/args';
import { composeHeadlessStartupHandlers } from './headless-startup-composer';

test('composeHeadlessStartupHandlers returns startup bootstrap handlers', () => {
  const calls: string[] = [];

  const handlers = composeHeadlessStartupHandlers<
    CliArgs,
    { mode: string },
    { startAppLifecycle: (args: CliArgs) => void }
  >({
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
      createAppLifecycleRuntimeRunner: () => (args) => {
        calls.push(`lifecycle:${(args as { command?: string }).command ?? 'unknown'}`);
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
        startAppLifecycle: (args) => startAppLifecycle(args as never),
      }),
      createStartupBootstrapRuntimeDeps: (deps) => ({
        startAppLifecycle: deps.startAppLifecycle,
      }),
      runStartupBootstrapRuntime: (deps) => {
        deps.startAppLifecycle({ command: 'start' } as unknown as CliArgs);
        return { mode: 'started' };
      },
      applyStartupState: (state) => {
        calls.push(`apply:${state.mode}`);
      },
    },
  });

  assert.equal(typeof handlers.runAndApplyStartupState, 'function');
  assert.deepEqual(handlers.runAndApplyStartupState(), { mode: 'started' });
  assert.deepEqual(calls, ['lifecycle:start', 'apply:started']);
});
