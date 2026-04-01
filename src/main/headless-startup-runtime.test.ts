import assert from 'node:assert/strict';
import test from 'node:test';
import type { CliArgs } from '../cli/args';
import type { LogLevelSource } from '../logger';
import type { StartupBootstrapRuntimeFactoryDeps } from './startup';

import { createHeadlessStartupRuntime } from './headless-startup-runtime';

test('headless startup runtime returns callable handlers and applies startup state', () => {
  const calls: string[] = [];

  const runtime = createHeadlessStartupRuntime<
    { mode: string },
    { startAppLifecycle: (args: CliArgs) => void }
  >({
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
    bootstrap: {
      argv: ['node', 'main.js'],
      parseArgs: () => ({ command: 'start' }) as never,
      setLogLevel: (_level: string, _source: LogLevelSource) => {},
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
      startAppLifecycle: (args: CliArgs) => {
        calls.push(`bootstrap:${(args as { command?: string }).command ?? 'unknown'}`);
      },
    },
    createAppLifecycleRuntimeRunner: () => (args: CliArgs) => {
      calls.push(`lifecycle:${(args as { command?: string }).command ?? 'unknown'}`);
    },
    createStartupBootstrapRuntimeDeps: (deps: StartupBootstrapRuntimeFactoryDeps) => ({
      startAppLifecycle: deps.startAppLifecycle,
    }),
    runStartupBootstrapRuntime: (deps) => {
      deps.startAppLifecycle({ command: 'start' } as unknown as CliArgs);
      return { mode: 'started' };
    },
    applyStartupState: (state: { mode: string }) => {
      calls.push(`apply:${state.mode}`);
    },
  });

  assert.equal(typeof runtime.appLifecycleRuntimeRunner, 'function');
  assert.equal(typeof runtime.runAndApplyStartupState, 'function');
  assert.deepEqual(runtime.runAndApplyStartupState(), { mode: 'started' });
  assert.deepEqual(calls, ['lifecycle:start', 'apply:started']);
});

test('headless startup runtime accepts grouped app lifecycle input', () => {
  const calls: string[] = [];

  const runtime = createHeadlessStartupRuntime<
    { mode: string },
    { startAppLifecycle: (args: CliArgs) => void }
  >({
    appLifecycle: {
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
    bootstrap: {
      argv: ['node', 'main.js'],
      parseArgs: () => ({ command: 'start' }) as never,
      setLogLevel: (_level: string, _source: LogLevelSource) => {},
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
      startAppLifecycle: (args: CliArgs) => {
        calls.push(`bootstrap:${(args as { command?: string }).command ?? 'unknown'}`);
      },
    },
    createAppLifecycleRuntimeRunner: () => (args: CliArgs) => {
      calls.push(`lifecycle:${(args as { command?: string }).command ?? 'unknown'}`);
    },
    createStartupBootstrapRuntimeDeps: (deps: StartupBootstrapRuntimeFactoryDeps) => ({
      startAppLifecycle: deps.startAppLifecycle,
    }),
    runStartupBootstrapRuntime: (deps) => {
      deps.startAppLifecycle({ command: 'start' } as unknown as CliArgs);
      return { mode: 'started' };
    },
    applyStartupState: (state: { mode: string }) => {
      calls.push(`apply:${state.mode}`);
    },
  });

  runtime.appLifecycleRuntimeRunner({ command: 'start' } as unknown as CliArgs);

  assert.deepEqual(runtime.runAndApplyStartupState(), { mode: 'started' });
  assert.deepEqual(calls, ['lifecycle:start', 'lifecycle:start', 'apply:started']);
});
