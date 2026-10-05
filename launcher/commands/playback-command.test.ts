import test from 'node:test';
import assert from 'node:assert/strict';
import { EventEmitter } from 'node:events';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import type { Args } from '../types.js';
import type { LauncherCommandContext } from './context.js';
import { registerCleanup, runPlaybackCommandWithDeps } from './playback-command.js';
import { state, type startMpv } from '../mpv.js';
import { makeLauncherArgs } from '../test-support/args.js';
import { withEnv } from '../test-support/env.js';

type PlaybackDeps = Parameters<typeof runPlaybackCommandWithDeps>[1];
type StartMpvOptions = NonNullable<Parameters<typeof startMpv>[6]>;

const YOUTUBE_URL = 'https://www.youtube.com/watch?v=65Ovd7t8sNw';

function makePluginConfig(
  overrides: Partial<LauncherCommandContext['pluginRuntimeConfig']> = {},
): LauncherCommandContext['pluginRuntimeConfig'] {
  return {
    socketPath: '/tmp/subminer.sock',
    binaryPath: '',
    backend: 'auto',
    autoStart: true,
    autoStartVisibleOverlay: true,
    autoStartPauseUntilReady: true,
    osdMessages: false,
    texthookerEnabled: false,
    ...overrides,
  };
}

/** A context for a YouTube URL target with logging at info and CLI texthooker off. */
function createContext(
  options: {
    args?: Partial<Args>;
    plugin?: Partial<LauncherCommandContext['pluginRuntimeConfig']>;
  } = {},
): LauncherCommandContext {
  return {
    args: makeLauncherArgs({
      target: YOUTUBE_URL,
      targetKind: 'url',
      logLevel: 'info',
      useTexthooker: false,
      ...options.args,
    }),
    scriptPath: '/tmp/subminer',
    scriptName: 'subminer',
    mpvSocketPath: '/tmp/subminer.sock',
    pluginRuntimeConfig: makePluginConfig(options.plugin),
    appPath: '/tmp/SubMiner.AppImage',
    launcherJellyfinConfig: {},
    processAdapter: {
      platform: () => 'linux',
      onSignal: () => {},
      writeStdout: () => {},
      exit: (_code: number): never => {
        throw new Error('unexpected exit');
      },
      setExitCode: () => {},
    },
  };
}

const FILE_TARGET = { target: '/tmp/movie.mkv', targetKind: 'file' } as const;

/**
 * Playback harness. Deps are inert by default and record what the tests assert on: the
 * options mpv was started with, the overlay launches, and an `events` log for the few cases
 * where relative order is the behavior.
 */
function makeHarness(
  options: {
    context?: LauncherCommandContext;
    deps?: Partial<PlaybackDeps>;
    events?: string[];
  } = {},
) {
  const context = options.context ?? createContext();
  const events = options.events ?? [];
  const startMpvOptions: StartMpvOptions[] = [];
  const overlayCalls: Array<{ extraAppArgs: string[]; configDir?: string }> = [];
  const deps: PlaybackDeps = {
    ensurePlaybackSetupReady: async () => {},
    ensureRuntimePluginReady: async () => {
      events.push('plugin');
    },
    chooseTarget: async () => ({
      target: context.args.target,
      kind: context.args.targetKind === 'file' ? 'file' : 'url',
    }),
    checkDependencies: () => {},
    registerCleanup: () => {},
    startMpv: async (_target, _kind, _args, _socket, _appPath, _subtitles, startOptions) => {
      events.push('startMpv');
      if (startOptions) startMpvOptions.push(startOptions);
    },
    waitForUnixSocketReady: async () => true,
    startOverlay: async (_appPath, _args, _socket, extraAppArgs = [], configDir) => {
      events.push('startOverlay');
      overlayCalls.push({ extraAppArgs, configDir });
    },
    launchAppCommandDetached: () => {},
    log: () => {},
    cleanupPlaybackSession: async () => {},
    getMpvProc: () => null,
    ...options.deps,
  };
  return {
    context,
    events,
    startMpvOptions,
    overlayCalls,
    run: () => runPlaybackCommandWithDeps(context, deps),
  };
}

function assertRanInOrder(events: string[], ...expected: string[]): void {
  const positions = expected.map((event) => events.indexOf(event));
  assert.ok(
    positions.every(
      (position, index) => position >= 0 && (index === 0 || position > positions[index - 1]!),
    ),
    `expected ${expected.join(' -> ')} in order, got ${events.join(', ')}`,
  );
}

test('playback cleanup signal handlers are registered once across repeated sessions', () => {
  const context = createContext();
  const registeredSignals: NodeJS.Signals[] = [];
  context.processAdapter.onSignal = (signal) => {
    registeredSignals.push(signal);
  };

  registerCleanup(context);
  registerCleanup(context);

  assert.deepEqual(registeredSignals, ['SIGINT', 'SIGTERM']);
});

test('youtube playback launches overlay with app-owned youtube flow args', async () => {
  const harness = makeHarness({
    context: createContext({
      plugin: {
        autoStart: false,
        autoStartVisibleOverlay: false,
        autoStartPauseUntilReady: false,
      },
    }),
  });

  await harness.run();

  assertRanInOrder(harness.events, 'startMpv', 'startOverlay');
  assert.deepEqual(harness.overlayCalls[0]?.extraAppArgs, ['--youtube-play', YOUTUBE_URL]);
  assert.equal(harness.startMpvOptions[0]?.startPaused, true);
  assert.equal(harness.startMpvOptions[0]?.disableYoutubeSubtitleAutoLoad, true);
});

test('youtube app-owned playback disables mpv plugin auto-start', async () => {
  const harness = makeHarness();

  await harness.run();

  const runtimeConfig = harness.startMpvOptions[0]?.runtimePluginConfig;
  assert.equal(runtimeConfig?.autoStart, false);
  assert.equal(runtimeConfig?.autoStartVisibleOverlay, false);
  assert.equal(runtimeConfig?.autoStartPauseUntilReady, false);
});

test('plugin auto-start playback leaves app lifetime to managed-playback owner', async () => {
  const context = createContext({
    args: { ...FILE_TARGET, useTexthooker: true },
    plugin: { autoStartVisibleOverlay: false, autoStartPauseUntilReady: false },
  });
  state.appPath = context.appPath ?? '';
  state.overlayManagedByLauncher = false;
  const mpvProc = new EventEmitter() as EventEmitter & {
    exitCode: number | null;
    killed: boolean;
    kill: () => boolean;
  };
  mpvProc.exitCode = null;
  mpvProc.killed = false;
  mpvProc.kill = () => true;
  let cleanupSawManagedOverlay = true;
  const harness = makeHarness({
    context,
    deps: {
      startMpv: async () => {
        setTimeout(() => {
          mpvProc.exitCode = 0;
          mpvProc.emit('exit', 0);
        }, 5);
      },
      startOverlay: async () => {
        throw new Error('startOverlay should not run when plugin auto-start is used');
      },
      cleanupPlaybackSession: async () => {
        cleanupSawManagedOverlay = state.overlayManagedByLauncher;
      },
      getMpvProc: () => mpvProc as NonNullable<typeof state.mpvProc>,
    },
  });

  try {
    await harness.run();

    assert.equal(cleanupSawManagedOverlay, false);
  } finally {
    state.appPath = '';
    state.overlayManagedByLauncher = false;
  }
});

test('plugin auto-start playback attaches a warm background app through the launcher', async () => {
  const harness = makeHarness({
    context: createContext({
      args: { ...FILE_TARGET, useTexthooker: true },
      plugin: { texthookerEnabled: true },
    }),
    deps: { isAppControlServerAvailable: async () => true },
  });

  await harness.run();

  assertRanInOrder(harness.events, 'startMpv', 'startOverlay');
  assert.deepEqual(harness.overlayCalls[0]?.extraAppArgs, [
    '--show-visible-overlay',
    '--texthooker',
  ]);
  assert.equal(harness.startMpvOptions[0]?.startPaused, true);
  assert.equal(harness.startMpvOptions[0]?.runtimePluginConfig?.autoStart, false);
});

test('plugin auto-start attach mode reuses launcher-resolved config dir for app control', async () => {
  const xdgConfigHome = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-test-xdg-'));
  const expectedConfigDir = path.join(xdgConfigHome, 'SubMiner');
  fs.mkdirSync(expectedConfigDir, { recursive: true });
  fs.writeFileSync(path.join(expectedConfigDir, 'config.jsonc'), '{}');
  let availabilityConfigDir: string | undefined;
  const harness = makeHarness({
    context: createContext({
      args: { ...FILE_TARGET, useTexthooker: true },
      plugin: { texthookerEnabled: true },
    }),
    deps: {
      isAppControlServerAvailable: async (_logLevel, configDir) => {
        availabilityConfigDir = configDir;
        return true;
      },
    },
  });

  try {
    await withEnv({ XDG_CONFIG_HOME: xdgConfigHome, APPDATA: xdgConfigHome }, () => harness.run());

    assert.equal(availabilityConfigDir, expectedConfigDir);
    assert.equal(harness.overlayCalls[0]?.configDir, expectedConfigDir);
    assert.equal(harness.startMpvOptions[0]?.runtimePluginConfig?.overlayLoadingOsd, true);
  } finally {
    fs.rmSync(xdgConfigHome, { recursive: true, force: true });
  }
});

test('plugin auto-start attach mode omits texthooker flag when CLI texthooker is disabled', async () => {
  const harness = makeHarness({
    context: createContext({ args: FILE_TARGET, plugin: { texthookerEnabled: true } }),
    deps: { isAppControlServerAvailable: async () => true },
  });

  await harness.run();

  assert.deepEqual(harness.overlayCalls[0]?.extraAppArgs, ['--show-visible-overlay']);
});

test('playback command ensures Linux runtime plugin before mpv launch', async () => {
  const harness = makeHarness({ context: createContext({ args: FILE_TARGET }) });

  await harness.run();

  assertRanInOrder(harness.events, 'plugin', 'startMpv');
});

test('rofi playback repairs support assets before opening the picker', async () => {
  const events: string[] = [];
  const harness = makeHarness({
    events,
    context: createContext({ args: { target: '', targetKind: '', useRofi: true } }),
    deps: {
      chooseTarget: async () => {
        events.push('picker');
        return { target: '/tmp/movie.mkv', kind: 'file' };
      },
      checkPickerDependencies: () => {},
    },
  });

  await harness.run();

  assertRanInOrder(events, 'plugin', 'picker', 'startMpv');
});
