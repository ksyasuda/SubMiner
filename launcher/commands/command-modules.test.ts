import test from 'node:test';
import assert from 'node:assert/strict';
import { parseArgs } from '../config.js';
import type { ProcessAdapter } from '../process-adapter.js';
import type { LauncherCommandContext } from './context.js';
import { runConfigCommand } from './config-command.js';
import { runDictionaryCommand } from './dictionary-command.js';
import { runDoctorCommand } from './doctor-command.js';
import { runLogsCommand } from './logs-command.js';
import { runMpvPreAppCommand } from './mpv-command.js';
import { runAppPassthroughCommand } from './app-command.js';
import { runStatsCommand } from './stats-command.js';
import { runUpdateCommand } from './update-command.js';

class ExitSignal extends Error {
  code: number;

  constructor(code: number) {
    super(`exit:${code}`);
    this.code = code;
  }
}

function createContext(overrides: Partial<LauncherCommandContext> = {}): LauncherCommandContext {
  const args = parseArgs([], 'subminer', {});
  const adapter: ProcessAdapter = {
    platform: () => 'linux',
    onSignal: () => {},
    writeStdout: () => {},
    exit: (code) => {
      throw new ExitSignal(code);
    },
    setExitCode: () => {},
  };

  return {
    args,
    scriptPath: '/tmp/subminer',
    scriptName: 'subminer',
    mpvSocketPath: '/tmp/subminer.sock',
    pluginRuntimeConfig: {
      socketPath: '/tmp/subminer.sock',
      binaryPath: '',
      backend: 'auto',
      autoStart: true,
      autoStartVisibleOverlay: true,
      autoStartPauseUntilReady: true,
      texthookerEnabled: false,
      aniskipEnabled: true,
      aniskipButtonKey: 'TAB',
    },
    appPath: '/tmp/subminer.app',
    launcherJellyfinConfig: {},
    processAdapter: adapter,
    ...overrides,
  };
}

type StatsTestArgOverrides = {
  stats?: boolean;
  statsBackground?: boolean;
  statsCleanup?: boolean;
  statsCleanupVocab?: boolean;
  statsCleanupLifetime?: boolean;
  statsStop?: boolean;
  logLevel?: LauncherCommandContext['args']['logLevel'];
};

function createStatsTestHarness(overrides: StatsTestArgOverrides = {}) {
  const context = createContext();
  const forwarded: string[][] = [];
  const removedPaths: string[] = [];
  const createTempDir = (_prefix: string) => {
    const created = `/tmp/subminer-stats-test`;
    return created;
  };
  const joinPath = (...parts: string[]) => parts.join('/');
  const removeDir = (targetPath: string) => {
    removedPaths.push(targetPath);
  };
  const runAppCommandAttachedStub = async (
    _appPath: string,
    appArgs: string[],
    _logLevel: LauncherCommandContext['args']['logLevel'],
    _label: string,
  ) => {
    forwarded.push(appArgs);
    return 0;
  };
  const waitForStatsResponseStub = async () => ({ ok: true, url: 'http://127.0.0.1:5175' });

  context.args = {
    ...context.args,
    stats: true,
    ...overrides,
  };

  return {
    context,
    forwarded,
    removedPaths,
    createTempDir,
    joinPath,
    removeDir,
    runAppCommandAttachedStub,
    waitForStatsResponseStub,
    commandDeps: {
      createTempDir,
      joinPath,
      runAppCommandAttached: runAppCommandAttachedStub,
      waitForStatsResponse: waitForStatsResponseStub,
      removeDir,
    },
  };
}

test('config command writes newline-terminated path via process adapter', () => {
  const writes: string[] = [];
  const context = createContext();
  context.args.configPath = true;
  context.processAdapter = {
    ...context.processAdapter,
    writeStdout: (text) => writes.push(text),
  };

  const handled = runConfigCommand(context, {
    existsSync: () => true,
    readFileSync: () => '',
    resolveMainConfigPath: () => '/tmp/SubMiner/config.jsonc',
  });

  assert.equal(handled, true);
  assert.deepEqual(writes, ['/tmp/SubMiner/config.jsonc\n']);
});

test('doctor command exits non-zero for missing hard dependencies', () => {
  const context = createContext({ appPath: null });
  context.args.doctor = true;

  assert.throws(
    () =>
      runDoctorCommand(context, {
        commandExists: () => false,
        configExists: () => true,
        resolveMainConfigPath: () => '/tmp/SubMiner/config.jsonc',
        runAppCommandWithInherit: () => {
          throw new Error('unexpected app handoff');
        },
      }),
    (error: unknown) => error instanceof ExitSignal && error.code === 1,
  );
});

test('doctor command forwards refresh-known-words to app binary', () => {
  const context = createContext();
  context.args.doctor = true;
  context.args.doctorRefreshKnownWords = true;
  const forwarded: string[][] = [];

  const handled = runDoctorCommand(context, {
    commandExists: () => false,
    configExists: () => true,
    resolveMainConfigPath: () => '/tmp/SubMiner/config.jsonc',
    runAppCommandWithInherit: (_appPath, appArgs) => {
      forwarded.push(appArgs);
    },
  });

  assert.equal(handled, true);
  assert.deepEqual(forwarded, [['--refresh-known-words']]);
});

test('logs command exports logs and writes archive path', () => {
  const writes: string[] = [];
  const context = createContext();
  context.args.logsExport = true;
  context.processAdapter = {
    ...context.processAdapter,
    writeStdout: (text) => writes.push(text),
  };

  const handled = runLogsCommand(context, {
    exportLogsArchive: () => ({
      zipPath: '/tmp/subminer-logs.zip',
      exportedFiles: ['/tmp/app.log'],
      mode: 'current-day',
    }),
  });

  assert.equal(handled, true);
  assert.deepEqual(writes, ['/tmp/subminer-logs.zip\n']);
});

test('logs command ignores unrelated launcher commands', () => {
  const context = createContext();

  assert.equal(runLogsCommand(context), false);
});

test('app command starts default macOS background app detached from launcher', () => {
  const context = createContext();
  context.args.appPassthrough = true;
  context.args.appArgs = [];
  const calls: string[] = [];

  const handled = runAppPassthroughCommand(context, {
    runAppCommandWithInherit: () => {
      calls.push('attached');
    },
    launchAppBackgroundDetached: (appPath, logLevel) => {
      calls.push(`detached:${appPath}:${logLevel}`);
    },
  });

  assert.equal(handled, true);
  assert.deepEqual(calls, ['detached:/tmp/subminer.app:warn']);
});

test('app command starts default Linux background app detached from launcher', () => {
  const context = createContext();
  context.args.appPassthrough = true;
  context.args.appArgs = [];
  const calls: string[] = [];

  const handled = runAppPassthroughCommand(context, {
    runAppCommandWithInherit: () => {
      calls.push('attached');
    },
    launchAppBackgroundDetached: (appPath, logLevel) => {
      calls.push(`detached:${appPath}:${logLevel}`);
    },
  });

  assert.equal(handled, true);
  assert.deepEqual(calls, ['detached:/tmp/subminer.app:warn']);
});

test('app command keeps explicit passthrough args attached', () => {
  const context = createContext();
  context.args.appPassthrough = true;
  context.args.appArgs = ['--settings'];
  const forwarded: string[][] = [];
  const detached: string[] = [];

  const handled = runAppPassthroughCommand(context, {
    runAppCommandWithInherit: (_appPath, appArgs) => {
      forwarded.push(appArgs);
    },
    launchAppBackgroundDetached: () => {
      detached.push('detached');
    },
  });

  assert.equal(handled, true);
  assert.deepEqual(forwarded, [['--settings']]);
  assert.deepEqual(detached, []);
});

test('mpv pre-app command exits non-zero when socket is not ready', async () => {
  const context = createContext();
  context.args.mpvStatus = true;

  await assert.rejects(
    async () => {
      await runMpvPreAppCommand(context, {
        waitForUnixSocketReady: async () => false,
        launchMpvIdleDetached: async () => {},
      });
    },
    (error: unknown) => error instanceof ExitSignal && error.code === 1,
  );
});

test('dictionary command forwards --dictionary and target path to app binary', () => {
  const context = createContext();
  context.args.dictionary = true;
  context.args.dictionaryTarget = '/tmp/anime';
  const forwarded: string[][] = [];

  const handled = runDictionaryCommand(context, {
    runAppCommandWithInherit: (_appPath, appArgs) => {
      forwarded.push(appArgs);
    },
  });

  assert.equal(handled, true);
  assert.deepEqual(forwarded, [['--start', '--dictionary', '--dictionary-target', '/tmp/anime']]);
});

test('dictionary command forwards candidate and selection modes to app binary', () => {
  const candidatesContext = createContext();
  candidatesContext.args.dictionary = true;
  candidatesContext.args.dictionaryCandidates = true;
  candidatesContext.args.dictionaryTarget = '/tmp/anime.mkv';
  const selectContext = createContext();
  selectContext.args.dictionary = true;
  selectContext.args.dictionarySelect = true;
  selectContext.args.dictionaryAnilistId = 21355;
  selectContext.args.dictionaryTarget = '/tmp/anime.mkv';
  const forwarded: string[][] = [];

  runDictionaryCommand(candidatesContext, {
    runAppCommandWithInherit: (_appPath, appArgs) => {
      forwarded.push(appArgs);
    },
  });
  runDictionaryCommand(selectContext, {
    runAppCommandWithInherit: (_appPath, appArgs) => {
      forwarded.push(appArgs);
    },
  });

  assert.deepEqual(forwarded, [
    ['--start', '--dictionary-candidates', '--dictionary-target', '/tmp/anime.mkv'],
    [
      '--start',
      '--dictionary-select',
      '--dictionary-anilist-id',
      '21355',
      '--dictionary-target',
      '/tmp/anime.mkv',
    ],
  ]);
});

test('dictionary command returns after app handoff starts', () => {
  const context = createContext();
  context.args.dictionary = true;

  const handled = runDictionaryCommand(context, {
    runAppCommandWithInherit: () => undefined,
  });

  assert.equal(handled, true);
});

test('update command runs direct Linux release update without launching Electron', async () => {
  const context = createContext();
  context.args.update = true;
  const calls: string[] = [];

  const handled = await runUpdateCommand(context, {
    runAppCommandCaptureOutput: () => {
      throw new Error('unexpected Electron launch');
    },
    runDirectReleaseUpdate: async (request) => {
      calls.push(`direct:${request.appPath}:${request.launcherPath}:${request.channel}`);
      return {
        appImage: { status: 'not-found' },
        launcher: { status: 'updated' },
        supportAssets: [{ status: 'skipped' }],
      };
    },
    readMainConfig: () => null,
    log: (level, _configured, message) => {
      calls.push(`${level}:${message}`);
    },
  });

  assert.equal(handled, true);
  assert.deepEqual(calls, [
    'direct:/tmp/subminer.app:/tmp/subminer:stable',
    'info:AppImage update: not-found',
    'info:Launcher update: updated',
    'info:Rofi theme update: skipped',
  ]);
});

test('stats command launches attached app command with response path', async () => {
  const harness = createStatsTestHarness({ stats: true, logLevel: 'debug' });
  const handled = await runStatsCommand(harness.context, harness.commandDeps);

  assert.equal(handled, true);
  assert.deepEqual(harness.forwarded, [
    [
      '--stats',
      '--stats-response-path',
      '/tmp/subminer-stats-test/response.json',
      '--log-level',
      'debug',
    ],
  ]);
  assert.equal(harness.removedPaths.length, 1);
});

test('stats background command launches attached daemon control command with response path', async () => {
  const harness = createStatsTestHarness({ stats: true, statsBackground: true });
  const handled = await runStatsCommand(harness.context, harness.commandDeps);

  assert.equal(handled, true);
  assert.deepEqual(harness.forwarded, [
    ['--stats-daemon-start', '--stats-response-path', '/tmp/subminer-stats-test/response.json'],
  ]);
  assert.equal(harness.removedPaths.length, 1);
});

test('stats command waits for attached app exit after startup response', async () => {
  const harness = createStatsTestHarness({ stats: true });
  const started = new Promise<number>((resolve) => setTimeout(() => resolve(0), 20));

  const statsCommand = runStatsCommand(harness.context, {
    ...harness.commandDeps,
    runAppCommandAttached: async (...args) => {
      await harness.runAppCommandAttachedStub(...args);
      return started;
    },
  });
  const result = await Promise.race([
    statsCommand.then(() => 'resolved'),
    new Promise<'timeout'>((resolve) => setTimeout(() => resolve('timeout'), 5)),
  ]);

  assert.equal(result, 'timeout');

  const final = await statsCommand;
  assert.equal(final, true);
  assert.deepEqual(harness.forwarded, [
    ['--stats', '--stats-response-path', '/tmp/subminer-stats-test/response.json'],
  ]);
  assert.equal(harness.removedPaths.length, 1);
});

test('stats command throws when attached app exits non-zero after startup response', async () => {
  const harness = createStatsTestHarness({ stats: true });

  await assert.rejects(async () => {
    await runStatsCommand(harness.context, {
      ...harness.commandDeps,
      runAppCommandAttached: async (...args) => {
        await harness.runAppCommandAttachedStub(...args);
        await new Promise((resolve) => setTimeout(resolve, 10));
        return 3;
      },
    });
  }, /Stats app exited with status 3\./);

  assert.equal(harness.removedPaths.length, 1);
});

test('stats cleanup command forwards cleanup vocab flags to the app', async () => {
  const harness = createStatsTestHarness({
    stats: true,
    statsCleanup: true,
    statsCleanupVocab: true,
  });
  const handled = await runStatsCommand(harness.context, {
    ...harness.commandDeps,
    waitForStatsResponse: async () => ({ ok: true }),
  });

  assert.equal(handled, true);
  assert.deepEqual(harness.forwarded, [
    [
      '--stats',
      '--stats-response-path',
      '/tmp/subminer-stats-test/response.json',
      '--stats-cleanup',
      '--stats-cleanup-vocab',
    ],
  ]);
  assert.equal(harness.removedPaths.length, 1);
});

test('stats stop command forwards stop flag to the app', async () => {
  const harness = createStatsTestHarness({ stats: true, statsStop: true });

  const handled = await runStatsCommand(harness.context, {
    ...harness.commandDeps,
    waitForStatsResponse: async () => ({ ok: true }),
  });

  assert.equal(handled, true);
  assert.deepEqual(harness.forwarded, [
    ['--stats-daemon-stop', '--stats-response-path', '/tmp/subminer-stats-test/response.json'],
  ]);
  assert.equal(harness.removedPaths.length, 1);
});

test('stats stop command exits on process exit without waiting for startup response', async () => {
  const harness = createStatsTestHarness({ stats: true, statsStop: true });
  let waitedForResponse = false;

  const handled = await runStatsCommand(harness.context, {
    ...harness.commandDeps,
    runAppCommandAttached: async (...args) => {
      await harness.runAppCommandAttachedStub(...args);
      return 0;
    },
    waitForStatsResponse: async () => {
      waitedForResponse = true;
      return { ok: true };
    },
  });

  assert.equal(handled, true);
  assert.equal(waitedForResponse, false);
  assert.equal(harness.removedPaths.length, 1);
});

test('stats cleanup command forwards lifetime rebuild flag to the app', async () => {
  const harness = createStatsTestHarness({
    stats: true,
    statsCleanup: true,
    statsCleanupLifetime: true,
  });
  const handled = await runStatsCommand(harness.context, {
    ...harness.commandDeps,
    waitForStatsResponse: async () => ({ ok: true }),
  });

  assert.equal(handled, true);
  assert.deepEqual(harness.forwarded, [
    [
      '--stats',
      '--stats-response-path',
      '/tmp/subminer-stats-test/response.json',
      '--stats-cleanup',
      '--stats-cleanup-lifetime',
    ],
  ]);
  assert.equal(harness.removedPaths.length, 1);
});

test('stats command throws when stats response reports an error', async () => {
  const harness = createStatsTestHarness({ stats: true });

  await assert.rejects(async () => {
    await runStatsCommand(harness.context, {
      ...harness.commandDeps,
      runAppCommandAttached: async (...args) => {
        await harness.runAppCommandAttachedStub(...args);
        return 0;
      },
      waitForStatsResponse: async () => ({
        ok: false,
        error: 'Immersion tracking is disabled in config.',
      }),
    });
  }, /Immersion tracking is disabled in config\./);

  assert.equal(harness.removedPaths.length, 1);
});

test('stats cleanup command fails if attached app exits before startup response', async () => {
  const harness = createStatsTestHarness({
    stats: true,
    statsCleanup: true,
    statsCleanupVocab: true,
  });

  await assert.rejects(async () => {
    await runStatsCommand(harness.context, {
      ...harness.commandDeps,
      runAppCommandAttached: async (...args) => {
        await harness.runAppCommandAttachedStub(...args);
        return 2;
      },
      waitForStatsResponse: async () => {
        await new Promise((resolve) => setTimeout(resolve, 25));
        return { ok: true, url: 'http://127.0.0.1:5175' };
      },
    });
  }, /Stats app exited before startup response \(status 2\)\./);

  assert.equal(harness.removedPaths.length, 1);
});

test('stats command aborts pending response wait when app exits before startup response', async () => {
  const harness = createStatsTestHarness({ stats: true });
  let aborted = false;

  await assert.rejects(async () => {
    await runStatsCommand(harness.context, {
      ...harness.commandDeps,
      runAppCommandAttached: async (...args) => {
        await harness.runAppCommandAttachedStub(...args);
        return 2;
      },
      waitForStatsResponse: async (_responsePath, signal) =>
        await new Promise((resolve) => {
          signal?.addEventListener(
            'abort',
            () => {
              aborted = true;
              resolve({ ok: false, error: 'aborted' });
            },
            { once: true },
          );
        }),
    });
  }, /Stats app exited before startup response \(status 2\)\./);

  assert.equal(aborted, true);
  assert.equal(harness.removedPaths.length, 1);
});

test('stats command aborts pending response wait when attached app fails to spawn', async () => {
  const harness = createStatsTestHarness({ stats: true });
  const spawnError = new Error('spawn failed');
  let aborted = false;

  await assert.rejects(
    async () => {
      await runStatsCommand(harness.context, {
        ...harness.commandDeps,
        runAppCommandAttached: async (...args) => {
          await harness.runAppCommandAttachedStub(...args);
          throw spawnError;
        },
        waitForStatsResponse: async (_responsePath, signal) =>
          await new Promise((resolve) => {
            signal?.addEventListener(
              'abort',
              () => {
                aborted = true;
                resolve({ ok: false, error: 'aborted' });
              },
              { once: true },
            );
          }),
      });
    },
    (error: unknown) => error === spawnError,
  );

  assert.equal(aborted, true);
  assert.equal(harness.removedPaths.length, 1);
});

test('stats cleanup command aborts pending response wait when app exits before startup response', async () => {
  const harness = createStatsTestHarness({
    stats: true,
    statsCleanup: true,
    statsCleanupVocab: true,
  });
  let aborted = false;

  await assert.rejects(async () => {
    await runStatsCommand(harness.context, {
      ...harness.commandDeps,
      runAppCommandAttached: async (...args) => {
        await harness.runAppCommandAttachedStub(...args);
        return 2;
      },
      waitForStatsResponse: async (_responsePath, signal) =>
        await new Promise((resolve) => {
          signal?.addEventListener(
            'abort',
            () => {
              aborted = true;
              resolve({ ok: false, error: 'aborted' });
            },
            { once: true },
          );
        }),
    });
  }, /Stats app exited before startup response \(status 2\)\./);

  assert.equal(aborted, true);
  assert.equal(harness.removedPaths.length, 1);
});
