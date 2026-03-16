import test from 'node:test';
import assert from 'node:assert/strict';
import { parseArgs } from '../config.js';
import type { ProcessAdapter } from '../process-adapter.js';
import type { LauncherCommandContext } from './context.js';
import { runConfigCommand } from './config-command.js';
import { runDictionaryCommand } from './dictionary-command.js';
import { runDoctorCommand } from './doctor-command.js';
import { runMpvPreAppCommand } from './mpv-command.js';
import { runStatsCommand } from './stats-command.js';

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
      autoStart: true,
      autoStartVisibleOverlay: true,
      autoStartPauseUntilReady: true,
    },
    appPath: '/tmp/subminer.app',
    launcherJellyfinConfig: {},
    processAdapter: adapter,
    ...overrides,
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
      }),
    (error: unknown) => error instanceof ExitSignal && error.code === 1,
  );
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

  assert.throws(
    () =>
      runDictionaryCommand(context, {
        runAppCommandWithInherit: (_appPath, appArgs) => {
          forwarded.push(appArgs);
          throw new ExitSignal(0);
        },
      }),
    (error: unknown) => error instanceof ExitSignal && error.code === 0,
  );

  assert.deepEqual(forwarded, [['--dictionary', '--dictionary-target', '/tmp/anime']]);
});

test('dictionary command throws if app handoff unexpectedly returns', () => {
  const context = createContext();
  context.args.dictionary = true;

  assert.throws(
    () =>
      runDictionaryCommand(context, {
        runAppCommandWithInherit: () => undefined as never,
      }),
    /unexpectedly returned/,
  );
});

test('stats command launches attached app command with response path', async () => {
  const context = createContext();
  context.args.stats = true;
  context.args.logLevel = 'debug';
  const forwarded: string[][] = [];

  const handled = await runStatsCommand(context, {
    createTempDir: () => '/tmp/subminer-stats-test',
    joinPath: (...parts) => parts.join('/'),
    runAppCommandAttached: async (_appPath, appArgs) => {
      forwarded.push(appArgs);
      return 0;
    },
    waitForStatsResponse: async () => ({ ok: true, url: 'http://127.0.0.1:5175' }),
    removeDir: () => {},
  });

  assert.equal(handled, true);
  assert.deepEqual(forwarded, [
    [
      '--stats',
      '--stats-response-path',
      '/tmp/subminer-stats-test/response.json',
      '--log-level',
      'debug',
    ],
  ]);
});

test('stats cleanup command forwards cleanup vocab flags to the app', async () => {
  const context = createContext();
  context.args.stats = true;
  context.args.statsCleanup = true;
  context.args.statsCleanupVocab = true;
  const forwarded: string[][] = [];

  const handled = await runStatsCommand(context, {
    createTempDir: () => '/tmp/subminer-stats-test',
    joinPath: (...parts) => parts.join('/'),
    runAppCommandAttached: async (_appPath, appArgs) => {
      forwarded.push(appArgs);
      return 0;
    },
    waitForStatsResponse: async () => ({ ok: true }),
    removeDir: () => {},
  });

  assert.equal(handled, true);
  assert.deepEqual(forwarded, [
    [
      '--stats',
      '--stats-response-path',
      '/tmp/subminer-stats-test/response.json',
      '--stats-cleanup',
      '--stats-cleanup-vocab',
    ],
  ]);
});

test('stats command throws when stats response reports an error', async () => {
  const context = createContext();
  context.args.stats = true;

  await assert.rejects(async () => {
    await runStatsCommand(context, {
      createTempDir: () => '/tmp/subminer-stats-test',
      joinPath: (...parts) => parts.join('/'),
      runAppCommandAttached: async () => 0,
      waitForStatsResponse: async () => ({
        ok: false,
        error: 'Immersion tracking is disabled in config.',
      }),
      removeDir: () => {},
    });
  }, /Immersion tracking is disabled in config\./);
});

test('stats command fails if attached app exits before startup response', async () => {
  const context = createContext();
  context.args.stats = true;

  await assert.rejects(async () => {
    await runStatsCommand(context, {
      createTempDir: () => '/tmp/subminer-stats-test',
      joinPath: (...parts) => parts.join('/'),
      runAppCommandAttached: async () => 2,
      waitForStatsResponse: async () => {
        await new Promise((resolve) => setTimeout(resolve, 25));
        return { ok: true, url: 'http://127.0.0.1:5175' };
      },
      removeDir: () => {},
    });
  }, /Stats app exited before startup response \(status 2\)\./);
});
