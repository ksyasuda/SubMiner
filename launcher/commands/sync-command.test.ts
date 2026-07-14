import test from 'node:test';
import assert from 'node:assert/strict';
import type { Args } from '../types.js';
import type { LauncherCommandContext } from './context.js';
import { runSyncCommand, type SyncCommandDeps } from './sync-command.js';

function makeContext(
  overrides: Partial<Args>,
  appPath: string | null = '/opt/SubMiner/subminer-app',
): LauncherCommandContext {
  return {
    args: {
      sync: true,
      syncCliTokens: [],
      logLevel: 'warn',
      ...overrides,
    } as Args,
    scriptPath: '/tmp/subminer',
    scriptName: 'subminer',
    mpvSocketPath: '',
    pluginRuntimeConfig: {},
    appPath,
    launcherJellyfinConfig: {},
    processAdapter: { platform: () => 'linux' },
  } as unknown as LauncherCommandContext;
}

test('runSyncCommand proxies sync argv to the app in --sync-cli mode', async () => {
  const spawned: Array<{ appPath: string; appArgs: string[] }> = [];
  const deps: Partial<SyncCommandDeps> = {
    runAppCommand: (appPath, appArgs) => {
      spawned.push({ appPath, appArgs });
    },
  };

  assert.equal(
    await runSyncCommand(makeContext({ syncCliTokens: ['media-box', '--json'] }), deps),
    true,
  );
  assert.deepEqual(spawned, [
    {
      appPath: '/opt/SubMiner/subminer-app',
      appArgs: ['--sync-cli', 'sync', 'media-box', '--json', '--log-level', 'warn'],
    },
  ]);

  assert.equal(await runSyncCommand(makeContext({ sync: false }), deps), false);
  assert.equal(spawned.length, 1);
});

test('runSyncCommand forwards tokens verbatim and appends the effective log level', async () => {
  const spawned: string[][] = [];
  const deps: Partial<SyncCommandDeps> = {
    runAppCommand: (_appPath, appArgs) => {
      spawned.push(appArgs);
    },
  };

  await runSyncCommand(
    makeContext({
      syncCliTokens: [
        'media-box',
        '--pull',
        '--remote-cmd',
        '/opt/SubMiner.AppImage',
        '--db',
        '/tmp/db.sqlite',
        '--force',
        '--json',
      ],
      logLevel: 'debug',
    }),
    deps,
  );

  assert.deepEqual(spawned, [
    [
      '--sync-cli',
      'sync',
      'media-box',
      '--pull',
      '--remote-cmd',
      '/opt/SubMiner.AppImage',
      '--db',
      '/tmp/db.sqlite',
      '--force',
      '--json',
      '--log-level',
      'debug',
    ],
  ]);
});

test('runSyncCommand fails with a clear message when the app binary is missing', async () => {
  const deps: Partial<SyncCommandDeps> = {
    runAppCommand: () => {
      throw new Error('should not spawn');
    },
    fail: (message: string): never => {
      throw new Error(message);
    },
  };

  await assert.rejects(
    () => runSyncCommand(makeContext({ syncCliTokens: ['media-box'] }, null), deps),
    /SubMiner app binary not found \(sync runs inside the app\)/,
  );
});
