import test from 'node:test';
import assert from 'node:assert/strict';
import type { Args } from '../types.js';
import type { LauncherCommandContext } from './context.js';
import { buildSyncCliArgv, runSyncCommand, type SyncCommandDeps } from './sync-command.js';

function makeContext(
  overrides: Partial<Args>,
  appPath: string | null = '/opt/SubMiner/subminer-app',
): LauncherCommandContext {
  return {
    args: {
      sync: true,
      syncHost: '',
      syncSnapshotPath: '',
      syncMergePath: '',
      syncDirection: 'both',
      syncRemoteCmd: '',
      syncDbPath: '',
      syncForce: false,
      syncJson: false,
      syncCheck: false,
      syncMakeTemp: false,
      syncRemoveTempPath: '',
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

function makeArgs(overrides: Partial<Parameters<typeof buildSyncCliArgv>[0]>) {
  return {
    syncHost: '',
    syncSnapshotPath: '',
    syncMergePath: '',
    syncDirection: 'both' as const,
    syncRemoteCmd: '',
    syncDbPath: '',
    syncForce: false,
    syncJson: false,
    syncCheck: false,
    syncMakeTemp: false,
    syncRemoveTempPath: '',
    logLevel: 'warn' as const,
    ...overrides,
  };
}

test('runSyncCommand proxies sync argv to the app in --sync-cli mode', async () => {
  const spawned: Array<{ appPath: string; appArgs: string[] }> = [];
  const deps: Partial<SyncCommandDeps> = {
    runAppCommand: (appPath, appArgs) => {
      spawned.push({ appPath, appArgs });
    },
  };

  assert.equal(
    await runSyncCommand(makeContext({ syncHost: 'media-box', syncJson: true }), deps),
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

test('buildSyncCliArgv forwards every sync option', () => {
  assert.deepEqual(
    buildSyncCliArgv(
      makeArgs({
        syncHost: 'media-box',
        syncDirection: 'pull',
        syncRemoteCmd: '/opt/SubMiner.AppImage',
        syncDbPath: '/tmp/db.sqlite',
        syncForce: true,
        syncJson: true,
        logLevel: 'debug',
      }),
    ),
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
  );

  assert.deepEqual(
    buildSyncCliArgv(makeArgs({ syncSnapshotPath: '/tmp/out.sqlite' })),
    ['--sync-cli', 'sync', '--snapshot', '/tmp/out.sqlite', '--log-level', 'warn'],
  );

  assert.deepEqual(
    buildSyncCliArgv(makeArgs({ syncHost: 'media-box', syncCheck: true })),
    ['--sync-cli', 'sync', 'media-box', '--check', '--log-level', 'warn'],
  );

  assert.deepEqual(
    buildSyncCliArgv(makeArgs({ syncMakeTemp: true })),
    ['--sync-cli', 'sync', '--make-temp', '--log-level', 'warn'],
  );

  assert.deepEqual(
    buildSyncCliArgv(makeArgs({ syncRemoveTempPath: '/tmp/subminer-sync-x' })),
    ['--sync-cli', 'sync', '--remove-temp', '/tmp/subminer-sync-x', '--log-level', 'warn'],
  );

  assert.deepEqual(
    buildSyncCliArgv(makeArgs({ syncMergePath: '/tmp/in.sqlite', syncForce: true })),
    ['--sync-cli', 'sync', '--merge', '/tmp/in.sqlite', '--force', '--log-level', 'warn'],
  );
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
    () => runSyncCommand(makeContext({ syncHost: 'media-box' }, null), deps),
    /SubMiner app binary not found \(sync runs inside the app\)/,
  );
});
