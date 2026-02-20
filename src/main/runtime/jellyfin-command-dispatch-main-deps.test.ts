import assert from 'node:assert/strict';
import test from 'node:test';
import type { CliArgs } from '../../cli/args';
import { createBuildRunJellyfinCommandMainDepsHandler } from './jellyfin-command-dispatch-main-deps';

test('run jellyfin command main deps builder maps callbacks', async () => {
  const calls: string[] = [];
  const args = { raw: [] } as unknown as CliArgs;
  const config = {
    serverUrl: 'http://localhost:8096',
    accessToken: 'token',
    userId: 'uid',
    username: 'alice',
  };
  const clientInfo = { clientName: 'SubMiner' };

  const deps = createBuildRunJellyfinCommandMainDepsHandler({
    getJellyfinConfig: () => config,
    defaultServerUrl: 'http://127.0.0.1:8096',
    getJellyfinClientInfo: () => clientInfo,
    handleAuthCommands: async () => {
      calls.push('auth');
      return false;
    },
    handleRemoteAnnounceCommand: async () => {
      calls.push('remote');
      return false;
    },
    handleListCommands: async () => {
      calls.push('list');
      return false;
    },
    handlePlayCommand: async () => {
      calls.push('play');
      return true;
    },
  })();

  assert.equal(deps.getJellyfinConfig(), config);
  assert.equal(deps.defaultServerUrl, 'http://127.0.0.1:8096');
  assert.equal(deps.getJellyfinClientInfo(config), clientInfo);
  await deps.handleAuthCommands({
    args,
    jellyfinConfig: config,
    serverUrl: config.serverUrl,
    clientInfo,
  });
  await deps.handleRemoteAnnounceCommand(args);
  await deps.handleListCommands({
    args,
    session: {
      serverUrl: config.serverUrl,
      accessToken: config.accessToken,
      userId: config.userId,
      username: config.username,
    },
    clientInfo,
    jellyfinConfig: config,
  });
  await deps.handlePlayCommand({
    args,
    session: {
      serverUrl: config.serverUrl,
      accessToken: config.accessToken,
      userId: config.userId,
      username: config.username,
    },
    clientInfo,
    jellyfinConfig: config,
  });
  assert.deepEqual(calls, ['auth', 'remote', 'list', 'play']);
});
