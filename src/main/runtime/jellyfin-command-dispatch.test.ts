import assert from 'node:assert/strict';
import test from 'node:test';
import { createRunJellyfinCommandHandler } from './jellyfin-command-dispatch';
import type { CliArgs } from '../../cli/args';

function createArgs(overrides: Partial<CliArgs> = {}): CliArgs {
  return {
    raw: [],
    target: null,
    start: false,
    stats: false,
    listRecent: false,
    listMediaInfo: false,
    subs: false,
    noAnki: false,
    noKnown: false,
    noAnilist: false,
    anilistStatus: false,
    clearAnilistToken: false,
    anilistSetup: false,
    anilistQueueStatus: false,
    anilistQueueRetry: false,
    yomitanSettings: false,
    toggleOverlay: false,
    hideOverlay: false,
    showOverlay: false,
    toggleInvisibleOverlay: false,
    hideInvisibleOverlay: false,
    showInvisibleOverlay: false,
    copyCurrentSubtitle: false,
    multiCopy: false,
    mineSentence: false,
    mineSentenceMultiple: false,
    updateLastCardFromClipboard: false,
    refreshKnownCache: false,
    triggerFieldGrouping: false,
    manualSubsync: false,
    markAudioCard: false,
    cycleSecondarySub: false,
    runtimeOptions: false,
    debugOverlay: false,
    jellyfinSetup: false,
    jellyfinLogin: false,
    jellyfinLibraries: false,
    jellyfinItems: false,
    jellyfinSubtitles: false,
    jellyfinPlay: false,
    jellyfinRemoteAnnounce: false,
    help: false,
    ...overrides,
  } as CliArgs;
}

test('run jellyfin command returns after auth branch handles command', async () => {
  const calls: string[] = [];
  const run = createRunJellyfinCommandHandler({
    getJellyfinConfig: () => ({ serverUrl: 'http://localhost:8096' }),
    defaultServerUrl: 'http://127.0.0.1:8096',
    getJellyfinClientInfo: () => ({ clientName: 'SubMiner' }),
    handleAuthCommands: async () => {
      calls.push('auth');
      return true;
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
      return false;
    },
  });

  await run(createArgs());
  assert.deepEqual(calls, ['auth']);
});

test('run jellyfin command throws when session missing after auth', async () => {
  const run = createRunJellyfinCommandHandler({
    getJellyfinConfig: () => ({ serverUrl: '', accessToken: '', userId: '' }),
    defaultServerUrl: '',
    getJellyfinClientInfo: () => ({ clientName: 'SubMiner' }),
    handleAuthCommands: async () => false,
    handleRemoteAnnounceCommand: async () => false,
    handleListCommands: async () => false,
    handlePlayCommand: async () => false,
  });

  await assert.rejects(() => run(createArgs()), /Missing Jellyfin session/);
});

test('run jellyfin command dispatches remote/list/play in order until handled', async () => {
  const calls: string[] = [];
  const seenServerUrls: string[] = [];
  const run = createRunJellyfinCommandHandler({
    getJellyfinConfig: () => ({
      serverUrl: 'http://localhost:8096',
      accessToken: 'token',
      userId: 'uid',
      username: 'alice',
    }),
    defaultServerUrl: 'http://127.0.0.1:8096',
    getJellyfinClientInfo: () => ({ clientName: 'SubMiner' }),
    handleAuthCommands: async ({ serverUrl }) => {
      calls.push('auth');
      seenServerUrls.push(serverUrl);
      return false;
    },
    handleRemoteAnnounceCommand: async () => {
      calls.push('remote');
      return false;
    },
    handleListCommands: async ({ session }) => {
      calls.push('list');
      seenServerUrls.push(session.serverUrl);
      return true;
    },
    handlePlayCommand: async () => {
      calls.push('play');
      return false;
    },
  });

  await run(createArgs({ jellyfinServer: 'http://override:8096' }));

  assert.deepEqual(calls, ['auth', 'remote', 'list']);
  assert.deepEqual(seenServerUrls, ['http://override:8096', 'http://override:8096']);
});
