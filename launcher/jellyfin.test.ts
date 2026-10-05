import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {
  buildForwardedJellyfinAppArgs,
  buildRootSearchGroups,
  classifyJellyfinChildSelection,
  deriveJellyfinTokenStorePath,
  hasStoredJellyfinSession,
  parseEpisodePathFromDisplay,
  parseJellyfinErrorFromAppOutput,
  parseJellyfinItemsFromAppOutput,
  parseJellyfinLibrariesFromAppOutput,
  parseJellyfinPreviewAuthResponse,
  readUtf8FileAppendedSince,
  runJellyfinPlayMenuWithDeps,
  shouldRetryWithStartForNoRunningInstance,
} from './jellyfin.js';
import { makeLauncherArgs } from './test-support/args.js';
import { withEnv } from './test-support/env.js';

function withTempDir<T>(fn: (dir: string) => T): T {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-jellyfin-test-'));
  try {
    return fn(dir);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

class HandoffComplete extends Error {}

type PlayMenuDeps = NonNullable<Parameters<typeof runJellyfinPlayMenuWithDeps>[4]>;

/**
 * Runs the play menu with fake deps and a direct env session. The app handoff never returns in
 * production, so the fake throws `HandoffComplete` after recording its args into `handoff`.
 */
async function runPlayMenu(options: {
  socketReady?: boolean;
  overrides?: Partial<PlayMenuDeps>;
  args?: Parameters<typeof makeLauncherArgs>[0];
}) {
  const launches: string[] = [];
  const handoff: { appArgs: string[] | null } = { appArgs: null };
  const deps: PlayMenuDeps = {
    loadLauncherJellyfinConfig: () => ({}),
    findRofiTheme: () => null,
    resolveJellyfinSelection: async () => 'item-123',
    resolveJellyfinSelectionViaApp: async () => {
      throw new Error('unexpected app-based selection');
    },
    hasStoredJellyfinSession: () => true,
    requestJellyfinPreviewAuthFromApp: async () => null,
    resolveLauncherMainConfigPath: () => '/tmp/SubMiner/config.jsonc',
    pathExists: () => options.socketReady ?? false,
    ensureRuntimePluginReady: async () => {
      launches.push('plugin');
    },
    waitForUnixSocketReady: async () => true,
    launchMpvIdleDetached: async () => {
      launches.push('launch');
    },
    resolveLauncherRuntimePluginPath: () => '/tmp/plugin/main.lua',
    runAppCommandWithInheritLogged: (_appPath, appArgs) => {
      handoff.appArgs = appArgs;
      throw new HandoffComplete();
    },
    log: () => {},
    ...options.overrides,
  };

  await withEnv(
    { SUBMINER_JELLYFIN_ACCESS_TOKEN: 'token', SUBMINER_JELLYFIN_USER_ID: 'user' },
    () =>
      assert.rejects(
        () =>
          runJellyfinPlayMenuWithDeps(
            '/tmp/SubMiner.AppImage',
            makeLauncherArgs({ jellyfinServer: 'https://jellyfin.example.test', ...options.args }),
            '/tmp/subminer',
            '/tmp/subminer.sock',
            deps,
          ),
        HandoffComplete,
      ),
  );

  return { launches, appArgs: handoff.appArgs };
}

test('Jellyfin playback launches the idle mpv with the runtime plugin, then hands the item to the app', async () => {
  const { launches, appArgs } = await runPlayMenu({});

  assert.deepEqual(launches, ['plugin', 'launch']);
  assert.deepEqual(appArgs, ['--start', '--jellyfin-play', '--jellyfin-item-id=item-123']);
});

test('Jellyfin playback reuses a ready mpv socket instead of launching idle mpv', async () => {
  const { launches, appArgs } = await runPlayMenu({ socketReady: true });

  assert.deepEqual(launches, []);
  assert.deepEqual(appArgs, ['--start', '--jellyfin-play', '--jellyfin-item-id=item-123']);
});

test('parseJellyfinLibrariesFromAppOutput parses prefixed library lines', () => {
  const parsed = parseJellyfinLibrariesFromAppOutput(`
[subminer] - 2026-03-01 13:10:34 - INFO - [main] Jellyfin library: Anime [lib1] (tvshows)
[subminer] - 2026-03-01 13:10:35 - INFO - [main] Jellyfin library: Movies [lib2] (movies)
`);

  assert.deepEqual(parsed, [
    { id: 'lib1', name: 'Anime', kind: 'tvshows' },
    { id: 'lib2', name: 'Movies', kind: 'movies' },
  ]);
});

test('parseJellyfinItemsFromAppOutput parses item title/id/type tuples', () => {
  const parsed = parseJellyfinItemsFromAppOutput(`
[subminer] - 2026-03-01 13:10:34 - INFO - [main] Jellyfin item: Solo Leveling S01E10 [item-10] (Episode)
[subminer] - 2026-03-01 13:10:35 - INFO - [main] Jellyfin item: Movie [Alt] [movie-1] (Movie)
`);

  assert.deepEqual(parsed, [
    {
      id: 'item-10',
      name: 'Solo Leveling S01E10',
      type: 'Episode',
      display: 'Solo Leveling S01E10',
    },
    {
      id: 'movie-1',
      name: 'Movie [Alt]',
      type: 'Movie',
      display: 'Movie [Alt]',
    },
  ]);
});

test('buildForwardedJellyfinAppArgs forces app log level for parseable list output', () => {
  const forwarded = buildForwardedJellyfinAppArgs(
    {
      jellyfinServer: 'https://jf.example.test/',
      passwordStore: 'gnome-libsecret',
      logLevel: 'info',
    } as never,
    ['--jellyfin-libraries'],
  );

  assert.deepEqual(forwarded, [
    '--jellyfin-libraries',
    '--jellyfin-server',
    'https://jf.example.test',
    '--password-store',
    'gnome-libsecret',
    '--log-level',
    'info',
  ]);
});

test('parseJellyfinErrorFromAppOutput extracts bracketed error lines', () => {
  const parsed = parseJellyfinErrorFromAppOutput(`
[subminer] - 2026-03-01 13:10:34 - WARN - [main] test warning
[2026-03-01T21:11:28.821Z] [ERROR] Missing Jellyfin session. Set SUBMINER_JELLYFIN_ACCESS_TOKEN and SUBMINER_JELLYFIN_USER_ID, then retry.
`);

  assert.equal(
    parsed,
    'Missing Jellyfin session. Set SUBMINER_JELLYFIN_ACCESS_TOKEN and SUBMINER_JELLYFIN_USER_ID, then retry.',
  );
});

test('parseJellyfinErrorFromAppOutput extracts main runtime error lines', () => {
  const parsed = parseJellyfinErrorFromAppOutput(`
[subminer] - 2026-03-01 13:10:34 - ERROR - [main] runJellyfinCommand failed: {"message":"Missing Jellyfin password."}
`);

  assert.equal(
    parsed,
    '[main] runJellyfinCommand failed: {"message":"Missing Jellyfin password."}',
  );
});

test('parseJellyfinPreviewAuthResponse parses valid structured response payload', () => {
  const parsed = parseJellyfinPreviewAuthResponse(
    JSON.stringify({
      serverUrl: 'http://pve-main:8096/',
      accessToken: 'token-123',
      userId: 'user-1',
    }),
  );

  assert.deepEqual(parsed, {
    serverUrl: 'http://pve-main:8096',
    accessToken: 'token-123',
    userId: 'user-1',
  });
});

test('parseJellyfinPreviewAuthResponse returns null for invalid payloads', () => {
  assert.equal(parseJellyfinPreviewAuthResponse(''), null);
  assert.equal(parseJellyfinPreviewAuthResponse('{not json}'), null);
  assert.equal(
    parseJellyfinPreviewAuthResponse(
      JSON.stringify({
        serverUrl: 'http://pve-main:8096',
        accessToken: '',
        userId: 'user-1',
      }),
    ),
    null,
  );
});

test('hasStoredJellyfinSession checks token-store existence', () => {
  const configPath = path.join('/home/test', '.config', 'SubMiner', 'config.jsonc');
  const tokenPath = deriveJellyfinTokenStorePath(configPath);
  const exists = (candidate: string): boolean => candidate === tokenPath;
  assert.equal(hasStoredJellyfinSession(configPath, exists), true);
  assert.equal(
    hasStoredJellyfinSession(path.join('/home/test', '.config', 'Other', 'alt.jsonc'), exists),
    false,
  );
});

test('shouldRetryWithStartForNoRunningInstance matches expected app lifecycle error', () => {
  assert.equal(
    shouldRetryWithStartForNoRunningInstance('No running instance. Use --start to launch the app.'),
    true,
  );
  assert.equal(
    shouldRetryWithStartForNoRunningInstance(
      'Missing Jellyfin session. Run --jellyfin-login first.',
    ),
    false,
  );
});

test('readUtf8FileAppendedSince treats offset as bytes and survives multibyte logs', () => {
  withTempDir((root) => {
    const logPath = path.join(root, 'SubMiner.log');
    const prefix = '[subminer] こんにちは\n';
    const suffix = '[subminer] Jellyfin library: Movies [lib2] (movies)\n';
    fs.writeFileSync(logPath, `${prefix}${suffix}`, 'utf8');

    const byteOffset = Buffer.byteLength(prefix, 'utf8');
    const fromByteOffset = readUtf8FileAppendedSince(logPath, byteOffset);
    assert.match(fromByteOffset, /Jellyfin library: Movies \[lib2\] \(movies\)/);

    const fromBeyondEnd = readUtf8FileAppendedSince(logPath, byteOffset + 9999);
    assert.match(fromBeyondEnd, /Jellyfin library: Movies \[lib2\] \(movies\)/);
  });
});

test('parseEpisodePathFromDisplay extracts series and season from episode display titles', () => {
  assert.deepEqual(
    parseEpisodePathFromDisplay('KONOSUBA S01E03 A Panty Treasure in This Right Hand!'),
    {
      seriesName: 'KONOSUBA',
      seasonNumber: 1,
    },
  );
  assert.deepEqual(parseEpisodePathFromDisplay('Frieren S2E10 Something'), {
    seriesName: 'Frieren',
    seasonNumber: 2,
  });
});

test('parseEpisodePathFromDisplay returns null for non-episode displays', () => {
  assert.equal(parseEpisodePathFromDisplay('Movie Title (Movie)'), null);
  assert.equal(parseEpisodePathFromDisplay('Just A Name'), null);
});

test('buildRootSearchGroups excludes episodes and keeps containers/movies', () => {
  const groups = buildRootSearchGroups([
    { id: 'series-1', name: 'The Eminence in Shadow', type: 'Series', display: 'x' },
    { id: 'movie-1', name: 'Spirited Away', type: 'Movie', display: 'x' },
    { id: 'episode-1', name: 'The Eminence in Shadow S01E01', type: 'Episode', display: 'x' },
  ]);

  assert.deepEqual(groups, [
    {
      id: 'series-1',
      name: 'The Eminence in Shadow',
      type: 'Series',
      display: 'The Eminence in Shadow (Series)',
    },
    {
      id: 'movie-1',
      name: 'Spirited Away',
      type: 'Movie',
      display: 'Spirited Away (Movie)',
    },
  ]);
});

test('classifyJellyfinChildSelection keeps container drilldown state instead of flattening', () => {
  const next = classifyJellyfinChildSelection({ id: 'season-2', type: 'Season' });
  assert.deepEqual(next, {
    kind: 'container',
    id: 'season-2',
  });
});
