import test from 'node:test';
import assert from 'node:assert/strict';
import path from 'node:path';
import {
  buildForwardedJellyfinAppArgs,
  buildRootSearchGroups,
  classifyJellyfinChildSelection,
  deriveJellyfinTokenStorePath,
  hasStoredJellyfinSession,
  parseEpisodePathFromDisplay,
  parseJellyfinErrorFromAppOutput,
  parseJellyfinAppReply,
  parseJellyfinItemsReply,
  parseJellyfinLibrariesReply,
  parseJellyfinPreviewAuthResponse,
  runJellyfinPlayMenuWithDeps,
  shouldRetryWithStartForNoRunningInstance,
} from './jellyfin.js';
import { makeLauncherArgs } from './test-support/args.js';
import { withEnv } from './test-support/env.js';

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

test('parseJellyfinAppReply maps library and item replies to picker entries', () => {
  const libraries = parseJellyfinAppReply(
    JSON.stringify({
      libraries: [
        { id: 'lib1', name: 'Anime', collectionType: 'tvshows' },
        { id: 'lib1', name: 'Duplicate', collectionType: 'tvshows' },
        { name: 'No id' },
      ],
    }),
  );
  assert.ok(libraries?.ok);
  assert.deepEqual(parseJellyfinLibrariesReply(libraries.payload), [
    { id: 'lib1', name: 'Anime', kind: 'tvshows' },
  ]);

  const items = parseJellyfinAppReply(
    JSON.stringify({ items: [{ id: 'movie-1', title: 'Movie [Alt]', type: 'Movie' }] }),
  );
  assert.ok(items?.ok);
  assert.deepEqual(parseJellyfinItemsReply(items.payload), [
    { id: 'movie-1', name: 'Movie [Alt]', type: 'Movie', display: 'Movie [Alt]' },
  ]);
  assert.equal(parseJellyfinLibrariesReply(items.payload), null);
});

test('parseJellyfinAppReply surfaces app errors and rejects partial files', () => {
  assert.deepEqual(
    parseJellyfinAppReply(
      JSON.stringify({ error: 'Missing Jellyfin session. Run --jellyfin-login first.' }),
    ),
    { ok: false, error: 'Missing Jellyfin session. Run `subminer jellyfin -l` to log in again.' },
  );
  assert.equal(parseJellyfinAppReply(''), null);
  assert.equal(parseJellyfinAppReply('{"libraries": ['), null);
});

test('buildForwardedJellyfinAppArgs appends server, password store, and log level', () => {
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
  const parsed = parseJellyfinPreviewAuthResponse({
    serverUrl: 'http://pve-main:8096/',
    accessToken: 'token-123',
    userId: 'user-1',
  });

  assert.deepEqual(parsed, {
    serverUrl: 'http://pve-main:8096',
    accessToken: 'token-123',
    userId: 'user-1',
  });
});

test('parseJellyfinPreviewAuthResponse returns null without a token', () => {
  assert.equal(
    parseJellyfinPreviewAuthResponse({
      serverUrl: 'http://pve-main:8096',
      accessToken: '',
      userId: 'user-1',
    }),
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
