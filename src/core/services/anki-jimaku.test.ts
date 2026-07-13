import test from 'node:test';
import assert from 'node:assert/strict';
import { AnkiJimakuIpcRuntimeOptions, registerAnkiJimakuIpcRuntime } from './anki-jimaku';

interface RuntimeHarness {
  options: AnkiJimakuIpcRuntimeOptions;
  registered: Record<string, (...args: unknown[]) => unknown>;
  state: {
    ankiIntegration: unknown;
    fieldGroupingResolver: ((choice: unknown) => void) | null;
    patches: boolean[];
    broadcasts: number;
    fetchCalls: Array<{ endpoint: string; query?: Record<string, unknown> }>;
    tsukihimeFetchCalls: Array<{ endpoint: string; query?: Record<string, unknown> }>;
    sentCommands: Array<{ command: (string | number)[] }>;
  };
}

function createHarness(): RuntimeHarness {
  const state = {
    ankiIntegration: null as unknown,
    fieldGroupingResolver: null as ((choice: unknown) => void) | null,
    patches: [] as boolean[],
    broadcasts: 0,
    fetchCalls: [] as Array<{
      endpoint: string;
      query?: Record<string, unknown>;
    }>,
    tsukihimeFetchCalls: [] as Array<{
      endpoint: string;
      query?: Record<string, unknown>;
    }>,
    sentCommands: [] as Array<{ command: (string | number)[] }>,
  };

  const options: AnkiJimakuIpcRuntimeOptions = {
    patchAnkiConnectEnabled: (enabled) => {
      state.patches.push(enabled);
    },
    getResolvedConfig: () => ({ tsukihime: { maxSearchResults: 2 } }),
    getRuntimeOptionsManager: () => null,
    tsukihimeFetchJson: async (endpoint, query) => {
      state.tsukihimeFetchCalls.push({
        endpoint,
        query: query as Record<string, unknown>,
      });
      if (endpoint.startsWith('/torrents/')) {
        return {
          ok: true,
          data: {
            id: 606713,
            files: [
              {
                id: 9,
                filename: 'episode.mkv',
                attachments: [
                  {
                    id: 1955356,
                    type: 1,
                    info: { codec: 'ASS', lang: 'en', name: 'English subs' },
                  },
                ],
              },
            ],
          } as never,
        };
      }
      return {
        ok: true,
        data: {
          results: [
            { id: 1, name: 'release a' },
            { id: 2, name: 'release b' },
            { id: 3, name: 'release c' },
          ],
        } as never,
      };
    },
    getSubtitleTimingTracker: () => null,
    getMpvClient: () => ({
      connected: true,
      send: (payload) => {
        state.sentCommands.push(payload);
      },
      request: async (command: unknown[]) => {
        state.sentCommands.push({ command } as never);
        return {
          data: [
            { id: 1, type: 'sub', selected: true },
            {
              id: 3,
              type: 'sub',
              lang: 'en',
              external: true,
              'external-filename': '/tmp/video.en.ass',
            },
          ],
        };
      },
    }),
    getAnkiIntegration: () => state.ankiIntegration as never,
    setAnkiIntegration: (integration) => {
      state.ankiIntegration = integration;
    },
    getKnownWordCacheStatePath: () => '/tmp/subminer-known-words-cache.json',
    showDesktopNotification: () => {},
    createFieldGroupingCallback: () => async () => ({
      keepNoteId: 1,
      deleteNoteId: 2,
      deleteDuplicate: false,
      cancelled: false,
    }),
    broadcastRuntimeOptionsChanged: () => {
      state.broadcasts += 1;
    },
    getFieldGroupingResolver: () => state.fieldGroupingResolver as never,
    setFieldGroupingResolver: (resolver) => {
      state.fieldGroupingResolver = resolver as never;
    },
    parseMediaInfo: () => ({
      title: 'video',
      confidence: 'high',
      rawTitle: 'video',
      filename: 'video.mkv',
      season: null,
      episode: null,
    }),
    getCurrentMediaPath: () => '/tmp/video.mkv',
    jimakuFetchJson: async (endpoint, query) => {
      state.fetchCalls.push({
        endpoint,
        query: query as Record<string, unknown>,
      });
      return {
        ok: true,
        data: [
          { id: 1, name: 'a' },
          { id: 2, name: 'b' },
          { id: 3, name: 'c' },
        ] as never,
      };
    },
    getJimakuMaxEntryResults: () => 2,
    getJimakuLanguagePreference: () => 'ja',
    resolveJimakuApiKey: async () => 'token',
    isRemoteMediaPath: () => false,
    downloadToFile: async (url, destPath) => ({
      ok: true,
      path: `${destPath}:${url}`,
    }),
  };

  let registered: Record<string, (...args: unknown[]) => unknown> = {};
  registerAnkiJimakuIpcRuntime(options, (deps) => {
    registered = deps as unknown as Record<string, (...args: unknown[]) => unknown>;
  });

  return { options, registered, state };
}

test('registerAnkiJimakuIpcRuntime provides full handler surface', () => {
  const { registered } = createHarness();
  const expected = [
    'setAnkiConnectEnabled',
    'clearAnkiHistory',
    'refreshKnownWords',
    'respondFieldGrouping',
    'buildKikuMergePreview',
    'getJimakuMediaInfo',
    'searchJimakuEntries',
    'listJimakuFiles',
    'resolveJimakuApiKey',
    'getCurrentMediaPath',
    'isRemoteMediaPath',
    'downloadToFile',
    'onDownloadedSubtitle',
    'searchTsukihimeEntries',
    'listTsukihimeFiles',
    'downloadTsukihimeSubtitle',
    'onDownloadedSecondarySubtitle',
  ];

  for (const key of expected) {
    assert.equal(typeof registered[key], 'function', `missing handler: ${key}`);
  }
});

test('refreshKnownWords throws when integration is unavailable', async () => {
  const { registered } = createHarness();

  await assert.rejects(
    async () => {
      await registered.refreshKnownWords!();
    },
    { message: 'AnkiConnect integration not enabled' },
  );
});

test('refreshKnownWords delegates to integration', async () => {
  const { registered, state } = createHarness();
  let refreshed = 0;
  state.ankiIntegration = {
    refreshKnownWordCache: async () => {
      refreshed += 1;
    },
  };

  await registered.refreshKnownWords!();

  assert.equal(refreshed, 1);
});

test('setAnkiConnectEnabled disables active integration and broadcasts changes', () => {
  const { registered, state } = createHarness();
  let destroyed = 0;
  state.ankiIntegration = {
    destroy: () => {
      destroyed += 1;
    },
  };

  registered.setAnkiConnectEnabled!(false);

  assert.deepEqual(state.patches, [false]);
  assert.equal(destroyed, 1);
  assert.equal(state.ankiIntegration, null);
  assert.equal(state.broadcasts, 1);
});

test('clearAnkiHistory and respondFieldGrouping execute runtime callbacks', () => {
  const { registered, state, options } = createHarness();
  let cleaned = 0;
  let resolvedChoice: unknown = null;
  state.fieldGroupingResolver = (choice) => {
    resolvedChoice = choice;
  };

  const originalGetTracker = options.getSubtitleTimingTracker;
  options.getSubtitleTimingTracker = () =>
    ({
      cleanup: () => {
        cleaned += 1;
      },
    }) as never;

  const choice = {
    keepNoteId: 10,
    deleteNoteId: 11,
    deleteDuplicate: true,
    cancelled: false,
  };
  registered.clearAnkiHistory!();
  registered.respondFieldGrouping!(choice);

  options.getSubtitleTimingTracker = originalGetTracker;

  assert.equal(cleaned, 1);
  assert.deepEqual(resolvedChoice, choice);
  assert.equal(state.fieldGroupingResolver, null);
});

test('buildKikuMergePreview returns guard error when integration is missing', async () => {
  const { registered } = createHarness();

  const result = await registered.buildKikuMergePreview!({
    keepNoteId: 1,
    deleteNoteId: 2,
    deleteDuplicate: false,
  });

  assert.deepEqual(result, {
    ok: false,
    error: 'AnkiConnect integration not enabled',
  });
});

test('buildKikuMergePreview delegates to integration when available', async () => {
  const { registered, state } = createHarness();
  const calls: unknown[] = [];
  state.ankiIntegration = {
    buildFieldGroupingPreview: async (
      keepNoteId: number,
      deleteNoteId: number,
      deleteDuplicate: boolean,
    ) => {
      calls.push([keepNoteId, deleteNoteId, deleteDuplicate]);
      return { ok: true };
    },
  };

  const result = await registered.buildKikuMergePreview!({
    keepNoteId: 3,
    deleteNoteId: 4,
    deleteDuplicate: true,
  });

  assert.deepEqual(calls, [[3, 4, true]]);
  assert.deepEqual(result, { ok: true });
});

test('searchJimakuEntries caps results and onDownloadedSubtitle sends sub-add to mpv', async () => {
  const { registered, state } = createHarness();

  const searchResult = await registered.searchJimakuEntries!({ query: 'test' });
  assert.deepEqual(state.fetchCalls, [
    {
      endpoint: '/api/entries/search',
      query: { anime: true, query: 'test' },
    },
  ]);
  assert.equal((searchResult as { ok: boolean }).ok, true);
  assert.equal((searchResult as { data: unknown[] }).data.length, 2);

  registered.onDownloadedSubtitle!('/tmp/subtitle.ass');
  assert.deepEqual(state.sentCommands, [{ command: ['sub-add', '/tmp/subtitle.ass', 'select'] }]);
});

test('onDownloadedSecondarySubtitle loads without stealing the primary track', async () => {
  const { registered, state } = createHarness();

  await registered.onDownloadedSecondarySubtitle!('/tmp/video.en.ass');

  assert.deepEqual(state.sentCommands[0], { command: ['sub-add', '/tmp/video.en.ass', 'auto'] });
  assert.deepEqual(state.sentCommands[1], { command: ['get_property', 'track-list'] });
  assert.deepEqual(state.sentCommands[2], { command: ['set_property', 'secondary-sid', 3] });
});

test('onDownloadedSecondarySubtitle retries until mpv reports the new track', async () => {
  const state = {
    sentCommands: [] as Array<{ command: (string | number)[] }>,
    trackListCalls: 0,
  };

  const options = {
    ...createHarness().options,
    getMpvClient: () => ({
      connected: true,
      send: (payload: { command: (string | number)[] }) => {
        state.sentCommands.push(payload);
      },
      request: async () => {
        state.trackListCalls += 1;
        // mpv has not registered the external file yet on the first poll.
        if (state.trackListCalls < 2) {
          return { data: [{ id: 1, type: 'sub', selected: true }] };
        }
        return {
          data: [
            { id: 1, type: 'sub', selected: true },
            { id: 4, type: 'sub', external: true, 'external-filename': '/tmp/video.en.ass' },
          ],
        };
      },
    }),
  } as unknown as AnkiJimakuIpcRuntimeOptions;

  let registered: Record<string, (...args: unknown[]) => unknown> = {};
  registerAnkiJimakuIpcRuntime(options, (deps) => {
    registered = deps as unknown as Record<string, (...args: unknown[]) => unknown>;
  });

  await registered.onDownloadedSecondarySubtitle!('/tmp/video.en.ass');

  assert.equal(state.trackListCalls, 2);
  assert.deepEqual(state.sentCommands.at(-1), {
    command: ['set_property', 'secondary-sid', 4],
  });
});

test('searchTsukihimeEntries caps results using tsukihime.maxSearchResults', async () => {
  const { registered, state } = createHarness();

  const searchResult = await registered.searchTsukihimeEntries!({ query: 'frieren 28' });
  assert.deepEqual(state.tsukihimeFetchCalls, [
    {
      endpoint: '/search/torrents',
      query: { q: 'frieren 28', limit: 2 },
    },
  ]);
  assert.equal((searchResult as { ok: boolean }).ok, true);
  const entries = (searchResult as { data: Array<{ id: number }> }).data;
  assert.equal(entries.length, 2);
  assert.deepEqual(
    entries.map((entry) => entry.id),
    [1, 2],
  );
});

test('listTsukihimeFiles extracts subtitle attachments from torrent detail', async () => {
  const { registered, state } = createHarness();

  const filesResult = await registered.listTsukihimeFiles!({ entryId: 606713 });
  assert.deepEqual(state.tsukihimeFetchCalls, [
    {
      endpoint: '/torrents/606713',
      query: {},
    },
  ]);
  assert.equal((filesResult as { ok: boolean }).ok, true);
  const files = (filesResult as { data: Array<Record<string, unknown>> }).data;
  assert.equal(files.length, 1);
  assert.equal(files[0]!.attachmentId, 1955356);
  assert.equal(files[0]!.filename, 'episode.en.ass');
  assert.equal(files[0]!.url, 'https://storage.tsukihime.org/attach/001dd61c/1955356.xz');
});
