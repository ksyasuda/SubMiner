import test from 'node:test';
import assert from 'node:assert/strict';
import {
  authenticateWithPassword,
  listItems,
  listLibraries,
  listSubtitleTracks,
  resolvePlaybackPlan,
  ticksToSeconds,
} from './jellyfin';

const clientInfo = {
  deviceId: 'subminer-test',
  clientName: 'SubMiner',
  clientVersion: '0.1.0-test',
};

test('authenticateWithPassword returns token and user', async () => {
  const originalFetch = globalThis.fetch;
  globalThis.fetch = (async (input) => {
    assert.match(String(input), /Users\/AuthenticateByName$/);
    return new Response(
      JSON.stringify({
        AccessToken: 'abc123',
        User: { Id: 'user-1' },
      }),
      { status: 200 },
    );
  }) as typeof fetch;

  try {
    const session = await authenticateWithPassword(
      'http://jellyfin.local:8096/',
      'kyle',
      'pw',
      clientInfo,
    );
    assert.equal(session.serverUrl, 'http://jellyfin.local:8096');
    assert.equal(session.accessToken, 'abc123');
    assert.equal(session.userId, 'user-1');
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('listLibraries maps server response', async () => {
  const originalFetch = globalThis.fetch;
  globalThis.fetch = (async () =>
    new Response(
      JSON.stringify({
        Items: [
          {
            Id: 'lib-1',
            Name: 'TV',
            CollectionType: 'tvshows',
            Type: 'CollectionFolder',
          },
        ],
      }),
      { status: 200 },
    )) as typeof fetch;

  try {
    const libraries = await listLibraries(
      {
        serverUrl: 'http://jellyfin.local',
        accessToken: 'token',
        userId: 'u1',
        username: 'kyle',
      },
      clientInfo,
    );
    assert.deepEqual(libraries, [
      {
        id: 'lib-1',
        name: 'TV',
        collectionType: 'tvshows',
        type: 'CollectionFolder',
      },
    ]);
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('listItems supports search and formats title', async () => {
  const originalFetch = globalThis.fetch;
  globalThis.fetch = (async (input) => {
    assert.match(String(input), /SearchTerm=planet/);
    assert.match(
      String(input),
      /IncludeItemTypes=Movie%2CEpisode%2CAudio%2CSeries%2CSeason%2CFolder%2CCollectionFolder/,
    );
    return new Response(
      JSON.stringify({
        Items: [
          {
            Id: 'ep-1',
            Name: 'Pilot',
            Type: 'Episode',
            SeriesName: 'Space Show',
            ParentIndexNumber: 1,
            IndexNumber: 2,
          },
        ],
      }),
      { status: 200 },
    );
  }) as typeof fetch;

  try {
    const items = await listItems(
      {
        serverUrl: 'http://jellyfin.local',
        accessToken: 'token',
        userId: 'u1',
        username: 'kyle',
      },
      clientInfo,
      {
        libraryId: 'lib-1',
        searchTerm: 'planet',
        limit: 25,
      },
    );
    assert.equal(items[0]!.title, 'Space Show S01E02 Pilot');
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('listItems keeps playable-only include types when search is empty', async () => {
  const originalFetch = globalThis.fetch;
  globalThis.fetch = (async (input) => {
    assert.match(String(input), /IncludeItemTypes=Movie%2CEpisode%2CAudio/);
    assert.doesNotMatch(String(input), /CollectionFolder|Series|Season|Folder/);
    return new Response(JSON.stringify({ Items: [] }), { status: 200 });
  }) as typeof fetch;

  try {
    const items = await listItems(
      {
        serverUrl: 'http://jellyfin.local',
        accessToken: 'token',
        userId: 'u1',
        username: 'kyle',
      },
      clientInfo,
      {
        libraryId: 'lib-1',
        limit: 25,
      },
    );
    assert.deepEqual(items, []);
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('listItems accepts explicit include types and recursive mode', async () => {
  const originalFetch = globalThis.fetch;
  globalThis.fetch = (async (input) => {
    assert.match(String(input), /Recursive=false/);
    assert.match(String(input), /IncludeItemTypes=Series%2CMovie%2CFolder/);
    return new Response(JSON.stringify({ Items: [] }), { status: 200 });
  }) as typeof fetch;

  try {
    const items = await listItems(
      {
        serverUrl: 'http://jellyfin.local',
        accessToken: 'token',
        userId: 'u1',
        username: 'kyle',
      },
      clientInfo,
      {
        libraryId: 'lib-1',
        includeItemTypes: 'Series,Movie,Folder',
        recursive: false,
        limit: 25,
      },
    );
    assert.deepEqual(items, []);
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('resolvePlaybackPlan chooses direct play when allowed', async () => {
  const originalFetch = globalThis.fetch;
  globalThis.fetch = (async () =>
    new Response(
      JSON.stringify({
        Id: 'movie-1',
        Name: 'Movie A',
        UserData: { PlaybackPositionTicks: 20_000_000 },
        MediaSources: [
          {
            Id: 'ms-1',
            Container: 'mkv',
            SupportsDirectStream: true,
            SupportsTranscoding: true,
            DefaultAudioStreamIndex: 1,
            DefaultSubtitleStreamIndex: 3,
          },
        ],
      }),
      { status: 200 },
    )) as typeof fetch;

  try {
    const plan = await resolvePlaybackPlan(
      {
        serverUrl: 'http://jellyfin.local',
        accessToken: 'token',
        userId: 'u1',
        username: 'kyle',
      },
      clientInfo,
      {
        enabled: true,
        directPlayPreferred: true,
        directPlayContainers: ['mkv'],
      },
      { itemId: 'movie-1' },
    );

    assert.equal(plan.mode, 'direct');
    assert.match(plan.url, /Videos\/movie-1\/stream\?/);
    assert.doesNotMatch(plan.url, /SubtitleStreamIndex=/);
    assert.equal(plan.subtitleStreamIndex, null);
    assert.equal(ticksToSeconds(plan.startTimeTicks), 2);
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('resolvePlaybackPlan prefers transcode when directPlayPreferred is disabled', async () => {
  const originalFetch = globalThis.fetch;
  globalThis.fetch = (async () =>
    new Response(
      JSON.stringify({
        Id: 'movie-2',
        Name: 'Movie B',
        UserData: { PlaybackPositionTicks: 10_000_000 },
        MediaSources: [
          {
            Id: 'ms-2',
            Container: 'mkv',
            SupportsDirectStream: true,
            SupportsTranscoding: true,
            DefaultAudioStreamIndex: 4,
          },
        ],
      }),
      { status: 200 },
    )) as typeof fetch;

  try {
    const plan = await resolvePlaybackPlan(
      {
        serverUrl: 'http://jellyfin.local',
        accessToken: 'token',
        userId: 'u1',
        username: 'kyle',
      },
      clientInfo,
      {
        enabled: true,
        directPlayPreferred: false,
        directPlayContainers: ['mkv'],
        transcodeVideoCodec: 'h264',
      },
      { itemId: 'movie-2' },
    );

    assert.equal(plan.mode, 'transcode');
    const url = new URL(plan.url);
    assert.match(url.pathname, /\/Videos\/movie-2\/master\.m3u8$/);
    assert.equal(url.searchParams.get('api_key'), 'token');
    assert.equal(url.searchParams.get('AudioStreamIndex'), '4');
    assert.equal(url.searchParams.get('StartTimeTicks'), '10000000');
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('resolvePlaybackPlan falls back to transcode when direct container not allowed', async () => {
  const originalFetch = globalThis.fetch;
  globalThis.fetch = (async () =>
    new Response(
      JSON.stringify({
        Id: 'movie-3',
        Name: 'Movie C',
        UserData: { PlaybackPositionTicks: 0 },
        MediaSources: [
          {
            Id: 'ms-3',
            Container: 'avi',
            SupportsDirectStream: true,
            SupportsTranscoding: true,
          },
        ],
      }),
      { status: 200 },
    )) as typeof fetch;

  try {
    const plan = await resolvePlaybackPlan(
      {
        serverUrl: 'http://jellyfin.local',
        accessToken: 'token',
        userId: 'u1',
        username: 'kyle',
      },
      clientInfo,
      {
        enabled: true,
        directPlayPreferred: true,
        directPlayContainers: ['mkv', 'mp4'],
        transcodeVideoCodec: 'h265',
      },
      {
        itemId: 'movie-3',
        audioStreamIndex: 2,
        subtitleStreamIndex: 5,
      },
    );

    assert.equal(plan.mode, 'transcode');
    const url = new URL(plan.url);
    assert.equal(url.searchParams.get('VideoCodec'), 'h265');
    assert.equal(url.searchParams.get('AudioStreamIndex'), '2');
    assert.equal(url.searchParams.get('SubtitleStreamIndex'), '5');
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('listSubtitleTracks returns all subtitle streams with delivery urls', async () => {
  const originalFetch = globalThis.fetch;
  globalThis.fetch = (async () =>
    new Response(
      JSON.stringify({
        Id: 'movie-1',
        MediaSources: [
          {
            Id: 'ms-1',
            MediaStreams: [
              {
                Type: 'Subtitle',
                Index: 2,
                Language: 'eng',
                DisplayTitle: 'English Full',
                IsDefault: true,
                DeliveryMethod: 'Embed',
              },
              {
                Type: 'Subtitle',
                Index: 3,
                Language: 'jpn',
                Title: 'Japanese Signs',
                IsForced: true,
                IsExternal: true,
                DeliveryMethod: 'External',
                DeliveryUrl: '/Videos/movie-1/ms-1/Subtitles/3/Stream.srt',
                IsExternalUrl: false,
              },
              {
                Type: 'Subtitle',
                Index: 4,
                Language: 'spa',
                Title: 'Spanish External',
                DeliveryMethod: 'External',
                DeliveryUrl: 'https://cdn.example.com/subs.srt',
                IsExternalUrl: true,
              },
            ],
          },
        ],
      }),
      { status: 200 },
    )) as typeof fetch;

  try {
    const tracks = await listSubtitleTracks(
      {
        serverUrl: 'http://jellyfin.local',
        accessToken: 'token',
        userId: 'u1',
        username: 'kyle',
      },
      clientInfo,
      'movie-1',
    );
    assert.equal(tracks.length, 3);
    assert.deepEqual(
      tracks.map((track) => track.index),
      [2, 3, 4],
    );
    assert.equal(
      tracks[0]!.deliveryUrl,
      'http://jellyfin.local/Videos/movie-1/ms-1/Subtitles/2/Stream.srt?api_key=token',
    );
    assert.equal(
      tracks[1]!.deliveryUrl,
      'http://jellyfin.local/Videos/movie-1/ms-1/Subtitles/3/Stream.srt?api_key=token',
    );
    assert.equal(tracks[2]!.deliveryUrl, 'https://cdn.example.com/subs.srt');
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('resolvePlaybackPlan falls back to transcode when direct play blocked', async () => {
  const originalFetch = globalThis.fetch;
  globalThis.fetch = (async () =>
    new Response(
      JSON.stringify({
        Id: 'movie-1',
        Name: 'Movie A',
        UserData: { PlaybackPositionTicks: 0 },
        MediaSources: [
          {
            Id: 'ms-1',
            Container: 'avi',
            SupportsDirectStream: true,
            SupportsTranscoding: true,
          },
        ],
      }),
      { status: 200 },
    )) as typeof fetch;

  try {
    const plan = await resolvePlaybackPlan(
      {
        serverUrl: 'http://jellyfin.local',
        accessToken: 'token',
        userId: 'u1',
        username: 'kyle',
      },
      clientInfo,
      {
        enabled: true,
        directPlayPreferred: true,
        directPlayContainers: ['mkv', 'mp4'],
        transcodeVideoCodec: 'h265',
      },
      { itemId: 'movie-1' },
    );

    assert.equal(plan.mode, 'transcode');
    assert.match(plan.url, /master\.m3u8\?/);
    assert.match(plan.url, /VideoCodec=h265/);
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('resolvePlaybackPlan reuses server transcoding url and appends missing params', async () => {
  const originalFetch = globalThis.fetch;
  globalThis.fetch = (async () =>
    new Response(
      JSON.stringify({
        Id: 'movie-4',
        Name: 'Movie D',
        UserData: { PlaybackPositionTicks: 50_000_000 },
        MediaSources: [
          {
            Id: 'ms-4',
            Container: 'mkv',
            SupportsDirectStream: false,
            SupportsTranscoding: true,
            DefaultAudioStreamIndex: 3,
            TranscodingUrl: '/Videos/movie-4/master.m3u8?VideoCodec=hevc',
          },
        ],
      }),
      { status: 200 },
    )) as typeof fetch;

  try {
    const plan = await resolvePlaybackPlan(
      {
        serverUrl: 'http://jellyfin.local',
        accessToken: 'token',
        userId: 'u1',
        username: 'kyle',
      },
      clientInfo,
      {
        enabled: true,
        directPlayPreferred: true,
      },
      {
        itemId: 'movie-4',
        subtitleStreamIndex: 8,
      },
    );

    assert.equal(plan.mode, 'transcode');
    const url = new URL(plan.url);
    assert.match(url.pathname, /\/Videos\/movie-4\/master\.m3u8$/);
    assert.equal(url.searchParams.get('VideoCodec'), 'hevc');
    assert.equal(url.searchParams.get('api_key'), 'token');
    assert.equal(url.searchParams.get('AudioStreamIndex'), '3');
    assert.equal(url.searchParams.get('SubtitleStreamIndex'), '8');
    assert.equal(url.searchParams.get('StartTimeTicks'), '50000000');
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('resolvePlaybackPlan preserves episode metadata, stream selection, and resume ticks', async () => {
  const originalFetch = globalThis.fetch;
  globalThis.fetch = (async () =>
    new Response(
      JSON.stringify({
        Id: 'ep-2',
        Type: 'Episode',
        Name: 'A New Hope',
        SeriesName: 'Galaxy Quest',
        ParentIndexNumber: 2,
        IndexNumber: 7,
        UserData: { PlaybackPositionTicks: 35_000_000 },
        MediaSources: [
          {
            Id: 'ms-ep-2',
            Container: 'mkv',
            SupportsDirectStream: true,
            SupportsTranscoding: true,
            DefaultAudioStreamIndex: 6,
          },
        ],
      }),
      { status: 200 },
    )) as typeof fetch;

  try {
    const plan = await resolvePlaybackPlan(
      {
        serverUrl: 'http://jellyfin.local',
        accessToken: 'token',
        userId: 'u1',
        username: 'kyle',
      },
      clientInfo,
      {
        enabled: true,
        directPlayPreferred: true,
        directPlayContainers: ['mkv'],
      },
      {
        itemId: 'ep-2',
        subtitleStreamIndex: 9,
      },
    );

    assert.equal(plan.mode, 'direct');
    assert.equal(plan.title, 'Galaxy Quest S02E07 A New Hope');
    assert.equal(plan.audioStreamIndex, 6);
    assert.equal(plan.subtitleStreamIndex, 9);
    assert.equal(plan.startTimeTicks, 35_000_000);
    const url = new URL(plan.url);
    assert.equal(url.searchParams.get('AudioStreamIndex'), '6');
    assert.equal(url.searchParams.get('SubtitleStreamIndex'), '9');
    assert.equal(url.searchParams.get('StartTimeTicks'), '35000000');
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('listSubtitleTracks falls back from PlaybackInfo to item media sources', async () => {
  const originalFetch = globalThis.fetch;
  let requestCount = 0;
  globalThis.fetch = (async (input) => {
    requestCount += 1;
    if (requestCount === 1) {
      assert.match(String(input), /\/Items\/movie-fallback\/PlaybackInfo\?/);
      return new Response('Playback info unavailable', { status: 500 });
    }
    return new Response(
      JSON.stringify({
        Id: 'movie-fallback',
        MediaSources: [
          {
            Id: 'ms-fallback',
            MediaStreams: [
              {
                Type: 'Subtitle',
                Index: 11,
                Language: 'eng',
                Title: 'English',
                DeliveryMethod: 'External',
                DeliveryUrl: '/Videos/movie-fallback/ms-fallback/Subtitles/11/Stream.srt',
                IsExternalUrl: false,
              },
            ],
          },
        ],
      }),
      { status: 200 },
    );
  }) as typeof fetch;

  try {
    const tracks = await listSubtitleTracks(
      {
        serverUrl: 'http://jellyfin.local',
        accessToken: 'token',
        userId: 'u1',
        username: 'kyle',
      },
      clientInfo,
      'movie-fallback',
    );
    assert.equal(requestCount, 2);
    assert.equal(tracks.length, 1);
    assert.equal(tracks[0]!.index, 11);
    assert.equal(
      tracks[0]!.deliveryUrl,
      'http://jellyfin.local/Videos/movie-fallback/ms-fallback/Subtitles/11/Stream.srt?api_key=token',
    );
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('authenticateWithPassword surfaces invalid credentials and server status failures', async () => {
  const originalFetch = globalThis.fetch;
  globalThis.fetch = (async () =>
    new Response('Unauthorized', { status: 401, statusText: 'Unauthorized' })) as typeof fetch;

  try {
    await assert.rejects(
      () => authenticateWithPassword('http://jellyfin.local:8096/', 'kyle', 'badpw', clientInfo),
      /Invalid Jellyfin username or password\./,
    );
  } finally {
    globalThis.fetch = originalFetch;
  }

  globalThis.fetch = (async () =>
    new Response('Oops', { status: 500, statusText: 'Internal Server Error' })) as typeof fetch;
  try {
    await assert.rejects(
      () => authenticateWithPassword('http://jellyfin.local:8096/', 'kyle', 'pw', clientInfo),
      /Jellyfin login failed \(500 Internal Server Error\)\./,
    );
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('authenticateWithPassword surfaces unreachable server failures', async () => {
  const originalFetch = globalThis.fetch;
  globalThis.fetch = (async () => {
    throw new TypeError('fetch failed');
  }) as typeof fetch;

  try {
    await assert.rejects(
      () => authenticateWithPassword('http://jellyfin.local:8096/', 'kyle', 'pw', clientInfo),
      /Could not reach Jellyfin server \(fetch failed\)\./,
    );
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('listLibraries surfaces token-expiry auth errors', async () => {
  const originalFetch = globalThis.fetch;
  globalThis.fetch = (async () =>
    new Response('Forbidden', { status: 403, statusText: 'Forbidden' })) as typeof fetch;

  try {
    await assert.rejects(
      () =>
        listLibraries(
          {
            serverUrl: 'http://jellyfin.local',
            accessToken: 'expired',
            userId: 'u1',
            username: 'kyle',
          },
          clientInfo,
        ),
      /Jellyfin authentication failed \(invalid or expired token\)\./,
    );
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('resolvePlaybackPlan surfaces no-source and no-stream fallback errors', async () => {
  const originalFetch = globalThis.fetch;

  globalThis.fetch = (async () =>
    new Response(
      JSON.stringify({
        Id: 'movie-empty',
        Name: 'Movie Empty',
        UserData: { PlaybackPositionTicks: 0 },
        MediaSources: [],
      }),
      { status: 200 },
    )) as typeof fetch;

  try {
    await assert.rejects(
      () =>
        resolvePlaybackPlan(
          {
            serverUrl: 'http://jellyfin.local',
            accessToken: 'token',
            userId: 'u1',
            username: 'kyle',
          },
          clientInfo,
          { enabled: true },
          { itemId: 'movie-empty' },
        ),
      /No playable media source found for Jellyfin item\./,
    );
  } finally {
    globalThis.fetch = originalFetch;
  }

  globalThis.fetch = (async () =>
    new Response(
      JSON.stringify({
        Id: 'movie-no-stream',
        Name: 'Movie No Stream',
        UserData: { PlaybackPositionTicks: 0 },
        MediaSources: [
          {
            Id: 'ms-none',
            Container: 'avi',
            SupportsDirectStream: false,
            SupportsTranscoding: false,
          },
        ],
      }),
      { status: 200 },
    )) as typeof fetch;

  try {
    await assert.rejects(
      () =>
        resolvePlaybackPlan(
          {
            serverUrl: 'http://jellyfin.local',
            accessToken: 'token',
            userId: 'u1',
            username: 'kyle',
          },
          clientInfo,
          { enabled: true },
          { itemId: 'movie-no-stream' },
        ),
      /Jellyfin item cannot be streamed by direct play or transcoding\./,
    );
  } finally {
    globalThis.fetch = originalFetch;
  }
});
