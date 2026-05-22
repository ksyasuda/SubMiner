import assert from 'node:assert/strict';
import test from 'node:test';
import { createPlayJellyfinItemInMpvHandler } from './jellyfin-playback-launch';

const baseSession = {
  serverUrl: 'http://localhost:8096',
  accessToken: 'token',
  userId: 'uid',
  username: 'alice',
};

const baseClientInfo = {
  clientName: 'SubMiner',
  clientVersion: '1.0.0',
  deviceId: 'did',
};

test('playback handler throws when mpv is not connected', async () => {
  const handler = createPlayJellyfinItemInMpvHandler({
    ensureMpvConnectedForPlayback: async () => false,
    getMpvClient: () => null,
    resolvePlaybackPlan: async () => {
      throw new Error('unreachable');
    },
    applyJellyfinMpvDefaults: () => {},
    showVisibleOverlay: () => {},
    sendMpvCommand: () => {},
    armQuitOnDisconnect: () => {},
    schedule: () => {},
    convertTicksToSeconds: (ticks) => ticks / 10_000_000,
    preloadExternalSubtitles: () => {},
    setActivePlayback: () => {},
    setLastProgressAtMs: () => {},
    reportPlaying: () => {},
    showMpvOsd: () => {},
  });

  await assert.rejects(
    () =>
      handler({
        session: baseSession,
        clientInfo: baseClientInfo,
        jellyfinConfig: {},
        itemId: 'item-1',
      }),
    /MPV not connected and auto-launch failed/,
  );
});

test('playback handler drives mpv commands and playback state', async () => {
  const commands: Array<Array<string | number>> = [];
  const scheduled: Array<{ delay: number; callback: () => void }> = [];
  const calls: string[] = [];
  const activeStates: Array<Record<string, unknown>> = [];
  const reportPayloads: Array<Record<string, unknown>> = [];
  const statsMetadata: Array<Record<string, unknown>> = [];
  const handler = createPlayJellyfinItemInMpvHandler({
    ensureMpvConnectedForPlayback: async () => true,
    getMpvClient: () => ({ connected: true, send: () => {} }),
    resolvePlaybackPlan: async () => ({
      url: 'https://stream.example/video.m3u8',
      mode: 'direct',
      title: 'Episode 1',
      itemTitle: 'Episode 1',
      seriesTitle: 'Show Title',
      seasonNumber: 1,
      episodeNumber: 1,
      startTimeTicks: 12_000_000,
      audioStreamIndex: 1,
      subtitleStreamIndex: 2,
    }),
    applyJellyfinMpvDefaults: () => calls.push('defaults'),
    showVisibleOverlay: () => calls.push('visible-overlay'),
    sendMpvCommand: (command) => commands.push(command),
    armQuitOnDisconnect: () => calls.push('arm'),
    schedule: (callback, delayMs) => {
      scheduled.push({ delay: delayMs, callback });
    },
    convertTicksToSeconds: (ticks) => ticks / 10_000_000,
    preloadExternalSubtitles: () => calls.push('preload'),
    setActivePlayback: (state) => activeStates.push(state as Record<string, unknown>),
    setLastProgressAtMs: (value) => calls.push(`progress:${value}`),
    reportPlaying: (payload) => reportPayloads.push(payload as Record<string, unknown>),
    showMpvOsd: (text) => calls.push(`osd:${text}`),
    recordJellyfinPlaybackMetadata: (metadata) =>
      statsMetadata.push(metadata as Record<string, unknown>),
  });

  await handler({
    session: baseSession,
    clientInfo: baseClientInfo,
    jellyfinConfig: {},
    itemId: 'item-1',
  });

  assert.deepEqual(commands.slice(0, 5), [
    ['set_property', 'sub-auto', 'no'],
    ['loadfile', 'https://stream.example/video.m3u8', 'replace'],
    ['set_property', 'force-media-title', 'Episode 1'],
    ['set_property', 'sid', 'no'],
    ['seek', 1.2, 'absolute+exact'],
  ]);
  assert.equal(scheduled.length, 0);
  assert.equal(
    commands.filter((command) => command[0] === 'set_property' && command[1] === 'sid').length,
    1,
  );

  assert.ok(calls.includes('defaults'));
  assert.ok(calls.includes('visible-overlay'));
  assert.ok(
    calls.indexOf('visible-overlay') < calls.indexOf('preload'),
    'visible overlay should be shown before Jellyfin subtitles are selected',
  );
  assert.ok(calls.includes('arm'));
  assert.ok(calls.includes('preload'));
  assert.ok(calls.includes('progress:0'));
  assert.ok(calls.includes('osd:Jellyfin direct: Episode 1'));

  assert.equal(activeStates.length, 1);
  assert.equal(activeStates[0]?.playMethod, 'DirectPlay');
  assert.equal(reportPayloads.length, 1);
  assert.equal(reportPayloads[0]?.eventName, 'start');
  assert.deepEqual(statsMetadata, [
    {
      mediaPath: 'https://stream.example/video.m3u8',
      displayTitle: 'Episode 1',
      itemTitle: 'Episode 1',
      seriesTitle: 'Show Title',
      seasonNumber: 1,
      episodeNumber: 1,
      itemId: 'item-1',
    },
  ]);
});

test('playback handler publishes Jellyfin title before loading tokenized stream url', async () => {
  const timeline: string[] = [];
  const handler = createPlayJellyfinItemInMpvHandler({
    ensureMpvConnectedForPlayback: async () => true,
    getMpvClient: () => ({ connected: true, send: () => {} }),
    resolvePlaybackPlan: async () => ({
      url: 'https://jellyfin.local/Videos/ep-1/stream?static=true&api_key=secret-token&MediaSourceId=ms-1',
      mode: 'direct',
      title: 'Galaxy Quest S02E07 A New Hope',
      itemTitle: 'A New Hope',
      seriesTitle: 'Galaxy Quest',
      seasonNumber: 2,
      episodeNumber: 7,
      startTimeTicks: 0,
      audioStreamIndex: null,
      subtitleStreamIndex: null,
    }),
    applyJellyfinMpvDefaults: () => {},
    showVisibleOverlay: () => {},
    sendMpvCommand: (command) => timeline.push(`cmd:${command[0]}:${String(command[1] ?? '')}`),
    armQuitOnDisconnect: () => {},
    schedule: () => {},
    convertTicksToSeconds: (ticks) => ticks / 10_000_000,
    preloadExternalSubtitles: () => {},
    setActivePlayback: () => {},
    setLastProgressAtMs: () => {},
    reportPlaying: () => {},
    showMpvOsd: () => {},
    updateCurrentMediaTitle: (title) => timeline.push(`title:${title}`),
  });

  await handler({
    session: baseSession,
    clientInfo: baseClientInfo,
    jellyfinConfig: {},
    itemId: 'ep-1',
  });

  const titleIndex = timeline.indexOf('title:Galaxy Quest S02E07 A New Hope');
  const loadIndex = timeline.findIndex((entry) => entry.startsWith('cmd:loadfile:'));
  assert.ok(titleIndex >= 0);
  assert.ok(loadIndex >= 0);
  assert.ok(titleIndex < loadIndex);
  assert.equal(timeline[titleIndex]?.includes('api_key'), false);
});

test('playback handler applies start override to stream url for remote resume', async () => {
  const commands: Array<Array<string | number>> = [];
  const handler = createPlayJellyfinItemInMpvHandler({
    ensureMpvConnectedForPlayback: async () => true,
    getMpvClient: () => ({ connected: true, send: () => {} }),
    resolvePlaybackPlan: async () => ({
      url: 'https://stream.example/video.m3u8?api_key=token',
      mode: 'transcode',
      title: 'Episode 2',
      itemTitle: 'Episode 2',
      seriesTitle: null,
      seasonNumber: null,
      episodeNumber: null,
      startTimeTicks: 0,
      audioStreamIndex: null,
      subtitleStreamIndex: null,
    }),
    applyJellyfinMpvDefaults: () => {},
    showVisibleOverlay: () => {},
    sendMpvCommand: (command) => commands.push(command),
    armQuitOnDisconnect: () => {},
    schedule: () => {},
    convertTicksToSeconds: (ticks) => ticks / 10_000_000,
    preloadExternalSubtitles: () => {},
    setActivePlayback: () => {},
    setLastProgressAtMs: () => {},
    reportPlaying: () => {},
    showMpvOsd: () => {},
  });

  await handler({
    session: baseSession,
    clientInfo: baseClientInfo,
    jellyfinConfig: {},
    itemId: 'item-2',
    startTimeTicksOverride: 55_000_000,
  });

  assert.equal(commands[1]?.[0], 'loadfile');
  const loadedUrl = String(commands[1]?.[1] ?? '');
  const parsed = new URL(loadedUrl);
  assert.equal(parsed.searchParams.get('StartTimeTicks'), '55000000');
  assert.deepEqual(commands[4], ['seek', 5.5, 'absolute+exact']);
});

test('playback handler does not let stats metadata failures block playback startup', async () => {
  const commands: Array<Array<string | number>> = [];
  const handler = createPlayJellyfinItemInMpvHandler({
    ensureMpvConnectedForPlayback: async () => true,
    getMpvClient: () => ({ connected: true, send: () => {} }),
    resolvePlaybackPlan: async () => ({
      url: 'https://stream.example/video.m3u8',
      mode: 'direct',
      title: 'Episode 3',
      itemTitle: 'Episode 3',
      seriesTitle: null,
      seasonNumber: null,
      episodeNumber: null,
      startTimeTicks: 0,
      audioStreamIndex: null,
      subtitleStreamIndex: null,
    }),
    applyJellyfinMpvDefaults: () => {},
    showVisibleOverlay: () => {},
    sendMpvCommand: (command) => commands.push(command),
    armQuitOnDisconnect: () => {},
    schedule: () => {},
    convertTicksToSeconds: (ticks) => ticks / 10_000_000,
    preloadExternalSubtitles: () => {},
    setActivePlayback: () => {},
    setLastProgressAtMs: () => {},
    reportPlaying: () => {},
    showMpvOsd: () => {},
    recordJellyfinPlaybackMetadata: () => {
      throw new Error('stats db unavailable');
    },
  });

  await handler({
    session: baseSession,
    clientInfo: baseClientInfo,
    jellyfinConfig: {},
    itemId: 'item-3',
  });

  assert.deepEqual(commands[1], ['loadfile', 'https://stream.example/video.m3u8', 'replace']);
});

test('playback handler does not let media title failures block playback startup', async () => {
  const commands: Array<Array<string | number>> = [];
  const handler = createPlayJellyfinItemInMpvHandler({
    ensureMpvConnectedForPlayback: async () => true,
    getMpvClient: () => ({ connected: true, send: () => {} }),
    resolvePlaybackPlan: async () => ({
      url: 'https://stream.example/video.m3u8',
      mode: 'direct',
      title: 'Episode 4',
      itemTitle: 'Episode 4',
      seriesTitle: null,
      seasonNumber: null,
      episodeNumber: null,
      startTimeTicks: 0,
      audioStreamIndex: null,
      subtitleStreamIndex: null,
    }),
    applyJellyfinMpvDefaults: () => {},
    showVisibleOverlay: () => {},
    sendMpvCommand: (command) => commands.push(command),
    armQuitOnDisconnect: () => {},
    schedule: () => {},
    convertTicksToSeconds: (ticks) => ticks / 10_000_000,
    preloadExternalSubtitles: () => {},
    setActivePlayback: () => {},
    setLastProgressAtMs: () => {},
    reportPlaying: () => {},
    showMpvOsd: () => {},
    updateCurrentMediaTitle: () => {
      throw new Error('title state unavailable');
    },
  });

  await handler({
    session: baseSession,
    clientInfo: baseClientInfo,
    jellyfinConfig: {},
    itemId: 'item-4',
  });

  assert.deepEqual(commands[1], ['loadfile', 'https://stream.example/video.m3u8', 'replace']);
});
