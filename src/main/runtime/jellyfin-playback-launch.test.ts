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
    preloadExternalSubtitles: () => {
      calls.push('preload');
    },
    setActivePlayback: (state) => activeStates.push(state as Record<string, unknown>),
    setLastProgressAtMs: (value) => calls.push(`progress:${value}`),
    reportPlaying: (payload) => reportPayloads.push(payload as Record<string, unknown>),
    showMpvOsd: (text) => calls.push(`osd:${text}`),
    recordJellyfinPlaybackMetadata: (metadata) => {
      statsMetadata.push(metadata as Record<string, unknown>);
    },
  });

  await handler({
    session: baseSession,
    clientInfo: baseClientInfo,
    jellyfinConfig: {},
    itemId: 'item-1',
  });

  assert.deepEqual(commands.slice(0, 8), [
    ['set_property', 'sub-auto', 'no'],
    ['set_property', 'sid', 'no'],
    ['set_property', 'secondary-sid', 'no'],
    ['set_property', 'sub-visibility', 'no'],
    ['set_property', 'secondary-sub-visibility', 'no'],
    ['script-message', 'subminer-managed-subtitles-loading'],
    [
      'loadfile',
      'https://stream.example/video.m3u8',
      'replace',
      -1,
      'sid=no,secondary-sid=no,sub-auto=no,sub-visibility=no,secondary-sub-visibility=no,start=1.2',
    ],
    ['set_property', 'force-media-title', 'Episode 1'],
  ]);
  assert.equal(scheduled.length, 0);
  assert.equal(
    commands.filter((command) => command[0] === 'set_property' && command[1] === 'sid').length,
    1,
  );

  assert.ok(calls.includes('defaults'));
  assert.ok(
    calls.indexOf('preload') < calls.indexOf('visible-overlay'),
    'visible overlay should be shown after Jellyfin subtitles are selected',
  );
  assert.ok(calls.includes('visible-overlay'));
  assert.ok(calls.includes('arm'));
  assert.ok(calls.includes('preload'));
  assert.ok(calls.includes('progress:0'));
  assert.ok(calls.includes('osd:Jellyfin direct: Episode 1'));

  assert.equal(activeStates.length, 1);
  assert.equal(activeStates[0]?.playMethod, 'DirectPlay');
  assert.equal(activeStates[0]?.lastKnownPositionSeconds, 1.2);
  assert.equal(reportPayloads.length, 1);
  assert.equal(reportPayloads[0]?.eventName, 'start');
  assert.equal(reportPayloads[0]?.positionTicks, 12_000_000);
  assert.equal(reportPayloads[0]?.isPaused, false);
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

test('playback handler waits for Jellyfin subtitle preload before showing visible overlay', async () => {
  const calls: string[] = [];
  let resolvePreload!: () => void;
  const preloadComplete = new Promise<void>((resolve) => {
    resolvePreload = resolve;
  });
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
      startTimeTicks: 0,
      audioStreamIndex: 1,
      subtitleStreamIndex: 2,
    }),
    applyJellyfinMpvDefaults: () => {},
    showVisibleOverlay: () => calls.push('visible-overlay'),
    sendMpvCommand: () => {},
    armQuitOnDisconnect: () => {},
    schedule: () => {},
    convertTicksToSeconds: (ticks) => ticks / 10_000_000,
    preloadExternalSubtitles: async () => {
      calls.push('preload-start');
      await preloadComplete;
      calls.push('preload-done');
    },
    setActivePlayback: () => {},
    setLastProgressAtMs: () => {},
    reportPlaying: () => {},
    showMpvOsd: () => {},
  });

  const playback = handler({
    session: baseSession,
    clientInfo: baseClientInfo,
    jellyfinConfig: {},
    itemId: 'item-1',
  });
  for (let i = 0; i < 5 && calls.length === 0; i += 1) {
    await Promise.resolve();
  }

  assert.equal(calls.length, 1);
  assert.equal(calls[0], 'preload-start');
  resolvePreload();
  await playback;

  assert.deepEqual(calls, ['preload-start', 'preload-done', 'visible-overlay']);
});

test('playback handler strips Jellyfin subtitle stream from mpv load URL', async () => {
  const commands: Array<Array<string | number>> = [];
  const reports: Array<Record<string, unknown>> = [];
  const handler = createPlayJellyfinItemInMpvHandler({
    ensureMpvConnectedForPlayback: async () => true,
    getMpvClient: () => ({ connected: true, send: () => {} }),
    resolvePlaybackPlan: async () => ({
      url: 'https://jellyfin.local/Videos/ep-1/stream?static=true&api_key=secret-token&MediaSourceId=ms-1&AudioStreamIndex=3&SubtitleStreamIndex=4',
      mode: 'direct',
      title: 'Episode 1',
      itemTitle: 'Episode 1',
      seriesTitle: null,
      seasonNumber: null,
      episodeNumber: null,
      startTimeTicks: 0,
      audioStreamIndex: 3,
      subtitleStreamIndex: 4,
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
    reportPlaying: (payload) => reports.push(payload),
    showMpvOsd: () => {},
  });

  await handler({
    session: baseSession,
    clientInfo: baseClientInfo,
    jellyfinConfig: {},
    itemId: 'ep-1',
  });

  const loadCommand = commands.find((command) => command[0] === 'loadfile');
  assert.ok(loadCommand);
  const url = new URL(String(loadCommand[1]));
  assert.equal(url.searchParams.get('AudioStreamIndex'), '3');
  assert.equal(url.searchParams.has('SubtitleStreamIndex'), false);
  assert.equal(reports[0]?.subtitleStreamIndex, 4);
});

test('playback handler starts remote Play from beginning when requested despite saved plan progress', async () => {
  const commands: Array<Array<string | number>> = [];
  const reportPayloads: Array<Record<string, unknown>> = [];
  const handler = createPlayJellyfinItemInMpvHandler({
    ensureMpvConnectedForPlayback: async () => true,
    getMpvClient: () => ({ connected: true, send: () => {} }),
    resolvePlaybackPlan: async () => ({
      url: 'https://stream.example/video.m3u8?api_key=token&StartTimeTicks=35000000',
      mode: 'transcode',
      title: 'Episode 2',
      itemTitle: 'Episode 2',
      seriesTitle: null,
      seasonNumber: null,
      episodeNumber: null,
      startTimeTicks: 35_000_000,
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
    reportPlaying: (payload) => reportPayloads.push(payload as Record<string, unknown>),
    showMpvOsd: () => {},
  });

  await handler({
    session: baseSession,
    clientInfo: baseClientInfo,
    jellyfinConfig: {},
    itemId: 'item-2',
    startTimeTicksOverride: 0,
    fallbackToPlanStartTimeOnZeroOverride: false,
  });

  const loadCommand = commands.find((command) => command[0] === 'loadfile');
  assert.ok(loadCommand);
  const loadedUrl = String(loadCommand[1] ?? '');
  const parsed = new URL(loadedUrl);
  assert.equal(parsed.searchParams.get('StartTimeTicks'), null);
  assert.equal(
    commands.some((command) => command[0] === 'seek'),
    false,
  );
  assert.equal(reportPayloads[0]?.positionTicks, 0);
});

test('playback handler disables mpv subtitle selection before Jellyfin media loads', async () => {
  const commands: Array<Array<string | number>> = [];
  const handler = createPlayJellyfinItemInMpvHandler({
    ensureMpvConnectedForPlayback: async () => true,
    getMpvClient: () => ({ connected: true, send: () => {} }),
    resolvePlaybackPlan: async () => ({
      url: 'https://stream.example/video.m3u8',
      mode: 'direct',
      title: 'Episode 1',
      itemTitle: 'Episode 1',
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
    itemId: 'item-1',
  });

  const loadIndex = commands.findIndex((command) => command[0] === 'loadfile');
  assert.ok(loadIndex > 0);
  assert.ok(
    commands.findIndex(
      (command, index) =>
        index < loadIndex &&
        command[0] === 'script-message' &&
        command[1] === 'subminer-managed-subtitles-loading',
    ) >= 0,
  );
  assert.ok(
    commands.findIndex(
      (command, index) =>
        index < loadIndex &&
        command[0] === 'set_property' &&
        command[1] === 'sid' &&
        command[2] === 'no',
    ) >= 0,
  );
  assert.ok(
    commands.findIndex(
      (command, index) =>
        index < loadIndex &&
        command[0] === 'set_property' &&
        command[1] === 'secondary-sid' &&
        command[2] === 'no',
    ) >= 0,
  );
  assert.ok(
    commands.findIndex(
      (command, index) =>
        index < loadIndex &&
        command[0] === 'set_property' &&
        command[1] === 'sub-visibility' &&
        command[2] === 'no',
    ) >= 0,
  );
  assert.ok(
    commands.findIndex(
      (command, index) =>
        index < loadIndex &&
        command[0] === 'set_property' &&
        command[1] === 'secondary-sub-visibility' &&
        command[2] === 'no',
    ) >= 0,
  );
  assert.equal(
    commands[loadIndex]?.[4],
    'sid=no,secondary-sid=no,sub-auto=no,sub-visibility=no,secondary-sub-visibility=no',
  );
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
    updateCurrentMediaTitle: (title) => {
      timeline.push(`title:${title}`);
    },
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

test('playback handler arms unloaded active playback before loading mpv media', async () => {
  const timeline: string[] = [];
  const handler = createPlayJellyfinItemInMpvHandler({
    ensureMpvConnectedForPlayback: async () => true,
    getMpvClient: () => ({ connected: true, send: () => {} }),
    resolvePlaybackPlan: async () => ({
      url: 'https://stream.example/video.m3u8',
      mode: 'direct',
      title: 'Episode 1',
      itemTitle: 'Episode 1',
      seriesTitle: null,
      seasonNumber: null,
      episodeNumber: null,
      startTimeTicks: 0,
      audioStreamIndex: null,
      subtitleStreamIndex: null,
    }),
    applyJellyfinMpvDefaults: () => {},
    showVisibleOverlay: () => {},
    sendMpvCommand: (command) => timeline.push(`cmd:${command[0]}`),
    armQuitOnDisconnect: () => {},
    schedule: () => {},
    convertTicksToSeconds: (ticks) => ticks / 10_000_000,
    preloadExternalSubtitles: () => {},
    setActivePlayback: (state) => timeline.push(`active:${String(state.loadedMediaPath)}`),
    setLastProgressAtMs: () => {},
    reportPlaying: () => {},
    showMpvOsd: () => {},
  });

  await handler({
    session: baseSession,
    clientInfo: baseClientInfo,
    jellyfinConfig: {},
    itemId: 'item-1',
  });

  assert.ok(timeline.indexOf('active:null') >= 0);
  assert.ok(timeline.indexOf('active:null') < timeline.indexOf('cmd:loadfile'));
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

  const loadCommand = commands.find((command) => command[0] === 'loadfile');
  assert.ok(loadCommand);
  const loadedUrl = String(loadCommand[1] ?? '');
  const parsed = new URL(loadedUrl);
  assert.equal(parsed.searchParams.get('StartTimeTicks'), '55000000');
  assert.equal(
    commands.some((command) => command[0] === 'seek'),
    false,
  );
});

test('playback handler keeps Jellyfin resume ticks when remote start override is zero', async () => {
  const commands: Array<Array<string | number>> = [];
  const reportPayloads: Array<Record<string, unknown>> = [];
  const handler = createPlayJellyfinItemInMpvHandler({
    ensureMpvConnectedForPlayback: async () => true,
    getMpvClient: () => ({ connected: true, send: () => {} }),
    resolvePlaybackPlan: async () => ({
      url: 'https://stream.example/video.m3u8?api_key=token&StartTimeTicks=35000000',
      mode: 'transcode',
      title: 'Episode 2',
      itemTitle: 'Episode 2',
      seriesTitle: null,
      seasonNumber: null,
      episodeNumber: null,
      startTimeTicks: 35_000_000,
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
    reportPlaying: (payload) => reportPayloads.push(payload as Record<string, unknown>),
    showMpvOsd: () => {},
  });

  await handler({
    session: baseSession,
    clientInfo: baseClientInfo,
    jellyfinConfig: {},
    itemId: 'item-2',
    startTimeTicksOverride: 0,
    fallbackToPlanStartTimeOnZeroOverride: true,
  });

  const loadCommand = commands.find((command) => command[0] === 'loadfile');
  assert.ok(loadCommand);
  const loadedUrl = String(loadCommand[1] ?? '');
  const parsed = new URL(loadedUrl);
  assert.equal(parsed.searchParams.get('StartTimeTicks'), '35000000');
  assert.equal(
    commands.some((command) => command[0] === 'seek'),
    false,
  );
  assert.equal(reportPayloads[0]?.positionTicks, 35_000_000);
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

  assert.deepEqual(
    commands.find((command) => command[0] === 'loadfile'),
    [
      'loadfile',
      'https://stream.example/video.m3u8',
      'replace',
      -1,
      'sid=no,secondary-sid=no,sub-auto=no,sub-visibility=no,secondary-sub-visibility=no',
    ],
  );
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

  assert.deepEqual(
    commands.find((command) => command[0] === 'loadfile'),
    [
      'loadfile',
      'https://stream.example/video.m3u8',
      'replace',
      -1,
      'sid=no,secondary-sid=no,sub-auto=no,sub-visibility=no,secondary-sub-visibility=no',
    ],
  );
});

test('playback handler handles rejected best-effort hook promises', async () => {
  const commands: Array<Array<string | number>> = [];
  const handler = createPlayJellyfinItemInMpvHandler({
    ensureMpvConnectedForPlayback: async () => true,
    getMpvClient: () => ({ connected: true, send: () => {} }),
    resolvePlaybackPlan: async () => ({
      url: 'https://stream.example/video.m3u8',
      mode: 'direct',
      title: 'Episode 5',
      itemTitle: 'Episode 5',
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
    updateCurrentMediaTitle: async () => {
      throw new Error('title async unavailable');
    },
    recordJellyfinPlaybackMetadata: async () => {
      throw new Error('stats async unavailable');
    },
  });

  await handler({
    session: baseSession,
    clientInfo: baseClientInfo,
    jellyfinConfig: {},
    itemId: 'item-5',
  });
  await Promise.resolve();
  await Promise.resolve();

  assert.deepEqual(
    commands.find((command) => command[0] === 'loadfile'),
    [
      'loadfile',
      'https://stream.example/video.m3u8',
      'replace',
      -1,
      'sid=no,secondary-sid=no,sub-auto=no,sub-visibility=no,secondary-sub-visibility=no',
    ],
  );
});
