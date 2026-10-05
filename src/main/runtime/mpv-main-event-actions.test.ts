import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createHandleMpvMediaPathChangeHandler,
  createHandleMpvPauseChangeHandler,
  createHandleMpvSubtitleChangeHandler,
  createHandleMpvTimePosChangeHandler,
} from './mpv-main-event-actions';

test('subtitle change handler updates state and forwards uncached text without raw broadcast', () => {
  const calls: string[] = [];
  const handler = createHandleMpvSubtitleChangeHandler({
    setCurrentSubText: (text) => calls.push(`set:${text}`),
    getImmediateSubtitlePayload: () => null,
    broadcastSubtitle: (payload) => calls.push(`broadcast:${payload.text}`),
    onSubtitleChange: (text) => calls.push(`process:${text}`),
    refreshDiscordPresence: () => calls.push('presence'),
  });

  handler({ text: 'line' });
  assert.deepEqual(calls, ['set:line', 'process:line', 'presence']);
});

test('subtitle change handler consistently forwards resolved canonical text', () => {
  const calls: string[] = [];
  const handler = createHandleMpvSubtitleChangeHandler({
    resolveSubtitleText: () => '今　手にある物差しでは',
    setCurrentSubText: (text) => calls.push(`set:${text}`),
    getImmediateSubtitlePayload: (text) => {
      calls.push(`lookup:${text}`);
      return null;
    },
    broadcastSubtitle: () => {},
    onSubtitleChange: (text) => calls.push(`process:${text}`),
    refreshDiscordPresence: () => {},
  });

  handler({ text: '今今今手手手ににに' });

  assert.deepEqual(calls, [
    'set:今　手にある物差しでは',
    'lookup:今　手にある物差しでは',
    'process:今　手にある物差しでは',
  ]);
});

test('subtitle change handler clears immediately for empty subtitle text', () => {
  const calls: string[] = [];
  const handler = createHandleMpvSubtitleChangeHandler({
    setCurrentSubText: (text) => calls.push(`set:${text}`),
    getImmediateSubtitlePayload: () => null,
    broadcastSubtitle: (payload) =>
      calls.push(`broadcast:${payload.text}:${payload.tokens === null ? 'plain' : 'annotated'}`),
    onSubtitleChange: (text) => calls.push(`process:${text}`),
    refreshDiscordPresence: () => calls.push('presence'),
  });

  handler({ text: '' });
  assert.deepEqual(calls, ['set:', 'broadcast::plain', 'process:', 'presence']);
});

test('subtitle change handler broadcasts cached annotated payload immediately when available', () => {
  const payloads: Array<{ text: string; tokens: unknown[] | null }> = [];
  const calls: string[] = [];
  const handler = createHandleMpvSubtitleChangeHandler({
    setCurrentSubText: (text) => calls.push(`set:${text}`),
    getImmediateSubtitlePayload: (text) => {
      calls.push(`lookup:${text}`);
      return { text, tokens: [] };
    },
    broadcastSubtitle: (payload) => {
      payloads.push(payload);
      calls.push(`broadcast:${payload.tokens === null ? 'plain' : 'annotated'}`);
    },
    onSubtitleChange: (text) => calls.push(`process:${text}`),
    refreshDiscordPresence: () => calls.push('presence'),
  });

  handler({ text: 'line' });

  assert.deepEqual(payloads, [{ text: 'line', tokens: [] }]);
  assert.deepEqual(calls, [
    'set:line',
    'lookup:line',
    'process:line',
    'broadcast:annotated',
    'presence',
  ]);
});

test('subtitle change handler logs debug when a cached payload is emitted immediately', () => {
  const debugs: string[] = [];
  const handler = createHandleMpvSubtitleChangeHandler({
    setCurrentSubText: () => {},
    getImmediateSubtitlePayload: (text) => (text ? { text, tokens: [] } : null),
    broadcastSubtitle: () => {},
    onSubtitleChange: () => {},
    refreshDiscordPresence: () => {},
    logDebug: (message) => debugs.push(message),
  });

  handler({ text: 'キャッシュ済みの行' });
  handler({ text: '' });

  assert.equal(debugs.length, 1);
  assert.match(debugs[0]!, /cached subtitle/);
});

test('subtitle change handler emits cached annotation after forwarding the subtitle change', () => {
  const calls: string[] = [];
  const handler = createHandleMpvSubtitleChangeHandler({
    setCurrentSubText: (text) => calls.push(`set:${text}`),
    getImmediateSubtitlePayload: (text) => {
      calls.push(`lookup:${text}`);
      return { text, tokens: [] };
    },
    emitImmediateSubtitle: (payload) => {
      calls.push(`emit:${payload.tokens === null ? 'plain' : 'annotated'}`);
    },
    broadcastSubtitle: (payload) => {
      calls.push(`broadcast:${payload.tokens === null ? 'plain' : 'annotated'}`);
    },
    onSubtitleChange: (text) => calls.push(`process:${text}`),
    refreshDiscordPresence: () => calls.push('presence'),
  });

  handler({ text: 'line' });

  assert.deepEqual(calls, [
    'set:line',
    'lookup:line',
    'process:line',
    'emit:annotated',
    'presence',
  ]);
});

type MediaPathDeps = Parameters<typeof createHandleMpvMediaPathChangeHandler>[0];

// Required deps record into `calls`; optional hooks are added per test via overrides.
function createMediaPathDeps(
  calls: string[],
  overrides: Partial<MediaPathDeps> = {},
): MediaPathDeps {
  return {
    updateCurrentMediaPath: (path) => calls.push(`path:${path}`),
    reportJellyfinRemoteStopped: () => calls.push('stopped'),
    restoreMpvSubVisibility: () => calls.push('restore-mpv-sub'),
    resetSubtitleSidebarEmbeddedLayout: () => calls.push('reset-sidebar-layout'),
    getCurrentAnilistMediaKey: () => null,
    resetAnilistMediaTracking: (mediaKey) => calls.push(`reset:${String(mediaKey)}`),
    maybeProbeAnilistDuration: (mediaKey) => calls.push(`probe:${mediaKey}`),
    ensureAnilistMediaGuess: (mediaKey) => calls.push(`guess:${mediaKey}`),
    syncImmersionMediaState: () => calls.push('sync'),
    refreshDiscordPresence: () => calls.push('presence'),
    ...overrides,
  };
}

test('media path change handler reports stop for empty path and probes media key', () => {
  const calls: string[] = [];
  const handler = createHandleMpvMediaPathChangeHandler(
    createMediaPathDeps(calls, {
      getCurrentAnilistMediaKey: () => 'show:1',
      flushPlaybackPositionOnMediaPathClear: () => calls.push('flush-playback'),
      scheduleCharacterDictionarySync: () => calls.push('dict-sync'),
    }),
  );

  handler({ path: '' });
  assert.deepEqual(calls, [
    'flush-playback',
    'path:',
    'reset-sidebar-layout',
    'stopped',
    'restore-mpv-sub',
    'reset:show:1',
    'probe:show:1',
    'guess:show:1',
    'sync',
    'presence',
  ]);
});

test('media path change handler signals autoplay readiness from warm media path', () => {
  const calls: string[] = [];
  const handler = createHandleMpvMediaPathChangeHandler(
    createMediaPathDeps(calls, {
      flushPlaybackPositionOnMediaPathClear: () => calls.push('flush-playback'),
      scheduleCharacterDictionarySync: () => calls.push('dict-sync'),
      signalAutoplayReadyIfWarm: (path) => calls.push(`autoplay:${path}`),
    }),
  );

  handler({ path: '/tmp/video.mkv' });

  assert.deepEqual(calls, [
    'path:/tmp/video.mkv',
    'reset-sidebar-layout',
    'reset:null',
    'sync',
    'dict-sync',
    'autoplay:/tmp/video.mkv',
    'presence',
  ]);
});

test('media path change handler schedules character dictionary once per media path', () => {
  const calls: string[] = [];
  const handler = createHandleMpvMediaPathChangeHandler(
    createMediaPathDeps(calls, {
      scheduleCharacterDictionarySync: () => calls.push('dict-sync'),
    }),
  );

  handler({ path: '/tmp/video.mkv' });
  handler({ path: '/tmp/video.mkv' });
  handler({ path: '/tmp/next-video.mkv' });
  handler({ path: '' });
  handler({ path: '/tmp/video.mkv' });

  assert.deepEqual(
    calls.filter((call) => call === 'dict-sync'),
    ['dict-sync', 'dict-sync', 'dict-sync'],
  );
});

test('media path change handler marks Jellyfin remote playback loaded from media path', () => {
  const calls: string[] = [];
  const handler = createHandleMpvMediaPathChangeHandler(
    createMediaPathDeps(calls, {
      markJellyfinRemotePlaybackLoaded: (path) => calls.push(`jellyfin-loaded:${path}`),
    }),
  );

  handler({ path: 'https://stream.example/video.m3u8' });

  assert.ok(calls.includes('jellyfin-loaded:https://stream.example/video.m3u8'));
  assert.equal(calls.includes('stopped'), false);
});

test('time-pos handler forces Jellyfin progress when mpv position jumps', () => {
  const calls: string[] = [];
  const timeHandler = createHandleMpvTimePosChangeHandler({
    recordPlaybackPosition: (time) => calls.push(`time:${time}`),
    reportJellyfinRemoteProgress: (force) => calls.push(`progress:${force ? 'force' : 'normal'}`),
    refreshDiscordPresence: () => calls.push('presence'),
    maybeRunAnilistPostWatchUpdate: async () => {},
  });

  timeHandler({ time: 10 });
  timeHandler({ time: 11 });
  timeHandler({ time: 90 });
  timeHandler({ time: 30 });

  assert.deepEqual(calls, [
    'time:10',
    'progress:normal',
    'presence',
    'time:11',
    'progress:normal',
    'presence',
    'time:90',
    'progress:force',
    'presence',
    'time:30',
    'progress:force',
    'presence',
  ]);
});

test('time-pos handler treats an explicit short jump as a seek', () => {
  const updateKinds: string[] = [];
  let explicitSeekPending = false;
  const timeHandler = createHandleMpvTimePosChangeHandler({
    recordPlaybackPosition: () => {},
    reportJellyfinRemoteProgress: () => {},
    refreshDiscordPresence: () => {},
    maybeRunAnilistPostWatchUpdate: async () => {},
    consumeExplicitSeek: () => {
      const pending = explicitSeekPending;
      explicitSeekPending = false;
      return pending;
    },
    onTimePosUpdate: (_time, kind) => updateKinds.push(kind),
  });

  timeHandler({ time: 10 });
  explicitSeekPending = true;
  timeHandler({ time: 11.5 });
  timeHandler({ time: 11.6 });

  assert.deepEqual(updateKinds, ['initial', 'seek', 'playback']);
});

test('time-pos handler passes fresh playback time to AniList post-watch', async () => {
  const watchedSeconds: unknown[] = [];
  const timeHandler = createHandleMpvTimePosChangeHandler({
    recordPlaybackPosition: () => {},
    reportJellyfinRemoteProgress: () => {},
    refreshDiscordPresence: () => {},
    maybeRunAnilistPostWatchUpdate: async (options) => {
      watchedSeconds.push(options?.watchedSeconds);
    },
  });

  timeHandler({ time: 850 });
  await Promise.resolve();

  assert.deepEqual(watchedSeconds, [850]);
});

test('time-pos handler logs post-watch update rejection without blocking later handlers', async () => {
  const calls: string[] = [];
  const timeHandler = createHandleMpvTimePosChangeHandler({
    recordPlaybackPosition: (time) => calls.push(`time:${time}`),
    reportJellyfinRemoteProgress: (force) => calls.push(`progress:${force ? 'force' : 'normal'}`),
    refreshDiscordPresence: () => calls.push('presence'),
    maybeRunAnilistPostWatchUpdate: async () => {
      calls.push('post-watch');
      throw new Error('boom');
    },
    logError: (message, error) => calls.push(`error:${message}:${(error as Error).message}`),
  });
  const pauseHandler = createHandleMpvPauseChangeHandler({
    recordPauseState: (paused) => calls.push(`pause:${paused ? 'yes' : 'no'}`),
    reportJellyfinRemoteProgress: (force) => calls.push(`progress:${force ? 'force' : 'normal'}`),
    refreshDiscordPresence: () => calls.push('presence'),
  });

  timeHandler({ time: 12.5 });
  pauseHandler({ paused: true });
  await Promise.resolve();
  await Promise.resolve();

  assert.deepEqual(calls, [
    'time:12.5',
    'progress:normal',
    'presence',
    'post-watch',
    'pause:yes',
    'progress:force',
    'presence',
    'error:AniList post-watch update failed unexpectedly:boom',
  ]);
});
