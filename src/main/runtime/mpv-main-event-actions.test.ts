import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createHandleMpvMediaPathChangeHandler,
  createHandleMpvMediaTitleChangeHandler,
  createHandleMpvPauseChangeHandler,
  createHandleMpvSecondarySubtitleChangeHandler,
  createHandleMpvSecondarySubtitleVisibilityHandler,
  createHandleMpvSubtitleAssChangeHandler,
  createHandleMpvSubtitleChangeHandler,
  createHandleMpvSubtitleMetricsChangeHandler,
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

test('subtitle ass change handler updates state and broadcasts', () => {
  const calls: string[] = [];
  const handler = createHandleMpvSubtitleAssChangeHandler({
    setCurrentSubAssText: (text) => calls.push(`set:${text}`),
    broadcastSubtitleAss: (text) => calls.push(`broadcast:${text}`),
  });

  handler({ text: '{\\an8}line' });
  assert.deepEqual(calls, ['set:{\\an8}line', 'broadcast:{\\an8}line']);
});

test('secondary subtitle change handler broadcasts text', () => {
  const seen: string[] = [];
  const handler = createHandleMpvSecondarySubtitleChangeHandler({
    broadcastSecondarySubtitle: (text) => seen.push(text),
  });

  handler({ text: 'secondary' });
  assert.deepEqual(seen, ['secondary']);
});

test('media path change handler reports stop for empty path and probes media key', () => {
  const calls: string[] = [];
  const handler = createHandleMpvMediaPathChangeHandler({
    updateCurrentMediaPath: (path) => calls.push(`path:${path}`),
    reportJellyfinRemoteStopped: () => calls.push('stopped'),
    restoreMpvSubVisibility: () => calls.push('restore-mpv-sub'),
    resetSubtitleSidebarEmbeddedLayout: () => calls.push('reset-sidebar-layout'),
    getCurrentAnilistMediaKey: () => 'show:1',
    resetAnilistMediaTracking: (mediaKey) => calls.push(`reset:${String(mediaKey)}`),
    maybeProbeAnilistDuration: (mediaKey) => calls.push(`probe:${mediaKey}`),
    ensureAnilistMediaGuess: (mediaKey) => calls.push(`guess:${mediaKey}`),
    syncImmersionMediaState: () => calls.push('sync'),
    flushPlaybackPositionOnMediaPathClear: () => calls.push('flush-playback'),
    scheduleCharacterDictionarySync: () => calls.push('dict-sync'),
    refreshDiscordPresence: () => calls.push('presence'),
  });

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
  const handler = createHandleMpvMediaPathChangeHandler({
    updateCurrentMediaPath: (path) => calls.push(`path:${path}`),
    reportJellyfinRemoteStopped: () => calls.push('stopped'),
    restoreMpvSubVisibility: () => calls.push('restore-mpv-sub'),
    resetSubtitleSidebarEmbeddedLayout: () => calls.push('reset-sidebar-layout'),
    getCurrentAnilistMediaKey: () => null,
    resetAnilistMediaTracking: (mediaKey) => calls.push(`reset:${String(mediaKey)}`),
    maybeProbeAnilistDuration: (mediaKey) => calls.push(`probe:${mediaKey}`),
    ensureAnilistMediaGuess: (mediaKey) => calls.push(`guess:${mediaKey}`),
    syncImmersionMediaState: () => calls.push('sync'),
    flushPlaybackPositionOnMediaPathClear: () => calls.push('flush-playback'),
    scheduleCharacterDictionarySync: () => calls.push('dict-sync'),
    signalAutoplayReadyIfWarm: (path) => calls.push(`autoplay:${path}`),
    refreshDiscordPresence: () => calls.push('presence'),
  });

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
  const handler = createHandleMpvMediaPathChangeHandler({
    updateCurrentMediaPath: (path) => calls.push(`path:${path}`),
    reportJellyfinRemoteStopped: () => calls.push('stopped'),
    restoreMpvSubVisibility: () => calls.push('restore-mpv-sub'),
    resetSubtitleSidebarEmbeddedLayout: () => calls.push('reset-sidebar-layout'),
    getCurrentAnilistMediaKey: () => null,
    resetAnilistMediaTracking: (mediaKey) => calls.push(`reset:${String(mediaKey)}`),
    maybeProbeAnilistDuration: (mediaKey) => calls.push(`probe:${mediaKey}`),
    ensureAnilistMediaGuess: (mediaKey) => calls.push(`guess:${mediaKey}`),
    syncImmersionMediaState: () => calls.push('sync'),
    scheduleCharacterDictionarySync: () => calls.push('dict-sync'),
    refreshDiscordPresence: () => calls.push('presence'),
  });

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
  const handler = createHandleMpvMediaPathChangeHandler({
    updateCurrentMediaPath: (path) => calls.push(`path:${path}`),
    reportJellyfinRemoteStopped: () => calls.push('stopped'),
    restoreMpvSubVisibility: () => calls.push('restore-mpv-sub'),
    resetSubtitleSidebarEmbeddedLayout: () => calls.push('reset-sidebar-layout'),
    getCurrentAnilistMediaKey: () => null,
    resetAnilistMediaTracking: (mediaKey) => calls.push(`reset:${String(mediaKey)}`),
    maybeProbeAnilistDuration: (mediaKey) => calls.push(`probe:${mediaKey}`),
    ensureAnilistMediaGuess: (mediaKey) => calls.push(`guess:${mediaKey}`),
    syncImmersionMediaState: () => calls.push('sync'),
    markJellyfinRemotePlaybackLoaded: (path) => calls.push(`jellyfin-loaded:${path}`),
    refreshDiscordPresence: () => calls.push('presence'),
  });

  handler({ path: 'https://stream.example/video.m3u8' });

  assert.ok(calls.includes('jellyfin-loaded:https://stream.example/video.m3u8'));
  assert.equal(calls.includes('stopped'), false);
});

test('media title change handler clears guess state without re-scheduling character dictionary sync', () => {
  const calls: string[] = [];
  const deps: Parameters<typeof createHandleMpvMediaTitleChangeHandler>[0] & {
    scheduleCharacterDictionarySync: () => void;
  } = {
    updateCurrentMediaTitle: (title) => calls.push(`title:${title}`),
    resetAnilistMediaGuessState: () => calls.push('reset-guess'),
    notifyImmersionTitleUpdate: (title) => calls.push(`notify:${title}`),
    syncImmersionMediaState: () => calls.push('sync'),
    scheduleCharacterDictionarySync: () => calls.push('dict-sync'),
    refreshDiscordPresence: () => calls.push('presence'),
  };
  const handler = createHandleMpvMediaTitleChangeHandler(deps);

  handler({ title: 'Episode 1' });
  assert.deepEqual(calls, [
    'title:Episode 1',
    'reset-guess',
    'notify:Episode 1',
    'sync',
    'presence',
  ]);
});

test('time-pos and pause handlers report progress with correct urgency', () => {
  const calls: string[] = [];
  const timeHandler = createHandleMpvTimePosChangeHandler({
    recordPlaybackPosition: (time) => calls.push(`time:${time}`),
    reportJellyfinRemoteProgress: (force) => calls.push(`progress:${force ? 'force' : 'normal'}`),
    refreshDiscordPresence: () => calls.push('presence'),
    maybeRunAnilistPostWatchUpdate: async () => {
      calls.push('post-watch');
    },
    logError: () => calls.push('post-watch-error'),
  });
  const pauseHandler = createHandleMpvPauseChangeHandler({
    recordPauseState: (paused) => calls.push(`pause:${paused ? 'yes' : 'no'}`),
    reportJellyfinRemoteProgress: (force) => calls.push(`progress:${force ? 'force' : 'normal'}`),
    refreshDiscordPresence: () => calls.push('presence'),
  });

  timeHandler({ time: 12.5 });
  pauseHandler({ paused: true });
  assert.deepEqual(calls, [
    'time:12.5',
    'progress:normal',
    'presence',
    'post-watch',
    'pause:yes',
    'progress:force',
    'presence',
  ]);
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

test('subtitle metrics change handler forwards patch payload', () => {
  let received: Record<string, unknown> | null = null;
  const handler = createHandleMpvSubtitleMetricsChangeHandler({
    updateSubtitleRenderMetrics: (patch) => {
      received = patch;
    },
  });

  const patch = { fontSize: 48 };
  handler({ patch });
  assert.deepEqual(received, patch);
});

test('secondary subtitle visibility handler stores visibility flag', () => {
  const seen: boolean[] = [];
  const handler = createHandleMpvSecondarySubtitleVisibilityHandler({
    setPreviousSecondarySubVisibility: (visible) => seen.push(visible),
  });

  handler({ visible: true });
  handler({ visible: false });
  assert.deepEqual(seen, [true, false]);
});
