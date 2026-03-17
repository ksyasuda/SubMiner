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

test('subtitle change handler updates state, broadcasts, and forwards', () => {
  const calls: string[] = [];
  const handler = createHandleMpvSubtitleChangeHandler({
    setCurrentSubText: (text) => calls.push(`set:${text}`),
    broadcastSubtitle: (payload) => calls.push(`broadcast:${payload.text}`),
    onSubtitleChange: (text) => calls.push(`process:${text}`),
    refreshDiscordPresence: () => calls.push('presence'),
  });

  handler({ text: 'line' });
  assert.deepEqual(calls, ['set:line', 'broadcast:line', 'process:line', 'presence']);
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
    getCurrentAnilistMediaKey: () => 'show:1',
    resetAnilistMediaTracking: (mediaKey) => calls.push(`reset:${String(mediaKey)}`),
    maybeProbeAnilistDuration: (mediaKey) => calls.push(`probe:${mediaKey}`),
    ensureAnilistMediaGuess: (mediaKey) => calls.push(`guess:${mediaKey}`),
    syncImmersionMediaState: () => calls.push('sync'),
    scheduleCharacterDictionarySync: () => calls.push('dict-sync'),
    signalAutoplayReadyIfWarm: (path) => calls.push(`autoplay:${path}`),
    refreshDiscordPresence: () => calls.push('presence'),
  });

  handler({ path: '' });
  assert.deepEqual(calls, [
    'path:',
    'stopped',
    'restore-mpv-sub',
    'reset:show:1',
    'probe:show:1',
    'guess:show:1',
    'sync',
    'presence',
  ]);
});

test('media path change handler signals autoplay-ready fast path for warm non-empty media', () => {
  const calls: string[] = [];
  const handler = createHandleMpvMediaPathChangeHandler({
    updateCurrentMediaPath: (path) => calls.push(`path:${path}`),
    reportJellyfinRemoteStopped: () => calls.push('stopped'),
    restoreMpvSubVisibility: () => calls.push('restore-mpv-sub'),
    getCurrentAnilistMediaKey: () => null,
    resetAnilistMediaTracking: (mediaKey) => calls.push(`reset:${String(mediaKey)}`),
    maybeProbeAnilistDuration: (mediaKey) => calls.push(`probe:${mediaKey}`),
    ensureAnilistMediaGuess: (mediaKey) => calls.push(`guess:${mediaKey}`),
    syncImmersionMediaState: () => calls.push('sync'),
    scheduleCharacterDictionarySync: () => calls.push('dict-sync'),
    signalAutoplayReadyIfWarm: (path) => calls.push(`autoplay:${path}`),
    refreshDiscordPresence: () => calls.push('presence'),
  });

  handler({ path: '/tmp/video.mkv' });

  assert.deepEqual(calls, [
    'path:/tmp/video.mkv',
    'reset:null',
    'sync',
    'dict-sync',
    'autoplay:/tmp/video.mkv',
    'presence',
  ]);
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
    'pause:yes',
    'progress:force',
    'presence',
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
