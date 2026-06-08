import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createBindMpvClientEventHandlers,
  createHandleMpvConnectionChangeHandler,
  createHandleMpvSubtitleTimingHandler,
} from './mpv-client-event-bindings';

test('mpv connection handler reports stop and quits when disconnect guard passes', () => {
  const calls: string[] = [];
  const handler = createHandleMpvConnectionChangeHandler({
    reportJellyfinRemoteStopped: () => calls.push('report-stop'),
    refreshDiscordPresence: () => calls.push('presence-refresh'),
    syncOverlayMpvSubtitleSuppression: () => calls.push('sync-overlay-mpv-sub'),
    hasInitialPlaybackQuitOnDisconnectArg: () => true,
    isOverlayRuntimeInitialized: () => false,
    shouldQuitOnDisconnectWhenOverlayRuntimeInitialized: () => false,
    isQuitOnDisconnectArmed: () => true,
    scheduleQuitCheck: (callback) => {
      calls.push('schedule');
      callback();
    },
    isMpvConnected: () => false,
    quitApp: () => calls.push('quit'),
  });

  handler({ connected: false });
  assert.deepEqual(calls, ['presence-refresh', 'report-stop', 'schedule', 'quit']);
});

test('mpv connection handler syncs overlay subtitle suppression on connect', () => {
  const calls: string[] = [];
  const deps: Parameters<typeof createHandleMpvConnectionChangeHandler>[0] & {
    scheduleCharacterDictionarySync: () => void;
  } = {
    reportJellyfinRemoteStopped: () => calls.push('report-stop'),
    refreshDiscordPresence: () => calls.push('presence-refresh'),
    syncOverlayMpvSubtitleSuppression: () => calls.push('sync-overlay-mpv-sub'),
    scheduleCharacterDictionarySync: () => calls.push('dict-sync'),
    hasInitialPlaybackQuitOnDisconnectArg: () => true,
    isOverlayRuntimeInitialized: () => false,
    shouldQuitOnDisconnectWhenOverlayRuntimeInitialized: () => false,
    isQuitOnDisconnectArmed: () => true,
    scheduleQuitCheck: () => {
      calls.push('schedule');
    },
    isMpvConnected: () => false,
    quitApp: () => calls.push('quit'),
  };
  const handler = createHandleMpvConnectionChangeHandler(deps);

  handler({ connected: true });

  assert.deepEqual(calls, ['presence-refresh', 'sync-overlay-mpv-sub']);
});

test('mpv connection handler runs connected hook on connect', () => {
  const calls: string[] = [];
  const handler = createHandleMpvConnectionChangeHandler({
    reportJellyfinRemoteStopped: () => calls.push('report-stop'),
    refreshDiscordPresence: () => calls.push('presence-refresh'),
    syncOverlayMpvSubtitleSuppression: () => calls.push('sync-overlay-mpv-sub'),
    onConnected: () => calls.push('connected-hook'),
    hasInitialPlaybackQuitOnDisconnectArg: () => false,
    isOverlayRuntimeInitialized: () => false,
    shouldQuitOnDisconnectWhenOverlayRuntimeInitialized: () => false,
    isQuitOnDisconnectArmed: () => false,
    scheduleQuitCheck: () => calls.push('schedule'),
    isMpvConnected: () => false,
    quitApp: () => calls.push('quit'),
  });

  handler({ connected: true });
  handler({ connected: false });

  assert.deepEqual(calls, [
    'presence-refresh',
    'sync-overlay-mpv-sub',
    'connected-hook',
    'presence-refresh',
    'report-stop',
  ]);
});

test('mpv connection handler quits standalone youtube playback even after overlay runtime init', () => {
  const calls: string[] = [];
  const handler = createHandleMpvConnectionChangeHandler({
    reportJellyfinRemoteStopped: () => calls.push('report-stop'),
    refreshDiscordPresence: () => calls.push('presence-refresh'),
    syncOverlayMpvSubtitleSuppression: () => calls.push('sync-overlay-mpv-sub'),
    hasInitialPlaybackQuitOnDisconnectArg: () => true,
    isOverlayRuntimeInitialized: () => true,
    shouldQuitOnDisconnectWhenOverlayRuntimeInitialized: () => true,
    isQuitOnDisconnectArmed: () => true,
    scheduleQuitCheck: (callback) => {
      calls.push('schedule');
      callback();
    },
    isMpvConnected: () => false,
    quitApp: () => calls.push('quit'),
  });

  handler({ connected: false });
  assert.deepEqual(calls, ['presence-refresh', 'report-stop', 'schedule', 'quit']);
});

test('mpv connection handler keeps overlay-initialized non-youtube sessions alive', () => {
  const calls: string[] = [];
  const handler = createHandleMpvConnectionChangeHandler({
    reportJellyfinRemoteStopped: () => calls.push('report-stop'),
    refreshDiscordPresence: () => calls.push('presence-refresh'),
    syncOverlayMpvSubtitleSuppression: () => calls.push('sync-overlay-mpv-sub'),
    hasInitialPlaybackQuitOnDisconnectArg: () => true,
    isOverlayRuntimeInitialized: () => true,
    shouldQuitOnDisconnectWhenOverlayRuntimeInitialized: () => false,
    isQuitOnDisconnectArmed: () => true,
    scheduleQuitCheck: () => {
      calls.push('schedule');
    },
    isMpvConnected: () => false,
    quitApp: () => calls.push('quit'),
  });

  handler({ connected: false });
  assert.deepEqual(calls, ['presence-refresh', 'report-stop']);
});

test('mpv subtitle timing handler skips blank subtitle recording but still checks AniList time', () => {
  const calls: string[] = [];
  const handler = createHandleMpvSubtitleTimingHandler({
    recordImmersionSubtitleLine: () => calls.push('immersion'),
    hasSubtitleTimingTracker: () => true,
    recordSubtitleTiming: () => calls.push('timing'),
    maybeRunAnilistPostWatchUpdate: async (options) => {
      calls.push(`post-watch:${options?.watchedSeconds}`);
    },
    logError: () => calls.push('error'),
  });

  handler({ text: '   ', start: 1, end: 2 });
  assert.deepEqual(calls, ['post-watch:2']);
});

test('mpv subtitle timing handler runs AniList without timing tracker and passes subtitle time', () => {
  const calls: string[] = [];
  const handler = createHandleMpvSubtitleTimingHandler({
    recordImmersionSubtitleLine: (text, start, end) =>
      calls.push(`immersion:${text}:${start}:${end}`),
    hasSubtitleTimingTracker: () => false,
    recordSubtitleTiming: () => calls.push('timing'),
    maybeRunAnilistPostWatchUpdate: async (options) => {
      calls.push(`post-watch:${options?.watchedSeconds}`);
    },
    logError: () => calls.push('error'),
  });

  handler({ text: 'line', start: 899, end: 901 });

  assert.deepEqual(calls, ['immersion:line:899:901', 'post-watch:901']);
});

test('mpv subtitle timing handler skips invalid cue pairs until timing is complete', () => {
  const calls: string[] = [];
  const handler = createHandleMpvSubtitleTimingHandler({
    recordImmersionSubtitleLine: (text, start, end) =>
      calls.push(`immersion:${text}:${start}:${end}`),
    hasSubtitleTimingTracker: () => true,
    recordSubtitleTiming: (text, start, end) => calls.push(`timing:${text}:${start}:${end}`),
    maybeRunAnilistPostWatchUpdate: async (options) => {
      calls.push(`post-watch:${options?.watchedSeconds}`);
    },
    logError: () => calls.push('error'),
  });

  handler({ text: 'line', start: 953.991, end: 953.891 });
  handler({ text: 'line', start: 953.991, end: 956.56 });

  assert.deepEqual(calls, [
    'post-watch:953.991',
    'immersion:line:953.991:956.56',
    'timing:line:953.991:956.56',
    'post-watch:956.56',
  ]);
});

test('mpv event bindings register all expected events', () => {
  const seenEvents: string[] = [];
  const bindHandlers = createBindMpvClientEventHandlers({
    onConnectionChange: () => {},
    onSubtitleChange: () => {},
    onSubtitleAssChange: () => {},
    onSecondarySubtitleChange: () => {},
    onSubtitleTrackChange: () => {},
    onSubtitleTrackListChange: () => {},
    onSubtitleTiming: () => {},
    onMediaPathChange: () => {},
    onMediaTitleChange: () => {},
    onTimePosChange: () => {},
    onDurationChange: () => {},
    onPauseChange: () => {},
    onFullscreenChange: () => {},
    onSubtitleMetricsChange: () => {},
    onSecondarySubtitleVisibility: () => {},
  });

  bindHandlers({
    on: (event) => {
      seenEvents.push(event);
    },
  });

  assert.deepEqual(seenEvents, [
    'connection-change',
    'subtitle-change',
    'subtitle-ass-change',
    'secondary-subtitle-change',
    'subtitle-track-change',
    'subtitle-track-list-change',
    'subtitle-timing',
    'media-path-change',
    'media-title-change',
    'time-pos-change',
    'duration-change',
    'pause-change',
    'fullscreen-change',
    'subtitle-metrics-change',
    'secondary-subtitle-visibility',
  ]);
});
