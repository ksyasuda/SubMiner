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

test('mpv subtitle timing handler ignores blank subtitle lines', () => {
  const calls: string[] = [];
  const handler = createHandleMpvSubtitleTimingHandler({
    recordImmersionSubtitleLine: () => calls.push('immersion'),
    hasSubtitleTimingTracker: () => true,
    recordSubtitleTiming: () => calls.push('timing'),
    maybeRunAnilistPostWatchUpdate: async () => {
      calls.push('post-watch');
    },
    logError: () => calls.push('error'),
  });

  handler({ text: '   ', start: 1, end: 2 });
  assert.deepEqual(calls, []);
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
    'subtitle-metrics-change',
    'secondary-subtitle-visibility',
  ]);
});
