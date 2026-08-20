import assert from 'node:assert/strict';
import test from 'node:test';
import { parseSubtitleCues } from '../../core/services/subtitle-cue-parser';
import { createBindMpvMainEventHandlersHandler } from './mpv-main-event-bindings';
import { resolvePrimarySubtitleText } from './primary-subtitle-text';

test('main mpv event binder wires callbacks through to runtime deps', () => {
  const handlers = new Map<string, (payload: unknown) => void>();
  const calls: string[] = [];
  let currentTime = 0;
  const seekLiveText = '少しだけ好きになる\n少しだけ好きになる';
  const seekCues = parseSubtitleCues(
    [
      '[Events]',
      'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
      'Dialogue: 1,0:01:29.00,0:01:32.00,EDJP,,0,0,0,,少しだけ好きになる',
      'Dialogue: 0,0:01:29.00,0:01:32.00,EDJP,,0,0,0,,少しだけ好きになる',
    ].join('\n'),
    'seek-ending.ass',
  );

  const bind = createBindMpvMainEventHandlersHandler({
    reportJellyfinRemoteStopped: () => calls.push('remote-stopped'),
    syncOverlayMpvSubtitleSuppression: () => calls.push('sync-overlay-mpv-sub'),
    resetSubtitleSidebarEmbeddedLayout: () => calls.push('reset-sidebar-layout'),
    hasInitialPlaybackQuitOnDisconnectArg: () => false,
    isOverlayRuntimeInitialized: () => false,
    shouldQuitOnDisconnectWhenOverlayRuntimeInitialized: () => false,
    isQuitOnDisconnectArmed: () => false,
    scheduleQuitCheck: () => {
      calls.push('schedule-quit-check');
    },
    isMpvConnected: () => false,
    quitApp: () => calls.push('quit-app'),

    recordImmersionSubtitleLine: (text) => calls.push(`immersion:${text}`),
    hasSubtitleTimingTracker: () => false,
    recordSubtitleTiming: () => calls.push('record-timing'),
    maybeRunAnilistPostWatchUpdate: async (options) => {
      calls.push(`post-watch:${options?.watchedSeconds ?? 'none'}`);
    },
    logSubtitleTimingError: () => calls.push('subtitle-error'),
    resolveSubtitleText: (liveText) =>
      resolvePrimarySubtitleText({ liveText, currentTimeSec: currentTime, cues: seekCues }),
    getCurrentLiveSubtitleText: () => seekLiveText,
    setCurrentSubText: (text) => calls.push(`set-sub:${text}`),
    getImmediateSubtitlePayload: (text) => ({ text, tokens: [] }),
    broadcastSubtitle: (payload) => calls.push(`broadcast-sub:${payload.text}`),
    onSubtitleChange: (text) => calls.push(`subtitle-change:${text}`),
    refreshDiscordPresence: () => calls.push('presence-refresh'),

    setCurrentSubAssText: (text) => calls.push(`set-ass:${text}`),
    broadcastSubtitleAss: (text) => calls.push(`broadcast-ass:${text}`),
    broadcastSecondarySubtitle: (text) => calls.push(`broadcast-secondary:${text}`),
    onSubtitleTrackChange: () => calls.push('subtitle-track-change'),
    onSecondarySubtitleTrackChange: () => calls.push('secondary-subtitle-track-change'),
    onSecondarySubtitleDelayChange: (delay) =>
      calls.push(`secondary-subtitle-delay-change:${delay}`),
    onSubtitleTrackListChange: () => calls.push('subtitle-track-list-change'),

    updateCurrentMediaPath: (path) => calls.push(`media-path:${path}`),
    restoreMpvSubVisibility: () => calls.push('restore-mpv-sub'),
    getCurrentAnilistMediaKey: () => 'media-key',
    resetAnilistMediaTracking: (key) => calls.push(`reset-media:${String(key)}`),
    maybeProbeAnilistDuration: (mediaKey) => calls.push(`probe:${mediaKey}`),
    ensureAnilistMediaGuess: (mediaKey) => calls.push(`guess:${mediaKey}`),
    syncImmersionMediaState: () => calls.push('sync-immersion'),
    signalAutoplayReadyIfWarm: (path) => calls.push(`autoplay:${path}`),
    flushPlaybackPositionOnMediaPathClear: () => calls.push('flush-playback'),

    updateCurrentMediaTitle: (title) => calls.push(`media-title:${title}`),
    resetAnilistMediaGuessState: () => calls.push('reset-guess-state'),
    notifyImmersionTitleUpdate: (title) => calls.push(`notify-title:${title}`),

    recordPlaybackPosition: (time) => calls.push(`time-pos:${time}`),
    recordMediaDuration: (duration) => calls.push(`duration:${duration}`),
    reportJellyfinRemoteProgress: (forceImmediate) =>
      calls.push(`progress:${forceImmediate ? 'force' : 'normal'}`),
    onTimePosUpdate: (time) => {
      currentTime = time;
    },
    recordPauseState: (paused) => calls.push(`pause:${paused ? 'yes' : 'no'}`),

    updateSubtitleRenderMetrics: () => calls.push('subtitle-metrics'),
    setPreviousSecondarySubVisibility: (visible) =>
      calls.push(`secondary-visible:${visible ? 'yes' : 'no'}`),
  });

  bind({
    on: (event, handler) => {
      handlers.set(event, handler as (payload: unknown) => void);
    },
  });

  handlers.get('connection-change')?.({ connected: true });
  handlers.get('subtitle-change')?.({ text: 'line' });
  handlers.get('subtitle-track-change')?.({ sid: 3 });
  handlers.get('secondary-subtitle-track-change')?.({ sid: 4 });
  handlers.get('secondary-subtitle-delay-change')?.({ delay: 0.5 });
  handlers.get('subtitle-track-list-change')?.({ trackList: [] });
  handlers.get('media-path-change')?.({ path: '/tmp/video.mkv' });
  handlers.get('media-path-change')?.({ path: '' });
  handlers.get('media-title-change')?.({ title: 'Episode 1' });
  handlers.get('subtitle-timing')?.({ text: 'timed line', start: 899, end: 901 });
  handlers.get('subtitle-change')?.({ text: seekLiveText });
  handlers.get('time-pos-change')?.({ time: 90 });
  assert.ok(calls.includes('set-sub:少しだけ好きになる'));

  handlers.get('time-pos-change')?.({ time: 2.5 });
  handlers.get('subtitle-change')?.({ text: seekLiveText });
  handlers.get('time-pos-change')?.({ time: 90 });
  handlers.get('pause-change')?.({ paused: true });

  assert.ok(calls.includes('set-sub:line'));
  assert.ok(calls.includes('reset-sidebar-layout'));
  assert.equal(calls.includes('broadcast-sub:line'), true);
  assert.ok(calls.includes('subtitle-change:line'));
  assert.ok(calls.includes('subtitle-track-change'));
  assert.ok(calls.includes('secondary-subtitle-track-change'));
  assert.ok(calls.includes('secondary-subtitle-delay-change:0.5'));
  assert.ok(calls.includes('subtitle-track-list-change'));
  assert.ok(calls.includes('media-title:Episode 1'));
  assert.ok(calls.includes('media-path:/tmp/video.mkv'));
  assert.ok(calls.includes('autoplay:/tmp/video.mkv'));
  assert.ok(calls.includes('reset-guess-state'));
  assert.ok(calls.includes('notify-title:Episode 1'));
  assert.ok(calls.includes('post-watch:901'));
  assert.ok(calls.includes('progress:normal'));
  assert.ok(calls.includes('progress:force'));
  assert.ok(calls.includes('presence-refresh'));
  assert.ok(calls.includes('sync-immersion'));
  assert.ok(calls.includes('flush-playback'));
});

test('main mpv event binder runs mpv-connected callback on connection', () => {
  const handlers = new Map<string, (payload: unknown) => void>();
  const calls: string[] = [];

  const bind = createBindMpvMainEventHandlersHandler({
    reportJellyfinRemoteStopped: () => calls.push('remote-stopped'),
    syncOverlayMpvSubtitleSuppression: () => calls.push('sync-overlay-mpv-sub'),
    onMpvConnected: () => calls.push('mpv-connected'),
    resetSubtitleSidebarEmbeddedLayout: () => calls.push('reset-sidebar-layout'),
    hasInitialPlaybackQuitOnDisconnectArg: () => false,
    isOverlayRuntimeInitialized: () => false,
    shouldQuitOnDisconnectWhenOverlayRuntimeInitialized: () => false,
    isQuitOnDisconnectArmed: () => false,
    scheduleQuitCheck: () => {},
    isMpvConnected: () => true,
    quitApp: () => {},

    recordImmersionSubtitleLine: () => {},
    hasSubtitleTimingTracker: () => false,
    recordSubtitleTiming: () => {},
    maybeRunAnilistPostWatchUpdate: async () => {},
    logSubtitleTimingError: () => {},
    setCurrentSubText: () => {},
    broadcastSubtitle: () => {},
    onSubtitleChange: () => {},
    refreshDiscordPresence: () => calls.push('presence-refresh'),

    setCurrentSubAssText: () => {},
    broadcastSubtitleAss: () => {},
    broadcastSecondarySubtitle: () => {},

    updateCurrentMediaPath: () => {},
    restoreMpvSubVisibility: () => {},
    getCurrentAnilistMediaKey: () => null,
    resetAnilistMediaTracking: () => {},
    maybeProbeAnilistDuration: () => {},
    ensureAnilistMediaGuess: () => {},
    syncImmersionMediaState: () => {},

    updateCurrentMediaTitle: () => {},
    resetAnilistMediaGuessState: () => {},
    notifyImmersionTitleUpdate: () => {},

    recordPlaybackPosition: () => {},
    recordMediaDuration: () => {},
    reportJellyfinRemoteProgress: () => {},
    recordPauseState: () => {},

    updateSubtitleRenderMetrics: () => {},
    setPreviousSecondarySubVisibility: () => {},
  });

  bind({
    on: (event, handler) => {
      handlers.set(event, handler as (payload: unknown) => void);
    },
  });

  handlers.get('connection-change')?.({ connected: true });

  assert.ok(calls.includes('mpv-connected'));
});

test('main mpv event binder clears media path on disconnect', () => {
  const handlers = new Map<string, (payload: unknown) => void>();
  const calls: string[] = [];

  const bind = createBindMpvMainEventHandlersHandler({
    reportJellyfinRemoteStopped: () => calls.push('remote-stopped'),
    syncOverlayMpvSubtitleSuppression: () => calls.push('sync-overlay-mpv-sub'),
    resetSubtitleSidebarEmbeddedLayout: () => calls.push('reset-sidebar-layout'),
    hasInitialPlaybackQuitOnDisconnectArg: () => false,
    isOverlayRuntimeInitialized: () => true,
    shouldQuitOnDisconnectWhenOverlayRuntimeInitialized: () => false,
    isQuitOnDisconnectArmed: () => false,
    scheduleQuitCheck: () => {},
    isMpvConnected: () => false,
    quitApp: () => {},

    recordImmersionSubtitleLine: () => {},
    hasSubtitleTimingTracker: () => false,
    recordSubtitleTiming: () => {},
    maybeRunAnilistPostWatchUpdate: async () => {},
    logSubtitleTimingError: () => {},
    setCurrentSubText: () => {},
    broadcastSubtitle: () => {},
    onSubtitleChange: () => {},
    refreshDiscordPresence: () => calls.push('presence-refresh'),

    setCurrentSubAssText: () => {},
    broadcastSubtitleAss: () => {},
    broadcastSecondarySubtitle: () => {},

    updateCurrentMediaPath: (path) => calls.push(`media-path:${path}`),
    restoreMpvSubVisibility: () => {},
    getCurrentAnilistMediaKey: () => null,
    resetAnilistMediaTracking: () => {},
    maybeProbeAnilistDuration: () => {},
    ensureAnilistMediaGuess: () => {},
    syncImmersionMediaState: () => {},

    updateCurrentMediaTitle: () => {},
    resetAnilistMediaGuessState: () => {},
    notifyImmersionTitleUpdate: () => {},

    recordPlaybackPosition: () => {},
    recordMediaDuration: () => {},
    reportJellyfinRemoteProgress: () => {},
    recordPauseState: () => {},

    updateSubtitleRenderMetrics: () => {},
    setPreviousSecondarySubVisibility: () => {},
  });

  bind({
    on: (event, handler) => {
      handlers.set(event, handler as (payload: unknown) => void);
    },
  });

  handlers.get('connection-change')?.({ connected: false });

  assert.ok(calls.includes('media-path:'));
  assert.ok(calls.includes('remote-stopped'));
  assert.ok(calls.includes('presence-refresh'));
});
