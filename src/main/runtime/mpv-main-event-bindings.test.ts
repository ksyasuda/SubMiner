import assert from 'node:assert/strict';
import test from 'node:test';
import { createBindMpvMainEventHandlersHandler } from './mpv-main-event-bindings';

test('main mpv event binder wires callbacks through to runtime deps', () => {
  const handlers = new Map<string, (payload: unknown) => void>();
  const calls: string[] = [];

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
    maybeRunAnilistPostWatchUpdate: async () => {
      calls.push('post-watch');
    },
    logSubtitleTimingError: () => calls.push('subtitle-error'),
    setCurrentSubText: (text) => calls.push(`set-sub:${text}`),
    broadcastSubtitle: (payload) => calls.push(`broadcast-sub:${payload.text}`),
    onSubtitleChange: (text) => calls.push(`subtitle-change:${text}`),
    refreshDiscordPresence: () => calls.push('presence-refresh'),

    setCurrentSubAssText: (text) => calls.push(`set-ass:${text}`),
    broadcastSubtitleAss: (text) => calls.push(`broadcast-ass:${text}`),
    broadcastSecondarySubtitle: (text) => calls.push(`broadcast-secondary:${text}`),
    onSubtitleTrackChange: () => calls.push('subtitle-track-change'),
    onSubtitleTrackListChange: () => calls.push('subtitle-track-list-change'),

    updateCurrentMediaPath: (path) => calls.push(`media-path:${path}`),
    restoreMpvSubVisibility: () => calls.push('restore-mpv-sub'),
    getCurrentAnilistMediaKey: () => 'media-key',
    resetAnilistMediaTracking: (key) => calls.push(`reset-media:${String(key)}`),
    maybeProbeAnilistDuration: (mediaKey) => calls.push(`probe:${mediaKey}`),
    ensureAnilistMediaGuess: (mediaKey) => calls.push(`guess:${mediaKey}`),
    syncImmersionMediaState: () => calls.push('sync-immersion'),
    flushPlaybackPositionOnMediaPathClear: () => calls.push('flush-playback'),

    updateCurrentMediaTitle: (title) => calls.push(`media-title:${title}`),
    resetAnilistMediaGuessState: () => calls.push('reset-guess-state'),
    notifyImmersionTitleUpdate: (title) => calls.push(`notify-title:${title}`),

    recordPlaybackPosition: (time) => calls.push(`time-pos:${time}`),
    recordMediaDuration: (duration) => calls.push(`duration:${duration}`),
    reportJellyfinRemoteProgress: (forceImmediate) =>
      calls.push(`progress:${forceImmediate ? 'force' : 'normal'}`),
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
  handlers.get('subtitle-track-list-change')?.({ trackList: [] });
  handlers.get('media-path-change')?.({ path: '' });
  handlers.get('media-title-change')?.({ title: 'Episode 1' });
  handlers.get('time-pos-change')?.({ time: 2.5 });
  handlers.get('pause-change')?.({ paused: true });

  assert.ok(calls.includes('set-sub:line'));
  assert.ok(calls.includes('reset-sidebar-layout'));
  assert.ok(calls.includes('broadcast-sub:line'));
  assert.ok(calls.includes('subtitle-change:line'));
  assert.ok(calls.includes('subtitle-track-change'));
  assert.ok(calls.includes('subtitle-track-list-change'));
  assert.ok(calls.includes('media-title:Episode 1'));
  assert.ok(calls.includes('restore-mpv-sub'));
  assert.ok(calls.includes('reset-guess-state'));
  assert.ok(calls.includes('notify-title:Episode 1'));
  assert.ok(calls.includes('progress:normal'));
  assert.ok(calls.includes('progress:force'));
  assert.ok(calls.includes('presence-refresh'));
  assert.ok(calls.includes('sync-immersion'));
  assert.ok(calls.includes('flush-playback'));
});
