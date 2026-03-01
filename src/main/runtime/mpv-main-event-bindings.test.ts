import assert from 'node:assert/strict';
import test from 'node:test';
import { createBindMpvMainEventHandlersHandler } from './mpv-main-event-bindings';

test('main mpv event binder wires callbacks through to runtime deps', () => {
  const handlers = new Map<string, (payload: unknown) => void>();
  const calls: string[] = [];

  const bind = createBindMpvMainEventHandlersHandler({
    reportJellyfinRemoteStopped: () => calls.push('remote-stopped'),
    syncOverlayMpvSubtitleSuppression: () => calls.push('sync-overlay-mpv-sub'),
    hasInitialJellyfinPlayArg: () => false,
    isOverlayRuntimeInitialized: () => false,
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

    updateCurrentMediaPath: (path) => calls.push(`media-path:${path}`),
    restoreMpvSubVisibility: () => calls.push('restore-mpv-sub'),
    getCurrentAnilistMediaKey: () => 'media-key',
    resetAnilistMediaTracking: (key) => calls.push(`reset-media:${String(key)}`),
    maybeProbeAnilistDuration: (mediaKey) => calls.push(`probe:${mediaKey}`),
    ensureAnilistMediaGuess: (mediaKey) => calls.push(`guess:${mediaKey}`),
    syncImmersionMediaState: () => calls.push('sync-immersion'),

    updateCurrentMediaTitle: (title) => calls.push(`media-title:${title}`),
    resetAnilistMediaGuessState: () => calls.push('reset-guess-state'),
    notifyImmersionTitleUpdate: (title) => calls.push(`notify-title:${title}`),

    recordPlaybackPosition: (time) => calls.push(`time-pos:${time}`),
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

  handlers.get('subtitle-change')?.({ text: 'line' });
  handlers.get('media-path-change')?.({ path: '' });
  handlers.get('media-title-change')?.({ title: 'Episode 1' });
  handlers.get('time-pos-change')?.({ time: 2.5 });
  handlers.get('pause-change')?.({ paused: true });

  assert.ok(calls.includes('set-sub:line'));
  assert.ok(calls.includes('broadcast-sub:line'));
  assert.ok(calls.includes('subtitle-change:line'));
  assert.ok(calls.includes('media-title:Episode 1'));
  assert.ok(calls.includes('restore-mpv-sub'));
  assert.ok(calls.includes('reset-guess-state'));
  assert.ok(calls.includes('notify-title:Episode 1'));
  assert.ok(calls.includes('progress:normal'));
  assert.ok(calls.includes('progress:force'));
  assert.ok(calls.includes('presence-refresh'));
});
