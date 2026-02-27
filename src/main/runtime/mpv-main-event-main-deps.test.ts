import assert from 'node:assert/strict';
import test from 'node:test';
import { createBuildBindMpvMainEventHandlersMainDepsHandler } from './mpv-main-event-main-deps';

test('mpv main event main deps map app state updates and delegate callbacks', async () => {
  const calls: string[] = [];
  const appState = {
    initialArgs: { jellyfinPlay: true },
    overlayRuntimeInitialized: true,
    mpvClient: { connected: true },
    immersionTracker: {
      recordSubtitleLine: (text: string) => calls.push(`immersion-sub:${text}`),
      handleMediaTitleUpdate: (title: string) => calls.push(`immersion-title:${title}`),
      recordPlaybackPosition: (time: number) => calls.push(`immersion-time:${time}`),
      recordPauseState: (paused: boolean) => calls.push(`immersion-pause:${paused}`),
    },
    subtitleTimingTracker: {
      recordSubtitle: (text: string) => calls.push(`timing:${text}`),
    },
    currentSubText: '',
    currentSubAssText: '',
    playbackPaused: null,
    previousSecondarySubVisibility: false,
  };

  const deps = createBuildBindMpvMainEventHandlersMainDepsHandler({
    appState,
    getQuitOnDisconnectArmed: () => true,
    scheduleQuitCheck: (callback) => {
      calls.push('schedule');
      callback();
    },
    quitApp: () => calls.push('quit'),
    reportJellyfinRemoteStopped: () => calls.push('remote-stopped'),
    syncOverlayMpvSubtitleSuppression: () => calls.push('sync-overlay-mpv-sub'),
    maybeRunAnilistPostWatchUpdate: async () => {
      calls.push('anilist-post-watch');
    },
    logSubtitleTimingError: (message) => calls.push(`subtitle-error:${message}`),
    broadcastToOverlayWindows: (channel, payload) =>
      calls.push(`broadcast:${channel}:${String(payload)}`),
    onSubtitleChange: (text) => calls.push(`subtitle-change:${text}`),
    updateCurrentMediaPath: (path) => calls.push(`path:${path}`),
    restoreMpvSubVisibilityForInvisibleOverlay: () => calls.push('restore-mpv-sub'),
    getCurrentAnilistMediaKey: () => 'media-key',
    resetAnilistMediaTracking: (mediaKey) => calls.push(`reset:${mediaKey}`),
    maybeProbeAnilistDuration: (mediaKey) => calls.push(`probe:${mediaKey}`),
    ensureAnilistMediaGuess: (mediaKey) => calls.push(`guess:${mediaKey}`),
    syncImmersionMediaState: () => calls.push('sync-immersion'),
    updateCurrentMediaTitle: (title) => calls.push(`title:${title}`),
    resetAnilistMediaGuessState: () => calls.push('reset-guess'),
    reportJellyfinRemoteProgress: (forceImmediate) => calls.push(`progress:${forceImmediate}`),
    updateSubtitleRenderMetrics: () => calls.push('metrics'),
    refreshDiscordPresence: () => calls.push('presence-refresh'),
  })();

  assert.equal(deps.hasInitialJellyfinPlayArg(), true);
  assert.equal(deps.isOverlayRuntimeInitialized(), true);
  assert.equal(deps.isQuitOnDisconnectArmed(), true);
  assert.equal(deps.isMpvConnected(), true);
  deps.scheduleQuitCheck(() => calls.push('scheduled-callback'));
  deps.quitApp();
  deps.reportJellyfinRemoteStopped();
  deps.syncOverlayMpvSubtitleSuppression();
  deps.recordImmersionSubtitleLine('x', 0, 1);
  assert.equal(deps.hasSubtitleTimingTracker(), true);
  deps.recordSubtitleTiming('y', 0, 1);
  await deps.maybeRunAnilistPostWatchUpdate();
  deps.logSubtitleTimingError('err', new Error('boom'));
  deps.setCurrentSubText('sub');
  deps.broadcastSubtitle({ text: 'sub', tokens: null });
  deps.onSubtitleChange('sub');
  deps.refreshDiscordPresence();
  deps.setCurrentSubAssText('ass');
  deps.broadcastSubtitleAss('ass');
  deps.broadcastSecondarySubtitle('sec');
  deps.updateCurrentMediaPath('/tmp/video');
  deps.restoreMpvSubVisibilityForInvisibleOverlay();
  assert.equal(deps.getCurrentAnilistMediaKey(), 'media-key');
  deps.resetAnilistMediaTracking('media-key');
  deps.maybeProbeAnilistDuration('media-key');
  deps.ensureAnilistMediaGuess('media-key');
  deps.syncImmersionMediaState();
  deps.updateCurrentMediaTitle('title');
  deps.resetAnilistMediaGuessState();
  deps.notifyImmersionTitleUpdate('title');
  deps.recordPlaybackPosition(10);
  deps.reportJellyfinRemoteProgress(true);
  deps.recordPauseState(true);
  deps.updateSubtitleRenderMetrics({});
  deps.setPreviousSecondarySubVisibility(true);

  assert.equal(appState.currentSubText, 'sub');
  assert.equal(appState.currentSubAssText, 'ass');
  assert.equal(appState.playbackPaused, true);
  assert.equal(appState.previousSecondarySubVisibility, true);
  assert.ok(calls.includes('remote-stopped'));
  assert.ok(calls.includes('sync-overlay-mpv-sub'));
  assert.ok(calls.includes('anilist-post-watch'));
  assert.ok(calls.includes('sync-immersion'));
  assert.ok(calls.includes('metrics'));
  assert.ok(calls.includes('presence-refresh'));
  assert.ok(calls.includes('restore-mpv-sub'));
});
