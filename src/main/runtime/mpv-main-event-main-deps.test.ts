import assert from 'node:assert/strict';
import test from 'node:test';
import { parseAssCues } from '../../core/services/subtitle-cue-parser';
import { createBuildBindMpvMainEventHandlersMainDepsHandler } from './mpv-main-event-main-deps';

test('mpv main event main deps map app state updates and delegate callbacks', async () => {
  const calls: string[] = [];
  const appState = {
    initialArgs: { jellyfinPlay: true },
    overlayRuntimeInitialized: true,
    mpvClient: {
      connected: true,
      currentSecondarySubText: 'secondary',
      currentTimePos: 12.25,
      requestProperty: async () => 18.75,
    },
    immersionTracker: {
      recordSubtitleLine: (text: string) => calls.push(`immersion-sub:${text}`),
      handleMediaTitleUpdate: (title: string) => calls.push(`immersion-title:${title}`),
      recordPlaybackPosition: (time: number) => calls.push(`immersion-time:${time}`),
      recordMediaDuration: (durationSec: number) => calls.push(`immersion-duration:${durationSec}`),
      recordPauseState: (paused: boolean) => calls.push(`immersion-pause:${paused}`),
    },
    subtitleTimingTracker: {
      recordSubtitle: (text: string, _start: number, _end: number, secondaryText?: string) =>
        calls.push(`timing:${text}:${secondaryText ?? ''}`),
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
    recordAnilistMediaDuration: (durationSec) => calls.push(`anilist-duration:${durationSec}`),
    logSubtitleTimingError: (message) => calls.push(`subtitle-error:${message}`),
    broadcastToOverlayWindows: (channel, payload) =>
      calls.push(`broadcast:${channel}:${String(payload)}`),
    onSecondarySubtitleChange: (text) => calls.push(`secondary:${text}`),
    onSecondarySubtitleTrackChange: (sid) => calls.push(`secondary-track:${String(sid)}`),
    onSecondarySubtitleDelayChange: (delay) => calls.push(`secondary-delay:${delay}`),
    onSubtitleChange: (text) => calls.push(`subtitle-change:${text}`),
    ensureImmersionTrackerInitialized: () => calls.push('ensure-immersion'),
    updateCurrentMediaPath: (path) => calls.push(`path:${path}`),
    restoreMpvSubVisibility: () => calls.push('restore-mpv-sub'),
    resetSubtitleSidebarEmbeddedLayout: () => calls.push('reset-sidebar-layout'),
    getCurrentAnilistMediaKey: () => 'media-key',
    resetAnilistMediaTracking: (mediaKey) => calls.push(`reset:${mediaKey}`),
    maybeProbeAnilistDuration: (mediaKey) => calls.push(`probe:${mediaKey}`),
    ensureAnilistMediaGuess: (mediaKey) => calls.push(`guess:${mediaKey}`),
    syncImmersionMediaState: () => calls.push('sync-immersion'),
    signalAutoplayReadyIfWarm: (path) => calls.push(`autoplay:${path}`),
    updateCurrentMediaTitle: (title) => calls.push(`title:${title}`),
    resetAnilistMediaGuessState: () => calls.push('reset-guess'),
    reportJellyfinRemoteProgress: (forceImmediate) => calls.push(`progress:${forceImmediate}`),
    onFullscreenChange: (fullscreen) => calls.push(`fullscreen:${fullscreen}`),
    updateSubtitleRenderMetrics: () => calls.push('metrics'),
    refreshDiscordPresence: () => calls.push('presence-refresh'),
  })();

  assert.equal(deps.hasInitialPlaybackQuitOnDisconnectArg(), true);
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
  deps.onSecondarySubtitleTrackChange?.(4);
  deps.onSecondarySubtitleDelayChange?.(0.5);
  deps.updateCurrentMediaPath('/tmp/video');
  deps.restoreMpvSubVisibility();
  deps.resetSubtitleSidebarEmbeddedLayout();
  assert.equal(deps.getCurrentAnilistMediaKey(), 'media-key');
  deps.resetAnilistMediaTracking('media-key');
  deps.maybeProbeAnilistDuration('media-key');
  deps.ensureAnilistMediaGuess('media-key');
  deps.syncImmersionMediaState();
  deps.signalAutoplayReadyIfWarm?.('/tmp/video');
  deps.updateCurrentMediaTitle('title');
  deps.resetAnilistMediaGuessState();
  deps.notifyImmersionTitleUpdate('title');
  deps.recordPlaybackPosition(10);
  deps.recordMediaDuration(1234);
  deps.reportJellyfinRemoteProgress(true);
  deps.onFullscreenChange?.(true);
  deps.recordPauseState(true);
  deps.updateSubtitleRenderMetrics({});
  deps.setPreviousSecondarySubVisibility(true);
  deps.flushPlaybackPositionOnMediaPathClear?.('');
  await Promise.resolve();

  assert.equal(appState.currentSubText, 'sub');
  assert.equal(appState.currentSubAssText, 'ass');
  assert.equal(appState.playbackPaused, true);
  assert.equal(appState.previousSecondarySubVisibility, true);
  assert.ok(calls.includes('remote-stopped'));
  assert.ok(calls.includes('sync-overlay-mpv-sub'));
  assert.ok(calls.includes('anilist-post-watch'));
  assert.ok(calls.includes('timing:y:secondary'));
  assert.ok(calls.includes('secondary:sec'));
  assert.ok(calls.includes('secondary-track:4'));
  assert.ok(calls.includes('secondary-delay:0.5'));
  assert.ok(!calls.includes('broadcast:secondary-subtitle:set:sec'));
  assert.ok(calls.includes('ensure-immersion'));
  assert.ok(calls.includes('sync-immersion'));
  assert.ok(calls.includes('autoplay:/tmp/video'));
  assert.ok(calls.includes('metrics'));
  assert.ok(calls.includes('fullscreen:true'));
  assert.ok(calls.includes('presence-refresh'));
  assert.ok(calls.includes('restore-mpv-sub'));
  assert.ok(calls.includes('reset-sidebar-layout'));
  assert.ok(calls.includes('immersion-duration:1234'));
  assert.ok(calls.includes('anilist-duration:1234'));
});

test('mpv main event main deps wire subtitle callbacks without suppression gate', () => {
  const deps = createBuildBindMpvMainEventHandlersMainDepsHandler({
    appState: {
      initialArgs: null,
      overlayRuntimeInitialized: true,
      mpvClient: null,
      immersionTracker: null,
      subtitleTimingTracker: null,
      currentSubText: '',
      currentSubAssText: '',
      playbackPaused: null,
      previousSecondarySubVisibility: false,
    },
    getQuitOnDisconnectArmed: () => false,
    scheduleQuitCheck: () => {},
    quitApp: () => {},
    reportJellyfinRemoteStopped: () => {},
    syncOverlayMpvSubtitleSuppression: () => {},
    maybeRunAnilistPostWatchUpdate: async () => {},
    logSubtitleTimingError: () => {},
    broadcastToOverlayWindows: () => {},
    onSubtitleChange: () => {},
    ensureImmersionTrackerInitialized: () => {},
    updateCurrentMediaPath: () => {},
    restoreMpvSubVisibility: () => {},
    resetSubtitleSidebarEmbeddedLayout: () => {},
    getCurrentAnilistMediaKey: () => null,
    resetAnilistMediaTracking: () => {},
    maybeProbeAnilistDuration: () => {},
    ensureAnilistMediaGuess: () => {},
    syncImmersionMediaState: () => {},
    updateCurrentMediaTitle: () => {},
    resetAnilistMediaGuessState: () => {},
    reportJellyfinRemoteProgress: () => {},
    updateSubtitleRenderMetrics: () => {},
    refreshDiscordPresence: () => {},
  })();

  deps.setCurrentSubText('sub');
  assert.equal(typeof deps.setCurrentSubText, 'function');
});

test('mpv main event main deps treat managed playback as quit-on-disconnect', () => {
  const deps = createBuildBindMpvMainEventHandlersMainDepsHandler({
    appState: {
      initialArgs: { managedPlayback: true },
      overlayRuntimeInitialized: false,
      mpvClient: null,
      immersionTracker: null,
      subtitleTimingTracker: null,
      currentSubText: '',
      currentSubAssText: '',
      playbackPaused: null,
      previousSecondarySubVisibility: false,
    },
    getQuitOnDisconnectArmed: () => true,
    scheduleQuitCheck: () => {},
    quitApp: () => {},
    reportJellyfinRemoteStopped: () => {},
    syncOverlayMpvSubtitleSuppression: () => {},
    maybeRunAnilistPostWatchUpdate: async () => {},
    logSubtitleTimingError: () => {},
    broadcastToOverlayWindows: () => {},
    onSubtitleChange: () => {},
    ensureImmersionTrackerInitialized: () => {},
    updateCurrentMediaPath: () => {},
    restoreMpvSubVisibility: () => {},
    resetSubtitleSidebarEmbeddedLayout: () => {},
    getCurrentAnilistMediaKey: () => null,
    resetAnilistMediaTracking: () => {},
    maybeProbeAnilistDuration: () => {},
    ensureAnilistMediaGuess: () => {},
    syncImmersionMediaState: () => {},
    updateCurrentMediaTitle: () => {},
    resetAnilistMediaGuessState: () => {},
    reportJellyfinRemoteProgress: () => {},
    updateSubtitleRenderMetrics: () => {},
    refreshDiscordPresence: () => {},
  })();

  assert.equal(deps.hasInitialPlaybackQuitOnDisconnectArg(), true);
  assert.equal(deps.shouldQuitOnDisconnectWhenOverlayRuntimeInitialized(), true);
});

test('flushPlaybackPositionOnMediaPathClear ignores disconnected mpv time-pos reads', async () => {
  const recorded: number[] = [];
  const deps = createBuildBindMpvMainEventHandlersMainDepsHandler({
    appState: {
      initialArgs: null,
      overlayRuntimeInitialized: true,
      mpvClient: {
        connected: false,
        currentTimePos: 42,
        requestProperty: async () => {
          throw new Error('disconnected');
        },
      },
      immersionTracker: {
        recordPlaybackPosition: (time: number) => {
          recorded.push(time);
        },
      },
      subtitleTimingTracker: null,
      currentMediaPath: '',
      currentSubText: '',
      currentSubAssText: '',
      playbackPaused: null,
      previousSecondarySubVisibility: false,
    },
    getQuitOnDisconnectArmed: () => false,
    scheduleQuitCheck: () => {},
    quitApp: () => {},
    reportJellyfinRemoteStopped: () => {},
    syncOverlayMpvSubtitleSuppression: () => {},
    maybeRunAnilistPostWatchUpdate: async () => {},
    logSubtitleTimingError: () => {},
    broadcastToOverlayWindows: () => {},
    onSubtitleChange: () => {},
    ensureImmersionTrackerInitialized: () => {},
    updateCurrentMediaPath: () => {},
    restoreMpvSubVisibility: () => {},
    resetSubtitleSidebarEmbeddedLayout: () => {},
    getCurrentAnilistMediaKey: () => null,
    resetAnilistMediaTracking: () => {},
    maybeProbeAnilistDuration: () => {},
    ensureAnilistMediaGuess: () => {},
    syncImmersionMediaState: () => {},
    updateCurrentMediaTitle: () => {},
    resetAnilistMediaGuessState: () => {},
    reportJellyfinRemoteProgress: () => {},
    updateSubtitleRenderMetrics: () => {},
    refreshDiscordPresence: () => {},
  })();

  deps.flushPlaybackPositionOnMediaPathClear?.('');
  await Promise.resolve();

  assert.deepEqual(recorded, [42]);
});

test('media and subtitle-track transitions reset live subtitle-line deduplication', () => {
  const recordedStarts: number[] = [];
  const handlers = createBuildBindMpvMainEventHandlersMainDepsHandler({
    appState: {
      initialArgs: null,
      overlayRuntimeInitialized: true,
      mpvClient: null,
      immersionTracker: {
        recordSubtitleLine: (_text: string, start: number) => recordedStarts.push(start),
      },
      subtitleTimingTracker: null,
      activeParsedSubtitleCues: null,
      currentMediaPath: '/video-a.mkv',
      currentSubText: '',
      currentSubAssText: '',
      playbackPaused: null,
      previousSecondarySubVisibility: false,
    },
    getQuitOnDisconnectArmed: () => false,
    scheduleQuitCheck: () => {},
    quitApp: () => {},
    reportJellyfinRemoteStopped: () => {},
    syncOverlayMpvSubtitleSuppression: () => {},
    maybeRunAnilistPostWatchUpdate: async () => {},
    logSubtitleTimingError: () => {},
    broadcastToOverlayWindows: () => {},
    onSubtitleChange: () => {},
    ensureImmersionTrackerInitialized: () => {},
    updateCurrentMediaPath: () => {},
    restoreMpvSubVisibility: () => {},
    resetSubtitleSidebarEmbeddedLayout: () => {},
    getCurrentAnilistMediaKey: () => null,
    resetAnilistMediaTracking: () => {},
    maybeProbeAnilistDuration: () => {},
    ensureAnilistMediaGuess: () => {},
    syncImmersionMediaState: () => {},
    updateCurrentMediaTitle: () => {},
    resetAnilistMediaGuessState: () => {},
    reportJellyfinRemoteProgress: () => {},
    updateSubtitleRenderMetrics: () => {},
    refreshDiscordPresence: () => {},
  })();

  for (let index = 0; index < 8; index += 1) {
    handlers.recordImmersionSubtitleLine('待って', index * 0.04, (index + 1) * 0.04);
  }
  assert.equal(recordedStarts.length, 4);

  handlers.updateCurrentMediaPath('/video-b.mkv');
  handlers.recordImmersionSubtitleLine('待って', 0.32, 0.36);
  assert.equal(recordedStarts.length, 5);

  for (let index = 9; index < 16; index += 1) {
    handlers.recordImmersionSubtitleLine('待って', index * 0.04, (index + 1) * 0.04);
  }
  assert.equal(recordedStarts.length, 8);

  assert.equal(typeof handlers.onSubtitleTrackChange, 'function');
  handlers.onSubtitleTrackChange?.(2);
  handlers.recordImmersionSubtitleLine('待って', 0.64, 0.68);
  assert.equal(recordedStarts.length, 9);
});

test('subtitle-track transitions ignore stale parsed cues until replacement cues arrive', () => {
  const recordedStarts: number[] = [];
  const appState = {
    initialArgs: null,
    overlayRuntimeInitialized: true,
    mpvClient: null,
    immersionTracker: {
      recordSubtitleLine: (_text: string, start: number) => recordedStarts.push(start),
    },
    subtitleTimingTracker: null,
    activeParsedSubtitleCues: [{ startTime: 10, endTime: 14, text: '飛び上がる' }],
    currentMediaPath: '/video-a.mkv',
    currentSubText: '',
    currentSubAssText: '',
    playbackPaused: null,
    previousSecondarySubVisibility: false,
  };
  const handlers = createBuildBindMpvMainEventHandlersMainDepsHandler({
    appState,
    getQuitOnDisconnectArmed: () => false,
    scheduleQuitCheck: () => {},
    quitApp: () => {},
    reportJellyfinRemoteStopped: () => {},
    syncOverlayMpvSubtitleSuppression: () => {},
    maybeRunAnilistPostWatchUpdate: async () => {},
    logSubtitleTimingError: () => {},
    broadcastToOverlayWindows: () => {},
    onSubtitleChange: () => {},
    ensureImmersionTrackerInitialized: () => {},
    updateCurrentMediaPath: () => {},
    restoreMpvSubVisibility: () => {},
    resetSubtitleSidebarEmbeddedLayout: () => {},
    getCurrentAnilistMediaKey: () => null,
    resetAnilistMediaTracking: () => {},
    maybeProbeAnilistDuration: () => {},
    ensureAnilistMediaGuess: () => {},
    syncImmersionMediaState: () => {},
    updateCurrentMediaTitle: () => {},
    resetAnilistMediaGuessState: () => {},
    reportJellyfinRemoteProgress: () => {},
    updateSubtitleRenderMetrics: () => {},
    refreshDiscordPresence: () => {},
  })();

  handlers.recordImmersionSubtitleLine('飛び上がる', 10, 10.04);
  handlers.onSubtitleTrackChange?.(2);
  for (let index = 1; index <= 8; index += 1) {
    handlers.recordImmersionSubtitleLine('飛び上がる', 10 + index * 0.04, 10 + (index + 1) * 0.04);
  }
  assert.equal(recordedStarts.length, 5);

  appState.activeParsedSubtitleCues = [{ startTime: 20, endTime: 24, text: '飛び上がる' }];
  handlers.recordImmersionSubtitleLine('飛び上がる', 20, 20.04);
  handlers.recordImmersionSubtitleLine('飛び上がる', 20.04, 20.08);
  assert.deepEqual(recordedStarts.slice(-1), [20]);
});

test('canonical ASS cues replace live glyph spam for display, history, and immersion', () => {
  const immersion: Array<{ text: string; start: number; end: number }> = [];
  const timing: Array<{ text: string; start: number; end: number }> = [];
  const handlers = createBuildBindMpvMainEventHandlersMainDepsHandler({
    appState: {
      initialArgs: null,
      overlayRuntimeInitialized: true,
      mpvClient: { currentTimePos: 2 },
      immersionTracker: {
        recordSubtitleLine: (text: string, start: number, end: number) =>
          immersion.push({ text, start, end }),
      },
      subtitleTimingTracker: {
        recordSubtitle: (text: string, start: number, end: number) =>
          timing.push({ text, start, end }),
      },
      activeParsedSubtitleCues: [
        {
          startTime: 1.2,
          endTime: 3.8,
          text: '今　手にある物差しでは',
          source: 'canonical-ass',
        },
        {
          startTime: 3,
          endTime: 6,
          text: '飛び越えてみたくて',
          source: 'canonical-ass',
        },
        {
          startTime: 10,
          endTime: 12,
          text: 'MaidCafeMaidCafe',
          source: 'reconstructed-ass',
          assLayout: { kind: 'fragment-grid', sourceOrder: 2 },
        },
      ],
      currentMediaPath: '/video.mkv',
      currentSubText: '',
      currentSubAssText: '',
      playbackPaused: null,
      previousSecondarySubVisibility: false,
    },
    getQuitOnDisconnectArmed: () => false,
    scheduleQuitCheck: () => {},
    quitApp: () => {},
    reportJellyfinRemoteStopped: () => {},
    syncOverlayMpvSubtitleSuppression: () => {},
    maybeRunAnilistPostWatchUpdate: async () => {},
    logSubtitleTimingError: () => {},
    broadcastToOverlayWindows: () => {},
    onSubtitleChange: () => {},
    ensureImmersionTrackerInitialized: () => {},
    updateCurrentMediaPath: () => {},
    restoreMpvSubVisibility: () => {},
    getCurrentAnilistMediaKey: () => null,
    resetAnilistMediaTracking: () => {},
    maybeProbeAnilistDuration: () => {},
    ensureAnilistMediaGuess: () => {},
    syncImmersionMediaState: () => {},
    updateCurrentMediaTitle: () => {},
    resetAnilistMediaGuessState: () => {},
    reportJellyfinRemoteProgress: () => {},
    updateSubtitleRenderMetrics: () => {},
    refreshDiscordPresence: () => {},
  })();

  assert.equal(handlers.resolveSubtitleText?.('今\n今\n今\n手\n手\n手'), '今　手にある物差しでは');
  handlers.recordImmersionSubtitleLine('今', 0.8, 1.5);
  handlers.recordImmersionSubtitleLine('手', 0.86, 1.56);
  handlers.recordSubtitleTiming('今', 0.8, 1.5);

  assert.deepEqual(immersion, [{ text: '今　手にある物差しでは', start: 1.2, end: 3.8 }]);
  assert.deepEqual(timing, [{ text: '今　手にある物差しでは', start: 1.2, end: 3.8 }]);

  // Concurrent dialogue during the song is not part of the animation: it must be
  // recorded as itself -- without the fragment lines beside it -- and must not cause
  // the song line to be recorded again when the animation frames resume.
  assert.equal(handlers.resolveSubtitleText?.('普通のセリフ\n今\n手'), '普通のセリフ\n今\n手');
  handlers.recordImmersionSubtitleLine('普通のセリフ\n今\n手', 1.9, 3.2);
  handlers.recordImmersionSubtitleLine('にある', 2.1, 2.9);
  handlers.recordSubtitleTiming('次のセリフ', 3.9, 5.0);

  assert.deepEqual(immersion.slice(1), [{ text: '普通のセリフ', start: 1.9, end: 3.2 }]);
  assert.deepEqual(timing.slice(1), [{ text: '次のセリフ', start: 3.9, end: 5 }]);

  // Overlapping canonical lines resolve as shifting subsets (A, then A+B, then A).
  // Every recorded cue is remembered, so each authored line still records exactly once.
  handlers.recordImmersionSubtitleLine('飛び越えて', 3.2, 3.4);
  handlers.recordImmersionSubtitleLine('手にある', 3.5, 3.7);
  handlers.recordSubtitleTiming('飛び越えて', 3.2, 3.4);
  handlers.recordSubtitleTiming('手にある', 3.5, 3.7);

  assert.deepEqual(immersion.slice(2), [{ text: '飛び越えてみたくて', start: 3, end: 6 }]);
  assert.deepEqual(timing.slice(2), [{ text: '飛び越えてみたくて', start: 3, end: 6 }]);

  // A backward seek means the user is rewatching: the timing history (a viewing log)
  // records the revisited line again, while immersion stays once-per-media.
  handlers.onTimePosUpdate?.(30);
  handlers.onTimePosUpdate?.(2);
  handlers.recordSubtitleTiming('今', 0.8, 1.5);
  handlers.recordImmersionSubtitleLine('今', 0.8, 1.5);

  assert.deepEqual(timing.slice(3), [{ text: '今　手にある物差しでは', start: 1.2, end: 3.8 }]);
  assert.equal(immersion.length, 3);

  // A jump of exactly the seek threshold counts as a seek, matching the time-pos
  // handler's own `>=` boundary.
  handlers.onTimePosUpdate?.(4.5);
  handlers.onTimePosUpdate?.(2);
  handlers.recordSubtitleTiming('今', 0.8, 1.5);

  assert.deepEqual(timing.slice(4), [{ text: '今　手にある物差しでは', start: 1.2, end: 3.8 }]);

  handlers.recordImmersionSubtitleLine('Maid\nCafe', 10, 12);
  handlers.recordSubtitleTiming('Maid\nCafe', 10, 12);
  assert.equal(immersion.length, 3);
  assert.equal(timing.length, 5);
});

test('subtitle-track changes stop stale canonical cues from substituting immediately', () => {
  const appState = {
    initialArgs: null,
    overlayRuntimeInitialized: true,
    mpvClient: { currentTimePos: 2 },
    immersionTracker: { recordSubtitleLine: () => {} },
    subtitleTimingTracker: { recordSubtitle: () => {} },
    activeParsedSubtitleCues: [
      {
        startTime: 1.2,
        endTime: 3.8,
        text: '今　手にある物差しでは',
        source: 'canonical-ass' as const,
      },
    ] as Array<{ startTime: number; endTime: number; text: string; source?: 'canonical-ass' }>,
    activeParsedSubtitleSource: 'track-a.ass' as string | null,
    currentMediaPath: '/video.mkv',
    currentSubText: '',
    currentSubAssText: '',
    playbackPaused: null,
    previousSecondarySubVisibility: false,
  };
  const handlers = createBuildBindMpvMainEventHandlersMainDepsHandler({
    appState,
    getQuitOnDisconnectArmed: () => false,
    scheduleQuitCheck: () => {},
    quitApp: () => {},
    reportJellyfinRemoteStopped: () => {},
    syncOverlayMpvSubtitleSuppression: () => {},
    maybeRunAnilistPostWatchUpdate: async () => {},
    logSubtitleTimingError: () => {},
    broadcastToOverlayWindows: () => {},
    onSubtitleChange: () => {},
    ensureImmersionTrackerInitialized: () => {},
    updateCurrentMediaPath: () => {},
    restoreMpvSubVisibility: () => {},
    getCurrentAnilistMediaKey: () => null,
    resetAnilistMediaTracking: () => {},
    maybeProbeAnilistDuration: () => {},
    ensureAnilistMediaGuess: () => {},
    syncImmersionMediaState: () => {},
    updateCurrentMediaTitle: () => {},
    resetAnilistMediaGuessState: () => {},
    reportJellyfinRemoteProgress: () => {},
    updateSubtitleRenderMetrics: () => {},
    refreshDiscordPresence: () => {},
  })();

  assert.equal(handlers.resolveSubtitleText?.('今\n手にある'), '今　手にある物差しでは');

  // The new track's cues arrive only after an async re-parse; until then, the old
  // track's canonical lyric must not replace the new track's live text.
  handlers.onSubtitleTrackChange?.(2);

  assert.deepEqual(appState.activeParsedSubtitleCues, []);
  assert.equal(appState.activeParsedSubtitleSource, null);
  assert.equal(handlers.resolveSubtitleText?.('今\n手にある'), '今\n手にある');
});

test('subtitle recorders drop ASS furigana events the same way the display does', () => {
  // Broadcast-caption ASS (Caption2Ass style): furigana are separate half-scale events
  // positioned above their base line, and mpv lists them as their own live lines.
  const cues = parseAssCues(
    [
      '[Script Info]',
      'PlayResX: 960',
      'PlayResY: 540',
      '',
      '[V4+ Styles]',
      'Format: Name, Fontname, Fontsize, PrimaryColour, SecondaryColour, OutlineColour, BackColour, Bold, Italic, Underline, StrikeOut, ScaleX, ScaleY, Spacing, Angle, BorderStyle, Outline, Shadow, Alignment, MarginL, MarginR, MarginV, Encoding',
      'Style: Default,Yu Gothic,46,&H00FFFFFF,&H000000FF,&H00000000,&H7F000000,1,0,0,0,100,100,4,0,1,2,2,1,0,0,0,1',
      '',
      '[Events]',
      'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
      'Dialogue: 0,0:04:56.26,0:04:59.63,Default,,0000,0000,0000,,{\\pos(472,443)\\fscx50\\fscy50}あくむ',
      'Dialogue: 0,0:04:56.26,0:04:59.63,Default,,0000,0000,0000,,{\\pos(172,497)}こんな短時間で{\\fscx50}　{\\fscx100}悪夢{\\fscx50}　{\\fscx100}見んなよ{\\fscx50}。',
      'Dialogue: 0,0:04:59.63,0:05:03.54,Default,,0000,0000,0000,,{\\pos(172,407)\\fscx50}（{\\fscx100}平{\\fscx50}）{\\fscx100}暗記教科は{\\fscx50}　{\\fscx100}もう',
      'Dialogue: 0,0:04:59.63,0:05:03.54,Default,,0000,0000,0000,,{\\pos(332,443)\\fscx50\\fscy50}かた',
      'Dialogue: 0,0:04:59.63,0:05:03.54,Default,,0000,0000,0000,,{\\pos(412,443)\\fscx50\\fscy50}ぱし',
      'Dialogue: 0,0:04:59.63,0:05:03.54,Default,,0000,0000,0000,,{\\pos(172,497)}とにかく片っ端から覚えるんだよ{\\fscx50}。',
    ].join('\n'),
  );
  assert.deepEqual(
    cues.map((cue) => cue.text),
    [
      'こんな短時間で　悪夢　見んなよ。',
      '（平）暗記教科は　もう',
      'とにかく片っ端から覚えるんだよ。',
    ],
  );

  const immersion: string[] = [];
  const timing: string[] = [];
  const handlers = createBuildBindMpvMainEventHandlersMainDepsHandler({
    appState: {
      initialArgs: null,
      overlayRuntimeInitialized: true,
      mpvClient: { currentTimePos: 299.7 },
      immersionTracker: { recordSubtitleLine: (text: string) => immersion.push(text) },
      subtitleTimingTracker: { recordSubtitle: (text: string) => timing.push(text) },
      activeParsedSubtitleCues: cues,
      currentSubText: '',
      currentSubAssText: '',
      playbackPaused: null,
      previousSecondarySubVisibility: false,
    },
    getQuitOnDisconnectArmed: () => false,
    scheduleQuitCheck: () => {},
    quitApp: () => {},
    reportJellyfinRemoteStopped: () => {},
    syncOverlayMpvSubtitleSuppression: () => {},
    maybeRunAnilistPostWatchUpdate: async () => {},
    logSubtitleTimingError: () => {},
    broadcastToOverlayWindows: () => {},
    onSubtitleChange: () => {},
    ensureImmersionTrackerInitialized: () => {},
    updateCurrentMediaPath: () => {},
    restoreMpvSubVisibility: () => {},
    getCurrentAnilistMediaKey: () => null,
    resetAnilistMediaTracking: () => {},
    maybeProbeAnilistDuration: () => {},
    ensureAnilistMediaGuess: () => {},
    syncImmersionMediaState: () => {},
    updateCurrentMediaTitle: () => {},
    resetAnilistMediaGuessState: () => {},
    reportJellyfinRemoteProgress: () => {},
    updateSubtitleRenderMetrics: () => {},
    refreshDiscordPresence: () => {},
  })();

  const liveText = '（平）暗記教科は　もう\nかた\nぱし\nとにかく片っ端から覚えるんだよ。';
  const expected = '（平）暗記教科は　もう\n\nとにかく片っ端から覚えるんだよ。';
  assert.equal(handlers.resolveSubtitleText?.(liveText), expected);
  handlers.recordImmersionSubtitleLine(liveText, 299.63, 303.54);
  handlers.recordSubtitleTiming(liveText, 299.63, 303.54);

  assert.deepEqual(immersion, [expected]);
  assert.deepEqual(timing, [expected]);
});

test('a resolved line survives recording while a fragment grid is on screen', () => {
  // Fragment stripping drops every line it can trace back to a cue, and returns nothing
  // at all when a fragment grid is nearby. Text the parsed cues already resolved is a
  // complete line, not raw mpv output, so it must not be fed through that path.
  const immersion: string[] = [];
  const timing: string[] = [];
  const handlers = createBuildBindMpvMainEventHandlersMainDepsHandler({
    appState: {
      initialArgs: null,
      overlayRuntimeInitialized: true,
      mpvClient: { currentTimePos: 3.2 },
      immersionTracker: { recordSubtitleLine: (text: string) => immersion.push(text) },
      subtitleTimingTracker: { recordSubtitle: (text: string) => timing.push(text) },
      activeParsedSubtitleCues: [
        { startTime: 3, endTime: 6, text: '飛び越えてみたくて', source: 'canonical-ass' },
        {
          startTime: 3,
          endTime: 6,
          text: 'MaidCafeMaidCafe',
          source: 'reconstructed-ass',
          assLayout: { kind: 'fragment-grid', sourceOrder: 2 },
        },
      ],
      currentSubText: '',
      currentSubAssText: '',
      playbackPaused: null,
      previousSecondarySubVisibility: false,
    },
    getQuitOnDisconnectArmed: () => false,
    scheduleQuitCheck: () => {},
    quitApp: () => {},
    reportJellyfinRemoteStopped: () => {},
    syncOverlayMpvSubtitleSuppression: () => {},
    maybeRunAnilistPostWatchUpdate: async () => {},
    logSubtitleTimingError: () => {},
    broadcastToOverlayWindows: () => {},
    onSubtitleChange: () => {},
    ensureImmersionTrackerInitialized: () => {},
    updateCurrentMediaPath: () => {},
    restoreMpvSubVisibility: () => {},
    getCurrentAnilistMediaKey: () => null,
    resetAnilistMediaTracking: () => {},
    maybeProbeAnilistDuration: () => {},
    ensureAnilistMediaGuess: () => {},
    syncImmersionMediaState: () => {},
    updateCurrentMediaTitle: () => {},
    resetAnilistMediaGuessState: () => {},
    reportJellyfinRemoteProgress: () => {},
    updateSubtitleRenderMetrics: () => {},
    refreshDiscordPresence: () => {},
  })();

  // The grid fragment beside the lyric keeps canonical substitution from applying, so
  // recording falls to the parsed view -- which is where the whole line is recovered.
  const liveText = '飛び越え\nMaid';
  assert.equal(handlers.resolveSubtitleText?.(liveText), '飛び越えてみたくて');
  handlers.recordImmersionSubtitleLine(liveText, 3, 6);
  handlers.recordSubtitleTiming(liveText, 3, 6);

  assert.deepEqual(immersion, ['飛び越えてみたくて']);
  assert.deepEqual(timing, ['飛び越えてみたくて']);

  // A spacer event left as literal control debris is not a subtitle line. The display
  // drops it, so no recorder may keep it either.
  assert.equal(handlers.resolveSubtitleText?.('\\'), '');
  handlers.recordImmersionSubtitleLine('\\', 3, 6);
  handlers.recordSubtitleTiming('\\', 3, 6);

  assert.equal(immersion.length, 1);
  assert.equal(timing.length, 1);
});
