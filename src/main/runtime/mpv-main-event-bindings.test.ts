import assert from 'node:assert/strict';
import test from 'node:test';
import { parseSubtitleCues } from '../../core/services/subtitle-cue-parser';
import { createBindMpvMainEventHandlersHandler } from './mpv-main-event-bindings';
import { resolvePrimarySubtitleText } from './primary-subtitle-text';

type Deps = Parameters<typeof createBindMpvMainEventHandlersHandler>[0];

function createDeps(overrides: Partial<Deps> = {}): Deps {
  return {
    reportJellyfinRemoteStopped: () => {},
    syncOverlayMpvSubtitleSuppression: () => {},
    resetSubtitleSidebarEmbeddedLayout: () => {},
    hasInitialPlaybackQuitOnDisconnectArg: () => false,
    isOverlayRuntimeInitialized: () => false,
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
    refreshDiscordPresence: () => {},

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
    ...overrides,
  };
}

/** Binds the main handlers onto a fake mpv client and returns an `emit(event, payload)` driver. */
function bindEvents(overrides: Partial<Deps> = {}) {
  const handlers = new Map<string, (payload: unknown) => void>();
  createBindMpvMainEventHandlersHandler(createDeps(overrides))({
    on: (event, handler) => {
      handlers.set(event, handler as (payload: unknown) => void);
    },
  });
  return (event: string, payload: unknown): void => handlers.get(event)?.(payload);
}

test('main mpv event binder re-resolves the live subtitle when time-pos seeks', () => {
  const calls: string[] = [];
  let currentTime = 0;
  const liveText = '少しだけ好きになる\n少しだけ好きになる';
  const cues = parseSubtitleCues(
    [
      '[Events]',
      'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
      'Dialogue: 1,0:01:29.00,0:01:32.00,EDJP,,0,0,0,,少しだけ好きになる',
      'Dialogue: 0,0:01:29.00,0:01:32.00,EDJP,,0,0,0,,少しだけ好きになる',
    ].join('\n'),
    'seek-ending.ass',
  );
  const emit = bindEvents({
    resolveSubtitleText: (text) =>
      resolvePrimarySubtitleText({ liveText: text, currentTimeSec: currentTime, cues }),
    getCurrentLiveSubtitleText: () => liveText,
    setCurrentSubText: (text) => calls.push(`set-sub:${text}`),
    onTimePosUpdate: (time) => {
      currentTime = time;
    },
  });

  emit('time-pos-change', { time: 2.5 });
  calls.length = 0;

  // A jump into the cue re-resolves the same live text at the new position.
  emit('time-pos-change', { time: 90 });
  assert.deepEqual(calls, ['set-sub:少しだけ好きになる']);

  // Ordinary playback ticks leave the displayed subtitle alone.
  emit('time-pos-change', { time: 90.5 });
  assert.deepEqual(calls, ['set-sub:少しだけ好きになる']);
});

test('main mpv event binder resets the sidebar layout on connect', () => {
  const calls: string[] = [];
  const emit = bindEvents({
    resetSubtitleSidebarEmbeddedLayout: () => calls.push('reset-sidebar-layout'),
    updateCurrentMediaPath: (path) => calls.push(`media-path:${path}`),
  });

  emit('connection-change', { connected: true });

  assert.deepEqual(calls, ['reset-sidebar-layout']);
});

test('main mpv event binder runs mpv-connected callback on connection', () => {
  const calls: string[] = [];
  const emit = bindEvents({ onMpvConnected: () => calls.push('mpv-connected') });

  emit('connection-change', { connected: true });

  assert.deepEqual(calls, ['mpv-connected']);
});

test('main mpv event binder clears media path on disconnect', () => {
  const calls: string[] = [];
  const emit = bindEvents({
    resetSubtitleSidebarEmbeddedLayout: () => calls.push('reset-sidebar-layout'),
    updateCurrentMediaPath: (path) => calls.push(`media-path:${path}`),
    reportJellyfinRemoteStopped: () => calls.push('remote-stopped'),
    refreshDiscordPresence: () => calls.push('presence-refresh'),
  });

  emit('connection-change', { connected: false });

  assert.ok(calls.includes('media-path:'));
  assert.ok(calls.includes('remote-stopped'));
  assert.ok(calls.includes('presence-refresh'));
  assert.equal(calls.includes('reset-sidebar-layout'), false);
});
