import assert from 'node:assert/strict';
import test from 'node:test';
import { DEFAULT_CONFIG } from '../../config/definitions';
import { OVERLAY_WINDOW_CONTENT_READY_FLAG } from '../../core/services/overlay-window-flags';
import type { SubtitlePrefetchService } from '../../core/services/subtitle-prefetch';
import type { OverlayContentMeasurement, SubtitleData } from '../../types';
import { createAppState } from '../state';
import { createPlaybackAssembly, type PlaybackAssemblyDeps } from './playback';

type Call = { name: string; args: unknown[] };

/** Records every dep call by name, in order; each dep still returns an inert value. */
function createRecorder() {
  const calls: Call[] = [];
  const spy =
    <A extends unknown[], R>(name: string, impl: (...args: A) => R) =>
    (...args: A): R => {
      calls.push({ name, args });
      return impl(...args);
    };
  const names = () => calls.map((call) => call.name);
  return { calls, spy, names };
}

type Recorder = ReturnType<typeof createRecorder>;

function createFakePrefetchService(recorder: Recorder): SubtitlePrefetchService {
  const { spy } = recorder;
  return {
    start: spy('prefetch.start', () => {}),
    stop: spy('prefetch.stop', () => {}),
    onSeek: spy('prefetch.onSeek', () => {}),
    pause: spy('prefetch.pause', () => {}),
    resume: spy('prefetch.resume', () => {}),
  };
}

function makeDeps(recorder: Recorder): PlaybackAssemblyDeps {
  const { spy } = recorder;
  const noop = () => {};
  const resolved = async () => {};
  return {
    appState: createAppState({
      mpvSocketPath: '/tmp/subminer-playback-test.sock',
      texthookerPort: 0,
    }),
    configService: { getConfig: spy('getConfig', () => DEFAULT_CONFIG) },
    logger: {
      debug: spy('logger.debug', noop),
      info: spy('logger.info', noop),
      warn: spy('logger.warn', noop),
      error: spy('logger.error', noop),
    },
    overlayManager: {
      broadcastToOverlayWindows: spy('overlay.broadcast', noop),
      getVisibleOverlayVisible: spy('overlay.getVisibleOverlayVisible', () => false),
      getMainWindow: spy('overlay.getMainWindow', () => null),
    },
    overlayContentMeasurementStore: {
      getLatestByLayer: spy('measurement.getLatest', () => null),
    },
    subtitleWsService: { broadcast: spy('subtitleWs.broadcast', noop) },
    annotationSubtitleWsService: { broadcast: spy('annotationWs.broadcast', noop) },
    startupOsdSequencer: {
      reset: spy('startupOsd.reset', noop),
      markTokenizationReady: spy('startupOsd.markTokenizationReady', noop),
      showTokenizationLoading: spy('startupOsd.showTokenizationLoading', noop),
      showAnnotationLoading: spy('startupOsd.showAnnotationLoading', noop),
      markAnnotationLoadingComplete: spy('startupOsd.markAnnotationLoadingComplete', noop),
    },
    mediaTimingReviewRuntime: {
      deferPlaybackResume: spy('timingReview.deferPlaybackResume', () => false),
      cancelPlaybackResume: spy('timingReview.cancelPlaybackResume', noop),
    },
    discordPresenceRuntime: { publishDiscordPresence: spy('discord.publish', noop) },
    immersionMediaRuntime: { syncFromCurrentMediaState: spy('immersion.sync', noop) },
    mediaRuntime: {
      updateCurrentMediaPath: spy('media.updateCurrentMediaPath', noop),
      updateCurrentMediaTitle: spy('media.updateCurrentMediaTitle', noop),
    },
    overlayNotificationsRuntime: {
      showConfiguredStatusNotification: spy('notifications.showStatus', noop),
      showConfiguredPlaybackFeedback: spy('notifications.showPlaybackFeedback', noop),
      getConfiguredStatusNotificationType: spy('notifications.getType', () => 'osd' as const),
    },
    overlayModalRuntime: { handleOverlayModalClosed: spy('overlayModal.handleClosed', noop) },
    youtube: {
      isPlaybackActiveNow: spy('youtube.isPlaybackActiveNow', () => false),
      shouldQuitWhenMpvCloses: spy('youtube.shouldQuitWhenMpvCloses', () => false),
      playback: { getQuitOnDisconnectArmed: spy('youtube.getQuitOnDisconnectArmed', () => false) },
      primarySubtitleNotification: {
        handleMediaPathChange: spy('youtube.notification.handleMediaPathChange', noop),
        handleSubtitleTrackChange: spy('youtube.notification.handleSubtitleTrackChange', noop),
        handleSubtitleTrackListChange: spy('youtube.notification.handleTrackListChange', noop),
        isAppOwnedFlowInFlight: spy('youtube.notification.isAppOwnedFlowInFlight', () => false),
      },
      mediaCachePlayback: { handleMediaPathChange: spy('youtube.mediaCache', resolved) },
      browserPlayback: {
        handleMediaPathChange: spy('youtube.browser', noop),
        handleMpvDisconnected: spy('youtube.browser.handleMpvDisconnected', noop),
      },
      flow: {
        handleMediaPathChange: spy('youtube.flow', noop),
        cancelWhisperGeneration: spy('youtube.flow.cancelWhisperGeneration', noop),
        hasActiveSession: spy('youtube.flow.hasActiveSession', () => false),
        cancelActivePicker: spy('youtube.flow.cancelActivePicker', () => false),
      },
    },
    jellyfin: {
      getQuitOnDisconnectArmed: spy('jellyfin.getQuitOnDisconnectArmed', () => false),
      reportJellyfinRemoteStopped: spy('jellyfin.reportStopped', resolved),
      reportJellyfinRemoteProgress: spy('jellyfin.reportProgress', resolved),
      getActivePlayback: spy('jellyfin.getActivePlayback', () => null),
      startJellyfinRemoteSession: spy('jellyfin.startRemoteSession', resolved),
    },
    anilist: {
      maybeRunAnilistPostWatchUpdate: spy('anilist.postWatch', resolved),
      recordAnilistMediaDuration: spy('anilist.recordDuration', noop),
      getCurrentAnilistMediaKey: spy('anilist.getMediaKey', () => null),
      resetAnilistMediaTracking: spy('anilist.resetTracking', noop),
      maybeProbeAnilistDuration: spy('anilist.probeDuration', async () => null),
      ensureAnilistMediaGuess: spy('anilist.ensureGuess', async () => null),
      resetAnilistMediaGuessState: spy('anilist.resetGuess', noop),
    },
    characterDictionary: {
      autoSync: {
        scheduleSync: spy('characterDictionary.scheduleSync', noop),
        getCurrentMediaId: spy('characterDictionary.getCurrentMediaId', () => null),
      },
      imageLookup: { get: spy('characterDictionary.imageLookup', () => null) },
      nameCandidates: { get: spy('characterDictionary.nameCandidates', () => null) },
    },
    dictionaries: {
      jlpt: { ensureJlptDictionaryLookup: spy('dictionaries.jlpt', resolved) },
      frequency: { ensureFrequencyDictionaryLookup: spy('dictionaries.frequency', resolved) },
    },
    getSubtitleGenerationRuntime: spy('getSubtitleGenerationRuntime', () => ({
      handleMediaPathChange: noop,
    })),
    getPluginHost: spy('getPluginHost', () => ({ attachMpvClient: noop })),
    isCharacterDictionaryEnabled: spy('isCharacterDictionaryEnabled', () => false),
    hasTray: spy('hasTray', () => false),
    loadSubtitleSourceText: spy('loadSubtitleSourceText', async () => ''),
    createSubtitlePrefetchService: spy('createSubtitlePrefetchService', () =>
      createFakePrefetchService(recorder),
    ),
    requestAppQuit: spy('requestAppQuit', noop),
    maybeStartOverlayLoadingOsd: spy('maybeStartOverlayLoadingOsd', noop),
    reportOverlayContentReadyToPlugin: spy('reportOverlayContentReadyToPlugin', noop),
    flushQueuedMpvOsdNotifications: spy('flushQueuedMpvOsdNotifications', noop),
    refreshMpvSessionBindings: spy('refreshMpvSessionBindings', resolved),
    ensureImmersionTrackerStarted: spy('ensureImmersionTrackerStarted', noop),
    restoreOverlayMpvSubtitles: spy('restoreOverlayMpvSubtitles', noop),
    syncOverlayMpvSubtitleSuppression: spy('syncOverlayMpvSubtitleSuppression', noop),
    setOverlayVisible: spy('setOverlayVisible', noop),
    ensureYomitanExtensionLoaded: spy('ensureYomitanExtensionLoaded', async () => null),
    shouldInitializeMecabForAnnotations: spy('shouldInitializeMecabForAnnotations', () => false),
    // Every annotation off, so tokenization warmups need no MeCab or dictionaries.
    getRuntimeBooleanOption: spy('getRuntimeBooleanOption', () => false),
    onFullscreenChange: spy('onFullscreenChange', noop),
    stealAppFocus: spy('stealAppFocus', noop),
    openJimaku: spy('overlayModals.openJimaku', async () => false),
  };
}

/** Binds the assembly's mpv event handlers to a fake client; returns an `emit` driver. */
function bindMpvEvents(playback: ReturnType<typeof createPlaybackAssembly>) {
  const handlers = new Map<string, (payload: unknown) => void>();
  playback.bindMpvClientEventHandlers({
    on: (event, handler) => {
      handlers.set(event, handler as (payload: unknown) => void);
    },
  });
  return (event: string, payload: unknown): void => {
    const handler = handlers.get(event);
    assert.ok(handler, `no handler bound for ${event}`);
    handler(payload);
  };
}

/**
 * Replaces the assembly's autoplay-ready and tokenization gates with recording stubs, so tests
 * pin playback's calls into them. Playback looks these methods up at call time.
 */
function stubGates(recorder: Recorder, playback: ReturnType<typeof createPlaybackAssembly>) {
  const { spy } = recorder;
  const noop = () => {};
  Object.assign(playback.autoplayReadyGate, {
    maybeSignalPluginAutoplayReady: spy('autoplay.maybeSignal', noop),
    flushPendingAutoplayReadySignal: spy('autoplay.flushPending', noop),
    invalidatePendingAutoplayReadyFallbacks: spy('autoplay.invalidateFallbacks', noop),
  } satisfies Partial<typeof playback.autoplayReadyGate>);
  Object.assign(playback.currentMediaTokenizationGate, {
    updateCurrentMediaPath: spy('tokenizationGate.updateCurrentMediaPath', noop),
    markReady: spy('tokenizationGate.markReady', noop),
  } satisfies Partial<typeof playback.currentMediaTokenizationGate>);
}

function setup(
  overrides: Partial<PlaybackAssemblyDeps> = {},
  options: { realGates?: boolean } = {},
) {
  const recorder = createRecorder();
  const deps = { ...makeDeps(recorder), ...overrides };
  const playback = createPlaybackAssembly(deps);
  if (!options.realGates) stubGates(recorder, playback);
  const emit = bindMpvEvents(playback);
  return { recorder, deps, appState: deps.appState, playback, emit };
}

/** Sets up with a fake prefetch service, installed by a real prefetch init over one SRT cue. */
async function setupWithFakePrefetch() {
  const ctx = setup({
    loadSubtitleSourceText: async () => '1\n00:00:01,000 --> 00:00:02,000\n一行目\n',
  });
  await ctx.playback.subtitlePrefetchRuntime.initSubtitlePrefetch('/subs/episode.srt', 0);
  assert.ok(ctx.playback.getSubtitlePrefetchService());
  ctx.recorder.calls.length = 0;
  return ctx;
}

const flushAsync = async (): Promise<void> => {
  for (let i = 0; i < 5; i += 1) await new Promise((resolve) => setTimeout(resolve, 0));
};

// main.ts builds the assembly before some siblings exist, so construction must stay lazy.
test('createPlaybackAssembly does not call its deps while constructing', () => {
  const recorder = createRecorder();
  createPlaybackAssembly(makeDeps(recorder));
  assert.deepEqual(recorder.names(), []);
});

test('mpv connection starts the loading OSD, resends overlay readiness, and flushes queued OSD notifications before session work', () => {
  const { recorder, emit } = setup();

  emit('connection-change', { connected: true });

  const names = recorder.names();
  const osdIndex = names.indexOf('maybeStartOverlayLoadingOsd');
  assert.ok(osdIndex >= 0);
  assert.deepEqual(recorder.calls[osdIndex]?.args, []);
  assert.ok(osdIndex < names.indexOf('reportOverlayContentReadyToPlugin'));
  assert.ok(
    names.indexOf('reportOverlayContentReadyToPlugin') <
      names.indexOf('flushQueuedMpvOsdNotifications'),
  );
  assert.ok(
    names.indexOf('flushQueuedMpvOsdNotifications') < names.indexOf('refreshMpvSessionBindings'),
  );
});

test('media path change starts the loading OSD, then clears rendered subtitle state before broadcasting', (t) => {
  const { recorder, deps, appState, playback, emit } = setup();
  t.after(() => playback.subtitlePrefetchRuntime.clearScheduledSubtitlePrefetchRefresh());
  appState.currentMediaPath = '/videos/a.mkv';
  appState.currentSubText = 'old line';
  appState.currentSubtitleData = { text: 'old line', tokens: null };
  appState.activeParsedSubtitleCues = [{ startTime: 1, endTime: 2, text: 'old line' }];
  appState.activeParsedSubtitleSource = '/subs/a.srt';
  appState.activeParsedSubtitleMediaPath = '/videos/a.mkv';
  emit('time-pos-change', { time: 42 });
  assert.equal(playback.getLastObservedTimePos(), 42);
  let subtitleDataAtBroadcast: SubtitleData | null | undefined;
  deps.overlayManager.broadcastToOverlayWindows = (channel) => {
    if (channel === 'subtitle:set') subtitleDataAtBroadcast = appState.currentSubtitleData;
  };
  recorder.calls.length = 0;

  emit('media-path-change', { path: '/videos/b.mkv' });

  assert.deepEqual(recorder.calls[0], {
    name: 'maybeStartOverlayLoadingOsd',
    args: ['/videos/b.mkv'],
  });
  assert.equal(subtitleDataAtBroadcast, null);
  assert.equal(appState.currentSubText, '');
  assert.equal(appState.currentSubtitleData, null);
  assert.deepEqual(appState.activeParsedSubtitleCues, []);
  assert.equal(appState.activeParsedSubtitleSource, null);
  assert.equal(appState.activeParsedSubtitleMediaPath, null);
  assert.equal(playback.getLastObservedTimePos(), 0);
  const resetPayload = { text: '', tokens: null };
  const wsCalls = recorder.calls.filter((call) => call.name.endsWith('Ws.broadcast'));
  assert.deepEqual(
    wsCalls.map((call) => [call.name, call.args[0]]),
    [
      ['subtitleWs.broadcast', resetPayload],
      ['annotationWs.broadcast', resetPayload],
    ],
  );
});

test('media path change to the same YouTube video keeps its parsed cues', (t) => {
  const { appState, playback, emit } = setup();
  t.after(() => playback.subtitlePrefetchRuntime.clearScheduledSubtitlePrefetchRefresh());
  const cues = [{ startTime: 1, endTime: 2, text: '字幕' }];
  appState.currentMediaPath = 'https://www.youtube.com/watch?v=dQw4w9WgXcQ';
  appState.currentSubText = '字幕';
  appState.activeParsedSubtitleCues = cues;
  appState.activeParsedSubtitleSource = 'youtube:ja';
  appState.activeParsedSubtitleMediaPath = 'https://www.youtube.com/watch?v=dQw4w9WgXcQ';

  emit('media-path-change', { path: 'https://www.youtube.com/watch?v=dQw4w9WgXcQ&t=30' });

  assert.equal(appState.currentSubText, '');
  assert.equal(appState.activeParsedSubtitleCues, cues);
  assert.equal(appState.activeParsedSubtitleSource, 'youtube:ja');
  assert.equal(
    appState.activeParsedSubtitleMediaPath,
    'https://www.youtube.com/watch?v=dQw4w9WgXcQ',
  );
});

test('same media path update keeps autoplay fallbacks; a new path invalidates them first', (t) => {
  const { recorder, appState, playback, emit } = setup();
  t.after(() => playback.subtitlePrefetchRuntime.clearScheduledSubtitlePrefetchRefresh());
  appState.currentMediaPath = '/videos/a.mkv';

  emit('media-path-change', { path: '/videos/a.mkv' });
  assert.ok(!recorder.names().includes('autoplay.invalidateFallbacks'));
  assert.ok(recorder.names().includes('tokenizationGate.updateCurrentMediaPath'));

  recorder.calls.length = 0;
  emit('media-path-change', { path: '/videos/b.mkv' });
  const names = recorder.names();
  assert.ok(names.includes('autoplay.invalidateFallbacks'));
  assert.ok(
    names.indexOf('autoplay.invalidateFallbacks') <
      names.indexOf('tokenizationGate.updateCurrentMediaPath'),
  );
});

test('subtitle change pauses prefetch without reseeking it, and resumes on a repeated line', async () => {
  const ctx = await setupWithFakePrefetch();
  const prefetchCalls = () => ctx.recorder.names().filter((name) => name.startsWith('prefetch.'));

  // The controller already holds the empty line, so nothing will settle to resume prefetch.
  ctx.emit('subtitle-change', { text: '' });
  assert.deepEqual(prefetchCalls(), ['prefetch.pause', 'prefetch.resume']);

  ctx.recorder.calls.length = 0;
  ctx.emit('subtitle-change', { text: '新しい行' });
  // A new line is still tokenizing, so its settle (not this handler) resumes prefetch.
  assert.deepEqual(prefetchCalls(), ['prefetch.pause']);
  await flushAsync();
});

test('annotation refresh invalidates cached tokenizations and reseeks prefetch', async () => {
  const ctx = await setupWithFakePrefetch();
  const generation = ctx.playback.subtitleProcessingController.getCacheGeneration();

  ctx.playback.refreshCurrentSubtitleAnnotations();

  assert.equal(ctx.playback.subtitleProcessingController.getCacheGeneration(), generation + 1);
  // No current line: nothing will settle, so the refresh resumes prefetch itself.
  assert.deepEqual(
    ctx.recorder.calls.filter((call) => call.name.startsWith('prefetch.')),
    [
      { name: 'prefetch.onSeek', args: [0] },
      { name: 'prefetch.resume', args: [] },
    ],
  );

  ctx.recorder.calls.length = 0;
  ctx.appState.currentSubText = '現在の行';
  ctx.playback.refreshCurrentSubtitleAnnotations();
  assert.deepEqual(
    ctx.recorder.names().filter((name) => name.startsWith('prefetch.')),
    ['prefetch.pause', 'prefetch.onSeek'],
  );
  await flushAsync();
});

test('annotation upgrade of the shown line skips the duplicate plain websocket event', () => {
  const { recorder, appState, playback, emit } = setup();
  // Emitted payloads always carry their timing, and an upgrade must match it.
  appState.currentSubtitleData = { text: '同じ行', tokens: null, startTime: null, endTime: null };
  playback.subtitleProcessingController.preCacheTokenization('同じ行', {
    text: '同じ行',
    tokens: [],
  });

  emit('subtitle-change', { text: '同じ行' });

  const names = recorder.names();
  assert.ok(!names.includes('subtitleWs.broadcast'));
  assert.equal(names.filter((name) => name === 'annotationWs.broadcast').length, 1);
  assert.equal(
    recorder.calls.filter(
      (call) => call.name === 'overlay.broadcast' && call.args[0] === 'subtitle:set',
    ).length,
    1,
  );
  assert.deepEqual(appState.currentSubtitleData?.tokens, []);
});

test('warm tokenization releases autoplay with a warm signal before the first subtitle', async (t) => {
  const { recorder, appState, playback, emit } = setup();
  t.after(() => playback.subtitlePrefetchRuntime.clearScheduledSubtitlePrefetchRefresh());
  appState.currentMediaPath = '/videos/a.mkv';
  appState.currentSubText = '';

  emit('media-path-change', { path: '/videos/a.mkv' });
  await flushAsync();

  const signals = recorder.calls.filter((call) => call.name === 'autoplay.maybeSignal');
  assert.deepEqual(signals.at(-1)?.args, [
    { text: '__warm__', tokens: null },
    { forceWhilePaused: true },
  ]);
  const names = recorder.names();
  assert.ok(names.lastIndexOf('autoplay.flushPending') < names.lastIndexOf('autoplay.maybeSignal'));
});

test('warm tokenization releases autoplay with the current subtitle when it is shown', async (t) => {
  const { recorder, appState, playback, emit } = setup();
  t.after(() => playback.subtitlePrefetchRuntime.clearScheduledSubtitlePrefetchRefresh());
  const current: SubtitleData = { text: '最初の行', tokens: null };
  appState.currentMediaPath = '/videos/a.mkv';
  appState.currentSubText = '最初の行';
  appState.currentSubtitleData = current;

  emit('media-path-change', { path: '/videos/a.mkv' });
  await flushAsync();

  const signals = recorder.calls.filter((call) => call.name === 'autoplay.maybeSignal');
  assert.deepEqual(signals.at(-1)?.args, [current, { forceWhilePaused: true }]);
});

/** An mpv client that answers the sidebar's property reads from a table. */
function fakeSidebarMpvClient(properties: Record<string, unknown>) {
  return {
    connected: true,
    requestProperty: async (name: string) => properties[name] ?? null,
  } as unknown as NonNullable<ReturnType<typeof setup>['appState']['mpvClient']>;
}

test('sidebar snapshot keeps cached YouTube cues instead of re-resolving the subtitle source', async () => {
  const { appState, playback, recorder } = setup();
  const cues = [{ startTime: 1, endTime: 2, text: '字幕' }];
  appState.activeParsedSubtitleCues = cues;
  appState.activeParsedSubtitleMediaPath = 'https://www.youtube.com/watch?v=dQw4w9WgXcQ';
  appState.mpvClient = fakeSidebarMpvClient({
    path: 'https://www.youtube.com/watch?v=dQw4w9WgXcQ&t=30',
    // Without the YouTube cache check this source would be re-read and re-parsed.
    'current-tracks/sub/external-filename': '/tmp/youtube-subs/dQw4w9WgXcQ.ja.vtt',
    'track-list': [],
    sid: 1,
  });

  const snapshot = await playback.ipc.getSubtitleSidebarSnapshot();

  assert.equal(snapshot.cues, cues);
  assert.equal(recorder.names().includes('loadSubtitleSourceText'), false);
});

test('sidebar snapshot falls back to cached cues while mpv is disconnected', async () => {
  const { appState, playback, recorder } = setup();
  const cues = [{ startTime: 1, endTime: 2, text: '字幕' }];
  appState.activeParsedSubtitleCues = cues;

  const snapshot = await playback.ipc.getSubtitleSidebarSnapshot();

  assert.equal(snapshot.cues, cues);
  assert.equal(snapshot.config, DEFAULT_CONFIG.subtitleSidebar);
  assert.equal(recorder.names().includes('loadSubtitleSourceText'), false);
});

/**
 * Real autoplay gate over a paused, connected mpv client. `overlayVisible` and `measurement`
 * control the visible-overlay readiness the gate waits for.
 */
function setupAutoplayGate(fixture: {
  overlayVisible: boolean;
  measurement?: OverlayContentMeasurement | null;
}) {
  const overlayWindow = {
    isDestroyed: () => false,
    // Focused already, so the macOS pointer recovery on release does not steal focus.
    isVisible: () => true,
    isFocused: () => true,
    [OVERLAY_WINDOW_CONTENT_READY_FLAG]: true,
  };
  const ctx = setup(
    {
      overlayManager: {
        broadcastToOverlayWindows: () => {},
        getVisibleOverlayVisible: () => fixture.overlayVisible,
        getMainWindow: () => overlayWindow as unknown as Electron.BrowserWindow,
      },
      overlayContentMeasurementStore: { getLatestByLayer: () => fixture.measurement ?? null },
    },
    { realGates: true },
  );
  let paused = true;
  const scriptMessages: string[] = [];
  ctx.appState.currentMediaPath = '/videos/a.mkv';
  ctx.appState.mpvClient = {
    connected: true,
    requestProperty: async (name: string) => (name === 'pause' ? paused : null),
    send: ({ command }: { command: unknown[] }) => {
      if (command[0] === 'script-message') scriptMessages.push(String(command[1]));
      if (command[0] === 'set_property' && command[1] === 'pause') paused = command[2] === true;
    },
  } as unknown as typeof ctx.appState.mpvClient;
  const released = () => scriptMessages.includes('subminer-autoplay-ready');
  return { ...ctx, released };
}

const TOKENIZED_LINE: SubtitleData = { text: '一行目', tokens: [] };

test('autoplay stays paused until the visible overlay has measured the subtitle', () => {
  const fixture: { overlayVisible: boolean; measurement: OverlayContentMeasurement | null } = {
    overlayVisible: true,
    measurement: null,
  };
  const { playback, released } = setupAutoplayGate(fixture);

  playback.autoplayReadyGate.maybeSignalPluginAutoplayReady(TOKENIZED_LINE);
  assert.equal(released(), false);

  fixture.measurement = {
    layer: 'visible',
    measuredAtMs: Date.now() + 1_000,
    viewport: { width: 1920, height: 1080 },
    contentRect: { x: 100, y: 900, width: 600, height: 60 },
  } as OverlayContentMeasurement;
  playback.autoplayReadyGate.flushPendingAutoplayReadySignal();
  assert.equal(released(), true);
});

// Untokenized lines (startup priming, tokenizer fallbacks) must not release the startup pause
// while tokenization warms up; a tokenized line must not wait on the warmup flag.
test('autoplay ignores an untokenized line during tokenization warmup but releases a tokenized one', () => {
  const { playback, released } = setupAutoplayGate({ overlayVisible: false });
  assert.equal(playback.isTokenizationWarmupReady(), false);

  playback.autoplayReadyGate.maybeSignalPluginAutoplayReady({ text: '一行目', tokens: null });
  playback.autoplayReadyGate.flushPendingAutoplayReadySignal();
  assert.equal(released(), false);

  playback.autoplayReadyGate.maybeSignalPluginAutoplayReady(TOKENIZED_LINE);
  assert.equal(released(), true);
});
