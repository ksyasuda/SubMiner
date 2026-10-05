import assert from 'node:assert/strict';
import test from 'node:test';
import { parseAssCues } from '../../core/services/subtitle-cue-parser';
import { createBuildBindMpvMainEventHandlersMainDepsHandler } from './mpv-main-event-main-deps';

type MainDeps = Parameters<typeof createBuildBindMpvMainEventHandlersMainDepsHandler>[0];

/**
 * Builds the handlers with inert deps. `appStateOverrides` is merged over a blank app
 * state; the full `appState` is returned so tests can mutate it and assert on it.
 */
function buildMainDeps(
  appStateOverrides: Partial<MainDeps['appState']> = {},
  depsOverrides: Partial<MainDeps> = {},
) {
  const appState: MainDeps['appState'] = {
    initialArgs: null,
    overlayRuntimeInitialized: true,
    mpvClient: null,
    immersionTracker: null,
    subtitleTimingTracker: null,
    currentSubText: '',
    currentSubAssText: '',
    playbackPaused: null,
    previousSecondarySubVisibility: false,
    ...appStateOverrides,
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
    ...depsOverrides,
  })();
  return { handlers, appState };
}

test('secondary subtitle changes go to the injected handler instead of the overlay broadcast', () => {
  const calls: string[] = [];
  const injected = buildMainDeps(
    {},
    {
      onSecondarySubtitleChange: (text) => calls.push(`secondary:${text}`),
      broadcastToOverlayWindows: (channel, payload) =>
        calls.push(`broadcast:${channel}:${String(payload)}`),
    },
  ).handlers;
  const fallback = buildMainDeps(
    {},
    {
      broadcastToOverlayWindows: (channel, payload) =>
        calls.push(`broadcast:${channel}:${String(payload)}`),
    },
  ).handlers;

  injected.broadcastSecondarySubtitle('sec');
  fallback.broadcastSecondarySubtitle('sec');

  assert.deepEqual(calls, ['secondary:sec', 'broadcast:secondary-subtitle:set:sec']);
});

test('media duration is recorded for both immersion and AniList', () => {
  const calls: string[] = [];
  const { handlers } = buildMainDeps(
    { immersionTracker: { recordMediaDuration: (sec) => calls.push(`immersion:${sec}`) } },
    { recordAnilistMediaDuration: (sec) => calls.push(`anilist:${sec}`) },
  );

  handlers.recordMediaDuration(1234);

  assert.deepEqual(calls, ['immersion:1234', 'anilist:1234']);
});

test('mpv main event main deps treat managed playback as quit-on-disconnect', () => {
  const deps = buildMainDeps({
    initialArgs: { managedPlayback: true },
    overlayRuntimeInitialized: false,
  }).handlers;

  assert.equal(deps.hasInitialPlaybackQuitOnDisconnectArg(), true);
  assert.equal(deps.shouldQuitOnDisconnectWhenOverlayRuntimeInitialized(), true);
});

test('flushPlaybackPositionOnMediaPathClear ignores disconnected mpv time-pos reads', async () => {
  const recorded: number[] = [];
  const deps = buildMainDeps({
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
    currentMediaPath: '',
  }).handlers;

  deps.flushPlaybackPositionOnMediaPathClear?.('');
  await Promise.resolve();

  assert.deepEqual(recorded, [42]);
});

test('media and subtitle-track transitions reset live subtitle-line deduplication', () => {
  const recordedStarts: number[] = [];
  const handlers = buildMainDeps({
    immersionTracker: {
      recordSubtitleLine: (_text: string, start: number) => recordedStarts.push(start),
    },
    activeParsedSubtitleCues: null,
    currentMediaPath: '/video-a.mkv',
  }).handlers;

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
  const { handlers, appState } = buildMainDeps({
    immersionTracker: {
      recordSubtitleLine: (_text: string, start: number) => recordedStarts.push(start),
    },
    activeParsedSubtitleCues: [{ startTime: 10, endTime: 14, text: '飛び上がる' }],
    currentMediaPath: '/video-a.mkv',
  });

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
  const handlers = buildMainDeps({
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
  }).handlers;

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

  // Jumping back to a brief previous line moves time-pos by less than the general
  // seek threshold. It is still a backward seek, so the revisited line records
  // again; otherwise multi-line copy would keep treating the later line as current.
  handlers.onTimePosUpdate?.(3.9);
  handlers.onTimePosUpdate?.(2.9);
  handlers.recordSubtitleTiming('今', 0.8, 1.5);

  assert.deepEqual(timing.slice(4), [{ text: '今　手にある物差しでは', start: 1.2, end: 3.8 }]);

  // Tiny time-pos jitter is not a seek and must not re-record the line.
  handlers.onTimePosUpdate?.(3.0);
  handlers.onTimePosUpdate?.(2.9);
  handlers.recordSubtitleTiming('今', 0.8, 1.5);
  assert.equal(timing.length, 5);

  handlers.recordImmersionSubtitleLine('Maid\nCafe', 10, 12);
  handlers.recordSubtitleTiming('Maid\nCafe', 10, 12);
  assert.equal(immersion.length, 3);
  assert.equal(timing.length, 5);
});

test('subtitle-track changes stop stale canonical cues from substituting immediately', () => {
  const { handlers, appState } = buildMainDeps({
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
  });

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
  const handlers = buildMainDeps({
    mpvClient: { currentTimePos: 299.7 },
    immersionTracker: { recordSubtitleLine: (text: string) => immersion.push(text) },
    subtitleTimingTracker: { recordSubtitle: (text: string) => timing.push(text) },
    activeParsedSubtitleCues: cues,
  }).handlers;

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
  const handlers = buildMainDeps({
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
  }).handlers;

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
