import test from 'node:test';
import assert from 'node:assert/strict';

import { createIpcDepsRuntime, registerIpcHandlers, type IpcServiceDeps } from './ipc';
import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import type {
  PlaylistBrowserSnapshot,
  SessionActionDispatchRequest,
  SubtitleSidebarSnapshot,
} from '../../types';

interface FakeIpcRegistrar {
  on: Map<string, (event: unknown, ...args: unknown[]) => void>;
  handle: Map<string, (event: unknown, ...args: unknown[]) => unknown>;
}

function createFakeIpcRegistrar(): {
  registrar: {
    on: (channel: string, listener: (event: unknown, ...args: unknown[]) => void) => void;
    handle: (channel: string, listener: (event: unknown, ...args: unknown[]) => unknown) => void;
  };
  handlers: FakeIpcRegistrar;
} {
  const handlers: FakeIpcRegistrar = {
    on: new Map(),
    handle: new Map(),
  };
  return {
    registrar: {
      on: (channel, listener) => {
        handlers.on.set(channel, listener);
      },
      handle: (channel, listener) => {
        handlers.handle.set(channel, listener);
      },
    },
    handlers,
  };
}

function createControllerConfigFixture() {
  return {
    enabled: true,
    preferredGamepadId: '',
    preferredGamepadLabel: '',
    smoothScroll: true,
    scrollPixelsPerSecond: 960,
    horizontalJumpPixels: 160,
    stickDeadzone: 0.2,
    triggerInputMode: 'auto' as const,
    triggerDeadzone: 0.5,
    repeatDelayMs: 220,
    repeatIntervalMs: 80,
    buttonIndices: {
      select: 6,
      buttonSouth: 0,
      buttonEast: 1,
      buttonWest: 2,
      buttonNorth: 3,
      leftShoulder: 4,
      rightShoulder: 5,
      leftStickPress: 9,
      rightStickPress: 10,
      leftTrigger: 6,
      rightTrigger: 7,
    },
    bindings: {
      toggleLookup: { kind: 'button' as const, buttonIndex: 0 },
      closeLookup: { kind: 'button' as const, buttonIndex: 1 },
      toggleKeyboardOnlyMode: { kind: 'button' as const, buttonIndex: 3 },
      mineCard: { kind: 'button' as const, buttonIndex: 2 },
      quitMpv: { kind: 'button' as const, buttonIndex: 6 },
      previousAudio: { kind: 'button' as const, buttonIndex: 4 },
      nextAudio: { kind: 'button' as const, buttonIndex: 5 },
      playCurrentAudio: { kind: 'button' as const, buttonIndex: 7 },
      toggleMpvPause: { kind: 'button' as const, buttonIndex: 6 },
      leftStickHorizontal: {
        kind: 'axis' as const,
        axisIndex: 0,
        dpadFallback: 'horizontal' as const,
      },
      leftStickVertical: { kind: 'axis' as const, axisIndex: 1, dpadFallback: 'vertical' as const },
      rightStickHorizontal: { kind: 'axis' as const, axisIndex: 3, dpadFallback: 'none' as const },
      rightStickVertical: { kind: 'axis' as const, axisIndex: 4, dpadFallback: 'none' as const },
    },
  };
}

function createSubtitleSidebarSnapshotFixture(): SubtitleSidebarSnapshot {
  return {
    cues: [],
    currentSubtitle: { text: '', startTime: null, endTime: null },
    config: {
      enabled: false,
      autoOpen: false,
      layout: 'overlay',
      toggleKey: 'Backslash',
      pauseVideoOnHover: false,
      autoScroll: true,
      maxWidth: 420,
      opacity: 0.92,
      backgroundColor: 'rgba(54, 58, 79, 0.88)',
      textColor: '#cad3f5',
      fontFamily: '"M PLUS 1", "Noto Sans CJK JP", sans-serif',
      fontSize: 16,
      timestampColor: '#a5adcb',
      activeLineColor: '#f5bde6',
      activeLineBackgroundColor: 'rgba(138, 173, 244, 0.22)',
      hoverLineBackgroundColor: 'rgba(54, 58, 79, 0.84)',
    },
  };
}

function createRegisterIpcDeps(overrides: Partial<IpcServiceDeps> = {}): IpcServiceDeps {
  return {
    onOverlayModalClosed: () => {},
    openYomitanSettings: () => {},
    quitApp: () => {},
    toggleDevTools: () => {},
    getVisibleOverlayVisibility: () => false,
    toggleVisibleOverlay: () => {},
    tokenizeCurrentSubtitle: async () => null,
    getCurrentSubtitleRaw: () => '',
    getCurrentSubtitleAss: () => '',
    getSubtitleSidebarSnapshot: async () => createSubtitleSidebarSnapshotFixture(),
    getPlaybackPaused: () => false,
    getSubtitlePosition: () => null,
    getSubtitleStyle: () => null,
    saveSubtitlePosition: () => {},
    getMecabStatus: () => ({ available: false, enabled: false, path: null }),
    setMecabEnabled: () => {},
    handleMpvCommand: () => {},
    getKeybindings: () => [],
    getSessionBindings: () => [],
    getConfiguredShortcuts: () => ({}),
    dispatchSessionAction: async () => {},
    getStatsToggleKey: () => 'Backquote',
    getMarkWatchedKey: () => 'KeyW',
    getControllerConfig: () => createControllerConfigFixture(),
    saveControllerConfig: async () => {},
    saveControllerPreference: async () => {},
    getSecondarySubMode: () => 'hover',
    getCurrentSecondarySub: () => '',
    focusMainWindow: () => {},
    runSubsyncManual: async () => ({ ok: true, message: 'ok' }),
    getAnkiConnectStatus: () => false,
    getRuntimeOptions: () => [],
    setRuntimeOption: () => ({ ok: true }),
    cycleRuntimeOption: () => ({ ok: true }),
    reportOverlayContentBounds: () => {},
    getAnilistStatus: () => ({}),
    clearAnilistToken: () => {},
    openAnilistSetup: () => {},
    getAnilistQueueStatus: () => ({}),
    retryAnilistQueueNow: async () => ({ ok: true, message: 'ok' }),
    appendClipboardVideoToQueue: () => ({ ok: true, message: 'ok' }),
    getPlaylistBrowserSnapshot: async () => ({
      directoryPath: null,
      directoryAvailable: false,
      directoryStatus: '',
      directoryItems: [],
      playlistItems: [],
      playingIndex: null,
      currentFilePath: null,
    }),
    appendPlaylistBrowserFile: async () => ({ ok: true, message: 'ok' }),
    playPlaylistBrowserIndex: async () => ({ ok: true, message: 'ok' }),
    removePlaylistBrowserIndex: async () => ({ ok: true, message: 'ok' }),
    movePlaylistBrowserIndex: async () => ({ ok: true, message: 'ok' }),
    onYoutubePickerResolve: async () => ({ ok: true, message: 'ok' }),
    immersionTracker: null,
    ...overrides,
  };
}

function createFakeImmersionTracker(
  overrides: Partial<NonNullable<IpcServiceDeps['immersionTracker']>> = {},
): NonNullable<IpcServiceDeps['immersionTracker']> {
  return {
    recordYomitanLookup: () => {},
    getSessionSummaries: async () => [],
    getDailyRollups: async () => [],
    getMonthlyRollups: async () => [],
    getQueryHints: async () => ({
      totalSessions: 0,
      activeSessions: 0,
      episodesToday: 0,
      activeAnimeCount: 0,
      totalActiveMin: 0,
      totalCards: 0,
      activeDays: 0,
      totalEpisodesWatched: 0,
      totalAnimeCompleted: 0,
      totalTokensSeen: 0,
      totalLookupCount: 0,
      totalLookupHits: 0,
      totalYomitanLookupCount: 0,
      newWordsToday: 0,
      newWordsThisWeek: 0,
    }),
    getSessionTimeline: async () => [],
    getSessionEvents: async () => [],
    getVocabularyStats: async () => [],
    getKanjiStats: async () => [],
    getMediaLibrary: async () => [],
    getMediaDetail: async () => null,
    getMediaSessions: async () => [],
    getMediaDailyRollups: async () => [],
    getCoverArt: async () => null,
    markActiveVideoWatched: async () => false,
    ...overrides,
  };
}

test('createIpcDepsRuntime wires AniList handlers', async () => {
  const calls: string[] = [];
  const deps = createIpcDepsRuntime({
    getMainWindow: () => null,
    getVisibleOverlayVisibility: () => false,
    onOverlayModalClosed: () => {},
    openYomitanSettings: () => {},
    quitApp: () => {},
    toggleVisibleOverlay: () => {},
    tokenizeCurrentSubtitle: async () => null,
    getCurrentSubtitleRaw: () => '',
    getCurrentSubtitleAss: () => '',
    getSubtitleSidebarSnapshot: async () => createSubtitleSidebarSnapshotFixture(),
    getPlaybackPaused: () => true,
    getSubtitlePosition: () => null,
    getSubtitleStyle: () => null,
    saveSubtitlePosition: () => {},
    getMecabTokenizer: () => null,
    handleMpvCommand: () => {},
    getKeybindings: () => [],
    getSessionBindings: () => [],
    getConfiguredShortcuts: () => ({}),
    dispatchSessionAction: async () => {},
    getStatsToggleKey: () => 'Backquote',
    getMarkWatchedKey: () => 'KeyW',
    getControllerConfig: () => createControllerConfigFixture(),
    saveControllerConfig: () => {},
    saveControllerPreference: () => {},
    getSecondarySubMode: () => 'hover',
    getMpvClient: () => null,
    focusMainWindow: () => {},
    runSubsyncManual: async () => ({ ok: true, message: 'ok' }),
    getAnkiConnectStatus: () => false,
    getRuntimeOptions: () => ({}),
    setRuntimeOption: () => ({ ok: true }),
    cycleRuntimeOption: () => ({ ok: true }),
    reportOverlayContentBounds: () => {},
    getAnilistStatus: () => ({ tokenStatus: 'resolved' }),
    clearAnilistToken: () => {
      calls.push('clearAnilistToken');
    },
    openAnilistSetup: () => {
      calls.push('openAnilistSetup');
    },
    getAnilistQueueStatus: () => ({ pending: 1, ready: 0, deadLetter: 0 }),
    retryAnilistQueueNow: async () => {
      calls.push('retryAnilistQueueNow');
      return { ok: true, message: 'done' };
    },
    appendClipboardVideoToQueue: () => ({ ok: true, message: 'queued' }),
    getPlaylistBrowserSnapshot: async () => ({
      directoryPath: '/tmp',
      directoryAvailable: true,
      directoryStatus: '/tmp',
      directoryItems: [],
      playlistItems: [],
      playingIndex: 0,
      currentFilePath: '/tmp/current.mkv',
    }),
    appendPlaylistBrowserFile: async () => ({ ok: true, message: 'append' }),
    playPlaylistBrowserIndex: async () => ({ ok: true, message: 'play' }),
    removePlaylistBrowserIndex: async () => ({ ok: true, message: 'remove' }),
    movePlaylistBrowserIndex: async () => ({ ok: true, message: 'move' }),
    onYoutubePickerResolve: async () => ({ ok: true, message: 'ok' }),
  });

  assert.deepEqual(deps.getAnilistStatus(), { tokenStatus: 'resolved' });
  deps.clearAnilistToken();
  deps.openAnilistSetup();
  assert.deepEqual(deps.getAnilistQueueStatus(), {
    pending: 1,
    ready: 0,
    deadLetter: 0,
  });
  assert.deepEqual(await deps.retryAnilistQueueNow(), {
    ok: true,
    message: 'done',
  });
  assert.equal((await deps.getPlaylistBrowserSnapshot()).directoryAvailable, true);
  assert.deepEqual(await deps.appendPlaylistBrowserFile('/tmp/new.mkv'), {
    ok: true,
    message: 'append',
  });
  assert.deepEqual(await deps.playPlaylistBrowserIndex(2), { ok: true, message: 'play' });
  assert.deepEqual(await deps.removePlaylistBrowserIndex(2), { ok: true, message: 'remove' });
  assert.deepEqual(await deps.movePlaylistBrowserIndex(2, -1), { ok: true, message: 'move' });
  assert.deepEqual(calls, ['clearAnilistToken', 'openAnilistSetup', 'retryAnilistQueueNow']);
  assert.equal(deps.getPlaybackPaused(), true);
});

test('registerIpcHandlers exposes playlist browser snapshot and mutations', async () => {
  const { registrar, handlers } = createFakeIpcRegistrar();
  const calls: Array<[string, unknown[]]> = [];
  registerIpcHandlers(
    createRegisterIpcDeps({
      getPlaylistBrowserSnapshot: async () => ({
        directoryPath: '/tmp/videos',
        directoryAvailable: true,
        directoryStatus: '/tmp/videos',
        directoryItems: [],
        playlistItems: [],
        playingIndex: 1,
        currentFilePath: '/tmp/videos/ep2.mkv',
      }),
      appendPlaylistBrowserFile: async (filePath) => {
        calls.push(['append', [filePath]]);
        return { ok: true, message: 'append-ok' };
      },
      playPlaylistBrowserIndex: async (index) => {
        calls.push(['play', [index]]);
        return { ok: true, message: 'play-ok' };
      },
      removePlaylistBrowserIndex: async (index) => {
        calls.push(['remove', [index]]);
        return { ok: true, message: 'remove-ok' };
      },
      movePlaylistBrowserIndex: async (index, direction) => {
        calls.push(['move', [index, direction]]);
        return { ok: true, message: 'move-ok' };
      },
    }),
    registrar,
  );

  const snapshot = (await handlers.handle.get(IPC_CHANNELS.request.getPlaylistBrowserSnapshot)?.(
    {},
  )) as PlaylistBrowserSnapshot | undefined;
  const append = await handlers.handle.get(IPC_CHANNELS.request.appendPlaylistBrowserFile)?.(
    {},
    '/tmp/videos/ep3.mkv',
  );
  const play = await handlers.handle.get(IPC_CHANNELS.request.playPlaylistBrowserIndex)?.({}, 2);
  const remove = await handlers.handle.get(IPC_CHANNELS.request.removePlaylistBrowserIndex)?.(
    {},
    2,
  );
  const move = await handlers.handle.get(IPC_CHANNELS.request.movePlaylistBrowserIndex)?.(
    {},
    2,
    -1,
  );

  assert.equal(snapshot?.playingIndex, 1);
  assert.deepEqual(append, { ok: true, message: 'append-ok' });
  assert.deepEqual(play, { ok: true, message: 'play-ok' });
  assert.deepEqual(remove, { ok: true, message: 'remove-ok' });
  assert.deepEqual(move, { ok: true, message: 'move-ok' });
  assert.deepEqual(calls, [
    ['append', ['/tmp/videos/ep3.mkv']],
    ['play', [2]],
    ['remove', [2]],
    ['move', [2, -1]],
  ]);
});

test('registerIpcHandlers rejects malformed runtime-option payloads', async () => {
  const { registrar, handlers } = createFakeIpcRegistrar();
  const calls: Array<{ id: string; value: unknown }> = [];
  const cycles: Array<{ id: string; direction: 1 | -1 }> = [];
  registerIpcHandlers(
    {
      onOverlayModalClosed: () => {},
      openYomitanSettings: () => {},
      quitApp: () => {},
      toggleDevTools: () => {},
      getVisibleOverlayVisibility: () => false,
      toggleVisibleOverlay: () => {},
      tokenizeCurrentSubtitle: async () => null,
      getCurrentSubtitleRaw: () => '',
      getCurrentSubtitleAss: () => '',
      getPlaybackPaused: () => null,
      getSubtitlePosition: () => null,
      getSubtitleStyle: () => null,
      saveSubtitlePosition: () => {},
      getMecabStatus: () => ({ available: false, enabled: false, path: null }),
      setMecabEnabled: () => {},
      handleMpvCommand: () => {},
      getKeybindings: () => [],
      getSessionBindings: () => [],
      getConfiguredShortcuts: () => ({}),
      dispatchSessionAction: async () => {},
      getStatsToggleKey: () => 'Backquote',
      getMarkWatchedKey: () => 'KeyW',
      getControllerConfig: () => createControllerConfigFixture(),
      saveControllerConfig: () => {},
      saveControllerPreference: () => {},
      getSecondarySubMode: () => 'hover',
      getCurrentSecondarySub: () => '',
      focusMainWindow: () => {},
      runSubsyncManual: async () => ({ ok: true, message: 'ok' }),
      getAnkiConnectStatus: () => false,
      getRuntimeOptions: () => [],
      setRuntimeOption: (id, value) => {
        calls.push({ id, value });
        return { ok: true };
      },
      cycleRuntimeOption: (id, direction) => {
        cycles.push({ id, direction });
        return { ok: true };
      },
      getSubtitleSidebarSnapshot: async () => createSubtitleSidebarSnapshotFixture(),
      reportOverlayContentBounds: () => {},
      getAnilistStatus: () => ({}),
      clearAnilistToken: () => {},
      openAnilistSetup: () => {},
      getAnilistQueueStatus: () => ({}),
      retryAnilistQueueNow: async () => ({ ok: true, message: 'ok' }),
      appendClipboardVideoToQueue: () => ({ ok: true, message: 'ok' }),
      getPlaylistBrowserSnapshot: async () => ({
        directoryPath: null,
        directoryAvailable: false,
        directoryStatus: '',
        directoryItems: [],
        playlistItems: [],
        playingIndex: null,
        currentFilePath: null,
      }),
      appendPlaylistBrowserFile: async () => ({ ok: true, message: 'ok' }),
      playPlaylistBrowserIndex: async () => ({ ok: true, message: 'ok' }),
      removePlaylistBrowserIndex: async () => ({ ok: true, message: 'ok' }),
      movePlaylistBrowserIndex: async () => ({ ok: true, message: 'ok' }),
      onYoutubePickerResolve: async () => ({ ok: true, message: 'ok' }),
    },
    registrar,
  );

  const setHandler = handlers.handle.get(IPC_CHANNELS.request.setRuntimeOption);
  assert.ok(setHandler);
  const invalidIdResult = await setHandler!({}, '__invalid__', true);
  assert.deepEqual(invalidIdResult, { ok: false, error: 'Invalid runtime option id' });
  const invalidValueResult = await setHandler!({}, 'anki.autoUpdateNewCards', 42);
  assert.deepEqual(invalidValueResult, {
    ok: false,
    error: 'Invalid runtime option value payload',
  });
  const validResult = await setHandler!({}, 'anki.autoUpdateNewCards', true);
  assert.deepEqual(validResult, { ok: true });
  const validSubtitleAnnotationResult = await setHandler!({}, 'subtitle.annotation.jlpt', false);
  assert.deepEqual(validSubtitleAnnotationResult, { ok: true });
  assert.deepEqual(calls, [
    { id: 'anki.autoUpdateNewCards', value: true },
    { id: 'subtitle.annotation.jlpt', value: false },
  ]);

  const cycleHandler = handlers.handle.get(IPC_CHANNELS.request.cycleRuntimeOption);
  assert.ok(cycleHandler);
  const invalidDirection = await cycleHandler!({}, 'anki.kikuFieldGrouping', 2);
  assert.deepEqual(invalidDirection, {
    ok: false,
    error: 'Invalid runtime option cycle direction',
  });
  await cycleHandler!({}, 'anki.kikuFieldGrouping', -1);
  assert.deepEqual(cycles, [{ id: 'anki.kikuFieldGrouping', direction: -1 }]);

  const getPlaybackPausedHandler = handlers.handle.get(IPC_CHANNELS.request.getPlaybackPaused);
  assert.ok(getPlaybackPausedHandler);
  assert.equal(getPlaybackPausedHandler!({}), null);

  const getControllerConfigHandler = handlers.handle.get(IPC_CHANNELS.request.getControllerConfig);
  assert.ok(getControllerConfigHandler);
  assert.equal(
    (getControllerConfigHandler!({}) as { scrollPixelsPerSecond: number }).scrollPixelsPerSecond,
    960,
  );
});

test('registerIpcHandlers exposes subtitle sidebar snapshot request', async () => {
  const { registrar, handlers } = createFakeIpcRegistrar();
  const snapshot = createSubtitleSidebarSnapshotFixture();
  snapshot.cues = [{ startTime: 1, endTime: 2, text: 'line-1' }];
  snapshot.config.enabled = true;

  registerIpcHandlers(
    createRegisterIpcDeps({
      getSubtitleSidebarSnapshot: async () => snapshot,
    }),
    registrar,
  );

  const handler = handlers.handle.get(IPC_CHANNELS.request.getSubtitleSidebarSnapshot);
  assert.ok(handler);
  assert.deepEqual(await handler!({}), snapshot);
});

test('registerIpcHandlers forwards yomitan lookup tracking commands to immersion tracker', () => {
  const { registrar, handlers } = createFakeIpcRegistrar();
  const calls: string[] = [];
  registerIpcHandlers(
    createRegisterIpcDeps({
      immersionTracker: createFakeImmersionTracker({
        recordYomitanLookup: () => {
          calls.push('lookup');
        },
      }),
    }),
    registrar,
  );

  const handler = handlers.on.get(IPC_CHANNELS.command.recordYomitanLookup);
  assert.equal(typeof handler, 'function');

  handler?.({}, null);

  assert.deepEqual(calls, ['lookup']);
});

test('registerIpcHandlers returns empty stats overview shape without a tracker', async () => {
  const { registrar, handlers } = createFakeIpcRegistrar();
  registerIpcHandlers(createRegisterIpcDeps(), registrar);

  const overviewHandler = handlers.handle.get(IPC_CHANNELS.request.statsGetOverview);
  assert.ok(overviewHandler);
  assert.deepEqual(await overviewHandler!({}), {
    sessions: [],
    rollups: [],
    hints: {
      totalSessions: 0,
      activeSessions: 0,
      episodesToday: 0,
      activeAnimeCount: 0,
      totalCards: 0,
      totalActiveMin: 0,
      activeDays: 0,
      totalEpisodesWatched: 0,
      totalAnimeCompleted: 0,
      totalTokensSeen: 0,
      totalLookupCount: 0,
      totalLookupHits: 0,
      totalYomitanLookupCount: 0,
      newWordsToday: 0,
      newWordsThisWeek: 0,
    },
  });
});

test('registerIpcHandlers validates and clamps stats request limits', async () => {
  const { registrar, handlers } = createFakeIpcRegistrar();
  const calls: Array<[string, number, number?]> = [];

  registerIpcHandlers(
    createRegisterIpcDeps({
      immersionTracker: {
        recordYomitanLookup: () => {},
        getSessionSummaries: async (limit = 0) => {
          calls.push(['sessions', limit]);
          return [];
        },
        getDailyRollups: async (limit = 0) => {
          calls.push(['daily', limit]);
          return [];
        },
        getMonthlyRollups: async (limit = 0) => {
          calls.push(['monthly', limit]);
          return [];
        },
        getQueryHints: async () => ({
          totalSessions: 0,
          activeSessions: 0,
          episodesToday: 0,
          activeAnimeCount: 0,
          totalCards: 0,
          totalActiveMin: 0,
          activeDays: 0,
          totalEpisodesWatched: 0,
          totalAnimeCompleted: 0,
          totalTokensSeen: 0,
          totalLookupCount: 0,
          totalLookupHits: 0,
          totalYomitanLookupCount: 0,
          newWordsToday: 0,
          newWordsThisWeek: 0,
        }),
        getSessionTimeline: async (sessionId: number, limit = 0) => {
          calls.push(['timeline', limit, sessionId]);
          return [];
        },
        getSessionEvents: async (sessionId: number, limit = 0) => {
          calls.push(['events', limit, sessionId]);
          return [];
        },
        getVocabularyStats: async (limit = 0) => {
          calls.push(['vocabulary', limit]);
          return [];
        },
        getKanjiStats: async (limit = 0) => {
          calls.push(['kanji', limit]);
          return [];
        },
        getMediaLibrary: async () => [],
        getMediaDetail: async () => null,
        getMediaSessions: async () => [],
        getMediaDailyRollups: async () => [],
        getCoverArt: async () => null,
        markActiveVideoWatched: async () => false,
      },
    }),
    registrar,
  );

  await handlers.handle.get(IPC_CHANNELS.request.statsGetDailyRollups)!({}, -1);
  await handlers.handle.get(IPC_CHANNELS.request.statsGetMonthlyRollups)!(
    {},
    Number.POSITIVE_INFINITY,
  );
  await handlers.handle.get(IPC_CHANNELS.request.statsGetSessions)!({}, 9999);
  await handlers.handle.get(IPC_CHANNELS.request.statsGetSessionTimeline)!({}, 7, 12.5);
  await handlers.handle.get(IPC_CHANNELS.request.statsGetSessionEvents)!({}, 7, 0);
  await handlers.handle.get(IPC_CHANNELS.request.statsGetVocabulary)!({}, 1000);
  await handlers.handle.get(IPC_CHANNELS.request.statsGetKanji)!({}, NaN);

  assert.deepEqual(calls, [
    ['daily', 60],
    ['monthly', 24],
    ['sessions', 500],
    ['timeline', 200, 7],
    ['events', 500, 7],
    ['vocabulary', 500],
    ['kanji', 100],
  ]);
});

test('registerIpcHandlers requests the full timeline when no limit is provided', async () => {
  const { registrar, handlers } = createFakeIpcRegistrar();
  const calls: Array<[string, number | undefined, number]> = [];

  registerIpcHandlers(
    createRegisterIpcDeps({
      immersionTracker: {
        recordYomitanLookup: () => {},
        getSessionSummaries: async () => [],
        getDailyRollups: async () => [],
        getMonthlyRollups: async () => [],
        getQueryHints: async () => ({
          totalSessions: 0,
          activeSessions: 0,
          episodesToday: 0,
          activeAnimeCount: 0,
          totalCards: 0,
          totalActiveMin: 0,
          activeDays: 0,
          totalEpisodesWatched: 0,
          totalAnimeCompleted: 0,
          totalTokensSeen: 0,
          totalLookupCount: 0,
          totalLookupHits: 0,
          totalYomitanLookupCount: 0,
          newWordsToday: 0,
          newWordsThisWeek: 0,
        }),
        getSessionTimeline: async (sessionId: number, limit?: number) => {
          calls.push(['timeline', limit, sessionId]);
          return [];
        },
        getSessionEvents: async () => [],
        getVocabularyStats: async () => [],
        getKanjiStats: async () => [],
        getMediaLibrary: async () => [],
        getMediaDetail: async () => null,
        getMediaSessions: async () => [],
        getMediaDailyRollups: async () => [],
        getCoverArt: async () => null,
        markActiveVideoWatched: async () => false,
      },
    }),
    registrar,
  );

  await handlers.handle.get(IPC_CHANNELS.request.statsGetSessionTimeline)!({}, 7, undefined);

  assert.deepEqual(calls, [['timeline', undefined, 7]]);
});

test('registerIpcHandlers ignores malformed fire-and-forget payloads', () => {
  const { registrar, handlers } = createFakeIpcRegistrar();
  const saves: unknown[] = [];
  const controllerSaves: unknown[] = [];
  const closedModals: unknown[] = [];
  const openedModals: unknown[] = [];
  registerIpcHandlers(
    {
      onOverlayModalClosed: (modal) => {
        closedModals.push(modal);
      },
      onOverlayModalOpened: (modal) => {
        openedModals.push(modal);
      },
      openYomitanSettings: () => {},
      quitApp: () => {},
      toggleDevTools: () => {},
      getVisibleOverlayVisibility: () => false,
      toggleVisibleOverlay: () => {},
      tokenizeCurrentSubtitle: async () => null,
      getCurrentSubtitleRaw: () => '',
      getCurrentSubtitleAss: () => '',
      getSubtitleSidebarSnapshot: async () => createSubtitleSidebarSnapshotFixture(),
      getPlaybackPaused: () => false,
      getSubtitlePosition: () => null,
      getSubtitleStyle: () => null,
      saveSubtitlePosition: (position) => {
        saves.push(position);
      },
      getMecabStatus: () => ({ available: false, enabled: false, path: null }),
      setMecabEnabled: () => {},
      handleMpvCommand: () => {},
      getKeybindings: () => [],
      getSessionBindings: () => [],
      getConfiguredShortcuts: () => ({}),
      dispatchSessionAction: async () => {},
      getStatsToggleKey: () => 'Backquote',
      getMarkWatchedKey: () => 'KeyW',
      getControllerConfig: () => createControllerConfigFixture(),
      saveControllerConfig: () => {},
      saveControllerPreference: (update) => {
        controllerSaves.push(update);
      },
      getSecondarySubMode: () => 'hover',
      getCurrentSecondarySub: () => '',
      focusMainWindow: () => {},
      runSubsyncManual: async () => ({ ok: true, message: 'ok' }),
      getAnkiConnectStatus: () => false,
      getRuntimeOptions: () => [],
      setRuntimeOption: () => ({ ok: true }),
      cycleRuntimeOption: () => ({ ok: true }),
      reportOverlayContentBounds: () => {},
      getAnilistStatus: () => ({}),
      clearAnilistToken: () => {},
      openAnilistSetup: () => {},
      getAnilistQueueStatus: () => ({}),
      retryAnilistQueueNow: async () => ({ ok: true, message: 'ok' }),
      appendClipboardVideoToQueue: () => ({ ok: true, message: 'ok' }),
      getPlaylistBrowserSnapshot: async () => ({
        directoryPath: null,
        directoryAvailable: false,
        directoryStatus: '',
        directoryItems: [],
        playlistItems: [],
        playingIndex: null,
        currentFilePath: null,
      }),
      appendPlaylistBrowserFile: async () => ({ ok: true, message: 'ok' }),
      playPlaylistBrowserIndex: async () => ({ ok: true, message: 'ok' }),
      removePlaylistBrowserIndex: async () => ({ ok: true, message: 'ok' }),
      movePlaylistBrowserIndex: async () => ({ ok: true, message: 'ok' }),
      onYoutubePickerResolve: async () => ({ ok: true, message: 'ok' }),
    },
    registrar,
  );

  handlers.on.get(IPC_CHANNELS.command.saveSubtitlePosition)!({}, { yPercent: 'bad' });
  handlers.on.get(IPC_CHANNELS.command.saveSubtitlePosition)!({}, { yPercent: 42 });
  assert.deepEqual(saves, [{ yPercent: 42 }]);

  handlers.on.get(IPC_CHANNELS.command.overlayModalClosed)!({}, 'not-a-modal');
  handlers.on.get(IPC_CHANNELS.command.overlayModalClosed)!({}, 'subsync');
  handlers.on.get(IPC_CHANNELS.command.overlayModalClosed)!({}, 'kiku');
  assert.deepEqual(closedModals, ['subsync', 'kiku']);

  handlers.on.get(IPC_CHANNELS.command.overlayModalOpened)!({}, 'bad');
  handlers.on.get(IPC_CHANNELS.command.overlayModalOpened)!({}, 'subsync');
  handlers.on.get(IPC_CHANNELS.command.overlayModalOpened)!({}, 'runtime-options');
  assert.deepEqual(openedModals, ['subsync', 'runtime-options']);
});

test('registerIpcHandlers awaits saveControllerPreference through request-response IPC', async () => {
  const { registrar, handlers } = createFakeIpcRegistrar();
  const controllerSaves: unknown[] = [];
  registerIpcHandlers(
    {
      onOverlayModalClosed: () => {},
      openYomitanSettings: () => {},
      quitApp: () => {},
      toggleDevTools: () => {},
      getVisibleOverlayVisibility: () => false,
      toggleVisibleOverlay: () => {},
      tokenizeCurrentSubtitle: async () => null,
      getCurrentSubtitleRaw: () => '',
      getCurrentSubtitleAss: () => '',
      getSubtitleSidebarSnapshot: async () => createSubtitleSidebarSnapshotFixture(),
      getPlaybackPaused: () => false,
      getSubtitlePosition: () => null,
      getSubtitleStyle: () => null,
      saveSubtitlePosition: () => {},
      getMecabStatus: () => ({ available: false, enabled: false, path: null }),
      setMecabEnabled: () => {},
      handleMpvCommand: () => {},
      getKeybindings: () => [],
      getSessionBindings: () => [],
      getConfiguredShortcuts: () => ({}),
      dispatchSessionAction: async () => {},
      getStatsToggleKey: () => 'Backquote',
      getMarkWatchedKey: () => 'KeyW',
      getControllerConfig: () => createControllerConfigFixture(),
      saveControllerConfig: async () => {},
      saveControllerPreference: async (update) => {
        await Promise.resolve();
        controllerSaves.push(update);
      },
      getSecondarySubMode: () => 'hover',
      getCurrentSecondarySub: () => '',
      focusMainWindow: () => {},
      runSubsyncManual: async () => ({ ok: true, message: 'ok' }),
      getAnkiConnectStatus: () => false,
      getRuntimeOptions: () => [],
      setRuntimeOption: () => ({ ok: true }),
      cycleRuntimeOption: () => ({ ok: true }),
      reportOverlayContentBounds: () => {},
      getAnilistStatus: () => ({}),
      clearAnilistToken: () => {},
      openAnilistSetup: () => {},
      getAnilistQueueStatus: () => ({}),
      retryAnilistQueueNow: async () => ({ ok: true, message: 'ok' }),
      appendClipboardVideoToQueue: () => ({ ok: true, message: 'ok' }),
      getPlaylistBrowserSnapshot: async () => ({
        directoryPath: null,
        directoryAvailable: false,
        directoryStatus: '',
        directoryItems: [],
        playlistItems: [],
        playingIndex: null,
        currentFilePath: null,
      }),
      appendPlaylistBrowserFile: async () => ({ ok: true, message: 'ok' }),
      playPlaylistBrowserIndex: async () => ({ ok: true, message: 'ok' }),
      removePlaylistBrowserIndex: async () => ({ ok: true, message: 'ok' }),
      movePlaylistBrowserIndex: async () => ({ ok: true, message: 'ok' }),
      onYoutubePickerResolve: async () => ({ ok: true, message: 'ok' }),
    },
    registrar,
  );

  const saveHandler = handlers.handle.get(IPC_CHANNELS.command.saveControllerPreference);
  assert.ok(saveHandler);

  await assert.rejects(async () => {
    await saveHandler!({}, { preferredGamepadId: 12 });
  }, /Invalid controller preference payload/);
  await saveHandler!(
    {},
    {
      preferredGamepadId: 'pad-1',
      preferredGamepadLabel: 'Pad 1',
    },
  );

  assert.deepEqual(controllerSaves, [
    {
      preferredGamepadId: 'pad-1',
      preferredGamepadLabel: 'Pad 1',
    },
  ]);
});

test('registerIpcHandlers validates dispatchSessionAction payloads', async () => {
  const { registrar, handlers } = createFakeIpcRegistrar();
  const dispatched: SessionActionDispatchRequest[] = [];
  registerIpcHandlers(
    createRegisterIpcDeps({
      dispatchSessionAction: async (request) => {
        dispatched.push(request);
      },
    }),
    registrar,
  );

  const dispatchHandler = handlers.handle.get(IPC_CHANNELS.command.dispatchSessionAction);
  assert.ok(dispatchHandler);

  await assert.rejects(async () => {
    await dispatchHandler!({}, { actionId: 'cycleRuntimeOption', payload: { direction: 1 } });
  }, /Invalid session action payload/);
  await assert.rejects(async () => {
    await dispatchHandler!({}, { actionId: 'unknown-action' });
  }, /Invalid session action payload/);

  await dispatchHandler!(
    {},
    {
      actionId: 'copySubtitleMultiple',
      payload: { count: 3 },
    },
  );
  await dispatchHandler!(
    {},
    {
      actionId: 'cycleRuntimeOption',
      payload: {
        runtimeOptionId: 'anki.autoUpdateNewCards',
        direction: -1,
      },
    },
  );
  await dispatchHandler!(
    {},
    {
      actionId: 'toggleSubtitleSidebar',
    },
  );
  await dispatchHandler!(
    {},
    {
      actionId: 'openSessionHelp',
    },
  );
  await dispatchHandler!(
    {},
    {
      actionId: 'openControllerSelect',
    },
  );
  await dispatchHandler!(
    {},
    {
      actionId: 'openControllerDebug',
    },
  );

  assert.deepEqual(dispatched, [
    {
      actionId: 'copySubtitleMultiple',
      payload: { count: 3 },
    },
    {
      actionId: 'cycleRuntimeOption',
      payload: {
        runtimeOptionId: 'anki.autoUpdateNewCards',
        direction: -1,
      },
    },
    {
      actionId: 'toggleSubtitleSidebar',
    },
    {
      actionId: 'openSessionHelp',
    },
    {
      actionId: 'openControllerSelect',
    },
    {
      actionId: 'openControllerDebug',
    },
  ]);
});

test('registerIpcHandlers rejects malformed controller preference payloads', async () => {
  const { registrar, handlers } = createFakeIpcRegistrar();
  registerIpcHandlers(
    {
      onOverlayModalClosed: () => {},
      openYomitanSettings: () => {},
      quitApp: () => {},
      toggleDevTools: () => {},
      getVisibleOverlayVisibility: () => false,
      toggleVisibleOverlay: () => {},
      tokenizeCurrentSubtitle: async () => null,
      getCurrentSubtitleRaw: () => '',
      getCurrentSubtitleAss: () => '',
      getSubtitleSidebarSnapshot: async () => createSubtitleSidebarSnapshotFixture(),
      getPlaybackPaused: () => false,
      getSubtitlePosition: () => null,
      getSubtitleStyle: () => null,
      saveSubtitlePosition: () => {},
      getMecabStatus: () => ({ available: false, enabled: false, path: null }),
      setMecabEnabled: () => {},
      handleMpvCommand: () => {},
      getKeybindings: () => [],
      getSessionBindings: () => [],
      getConfiguredShortcuts: () => ({}),
      dispatchSessionAction: async () => {},
      getStatsToggleKey: () => 'Backquote',
      getMarkWatchedKey: () => 'KeyW',
      getControllerConfig: () => createControllerConfigFixture(),
      saveControllerConfig: async () => {},
      saveControllerPreference: async () => {},
      getSecondarySubMode: () => 'hover',
      getCurrentSecondarySub: () => '',
      focusMainWindow: () => {},
      runSubsyncManual: async () => ({ ok: true, message: 'ok' }),
      getAnkiConnectStatus: () => false,
      getRuntimeOptions: () => [],
      setRuntimeOption: () => ({ ok: true }),
      cycleRuntimeOption: () => ({ ok: true }),
      reportOverlayContentBounds: () => {},
      getAnilistStatus: () => ({}),
      clearAnilistToken: () => {},
      openAnilistSetup: () => {},
      getAnilistQueueStatus: () => ({}),
      retryAnilistQueueNow: async () => ({ ok: true, message: 'ok' }),
      appendClipboardVideoToQueue: () => ({ ok: true, message: 'ok' }),
      getPlaylistBrowserSnapshot: async () => ({
        directoryPath: null,
        directoryAvailable: false,
        directoryStatus: '',
        directoryItems: [],
        playlistItems: [],
        playingIndex: null,
        currentFilePath: null,
      }),
      appendPlaylistBrowserFile: async () => ({ ok: true, message: 'ok' }),
      playPlaylistBrowserIndex: async () => ({ ok: true, message: 'ok' }),
      removePlaylistBrowserIndex: async () => ({ ok: true, message: 'ok' }),
      movePlaylistBrowserIndex: async () => ({ ok: true, message: 'ok' }),
      onYoutubePickerResolve: async () => ({ ok: true, message: 'ok' }),
    },
    registrar,
  );

  const saveHandler = handlers.handle.get(IPC_CHANNELS.command.saveControllerPreference);
  await assert.rejects(async () => {
    await saveHandler!({}, { preferredGamepadId: 12 });
  }, /Invalid controller preference payload/);
});

test('registerIpcHandlers exposes character dictionary selection handlers', async () => {
  const { registrar, handlers } = createFakeIpcRegistrar();
  const calls: number[] = [];

  registerIpcHandlers(
    createRegisterIpcDeps({
      getCharacterDictionarySelection: async () => ({
        seriesKey: 're-zero-starting-life-in-another-world-2016',
        guessTitle: 'Re ZERO, Starting Life in Another World',
        current: { id: 10607, title: 'Rerere no Tensai Bakabon', episodes: 24 },
        override: null,
        candidates: [
          { id: 21355, title: 'Re:ZERO -Starting Life in Another World-', episodes: 25 },
        ],
      }),
      setCharacterDictionarySelection: async (mediaId) => {
        calls.push(mediaId);
        return {
          ok: true,
          seriesKey: 're-zero-starting-life-in-another-world-2016',
          selected: {
            id: mediaId,
            title: 'Re:ZERO -Starting Life in Another World-',
            episodes: 25,
          },
          staleMediaIds: [10607],
        };
      },
    }),
    registrar,
  );

  const getHandler = handlers.handle.get(IPC_CHANNELS.request.getCharacterDictionarySelection);
  const setHandler = handlers.handle.get(IPC_CHANNELS.request.setCharacterDictionarySelection);

  assert.deepEqual(await getHandler!({}), {
    seriesKey: 're-zero-starting-life-in-another-world-2016',
    guessTitle: 'Re ZERO, Starting Life in Another World',
    current: { id: 10607, title: 'Rerere no Tensai Bakabon', episodes: 24 },
    override: null,
    candidates: [{ id: 21355, title: 'Re:ZERO -Starting Life in Another World-', episodes: 25 }],
  });
  assert.deepEqual(await setHandler!({}, 0), {
    ok: false,
    message: 'Invalid AniList media ID.',
  });
  assert.deepEqual(await setHandler!({}, 21355), {
    ok: true,
    seriesKey: 're-zero-starting-life-in-another-world-2016',
    selected: { id: 21355, title: 'Re:ZERO -Starting Life in Another World-', episodes: 25 },
    staleMediaIds: [10607],
  });
  assert.deepEqual(calls, [21355]);
});
