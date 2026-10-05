import test from 'node:test';
import assert from 'node:assert/strict';

import { createIpcDepsRuntime, registerIpcHandlers, type IpcServiceDeps } from './ipc';
import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import type { SessionActionDispatchRequest, SubtitleSidebarSnapshot } from '../../types';

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
    profiles: {},
  };
}

function createSubtitleSidebarSnapshotFixture(): SubtitleSidebarSnapshot {
  return {
    sourceKey: 'test-subtitles',
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
    getOverlayNotificationPosition: () => 'top-right',
    getControllerConfig: () => createControllerConfigFixture(),
    saveControllerConfig: async () => {},
    saveControllerPreference: async () => {},
    getSecondarySubMode: () => 'hover',
    getCurrentSecondarySub: () => '',
    focusMainWindow: () => {},
    activatePlaybackWindowForOverlayInteraction: () => false,
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
    markActiveVideoWatched: async () => false,
    ...overrides,
  };
}

test('createIpcDepsRuntime ignores overlay content reports from stale visible renderers', () => {
  const mainWindow = { id: 'main', isDestroyed: () => false } as never;
  const staleWindow = { id: 'stale', isDestroyed: () => false } as never;
  const reports: unknown[] = [];
  const deps = createIpcDepsRuntime({
    getMainWindow: () => mainWindow,
    reportOverlayContentBounds: (payload: unknown) => {
      reports.push(payload);
    },
  } as unknown as Parameters<typeof createIpcDepsRuntime>[0]);

  const report = deps.reportOverlayContentBounds as (
    payload: unknown,
    senderWindow: unknown,
  ) => void;
  report({ source: 'stale' }, staleWindow);
  report({ source: 'main' }, mainWindow);
  report({ source: 'missing' }, null);

  assert.deepEqual(reports, [{ source: 'main' }]);
});

test('registerIpcHandlers maps setIgnoreMouseEvents to overlay interaction active state', () => {
  const { registrar, handlers } = createFakeIpcRegistrar();
  const calls: string[] = [];

  registerIpcHandlers(
    createRegisterIpcDeps({
      onOverlayMouseInteractionChanged: (active) => {
        calls.push(`overlay-interaction:${active}`);
      },
    }),
    registrar,
  );

  const handler = handlers.on.get(IPC_CHANNELS.command.setIgnoreMouseEvents);
  assert.equal(typeof handler, 'function');

  handler?.({}, true, { forward: true });
  handler?.({}, false, {});

  assert.deepEqual(calls, ['overlay-interaction:false', 'overlay-interaction:true']);
});

test('registerIpcHandlers runs AniList update after manual mark watched succeeds', async () => {
  const { registrar, handlers } = createFakeIpcRegistrar();
  const calls: string[] = [];
  registerIpcHandlers(
    createRegisterIpcDeps({
      immersionTracker: createFakeImmersionTracker({
        markActiveVideoWatched: async () => {
          calls.push('mark');
          return true;
        },
      }),
      runAnilistPostWatchUpdateOnManualMark: async () => {
        calls.push('anilist');
      },
    }),
    registrar,
  );

  const result = await handlers.handle.get(IPC_CHANNELS.command.markActiveVideoWatched)?.({});

  assert.equal(result, true);
  assert.deepEqual(calls, ['mark', 'anilist']);
});

test('registerIpcHandlers isolates AniList update failures after manual mark watched succeeds', async () => {
  const { registrar, handlers } = createFakeIpcRegistrar();
  const calls: string[] = [];
  const originalWarn = console.warn;
  console.warn = () => undefined;

  try {
    registerIpcHandlers(
      createRegisterIpcDeps({
        immersionTracker: createFakeImmersionTracker({
          markActiveVideoWatched: async () => {
            calls.push('mark');
            return true;
          },
        }),
        runAnilistPostWatchUpdateOnManualMark: async () => {
          calls.push('anilist');
          throw new Error('post-watch failed');
        },
      }),
      registrar,
    );

    const result = await handlers.handle.get(IPC_CHANNELS.command.markActiveVideoWatched)?.({});

    assert.equal(result, true);
    assert.deepEqual(calls, ['mark', 'anilist']);
  } finally {
    console.warn = originalWarn;
  }
});

test('registerIpcHandlers skips AniList update when manual mark watched has no active session', async () => {
  const { registrar, handlers } = createFakeIpcRegistrar();
  const calls: string[] = [];
  registerIpcHandlers(
    createRegisterIpcDeps({
      immersionTracker: createFakeImmersionTracker({
        markActiveVideoWatched: async () => {
          calls.push('mark');
          return false;
        },
      }),
      runAnilistPostWatchUpdateOnManualMark: async () => {
        calls.push('anilist');
      },
    }),
    registrar,
  );

  const result = await handlers.handle.get(IPC_CHANNELS.command.markActiveVideoWatched)?.({});

  assert.equal(result, false);
  assert.deepEqual(calls, ['mark']);
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
      getOverlayNotificationPosition: () => 'top-right',
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

test('registerIpcHandlers accepts the keep-without-media timing decision', async () => {
  const { registrar, handlers } = createFakeIpcRegistrar();
  const requests: unknown[] = [];
  registerIpcHandlers(
    createRegisterIpcDeps({
      resolveMediaTimingReview: async (request) => {
        requests.push(request);
        return { ok: true };
      },
    }),
    registrar,
  );

  const handler = handlers.handle.get(IPC_CHANNELS.request.mediaTimingReviewResolve);
  assert.ok(handler);
  assert.deepEqual(
    await handler!({}, { reviewId: 'review-1', decision: { action: 'skip-media' } }),
    { ok: true },
  );
  assert.deepEqual(requests, [{ reviewId: 'review-1', decision: { action: 'skip-media' } }]);
});

test('frame IPC validates timestamps and directions', async () => {
  const { registrar, handlers } = createFakeIpcRegistrar();
  const requests: unknown[] = [];
  registerIpcHandlers(
    createRegisterIpcDeps({
      getMediaTimingReviewFrame: async (request) => {
        requests.push(request);
        return { ok: true };
      },
    }),
    registrar,
  );
  const frame = handlers.handle.get(IPC_CHANNELS.request.mediaTimingReviewFrame)!;
  const valid = { reviewId: 'r', timestamp: 13, direction: 1 };
  assert.deepEqual(await frame({}, valid), { ok: true });
  for (const invalid of [
    null,
    {},
    { ...valid, timestamp: NaN },
    { ...valid, timestamp: '13' },
    { ...valid, direction: 2 },
  ]) {
    assert.equal(((await frame({}, invalid)) as { ok: boolean }).ok, false);
  }
  assert.deepEqual(requests, [valid]);
});

test('registerIpcHandlers validates and forwards timing review text and screenshot selection', async () => {
  const { registrar, handlers } = createFakeIpcRegistrar();
  const requests: unknown[] = [];
  registerIpcHandlers(
    createRegisterIpcDeps({
      resolveMediaTimingReview: async (request) => {
        requests.push(request);
        return { ok: true };
      },
    }),
    registrar,
  );

  const handler = handlers.handle.get(IPC_CHANNELS.request.mediaTimingReviewResolve);
  assert.ok(handler);
  assert.deepEqual(
    await handler!(
      {},
      {
        reviewId: 'review-1',
        decision: {
          action: 'confirm',
          startTime: 10,
          endTime: 12,
          text: '前の行 対象の行',
          screenshotTime: 13,
        },
      },
    ),
    { ok: true },
  );
  assert.deepEqual(requests, [
    {
      reviewId: 'review-1',
      decision: {
        action: 'confirm',
        startTime: 10,
        endTime: 12,
        text: '前の行 対象の行',
        screenshotTime: 13,
      },
    },
  ]);

  for (const invalid of [{ text: '   ' }, { screenshotTime: Infinity }]) {
    assert.deepEqual(
      await handler!(
        {},
        {
          reviewId: 'review-1',
          decision: { action: 'confirm', startTime: 10, endTime: 12, ...invalid },
        },
      ),
      { ok: false, message: 'Timing review is unavailable.' },
    );
  }
  assert.equal(requests.length, 1);
});

test('registerIpcHandlers forwards valid subtitle sidebar mining context', () => {
  const { registrar, handlers } = createFakeIpcRegistrar();
  const contexts: unknown[] = [];
  const deps = createRegisterIpcDeps() as IpcServiceDeps & {
    recordSubtitleMiningContext: (context: unknown | null) => void;
  };
  deps.recordSubtitleMiningContext = (context) => {
    contexts.push(context);
  };

  registerIpcHandlers(deps, registrar);

  const handler = handlers.on.get(IPC_CHANNELS.command.recordYomitanLookup);
  assert.equal(typeof handler, 'function');

  handler?.(
    {},
    {
      source: 'subtitle-sidebar',
      text: 'sidebar previous line',
      startTime: 10,
      endTime: 12,
      capturedAtMs: 123,
    },
  );

  assert.deepEqual(contexts, [
    {
      source: 'subtitle-sidebar',
      text: 'sidebar previous line',
      startTime: 10,
      endTime: 12,
      capturedAtMs: 123,
    },
  ]);
});

test('registerIpcHandlers records yomitan lookup when subtitle context recording fails', () => {
  const { registrar, handlers } = createFakeIpcRegistrar();
  const calls: string[] = [];
  const warnings: unknown[][] = [];
  const originalWarn = console.warn;
  console.warn = (...args: unknown[]) => {
    warnings.push(args);
  };
  const deps = createRegisterIpcDeps({
    immersionTracker: createFakeImmersionTracker({
      recordYomitanLookup: () => {
        calls.push('lookup');
      },
    }),
  }) as IpcServiceDeps & {
    recordSubtitleMiningContext: (context: unknown | null) => void;
  };
  deps.recordSubtitleMiningContext = () => {
    throw new Error('context write failed');
  };

  try {
    registerIpcHandlers(deps, registrar);

    const handler = handlers.on.get(IPC_CHANNELS.command.recordYomitanLookup);
    assert.equal(typeof handler, 'function');

    assert.doesNotThrow(() => {
      handler?.({}, { source: 'subtitle-sidebar', text: 'line', startTime: 1, endTime: 2 });
    });

    assert.deepEqual(calls, ['lookup']);
    assert.equal(warnings.length, 1);
    assert.equal(warnings[0]?.[0], 'Failed to record subtitle mining context:');
    assert.equal(warnings[0]?.[1], 'context write failed');
  } finally {
    console.warn = originalWarn;
  }
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
      getOverlayNotificationPosition: () => 'top-right',
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
      getOverlayNotificationPosition: () => 'top-right',
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

test('registerIpcHandlers accepts per-controller profile config updates', async () => {
  const { registrar, handlers } = createFakeIpcRegistrar();
  const controllerSaves: unknown[] = [];
  registerIpcHandlers(
    createRegisterIpcDeps({
      saveControllerConfig: async (update) => {
        controllerSaves.push(update);
      },
    }),
    registrar,
  );

  const saveHandler = handlers.handle.get(IPC_CHANNELS.command.saveControllerConfig);
  assert.ok(saveHandler);

  const update = {
    profiles: {
      'pad-1': {
        label: 'Pad One',
        buttonIndices: {
          buttonSouth: 11,
        },
        bindings: {
          toggleLookup: { kind: 'button', buttonIndex: 11 },
          leftStickHorizontal: { kind: 'axis', axisIndex: 6, dpadFallback: 'horizontal' },
        },
      },
    },
  };
  await saveHandler({}, update);
  assert.deepEqual(controllerSaves, [
    {
      ...update,
      // Validation uses a null prototype to safely store arbitrary profile IDs.
      profiles: { __proto__: null, ...update.profiles },
    },
  ]);

  await assert.rejects(async () => {
    await saveHandler(
      {},
      {
        profiles: {
          'pad-1': {
            bindings: {
              toggleLookup: { kind: 'axis', axisIndex: 0 },
            },
          },
        },
      },
    );
  }, /Invalid controller config payload/);

  await assert.rejects(async () => {
    await saveHandler({}, JSON.parse('{"profiles":{"__proto__":{"label":"polluted"}}}'));
  }, /Invalid controller config payload/);
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

test('registerIpcHandlers forwards valid overlay notification actions', () => {
  const { registrar, handlers } = createFakeIpcRegistrar();
  const actions: Array<{ notificationId: string; actionId: string; noteId?: number }> = [];
  registerIpcHandlers(
    createRegisterIpcDeps({
      handleOverlayNotificationAction: ((
        notificationId: string,
        actionId: string,
        noteId?: number,
      ) => {
        actions.push({ notificationId, actionId, noteId });
      }) as IpcServiceDeps['handleOverlayNotificationAction'],
    } as Partial<IpcServiceDeps>),
    registrar,
  );

  const actionHandler = handlers.on.get(IPC_CHANNELS.command.overlayNotificationAction);
  assert.ok(actionHandler);

  actionHandler({}, null);
  actionHandler({}, { notificationId: '', actionId: 'install-update' });
  actionHandler({}, { notificationId: 'subminer-update-available', actionId: 42 });
  actionHandler(
    {},
    { notificationId: 'anki-update-progress', actionId: 'open-anki-card', noteId: -1 },
  );
  actionHandler({}, { notificationId: 'subminer-update-available', actionId: 'install-update' });
  actionHandler(
    {},
    { notificationId: 'anki-update-progress', actionId: 'open-anki-card', noteId: 42 },
  );

  assert.deepEqual(actions, [
    { notificationId: 'subminer-update-available', actionId: 'install-update', noteId: undefined },
    { notificationId: 'anki-update-progress', actionId: 'open-anki-card', noteId: 42 },
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
      getOverlayNotificationPosition: () => 'top-right',
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
  const searches: Array<string | undefined> = [];

  registerIpcHandlers(
    createRegisterIpcDeps({
      getCharacterDictionarySelection: async (searchTitle) => {
        searches.push(searchTitle);
        return {
          seriesKey: 're-zero-starting-life-in-another-world-2016',
          guessTitle: 'Re ZERO, Starting Life in Another World',
          current: { id: 10607, title: 'Rerere no Tensai Bakabon', episodes: 24 },
          override: null,
          candidates: [
            { id: 21355, title: 'Re:ZERO -Starting Life in Another World-', episodes: 25 },
          ],
        };
      },
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

  assert.deepEqual(await getHandler!({}, '  Re:ZERO  '), {
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
  assert.deepEqual(searches, ['Re:ZERO']);
});

test('mpv discovery has its own request and does not change session bindings', async () => {
  const { registrar, handlers } = createFakeIpcRegistrar();
  const snapshot = { keys: ['r'], blockedKeys: [] };
  registerIpcHandlers(
    createRegisterIpcDeps({ getMpvInputBindings: async () => snapshot }),
    registrar,
  );
  const discovery = handlers.handle.get(IPC_CHANNELS.request.getMpvInputBindings);
  const session = handlers.handle.get(IPC_CHANNELS.request.getSessionBindings);
  assert.ok(discovery);
  assert.ok(session);
  assert.deepEqual(await discovery({}), snapshot);
  assert.deepEqual(await session({}), []);
});

test('registerIpcHandlers validates playlist browser mutation payloads before forwarding', async () => {
  const { registrar, handlers } = createFakeIpcRegistrar();
  const forwarded: unknown[][] = [];
  const ok = { ok: true, message: 'ok' };
  registerIpcHandlers(
    createRegisterIpcDeps({
      appendPlaylistBrowserFile: async (...args) => (forwarded.push(args), ok),
      playPlaylistBrowserIndex: async (...args) => (forwarded.push(args), ok),
      removePlaylistBrowserIndex: async (...args) => (forwarded.push(args), ok),
      movePlaylistBrowserIndex: async (...args) => (forwarded.push(args), ok),
    }),
    registrar,
  );
  const { request } = IPC_CHANNELS;
  const cases: Array<{ channel: string; args: unknown[]; valid: boolean }> = [
    { channel: request.appendPlaylistBrowserFile, args: ['/media/ep01.mkv'], valid: true },
    { channel: request.appendPlaylistBrowserFile, args: ['   '], valid: false },
    { channel: request.appendPlaylistBrowserFile, args: [42], valid: false },
    { channel: request.playPlaylistBrowserIndex, args: [0], valid: true },
    { channel: request.playPlaylistBrowserIndex, args: [-1], valid: false },
    { channel: request.removePlaylistBrowserIndex, args: [1.5], valid: false },
    { channel: request.removePlaylistBrowserIndex, args: ['2'], valid: false },
    { channel: request.movePlaylistBrowserIndex, args: [2, -1], valid: true },
    { channel: request.movePlaylistBrowserIndex, args: [2, 2], valid: false },
    { channel: request.movePlaylistBrowserIndex, args: [Number.NaN, 1], valid: false },
  ];

  for (const { channel, args, valid } of cases) {
    forwarded.length = 0;
    const result = (await handlers.handle.get(channel)?.({}, ...args)) as { ok: boolean };
    const label = `${channel}(${args.map(String).join(', ')})`;
    assert.equal(result.ok, valid, label);
    assert.deepEqual(forwarded, valid ? [args] : [], label);
  }
});
