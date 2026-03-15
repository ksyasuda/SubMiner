import test from 'node:test';
import assert from 'node:assert/strict';

import { createIpcDepsRuntime, registerIpcHandlers } from './ipc';
import { IPC_CHANNELS } from '../../shared/ipc/contracts';

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
      leftStickHorizontal: { kind: 'axis' as const, axisIndex: 0, dpadFallback: 'horizontal' as const },
      leftStickVertical: { kind: 'axis' as const, axisIndex: 1, dpadFallback: 'vertical' as const },
      rightStickHorizontal: { kind: 'axis' as const, axisIndex: 3, dpadFallback: 'none' as const },
      rightStickVertical: { kind: 'axis' as const, axisIndex: 4, dpadFallback: 'none' as const },
    },
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
    getPlaybackPaused: () => true,
    getSubtitlePosition: () => null,
    getSubtitleStyle: () => null,
    saveSubtitlePosition: () => {},
    getMecabTokenizer: () => null,
    handleMpvCommand: () => {},
    getKeybindings: () => [],
    getConfiguredShortcuts: () => ({}),
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
  assert.deepEqual(calls, ['clearAnilistToken', 'openAnilistSetup', 'retryAnilistQueueNow']);
  assert.equal(deps.getPlaybackPaused(), true);
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
      getConfiguredShortcuts: () => ({}),
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
      reportOverlayContentBounds: () => {},
      getAnilistStatus: () => ({}),
      clearAnilistToken: () => {},
      openAnilistSetup: () => {},
      getAnilistQueueStatus: () => ({}),
      retryAnilistQueueNow: async () => ({ ok: true, message: 'ok' }),
      appendClipboardVideoToQueue: () => ({ ok: true, message: 'ok' }),
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
      getConfiguredShortcuts: () => ({}),
      getControllerConfig: () => createControllerConfigFixture(),
      saveControllerConfig: (update) => {
        controllerSaves.push(update);
      },
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
      getPlaybackPaused: () => false,
      getSubtitlePosition: () => null,
      getSubtitleStyle: () => null,
      saveSubtitlePosition: () => {},
      getMecabStatus: () => ({ available: false, enabled: false, path: null }),
      setMecabEnabled: () => {},
      handleMpvCommand: () => {},
      getKeybindings: () => [],
      getConfiguredShortcuts: () => ({}),
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

test('registerIpcHandlers awaits saveControllerConfig through request-response IPC', async () => {
  const { registrar, handlers } = createFakeIpcRegistrar();
  const controllerConfigSaves: unknown[] = [];
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
      getPlaybackPaused: () => false,
      getSubtitlePosition: () => null,
      getSubtitleStyle: () => null,
      saveSubtitlePosition: () => {},
      getMecabStatus: () => ({ available: false, enabled: false, path: null }),
      setMecabEnabled: () => {},
      handleMpvCommand: () => {},
      getKeybindings: () => [],
      getConfiguredShortcuts: () => ({}),
      getControllerConfig: () => createControllerConfigFixture(),
      saveControllerConfig: async (update) => {
        await Promise.resolve();
        controllerConfigSaves.push(update);
      },
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
    },
    registrar,
  );

  const saveHandler = handlers.handle.get(IPC_CHANNELS.command.saveControllerConfig);
  assert.ok(saveHandler);

  await assert.rejects(
    async () => {
      await saveHandler!({}, { bindings: { toggleLookup: { kind: 'button', buttonIndex: -1 } } });
    },
    /Invalid controller config payload/,
  );

  await saveHandler!({}, {
    preferredGamepadId: 'pad-2',
    bindings: {
      toggleLookup: { kind: 'button', buttonIndex: 11 },
      closeLookup: { kind: 'axis', axisIndex: 4, direction: 'negative' },
      leftStickHorizontal: { kind: 'axis', axisIndex: 7, dpadFallback: 'none' },
    },
  });

  assert.deepEqual(controllerConfigSaves, [
    {
      preferredGamepadId: 'pad-2',
      bindings: {
        toggleLookup: { kind: 'button', buttonIndex: 11 },
        closeLookup: { kind: 'axis', axisIndex: 4, direction: 'negative' },
        leftStickHorizontal: { kind: 'axis', axisIndex: 7, dpadFallback: 'none' },
      },
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
      getPlaybackPaused: () => false,
      getSubtitlePosition: () => null,
      getSubtitleStyle: () => null,
      saveSubtitlePosition: () => {},
      getMecabStatus: () => ({ available: false, enabled: false, path: null }),
      setMecabEnabled: () => {},
      handleMpvCommand: () => {},
      getKeybindings: () => [],
      getConfiguredShortcuts: () => ({}),
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
    },
    registrar,
  );

  const saveHandler = handlers.handle.get(IPC_CHANNELS.command.saveControllerPreference);
  await assert.rejects(async () => {
    await saveHandler!({}, { preferredGamepadId: 12 });
  }, /Invalid controller preference payload/);
});
