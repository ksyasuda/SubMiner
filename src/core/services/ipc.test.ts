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

test('createIpcDepsRuntime wires AniList handlers', async () => {
  const calls: string[] = [];
  const deps = createIpcDepsRuntime({
    getInvisibleWindow: () => null,
    getMainWindow: () => null,
    getVisibleOverlayVisibility: () => false,
    getInvisibleOverlayVisibility: () => false,
    onOverlayModalClosed: () => {},
    openYomitanSettings: () => {},
    quitApp: () => {},
    toggleVisibleOverlay: () => {},
    tokenizeCurrentSubtitle: async () => null,
    getCurrentSubtitleRaw: () => '',
    getCurrentSubtitleAss: () => '',
    getMpvSubtitleRenderMetrics: () => null,
    getSubtitlePosition: () => null,
    getSubtitleStyle: () => null,
    saveSubtitlePosition: () => {},
    getMecabTokenizer: () => null,
    handleMpvCommand: () => {},
    getKeybindings: () => [],
    getConfiguredShortcuts: () => ({}),
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
});

test('registerIpcHandlers rejects malformed runtime-option payloads', async () => {
  const { registrar, handlers } = createFakeIpcRegistrar();
  const calls: Array<{ id: string; value: unknown }> = [];
  const cycles: Array<{ id: string; direction: 1 | -1 }> = [];
  registerIpcHandlers(
    {
      getInvisibleWindow: () => null,
      isVisibleOverlayVisible: () => false,
      setInvisibleIgnoreMouseEvents: () => {},
      onOverlayModalClosed: () => {},
      openYomitanSettings: () => {},
      quitApp: () => {},
      toggleDevTools: () => {},
      getVisibleOverlayVisibility: () => false,
      toggleVisibleOverlay: () => {},
      getInvisibleOverlayVisibility: () => false,
      tokenizeCurrentSubtitle: async () => null,
      getCurrentSubtitleRaw: () => '',
      getCurrentSubtitleAss: () => '',
      getMpvSubtitleRenderMetrics: () => null,
      getSubtitlePosition: () => null,
      getSubtitleStyle: () => null,
      saveSubtitlePosition: () => {},
      getMecabStatus: () => ({ available: false, enabled: false, path: null }),
      setMecabEnabled: () => {},
      handleMpvCommand: () => {},
      getKeybindings: () => [],
      getConfiguredShortcuts: () => ({}),
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
  assert.deepEqual(calls, [{ id: 'anki.autoUpdateNewCards', value: true }]);

  const cycleHandler = handlers.handle.get(IPC_CHANNELS.request.cycleRuntimeOption);
  assert.ok(cycleHandler);
  const invalidDirection = await cycleHandler!({}, 'anki.kikuFieldGrouping', 2);
  assert.deepEqual(invalidDirection, {
    ok: false,
    error: 'Invalid runtime option cycle direction',
  });
  await cycleHandler!({}, 'anki.kikuFieldGrouping', -1);
  assert.deepEqual(cycles, [{ id: 'anki.kikuFieldGrouping', direction: -1 }]);
});

test('registerIpcHandlers ignores malformed fire-and-forget payloads', () => {
  const { registrar, handlers } = createFakeIpcRegistrar();
  const saves: unknown[] = [];
  const modals: unknown[] = [];
  registerIpcHandlers(
    {
      getInvisibleWindow: () => null,
      isVisibleOverlayVisible: () => false,
      setInvisibleIgnoreMouseEvents: () => {},
      onOverlayModalClosed: (modal) => {
        modals.push(modal);
      },
      openYomitanSettings: () => {},
      quitApp: () => {},
      toggleDevTools: () => {},
      getVisibleOverlayVisibility: () => false,
      toggleVisibleOverlay: () => {},
      getInvisibleOverlayVisibility: () => false,
      tokenizeCurrentSubtitle: async () => null,
      getCurrentSubtitleRaw: () => '',
      getCurrentSubtitleAss: () => '',
      getMpvSubtitleRenderMetrics: () => null,
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
  assert.deepEqual(saves, [
    { yPercent: 42, invisibleOffsetXPx: undefined, invisibleOffsetYPx: undefined },
  ]);

  handlers.on.get(IPC_CHANNELS.command.overlayModalClosed)!({}, 'not-a-modal');
  handlers.on.get(IPC_CHANNELS.command.overlayModalClosed)!({}, 'subsync');
  assert.deepEqual(modals, ['subsync']);
});
