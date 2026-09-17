import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createHandleMpvCommandFromIpcHandler,
  createRunSubsyncManualFromIpcHandler,
} from './ipc-bridge-actions';

test('handle mpv command handler forwards command and built deps', () => {
  const calls: string[] = [];
  const deps = {
    triggerSubsyncFromConfig: () => {},
    openRuntimeOptionsPalette: () => {},
    openJimaku: () => {},
    openTsukihime: () => {},
    openYoutubeTrackPicker: () => {},
    openPlaylistBrowser: () => {},
    openAnimeBrowser: () => {},
    cycleRuntimeOption: () => ({ ok: false as const, error: 'x' }),
    showMpvOsd: () => {},
    replayCurrentSubtitle: () => {},
    playNextSubtitle: () => {},
    sendMpvCommand: () => {},
    getMpvClient: () => null,
    isMpvConnected: () => true,
    hasRuntimeOptionsManager: () => true,
  };
  const handle = createHandleMpvCommandFromIpcHandler({
    handleMpvCommandFromIpcRuntime: (command, nextDeps) => {
      calls.push(`command:${command.join(':')}`);
      assert.equal(nextDeps, deps);
    },
    buildMpvCommandDeps: () => deps,
  });

  handle(['show-text', 'hello']);
  assert.deepEqual(calls, ['command:show-text:hello']);
});

test('run subsync manual handler forwards request and result', async () => {
  const calls: string[] = [];
  const run = createRunSubsyncManualFromIpcHandler({
    runManualFromIpc: async (request: { id: string }) => {
      calls.push(`request:${request.id}`);
      return { ok: true as const };
    },
  });

  const result = await run({ id: 'job-1' });
  assert.deepEqual(result, { ok: true });
  assert.deepEqual(calls, ['request:job-1']);
});
