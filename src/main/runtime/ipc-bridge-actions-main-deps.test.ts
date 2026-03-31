import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildHandleMpvCommandFromIpcMainDepsHandler,
  createBuildRunSubsyncManualFromIpcMainDepsHandler,
} from './ipc-bridge-actions-main-deps';

test('ipc bridge action main deps builders map callbacks', async () => {
  const calls: string[] = [];

  const handleMpv = createBuildHandleMpvCommandFromIpcMainDepsHandler({
    handleMpvCommandFromIpcRuntime: (command) => calls.push(`mpv:${command.join(':')}`),
    buildMpvCommandDeps: () => ({
      triggerSubsyncFromConfig: async () => {},
      openRuntimeOptionsPalette: () => {},
      openYoutubeTrackPicker: () => {},
      openPlaylistBrowser: () => {},
      cycleRuntimeOption: () => ({ ok: false as const, error: 'x' }),
      showMpvOsd: () => {},
      replayCurrentSubtitle: () => {},
      playNextSubtitle: () => {},
      shiftSubDelayToAdjacentSubtitle: async () => {},
      sendMpvCommand: () => {},
      getMpvClient: () => null,
      isMpvConnected: () => true,
      hasRuntimeOptionsManager: () => true,
    }),
  })();
  handleMpv.handleMpvCommandFromIpcRuntime(['show-text', 'hello'], handleMpv.buildMpvCommandDeps());
  assert.equal(handleMpv.buildMpvCommandDeps().isMpvConnected(), true);

  const runSubsync = createBuildRunSubsyncManualFromIpcMainDepsHandler({
    runManualFromIpc: async (request: { id: string }) => {
      calls.push(`subsync:${request.id}`);
      return { ok: true as const };
    },
  })();
  assert.deepEqual(await runSubsync.runManualFromIpc({ id: 'job-1' }), { ok: true });

  assert.deepEqual(calls, ['mpv:show-text:hello', 'subsync:job-1']);
});
