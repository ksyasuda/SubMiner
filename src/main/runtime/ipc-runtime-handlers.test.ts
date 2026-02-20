import assert from 'node:assert/strict';
import test from 'node:test';
import { createIpcRuntimeHandlers } from './ipc-runtime-handlers';

test('ipc runtime handlers wire command and subsync handlers through built deps', async () => {
  let receivedCommand: (string | number)[] | null = null;
  let receivedCommandDeps: { tag: string } | null = null;
  let buildMpvCommandDepsCalls = 0;
  let receivedSubsyncRequest: { id: string } | null = null;

  const runtime = createIpcRuntimeHandlers({
    handleMpvCommandFromIpcDeps: {
      handleMpvCommandFromIpcRuntime: (command, deps) => {
        receivedCommand = command;
        receivedCommandDeps = deps as unknown as { tag: string };
      },
      buildMpvCommandDeps: () => {
        buildMpvCommandDepsCalls += 1;
        return { tag: 'mpv-deps' } as never;
      },
    },
    runSubsyncManualFromIpcDeps: {
      runManualFromIpc: async (request: { id: string }) => {
        receivedSubsyncRequest = request;
        return { ok: true, id: request.id };
      },
    },
  });

  runtime.handleMpvCommandFromIpc(['set_property', 'pause', 'yes']);
  assert.deepEqual(receivedCommand, ['set_property', 'pause', 'yes']);
  assert.deepEqual(receivedCommandDeps, { tag: 'mpv-deps' });
  assert.equal(buildMpvCommandDepsCalls, 1);

  const response = await runtime.runSubsyncManualFromIpc({ id: 'abc' });
  assert.deepEqual(receivedSubsyncRequest, { id: 'abc' });
  assert.deepEqual(response, { ok: true, id: 'abc' });
});
