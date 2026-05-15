import test from 'node:test';
import assert from 'node:assert/strict';
import type { CliArgs } from '../../../cli/args';
import { runUpdateCliCommand } from './update-cli-command';

test('runUpdateCliCommand writes launcher response for second-instance update handoff', async () => {
  const writes: Array<{ path: string; payload: unknown }> = [];

  await runUpdateCliCommand(
    {
      update: true,
      updateLauncherPath: '/home/kyle/.local/bin/subminer',
      updateResponsePath: '/tmp/subminer-update-response.json',
    } as CliArgs,
    'second-instance',
    {
      checkForUpdates: async (request) => {
        assert.deepEqual(request, {
          source: 'launcher',
          launcherPath: '/home/kyle/.local/bin/subminer',
        });
        return { status: 'up-to-date', version: '0.15.0' };
      },
      writeResponse: (responsePath, payload) => {
        writes.push({ path: responsePath, payload });
      },
      logWarn: () => {},
    },
  );

  assert.deepEqual(writes, [
    {
      path: '/tmp/subminer-update-response.json',
      payload: { ok: true, status: 'up-to-date', version: '0.15.0' },
    },
  ]);
});
