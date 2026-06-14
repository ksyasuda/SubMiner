import test from 'node:test';
import assert from 'node:assert/strict';
import type { CliArgs } from '../../cli/args';
import { runEnsureLinuxRuntimePluginAssetsCliCommand } from './linux-runtime-plugin-assets-cli-command';

test('runEnsureLinuxRuntimePluginAssetsCliCommand writes success response for install', async () => {
  const writes: Array<{ path: string; payload: unknown }> = [];

  await runEnsureLinuxRuntimePluginAssetsCliCommand(
    {
      ensureLinuxRuntimePluginAssets: true,
      ensureLinuxRuntimePluginAssetsResponsePath: '/tmp/subminer-plugin-response.json',
    } as CliArgs,
    {
      ensureLinuxRuntimePluginAssets: async () => ({
        ok: true,
        status: 'installed',
        path: '/home/tester/.local/share/SubMiner/plugin/subminer/main.lua',
      }),
      writeResponse: (responsePath, payload) => {
        writes.push({ path: responsePath, payload });
      },
      logWarn: () => {},
    },
  );

  assert.deepEqual(writes, [
    {
      path: '/tmp/subminer-plugin-response.json',
      payload: {
        ok: true,
        status: 'installed',
        path: '/home/tester/.local/share/SubMiner/plugin/subminer/main.lua',
      },
    },
  ]);
});

test('runEnsureLinuxRuntimePluginAssetsCliCommand writes failure response on error', async () => {
  const writes: Array<{ path: string; payload: unknown }> = [];

  await assert.rejects(
    () =>
      runEnsureLinuxRuntimePluginAssetsCliCommand(
        {
          ensureLinuxRuntimePluginAssets: true,
          ensureLinuxRuntimePluginAssetsResponsePath: '/tmp/subminer-plugin-response.json',
        } as CliArgs,
        {
          ensureLinuxRuntimePluginAssets: async () => {
            throw new Error('copy failed');
          },
          writeResponse: (responsePath, payload) => {
            writes.push({ path: responsePath, payload });
          },
          logWarn: () => {},
        },
      ),
    /copy failed/,
  );

  assert.deepEqual(writes, [
    {
      path: '/tmp/subminer-plugin-response.json',
      payload: {
        ok: false,
        status: 'failed',
        error: 'copy failed',
      },
    },
  ]);
});
