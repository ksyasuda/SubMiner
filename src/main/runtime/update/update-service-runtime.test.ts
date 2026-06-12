import assert from 'node:assert/strict';
import test from 'node:test';
import { runSupportAssetUpdatesForLauncherResult } from './update-support-assets-runtime';

test('runSupportAssetUpdatesForLauncherResult logs support-asset errors and preserves launcher result', async () => {
  const warnings: string[] = [];
  const launcherResult = { status: 'updated' } as const;
  const result = await runSupportAssetUpdatesForLauncherResult({
    launcherResult,
    updateSupportAssets: async () => {
      throw new Error('archive failed');
    },
    logWarn: (message, details) => {
      warnings.push(`${message}:${details instanceof Error ? details.message : String(details)}`);
    },
  });

  assert.equal(result, launcherResult);
  assert.deepEqual(warnings, ['Support asset update failed after launcher update:archive failed']);
});
