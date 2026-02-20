import assert from 'node:assert/strict';
import test from 'node:test';
import { createBuildAppLifecycleRuntimeRunnerMainDepsHandler } from './startup-lifecycle-main-deps';

test('app lifecycle runtime runner main deps builder maps lifecycle callbacks', async () => {
  const calls: string[] = [];
  const deps = createBuildAppLifecycleRuntimeRunnerMainDepsHandler({
    app: {} as never,
    platform: 'darwin',
    shouldStartApp: () => true,
    parseArgs: () => ({}) as never,
    handleCliCommand: () => calls.push('handle-cli'),
    printHelp: () => calls.push('help'),
    logNoRunningInstance: () => calls.push('no-instance'),
    onReady: async () => {
      calls.push('ready');
    },
    onWillQuitCleanup: () => calls.push('cleanup'),
    shouldRestoreWindowsOnActivate: () => true,
    restoreWindowsOnActivate: () => calls.push('restore'),
    shouldQuitOnWindowAllClosed: () => false,
  })();

  assert.equal(deps.platform, 'darwin');
  assert.equal(deps.shouldStartApp({} as never), true);
  deps.handleCliCommand({} as never, 'initial');
  deps.printHelp();
  deps.logNoRunningInstance();
  await deps.onReady();
  deps.onWillQuitCleanup();
  deps.restoreWindowsOnActivate();
  assert.equal(deps.shouldRestoreWindowsOnActivate(), true);
  assert.equal(deps.shouldQuitOnWindowAllClosed(), false);
  assert.deepEqual(calls, ['handle-cli', 'help', 'no-instance', 'ready', 'cleanup', 'restore']);
});
