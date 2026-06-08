import test from 'node:test';
import assert from 'node:assert/strict';
import { shouldFetchReleaseMetadataForPlatform } from './release-metadata-policy';
import { createUpdateService, type UpdateServiceDeps, type UpdateState } from './update-service';

function createDeps(overrides: Partial<UpdateServiceDeps> = {}) {
  let state: UpdateState = {};
  const calls: string[] = [];
  const deps: UpdateServiceDeps = {
    getConfig: () => ({
      enabled: true,
      checkIntervalHours: 24,
      notificationType: 'system',
      channel: 'stable',
    }),
    getCurrentVersion: () => '0.14.0',
    now: () => 1_000_000,
    readState: async () => state,
    writeState: async (nextState) => {
      state = nextState;
      calls.push(`state:${JSON.stringify(nextState)}`);
    },
    checkAppUpdate: async () => ({ available: false, version: '0.14.0' }),
    fetchLatestStableRelease: async () => ({
      tag_name: 'v0.14.0',
      prerelease: false,
      draft: false,
      assets: [],
    }),
    updateLauncher: async () => ({ status: 'skipped' }),
    showNoUpdateDialog: async (version) => {
      calls.push(`no-update:${version}`);
    },
    showUpdateAvailableDialog: async (version) => {
      calls.push(`available-dialog:${version}`);
      return 'close';
    },
    showUpdateFailedDialog: async (message) => {
      calls.push(`failed:${message}`);
    },
    showManualUpdateRequiredDialog: async (version) => {
      calls.push(`manual-install:${version}`);
    },
    downloadAppUpdate: async () => {
      calls.push('download');
    },
    showRestartDialog: async () => {
      calls.push('restart-dialog');
      return 'later';
    },
    quitAndInstall: () => {
      calls.push('quit-install');
    },
    notifyUpdateAvailable: async (version) => {
      calls.push(`notify:${version}`);
    },
    log: (message) => calls.push(`log:${message}`),
    ...overrides,
  };

  return {
    deps,
    calls,
    getState: () => state,
    setState: (nextState: UpdateState) => (state = nextState),
  };
}

test('manual update check shows latest-version dialog when already current', async () => {
  const { deps, calls } = createDeps();
  const service = createUpdateService(deps);

  const result = await service.checkForUpdates({ source: 'manual' });

  assert.equal(result.status, 'up-to-date');
  assert.deepEqual(calls, ['no-update:0.14.0']);
});

test('manual update check falls back to GitHub release when app metadata is unavailable', async () => {
  const { deps, calls } = createDeps({
    checkAppUpdate: async () => {
      throw new Error('latest-linux.yml missing');
    },
    fetchLatestStableRelease: async () => ({
      tag_name: 'v0.15.0',
      prerelease: false,
      draft: false,
      assets: [],
    }),
  });
  const service = createUpdateService(deps);

  const result = await service.checkForUpdates({ source: 'manual' });

  assert.equal(result.status, 'update-available');
  assert.deepEqual(calls, ['available-dialog:0.15.0']);
});

test('manual update install request skips available dialog and updates app', async () => {
  const { deps, calls } = createDeps({
    checkAppUpdate: async () => ({ available: true, version: '0.15.0' }),
    showUpdateAvailableDialog: async () => {
      throw new Error('unexpected update confirmation');
    },
    updateLauncher: async (_launcherPath, channel) => {
      calls.push(`launcher:${channel}`);
      return { status: 'skipped' };
    },
  });
  const service = createUpdateService(deps);

  const result = await service.checkForUpdates({
    source: 'manual',
    installWhenAvailable: true,
  });

  assert.equal(result.status, 'updated');
  assert.deepEqual(calls, ['download', 'launcher:stable', 'restart-dialog']);
});

test('manual update check reports available when no update asset was applied', async () => {
  const { deps, calls } = createDeps({
    checkAppUpdate: async () => ({ available: false, version: '0.14.0', canUpdate: false }),
    fetchLatestStableRelease: async () => ({
      tag_name: 'v0.15.0',
      prerelease: false,
      draft: false,
      assets: [],
    }),
    showUpdateAvailableDialog: async (version) => {
      calls.push(`available-dialog:${version}`);
      return 'update';
    },
    updateLauncher: async (_launcherPath, channel) => {
      calls.push(`launcher:${channel}`);
      return { status: 'skipped' };
    },
  });
  const service = createUpdateService(deps);

  const result = await service.checkForUpdates({ source: 'manual' });

  assert.equal(result.status, 'update-available');
  assert.deepEqual(calls, ['available-dialog:0.15.0', 'launcher:stable', 'manual-install:0.15.0']);
});

test('manual update check does not prompt restart when only launcher updates', async () => {
  const { deps, calls } = createDeps({
    checkAppUpdate: async () => ({ available: false, version: '0.14.0', canUpdate: false }),
    fetchLatestStableRelease: async () => ({
      tag_name: 'v0.15.0',
      prerelease: false,
      draft: false,
      assets: [],
    }),
    showUpdateAvailableDialog: async (version) => {
      calls.push(`available-dialog:${version}`);
      return 'update';
    },
    updateLauncher: async (_launcherPath, channel) => {
      calls.push(`launcher:${channel}`);
      return { status: 'updated' };
    },
    showRestartDialog: async () => {
      calls.push('restart-dialog');
      return 'restart';
    },
    quitAndInstall: () => {
      calls.push('quit-install');
    },
  });
  const service = createUpdateService(deps);

  const result = await service.checkForUpdates({ source: 'manual' });

  assert.equal(result.status, 'update-available');
  assert.deepEqual(calls, ['available-dialog:0.15.0', 'launcher:stable', 'manual-install:0.15.0']);
});

test('manual update check can skip release metadata after unsupported app updater', async () => {
  const { deps, calls } = createDeps({
    checkAppUpdate: async () => ({ available: false, version: '0.14.0', canUpdate: false }),
    shouldFetchReleaseMetadata: ({ appUpdate }) => appUpdate.canUpdate !== false,
    fetchLatestStableRelease: async () => {
      calls.push('fetch-release');
      return {
        tag_name: 'v0.15.0',
        prerelease: false,
        draft: false,
        assets: [],
      };
    },
  } as Partial<UpdateServiceDeps>);
  const service = createUpdateService(deps);

  const result = await service.checkForUpdates({ source: 'manual' });

  assert.equal(result.status, 'up-to-date');
  assert.deepEqual(calls, ['no-update:0.14.0']);
});

test('manual update check fetches release metadata after app metadata errors', async () => {
  const { deps, calls } = createDeps({
    checkAppUpdate: async () => {
      throw new Error('latest-mac.yml missing');
    },
    shouldFetchReleaseMetadata: ({ appUpdate }) => appUpdate.canUpdate !== false,
    fetchLatestStableRelease: async () => {
      calls.push('fetch-release');
      return {
        tag_name: 'v0.15.0',
        prerelease: false,
        draft: false,
        assets: [],
      };
    },
  } as Partial<UpdateServiceDeps>);
  const service = createUpdateService(deps);

  const result = await service.checkForUpdates({ source: 'manual' });

  assert.equal(result.status, 'update-available');
  assert.deepEqual(calls, ['fetch-release', 'available-dialog:0.15.0']);
});

test('manual update check reports non-Error failures safely', async () => {
  const { deps, calls } = createDeps({
    checkAppUpdate: async () => ({ available: true, version: '0.15.0' }),
    showUpdateAvailableDialog: async (version) => {
      calls.push(`available-dialog:${version}`);
      return 'update';
    },
    downloadAppUpdate: async () => {
      throw 'download rejected';
    },
  });
  const service = createUpdateService(deps);

  const result = await service.checkForUpdates({ source: 'manual' });

  assert.deepEqual(result, { status: 'failed', error: 'download rejected' });
  assert.deepEqual(calls, ['available-dialog:0.15.0', 'failed:download rejected']);
});

test('automatic update check skips inside configured interval', async () => {
  const { deps, calls, setState } = createDeps();
  setState({ lastAutomaticCheckAt: 1_000_000 - 60 * 60 * 1000 });
  const service = createUpdateService(deps);

  const result = await service.checkForUpdates({ source: 'automatic' });

  assert.equal(result.status, 'skipped');
  assert.deepEqual(calls, []);
});

test('automatic update check notifies once per version and records check time', async () => {
  const { deps, calls, getState } = createDeps({
    checkAppUpdate: async () => ({ available: true, version: '0.15.0' }),
  });
  const service = createUpdateService(deps);

  const first = await service.checkForUpdates({ source: 'automatic' });
  const second = await service.checkForUpdates({ source: 'automatic', force: true });

  assert.equal(first.status, 'update-available');
  assert.equal(second.status, 'update-available');
  assert.deepEqual(
    calls.filter((call) => call === 'notify:0.15.0'),
    ['notify:0.15.0'],
  );
  assert.equal(getState().lastNotifiedVersion, '0.15.0');
  assert.equal(getState().lastAutomaticCheckAt, 1_000_000);
});

test('concurrent update checks share one in-flight check', async () => {
  let checkCount = 0;
  let resolveCheck: (value: { available: boolean; version: string }) => void = () => {};
  const { deps } = createDeps({
    checkAppUpdate: () =>
      new Promise((resolve) => {
        checkCount += 1;
        resolveCheck = resolve;
      }),
  });
  const service = createUpdateService(deps);
  const first = service.checkForUpdates({ source: 'manual' });
  const second = service.checkForUpdates({ source: 'manual' });

  await Promise.resolve();
  resolveCheck({ available: false, version: '0.14.0' });
  await Promise.all([first, second]);

  assert.equal(checkCount, 1);
});

test('manual install request does not reuse in-flight manual check', async () => {
  let checkCount = 0;
  const resolveChecks: Array<(value: { available: boolean; version: string }) => void> = [];
  const { deps } = createDeps({
    checkAppUpdate: () =>
      new Promise((resolve) => {
        checkCount += 1;
        resolveChecks.push(resolve);
      }),
  });
  const service = createUpdateService(deps);
  const manualCheck = service.checkForUpdates({ source: 'manual' });
  const manualInstall = service.checkForUpdates({ source: 'manual', installWhenAvailable: true });

  await Promise.resolve();
  assert.equal(checkCount, 2);
  for (const resolve of resolveChecks) {
    resolve({ available: false, version: '0.14.0' });
  }
  await Promise.all([manualCheck, manualInstall]);
});

test('manual update check does not reuse in-flight automatic check', async () => {
  let checkCount = 0;
  const resolveChecks: Array<(value: { available: boolean; version: string }) => void> = [];
  const { deps } = createDeps({
    checkAppUpdate: () =>
      new Promise((resolve) => {
        checkCount += 1;
        resolveChecks.push(resolve);
      }),
  });
  const service = createUpdateService(deps);
  const automatic = service.checkForUpdates({ source: 'automatic', force: true });
  const manual = service.checkForUpdates({ source: 'manual' });

  await Promise.resolve();
  assert.equal(checkCount, 2);
  for (const resolve of resolveChecks) {
    resolve({ available: false, version: '0.14.0' });
  }
  await Promise.all([automatic, manual]);
});

test('manual update check passes selected GitHub release to launcher update', async () => {
  const selectedRelease = {
    tag_name: 'v0.15.0',
    prerelease: false,
    draft: false,
    assets: [],
  };
  let forwardedRelease: unknown;
  const { deps, calls } = createDeps({
    checkAppUpdate: async () => ({ available: true, version: '0.15.0' }),
    fetchLatestStableRelease: async () => selectedRelease,
    showUpdateAvailableDialog: async (version) => {
      calls.push(`available-dialog:${version}`);
      return 'update';
    },
    updateLauncher: (async (...args: unknown[]) => {
      calls.push(`launcher:${args[1]}`);
      forwardedRelease = args[2];
      return { status: 'updated' };
    }) as UpdateServiceDeps['updateLauncher'],
  });
  const service = createUpdateService(deps);

  const result = await service.checkForUpdates({ source: 'manual' });

  assert.equal(result.status, 'updated');
  assert.equal(forwardedRelease, selectedRelease);
});

test('manual prerelease update check uses prerelease release and launcher channel', async () => {
  const { deps, calls } = createDeps({
    getConfig: () => ({
      enabled: true,
      checkIntervalHours: 24,
      notificationType: 'system',
      channel: 'prerelease',
    }),
    checkAppUpdate: async () => ({ available: true, version: '0.15.0-beta.1' }),
    fetchLatestStableRelease: async (channel) => {
      calls.push(`fetch:${channel}`);
      return {
        tag_name: 'v0.15.0-beta.1',
        prerelease: true,
        draft: false,
        assets: [],
      };
    },
    showUpdateAvailableDialog: async (version) => {
      calls.push(`available-dialog:${version}`);
      return 'update';
    },
    updateLauncher: async (_launcherPath, channel) => {
      calls.push(`launcher:${channel}`);
      return { status: 'skipped' };
    },
  });
  const service = createUpdateService(deps);

  const result = await service.checkForUpdates({ source: 'manual' });

  assert.equal(result.status, 'updated');
  assert.deepEqual(calls, [
    'fetch:prerelease',
    'available-dialog:0.15.0-beta.1',
    'download',
    'launcher:prerelease',
    'restart-dialog',
  ]);
});

test('manual macOS prerelease check reports GitHub update when native updater is unsupported', async () => {
  const { deps, calls } = createDeps({
    getConfig: () => ({
      enabled: true,
      checkIntervalHours: 24,
      notificationType: 'system',
      channel: 'prerelease',
    }),
    getCurrentVersion: () => '0.15.0-beta.4',
    checkAppUpdate: async (channel) => {
      calls.push(`app:${channel}`);
      return {
        available: false,
        version: '0.15.0-beta.4',
        canUpdate: false,
      };
    },
    shouldFetchReleaseMetadata: ({ request, appUpdate }) =>
      shouldFetchReleaseMetadataForPlatform('darwin', appUpdate, request),
    fetchLatestStableRelease: async (channel) => {
      calls.push(`fetch:${channel}`);
      return {
        tag_name: 'v0.15.0-beta.5',
        prerelease: true,
        draft: false,
        assets: [],
      };
    },
    showUpdateAvailableDialog: async (version) => {
      calls.push(`available-dialog:${version}`);
      return 'update';
    },
    updateLauncher: async (_launcherPath, channel, release) => {
      calls.push(`launcher:${channel}:${release?.tag_name ?? 'none'}`);
      return { status: 'skipped' };
    },
  });
  const service = createUpdateService(deps);

  const result = await service.checkForUpdates({ source: 'manual' });

  assert.equal(result.status, 'update-available');
  assert.deepEqual(calls, [
    'app:prerelease',
    'fetch:prerelease',
    'available-dialog:0.15.0-beta.5',
    'launcher:prerelease:v0.15.0-beta.5',
    'manual-install:0.15.0-beta.5',
  ]);
});

test('manual update check keeps current prerelease builds on configured stable channel', async () => {
  const { deps, calls } = createDeps({
    getCurrentVersion: () => '0.15.0-beta.3',
    checkAppUpdate: async (channel) => {
      calls.push(`app:${channel}`);
      return { available: false, version: '0.15.0-beta.3' };
    },
    fetchLatestStableRelease: async (channel) => {
      calls.push(`fetch:${channel}`);
      return {
        tag_name: 'v0.14.0',
        prerelease: false,
        draft: false,
        assets: [],
      };
    },
    showNoUpdateDialog: async (version) => {
      calls.push(`no-update:${version}`);
    },
    showUpdateAvailableDialog: async () => {
      throw new Error('unexpected update dialog');
    },
  });
  const service = createUpdateService(deps);

  const result = await service.checkForUpdates({ source: 'manual' });

  assert.equal(result.status, 'up-to-date');
  assert.deepEqual(calls, ['app:stable', 'fetch:stable', 'no-update:0.15.0-beta.3']);
});
