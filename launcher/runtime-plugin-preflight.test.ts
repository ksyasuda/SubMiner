import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import {
  ensureLinuxRuntimePluginAvailable,
  installManagedPluginAssetsViaApp,
} from './runtime-plugin-preflight';

test('ensureLinuxRuntimePluginAvailable is a no-op on non-Linux platforms', async () => {
  const calls: string[] = [];

  await ensureLinuxRuntimePluginAvailable({
    platform: 'darwin',
    detectInstalledPlugin: () => {
      calls.push('detect');
      return false;
    },
    resolveRuntimePluginPath: () => {
      calls.push('resolve');
      return null;
    },
    installManagedPluginAssets: async () => {
      calls.push('install');
      return { ok: true, status: 'installed', path: '/tmp/plugin/main.lua' };
    },
    log: () => {
      calls.push('log');
    },
  });

  assert.deepEqual(calls, []);
});

test('ensureLinuxRuntimePluginAvailable skips install when installed global plugin and managed theme exist', async () => {
  const calls: string[] = [];

  await ensureLinuxRuntimePluginAvailable({
    platform: 'linux',
    detectInstalledPlugin: () => {
      calls.push('detect');
      return true;
    },
    resolveRuntimePluginPath: () => {
      calls.push('resolve');
      return null;
    },
    installManagedPluginAssets: async () => {
      calls.push('install');
      return { ok: true, status: 'installed', path: '/tmp/plugin/main.lua' };
    },
    isManagedThemeAvailable: () => {
      calls.push('theme');
      return true;
    },
    log: () => {},
  });

  assert.deepEqual(calls, ['detect', 'theme']);
});

test('ensureLinuxRuntimePluginAvailable skips install when managed runtime path and theme already resolve', async () => {
  const calls: string[] = [];

  await ensureLinuxRuntimePluginAvailable({
    platform: 'linux',
    xdgDataHome: '/tmp/xdg-data',
    detectInstalledPlugin: () => {
      calls.push('detect');
      return false;
    },
    resolveRuntimePluginPath: () => {
      calls.push('resolve');
      return '/tmp/plugin/main.lua';
    },
    installManagedPluginAssets: async () => {
      calls.push('install');
      return { ok: true, status: 'installed', path: '/tmp/plugin/main.lua' };
    },
    isManagedThemeAvailable: () => {
      calls.push('theme');
      return true;
    },
    log: () => {},
  });

  assert.deepEqual(calls, ['detect', 'resolve', 'theme']);
});

test('ensureLinuxRuntimePluginAvailable installs managed assets when rofi theme is missing', async () => {
  const calls: string[] = [];

  await ensureLinuxRuntimePluginAvailable({
    platform: 'linux',
    xdgDataHome: '/tmp/xdg-data',
    detectInstalledPlugin: () => {
      calls.push('detect');
      return false;
    },
    resolveRuntimePluginPath: () => {
      calls.push('resolve');
      return '/tmp/plugin/main.lua';
    },
    isManagedThemeAvailable: () => {
      calls.push('theme');
      return false;
    },
    installManagedPluginAssets: async () => {
      calls.push('install');
      return { ok: true, status: 'installed', path: '/tmp/plugin/main.lua' };
    },
    log: (level, _configured, message) => {
      calls.push(`${level}:${message}`);
    },
  });

  assert.deepEqual(calls, [
    'detect',
    'resolve',
    'theme',
    'info:Linux runtime support assets missing; installing managed plugin/theme assets.',
    'install',
    'info:Managed Linux runtime support assets installed: plugin=/tmp/plugin/main.lua theme=/tmp/xdg-data/SubMiner/themes/subminer.rasi',
    'resolve',
  ]);
});

test('ensureLinuxRuntimePluginAvailable installs managed assets and re-resolves plugin path', async () => {
  const calls: string[] = [];
  let resolveCount = 0;

  await ensureLinuxRuntimePluginAvailable({
    platform: 'linux',
    xdgDataHome: '/tmp/xdg-data',
    detectInstalledPlugin: () => false,
    resolveRuntimePluginPath: () => {
      resolveCount += 1;
      calls.push(`resolve:${resolveCount}`);
      return resolveCount === 1 ? null : '/tmp/plugin/main.lua';
    },
    installManagedPluginAssets: async () => {
      calls.push('install');
      return { ok: true, status: 'installed', path: '/tmp/plugin/main.lua' };
    },
    log: (level, _configured, message) => {
      calls.push(`${level}:${message}`);
    },
  });

  assert.deepEqual(calls, [
    'resolve:1',
    'info:Linux runtime support assets missing; installing managed plugin/theme assets.',
    'install',
    'info:Managed Linux runtime support assets installed: plugin=/tmp/plugin/main.lua theme=/tmp/xdg-data/SubMiner/themes/subminer.rasi',
    'resolve:2',
  ]);
});

test('ensureLinuxRuntimePluginAvailable fails when install result is not ok', async () => {
  await assert.rejects(
    () =>
      ensureLinuxRuntimePluginAvailable({
        platform: 'linux',
        detectInstalledPlugin: () => false,
        resolveRuntimePluginPath: () => null,
        installManagedPluginAssets: async () => ({
          ok: false,
          status: 'failed',
          error: 'copy failed',
        }),
        log: () => {},
      }),
    /copy failed/,
  );
});

test('ensureLinuxRuntimePluginAvailable fails when runtime path remains unresolved after install', async () => {
  await assert.rejects(
    () =>
      ensureLinuxRuntimePluginAvailable({
        platform: 'linux',
        detectInstalledPlugin: () => false,
        resolveRuntimePluginPath: () => null,
        installManagedPluginAssets: async () => ({
          ok: true,
          status: 'installed',
          path: '/tmp/plugin/main.lua',
        }),
        log: () => {},
      }),
    /managed runtime plugin assets could not be installed/i,
  );
});

test('installManagedPluginAssetsViaApp returns launch errors without waiting for a response file', async () => {
  let waited = false;

  const result = await installManagedPluginAssetsViaApp(
    {
      appPath: '/opt/SubMiner/subminer',
    },
    {
      runAppCommandCaptureOutput: () => ({
        status: 1,
        stdout: '',
        stderr: '',
        error: new Error('spawn failed'),
      }),
      waitForInstallResponse: async () => {
        waited = true;
        return null;
      },
    },
  );

  assert.deepEqual(result, {
    ok: false,
    status: 'failed',
    error: 'spawn failed',
  });
  assert.equal(waited, false);
});

test('installManagedPluginAssetsViaApp does not let temp cleanup errors mask install result', async () => {
  const originalRmSync = fs.rmSync;
  fs.rmSync = ((targetPath, options) => {
    if (String(targetPath).includes('subminer-runtime-plugin-')) {
      throw new Error('cleanup failed');
    }
    return originalRmSync(targetPath, options);
  }) as typeof fs.rmSync;

  try {
    const result = await installManagedPluginAssetsViaApp(
      {
        appPath: '/opt/SubMiner/subminer',
      },
      {
        runAppCommandCaptureOutput: () => ({
          status: 0,
          stdout: '',
          stderr: '',
        }),
        waitForInstallResponse: async () => ({
          ok: true,
          status: 'installed',
          path: '/tmp/plugin/main.lua',
        }),
      },
    );

    assert.deepEqual(result, {
      ok: true,
      status: 'installed',
      path: '/tmp/plugin/main.lua',
    });
  } finally {
    fs.rmSync = originalRmSync;
  }
});
