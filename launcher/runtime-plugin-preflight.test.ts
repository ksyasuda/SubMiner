import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
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

test('ensureLinuxRuntimePluginAvailable skips install when plugin, theme, and thumbnailer exist', async () => {
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
    isManagedThumbnailerAvailable: () => {
      calls.push('thumbnailer');
      return true;
    },
    log: () => {},
  });

  assert.deepEqual(calls, ['detect', 'theme', 'thumbnailer']);
});

test('ensureLinuxRuntimePluginAvailable skips install when all managed assets resolve', async () => {
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
    isManagedThumbnailerAvailable: () => {
      calls.push('thumbnailer');
      return true;
    },
    log: () => {},
  });

  assert.deepEqual(calls, ['detect', 'resolve', 'theme', 'thumbnailer']);
});

test('ensureLinuxRuntimePluginAvailable installs managed assets when rofi theme is missing', async () => {
  const calls: string[] = [];
  let themeAvailable = false;

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
      return themeAvailable;
    },
    isManagedThumbnailerAvailable: () => {
      calls.push('thumbnailer');
      return true;
    },
    installManagedPluginAssets: async () => {
      calls.push('install');
      themeAvailable = true;
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
    'info:Linux runtime support assets missing; installing managed plugin/theme/thumbnailer assets.',
    'install',
    'info:Managed Linux runtime support assets installed: plugin=/tmp/plugin/main.lua theme=/tmp/xdg-data/SubMiner/themes/subminer.rasi thumbnailer=/tmp/xdg-data/SubMiner/thumbnailers/subminer-ffmpegthumbnailer.thumbnailer',
    'resolve',
    'theme',
    'thumbnailer',
  ]);
});

test('ensureLinuxRuntimePluginAvailable installs managed assets when thumbnailer is missing', async () => {
  const calls: string[] = [];
  let thumbnailerAvailable = false;

  await ensureLinuxRuntimePluginAvailable({
    platform: 'linux',
    xdgDataHome: '/tmp/xdg-data',
    detectInstalledPlugin: () => true,
    resolveRuntimePluginPath: () => '/tmp/plugin/main.lua',
    isManagedThemeAvailable: () => true,
    isManagedThumbnailerAvailable: () => thumbnailerAvailable,
    installManagedPluginAssets: async () => {
      calls.push('install');
      thumbnailerAvailable = true;
      return { ok: true, status: 'installed', path: '/tmp/plugin/main.lua' };
    },
    log: (_level, _configured, message) => {
      calls.push(message);
    },
  });

  assert.deepEqual(calls, [
    'Linux runtime support assets missing; installing managed plugin/theme/thumbnailer assets.',
    'install',
    'Managed Linux runtime support assets installed: plugin=/tmp/plugin/main.lua theme=/tmp/xdg-data/SubMiner/themes/subminer.rasi thumbnailer=/tmp/xdg-data/SubMiner/thumbnailers/subminer-ffmpegthumbnailer.thumbnailer',
  ]);
});

test('ensureLinuxRuntimePluginAvailable retains an installed plugin after installing support assets', async () => {
  const calls: string[] = [];
  let thumbnailerAvailable = false;

  await ensureLinuxRuntimePluginAvailable({
    platform: 'linux',
    xdgDataHome: '/tmp/xdg-data',
    detectInstalledPlugin: () => true,
    resolveRuntimePluginPath: () => {
      calls.push('resolve');
      return null;
    },
    isManagedThemeAvailable: () => true,
    isManagedThumbnailerAvailable: () => thumbnailerAvailable,
    installManagedPluginAssets: async () => {
      calls.push('install');
      thumbnailerAvailable = true;
      return { ok: true, status: 'installed', path: '/tmp/plugin/main.lua' };
    },
    log: () => {},
  });

  assert.deepEqual(calls, ['install']);
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
    isManagedThemeAvailable: () => true,
    isManagedThumbnailerAvailable: () => true,
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
    'info:Linux runtime support assets missing; installing managed plugin/theme/thumbnailer assets.',
    'install',
    'info:Managed Linux runtime support assets installed: plugin=/tmp/plugin/main.lua theme=/tmp/xdg-data/SubMiner/themes/subminer.rasi thumbnailer=/tmp/xdg-data/SubMiner/thumbnailers/subminer-ffmpegthumbnailer.thumbnailer',
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

test('ensureLinuxRuntimePluginAvailable fails when thumbnailer remains missing after install', async () => {
  await assert.rejects(
    () =>
      ensureLinuxRuntimePluginAvailable({
        platform: 'linux',
        xdgDataHome: '/tmp/xdg-data',
        detectInstalledPlugin: () => true,
        resolveRuntimePluginPath: () => '/tmp/plugin/main.lua',
        isManagedThemeAvailable: () => true,
        isManagedThumbnailerAvailable: () => false,
        installManagedPluginAssets: async () => ({
          ok: true,
          status: 'installed',
          path: '/tmp/plugin/main.lua',
        }),
        log: () => {},
      }),
    /thumbnailer=.*subminer-ffmpegthumbnailer\.thumbnailer/i,
  );
});

test('ensureLinuxRuntimePluginAvailable rejects a thumbnailer directory before and after install', async () => {
  const xdgDataHome = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-thumbnailer-directory-'));
  const thumbnailerPath = path.join(
    xdgDataHome,
    'SubMiner',
    'thumbnailers',
    'subminer-ffmpegthumbnailer.thumbnailer',
  );
  fs.mkdirSync(thumbnailerPath, { recursive: true });
  const calls: string[] = [];

  try {
    await assert.rejects(
      () =>
        ensureLinuxRuntimePluginAvailable({
          platform: 'linux',
          xdgDataHome,
          detectInstalledPlugin: () => true,
          isManagedThemeAvailable: () => true,
          installManagedPluginAssets: async () => {
            calls.push('install');
            return { ok: true, status: 'installed', path: '/tmp/plugin/main.lua' };
          },
          log: () => {},
        }),
      /thumbnailer=.*subminer-ffmpegthumbnailer\.thumbnailer/i,
    );
    assert.deepEqual(calls, ['install']);
  } finally {
    fs.rmSync(xdgDataHome, { recursive: true, force: true });
  }
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
