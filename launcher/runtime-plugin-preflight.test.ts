import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {
  ensureLinuxRuntimePluginAvailable,
  installManagedPluginAssetsViaApp,
} from './runtime-plugin-preflight';

type EnsureOptions = Parameters<typeof ensureLinuxRuntimePluginAvailable>[0];

/**
 * Linux preflight harness: by default every managed asset resolves, so tests override only
 * the probes that should report something missing. Counts installs and collects log messages.
 */
function makeEnsureHarness(overrides: Partial<EnsureOptions> = {}) {
  const state = { installs: 0, logs: [] as string[] };
  const options: EnsureOptions = {
    platform: 'linux',
    xdgDataHome: '/tmp/xdg-data',
    detectInstalledPlugin: () => false,
    resolveRuntimePluginPath: () => '/tmp/plugin/main.lua',
    isManagedThemeAvailable: () => true,
    isManagedThumbnailerAvailable: () => true,
    installManagedPluginAssets: async () => {
      state.installs += 1;
      return { ok: true, status: 'installed', path: '/tmp/plugin/main.lua' };
    },
    log: (_level, _configured, message) => {
      state.logs.push(message);
    },
    ...overrides,
  };
  return { state, run: () => ensureLinuxRuntimePluginAvailable(options) };
}

test('ensureLinuxRuntimePluginAvailable is a no-op on non-Linux platforms', async () => {
  const unexpected = () => {
    throw new Error('probed on a non-Linux platform');
  };
  const harness = makeEnsureHarness({
    platform: 'darwin',
    detectInstalledPlugin: unexpected,
    resolveRuntimePluginPath: unexpected,
    isManagedThemeAvailable: unexpected,
    isManagedThumbnailerAvailable: unexpected,
  });

  await harness.run();

  assert.equal(harness.state.installs, 0);
  assert.deepEqual(harness.state.logs, []);
});

const skipInstallCases: Array<{ name: string; overrides: Partial<EnsureOptions> }> = [
  {
    name: 'plugin, theme, and thumbnailer exist',
    overrides: { detectInstalledPlugin: () => true, resolveRuntimePluginPath: () => null },
  },
  { name: 'all managed assets resolve', overrides: {} },
];

for (const c of skipInstallCases) {
  test(`ensureLinuxRuntimePluginAvailable skips install when ${c.name}`, async () => {
    const harness = makeEnsureHarness(c.overrides);

    await harness.run();

    assert.equal(harness.state.installs, 0);
    assert.deepEqual(harness.state.logs, []);
  });
}

const installWhenMissingCases: Array<{ name: string; missing: 'theme' | 'thumbnailer' }> = [
  { name: 'rofi theme is missing', missing: 'theme' },
  { name: 'thumbnailer is missing', missing: 'thumbnailer' },
];

for (const c of installWhenMissingCases) {
  test(`ensureLinuxRuntimePluginAvailable installs managed assets when ${c.name}`, async () => {
    let available = false;
    const probe = {
      [c.missing === 'theme' ? 'isManagedThemeAvailable' : 'isManagedThumbnailerAvailable']: () =>
        available,
    };
    const harness = makeEnsureHarness({
      ...probe,
      installManagedPluginAssets: async () => {
        harness.state.installs += 1;
        available = true;
        return { ok: true, status: 'installed', path: '/tmp/plugin/main.lua' };
      },
    });

    await harness.run();

    assert.equal(harness.state.installs, 1);
    assert.equal(harness.state.logs.length, 2);
    assert.match(harness.state.logs[0]!, /support assets missing; installing/);
    assert.match(
      harness.state.logs[1]!,
      /installed: plugin=\/tmp\/plugin\/main\.lua theme=.*subminer\.rasi thumbnailer=.*\.thumbnailer/,
    );
  });
}

test('ensureLinuxRuntimePluginAvailable retains an installed plugin after installing support assets', async () => {
  let thumbnailerAvailable = false;
  const harness = makeEnsureHarness({
    detectInstalledPlugin: () => true,
    // An installed plugin means the bundled path is never needed, even after the install.
    resolveRuntimePluginPath: () => null,
    isManagedThumbnailerAvailable: () => thumbnailerAvailable,
    installManagedPluginAssets: async () => {
      harness.state.installs += 1;
      thumbnailerAvailable = true;
      return { ok: true, status: 'installed', path: '/tmp/plugin/main.lua' };
    },
  });

  await harness.run();

  assert.equal(harness.state.installs, 1);
});

test('ensureLinuxRuntimePluginAvailable installs managed assets and re-resolves plugin path', async () => {
  let installed = false;
  const harness = makeEnsureHarness({
    resolveRuntimePluginPath: () => (installed ? '/tmp/plugin/main.lua' : null),
    installManagedPluginAssets: async () => {
      harness.state.installs += 1;
      installed = true;
      return { ok: true, status: 'installed', path: '/tmp/plugin/main.lua' };
    },
  });

  await harness.run();

  assert.equal(harness.state.installs, 1);
});

const failureCases: Array<{ name: string; overrides: Partial<EnsureOptions>; error: RegExp }> = [
  {
    name: 'install result is not ok',
    overrides: {
      resolveRuntimePluginPath: () => null,
      installManagedPluginAssets: async () => ({
        ok: false,
        status: 'failed',
        error: 'copy failed',
      }),
    },
    error: /copy failed/,
  },
  {
    name: 'runtime path remains unresolved after install',
    overrides: { resolveRuntimePluginPath: () => null },
    error: /managed runtime plugin assets could not be installed/i,
  },
  {
    name: 'thumbnailer remains missing after install',
    overrides: { detectInstalledPlugin: () => true, isManagedThumbnailerAvailable: () => false },
    error: /thumbnailer=.*subminer-ffmpegthumbnailer\.thumbnailer/i,
  },
];

for (const c of failureCases) {
  test(`ensureLinuxRuntimePluginAvailable fails when ${c.name}`, async () => {
    await assert.rejects(() => makeEnsureHarness(c.overrides).run(), c.error);
  });
}

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
