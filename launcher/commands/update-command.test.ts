import test from 'node:test';
import assert from 'node:assert/strict';
import { runUpdateCommand } from './update-command';
import type { LauncherCommandContext } from './context';

function makeContext(overrides: Partial<LauncherCommandContext> = {}): LauncherCommandContext {
  return {
    args: {
      update: true,
      logLevel: 'warn',
    } as LauncherCommandContext['args'],
    scriptPath: '/home/kyle/.local/bin/subminer',
    scriptName: 'subminer',
    mpvSocketPath: '/tmp/subminer.sock',
    pluginRuntimeConfig: {} as LauncherCommandContext['pluginRuntimeConfig'],
    appPath: '/home/kyle/.local/bin/SubMiner.AppImage',
    launcherJellyfinConfig: {} as LauncherCommandContext['launcherJellyfinConfig'],
    processAdapter: {
      platform: () => 'linux',
    } as LauncherCommandContext['processAdapter'],
    ...overrides,
  };
}

test('runUpdateCommand updates directly on Linux without launching Electron', async () => {
  const calls: string[] = [];

  const handled = await runUpdateCommand(makeContext(), {
    runAppCommandCaptureOutput: () => {
      throw new Error('unexpected Electron launch');
    },
    runDirectReleaseUpdate: async (request) => {
      calls.push(`direct:${request.appPath}:${request.launcherPath}:${request.channel}`);
      return {
        appImage: { status: 'updated' },
        launcher: { status: 'updated' },
        supportAssets: [{ status: 'skipped' }],
      };
    },
    readMainConfig: () => ({ updates: { channel: 'prerelease' } }),
    log: (level, _configured, message) => {
      calls.push(`${level}:${message}`);
    },
  });

  assert.equal(handled, true);
  assert.deepEqual(calls, [
    'direct:/home/kyle/.local/bin/SubMiner.AppImage:/home/kyle/.local/bin/subminer:prerelease',
    'info:AppImage update: updated',
    'info:Launcher update: updated',
    'info:Rofi theme update: skipped',
  ]);
});

test('runUpdateCommand skips Linux asset replacement when release is not newer', async () => {
  const calls: string[] = [];
  const originalFetch = globalThis.fetch;
  globalThis.fetch = (async (url: string) => {
    calls.push(`fetch:${url}`);
    if (!url.endsWith('/releases')) {
      throw new Error(`unexpected asset fetch: ${url}`);
    }
    return {
      ok: true,
      status: 200,
      json: async () => [
        {
          tag_name: 'v0.14.0',
          prerelease: false,
          draft: false,
          assets: [
            {
              name: 'SHA256SUMS.txt',
              browser_download_url: 'https://example.test/SHA256SUMS.txt',
            },
            {
              name: 'SubMiner.AppImage',
              browser_download_url: 'https://example.test/SubMiner.AppImage',
            },
          ],
        },
      ],
      text: async () => '',
      arrayBuffer: async () => new ArrayBuffer(0),
    };
  }) as typeof fetch;

  try {
    const handled = await runUpdateCommand(makeContext(), {
      runAppCommandCaptureOutput: () => {
        throw new Error('unexpected Electron launch');
      },
      readMainConfig: () => null,
      log: (level, _configured, message) => {
        calls.push(`${level}:${message}`);
      },
    });

    assert.equal(handled, true);
    assert.deepEqual(calls, [
      'fetch:https://api.github.com/repos/ksyasuda/SubMiner/releases',
      'info:AppImage update: up to date',
      'info:Launcher update: up to date',
      'info:Rofi theme update: up to date',
    ]);
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('runUpdateCommand keeps app-mediated update path on non-Linux', async () => {
  const calls: string[] = [];

  const handled = await runUpdateCommand(
    makeContext({
      processAdapter: {
        platform: () => 'darwin',
      } as LauncherCommandContext['processAdapter'],
      appPath: '/Applications/SubMiner.app/Contents/MacOS/SubMiner',
    }),
    {
      createTempDir: () => '/tmp/subminer-update-test',
      joinPath: (...parts) => parts.join('/'),
      runAppCommandCaptureOutput: (appPath, appArgs) => {
        calls.push(`app:${appPath}:${appArgs.join(' ')}`);
        return { status: 0, stdout: '', stderr: '' };
      },
      waitForUpdateResponse: async () => ({ ok: true, status: 'up-to-date' }),
      removeDir: (targetPath) => {
        calls.push(`remove:${targetPath}`);
      },
    },
  );

  assert.equal(handled, true);
  assert.deepEqual(calls, [
    'app:/Applications/SubMiner.app/Contents/MacOS/SubMiner:--update --update-launcher-path /home/kyle/.local/bin/subminer --update-response-path /tmp/subminer-update-test/response.json',
    'remove:/tmp/subminer-update-test',
  ]);
});
