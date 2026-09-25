import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
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
        supportAssets: [
          { status: 'updated', component: 'theme', message: 'Installed theme.' },
          {
            status: 'updated',
            component: 'thumbnailer',
            message: 'Installed rofi thumbnailer.',
          },
          { status: 'skipped', component: 'plugin', message: 'Plugin already up to date.' },
        ],
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
    'info:Support assets (theme) update: updated - Installed theme.',
    'info:Support assets (thumbnailer) update: updated - Installed rofi thumbnailer.',
    'info:Support assets (plugin) update: skipped - Plugin already up to date.',
  ]);
});

test('runUpdateCommand sends symlinked AUR installs to the package helper without network access', async () => {
  const calls: string[] = [];
  const handled = await runUpdateCommand(makeContext({ appPath: '/usr/bin/SubMiner.AppImage' }), {
    resolveRealPath: () => '/opt/SubMiner/SubMiner.AppImage',
    runDirectReleaseUpdate: async () => {
      throw new Error('must not check GitHub releases for an AUR install');
    },
    log: (level, _configured, message) => {
      calls.push(`${level}:${message}`);
    },
  });

  assert.equal(handled, true);
  assert.deepEqual(calls, [
    'warn:SubMiner is installed through subminer-bin. Update it with your AUR helper, for example: yay -S subminer-bin.',
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
      'info:Support assets update: up to date',
    ]);
  } finally {
    globalThis.fetch = originalFetch;
  }
});

test('Linux update does not replace the launcher after an AppImage hash failure', async () => {
  const workspace = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-update-order-'));
  const appImagePath = path.join(workspace, 'SubMiner.AppImage');
  const launcherPath = path.join(workspace, 'subminer');
  fs.writeFileSync(appImagePath, 'old app');
  fs.writeFileSync(launcherPath, '#!/bin/sh\n# SubMiner launcher\n');

  const fetched: string[] = [];
  const originalFetch = globalThis.fetch;
  globalThis.fetch = (async (input: string | URL | Request) => {
    const url = input instanceof Request ? input.url : String(input);
    fetched.push(url);
    if (url.endsWith('/releases')) {
      return Response.json([
        {
          tag_name: 'v999.0.0',
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
            { name: 'subminer', browser_download_url: 'https://example.test/subminer' },
          ],
        },
      ]);
    }
    if (url.endsWith('/SHA256SUMS.txt')) {
      return new Response(
        `${createHash('sha256').update('expected app').digest('hex')}  SubMiner.AppImage\n${createHash('sha256').update('new launcher').digest('hex')}  subminer\n`,
      );
    }
    if (url.endsWith('/SubMiner.AppImage')) {
      return new Response('corrupt app');
    }
    throw new Error(`launcher asset should not be fetched: ${url}`);
  }) as typeof globalThis.fetch;

  try {
    const handled = await runUpdateCommand(
      makeContext({ appPath: appImagePath, scriptPath: launcherPath }),
      { readMainConfig: () => null, log: () => {} },
    );

    assert.equal(handled, true);
    assert.equal(fs.readFileSync(appImagePath, 'utf8'), 'old app');
    assert.equal(fs.readFileSync(launcherPath, 'utf8'), '#!/bin/sh\n# SubMiner launcher\n');
    assert.equal(fetched.includes('https://example.test/subminer'), false);
  } finally {
    globalThis.fetch = originalFetch;
    fs.rmSync(workspace, { recursive: true, force: true });
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

test('managed launcher passes its wrapper to app updates, protecting signed resources', async () => {
  const previous = process.env.SUBMINER_LAUNCHER_PATH;
  process.env.SUBMINER_LAUNCHER_PATH = '/Users/tester/.local/bin/subminer';
  try {
    let forwarded: string[] = [];
    await runUpdateCommand(
      makeContext({
        processAdapter: { ...makeContext().processAdapter, platform: () => 'darwin' },
        scriptPath: '/Applications/SubMiner.app/Contents/Resources/launcher/subminer',
        appPath: '/Applications/SubMiner.app/Contents/MacOS/SubMiner',
      }),
      {
        createTempDir: () => '/tmp/subminer-update-test',
        joinPath: (...parts) => parts.join('/'),
        runAppCommandCaptureOutput: (_app, args) => {
          forwarded = args;
          return { status: 0, stdout: '', stderr: '' };
        },
        waitForUpdateResponse: async () => ({ ok: true }),
        removeDir: () => {},
      },
    );
    assert.equal(forwarded[2], '/Users/tester/.local/bin/subminer');
  } finally {
    if (previous === undefined) delete process.env.SUBMINER_LAUNCHER_PATH;
    else process.env.SUBMINER_LAUNCHER_PATH = previous;
  }
});
