import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import {
  handleLaunchMpvEntry,
  readConfiguredWindowsMpvLaunchConfig,
} from './main-entry-launch-mpv';
import type { WindowsMpvLaunchDeps } from './main/runtime/windows-mpv-launch';

function createTempConfigDir(): string {
  const baseDir = path.join(process.cwd(), '.tmp', 'main-entry-launch-mpv-tests');
  fs.mkdirSync(baseDir, { recursive: true });
  return fs.mkdtempSync(path.join(baseDir, 'case-'));
}

test('readConfiguredWindowsMpvLaunchConfig parses launchMode from disk config', () => {
  const configDir = createTempConfigDir();
  try {
    fs.writeFileSync(
      path.join(configDir, 'config.jsonc'),
      `{
  "mpv": {
    "executablePath": "  C:\\\\Tools\\\\mpv.exe  ",
    "launchMode": " maximized "
  }
}
`,
    );

    assert.deepEqual(readConfiguredWindowsMpvLaunchConfig(configDir), {
      executablePath: 'C:\\Tools\\mpv.exe',
      launchMode: 'maximized',
    });
  } finally {
    fs.rmSync(configDir, { recursive: true, force: true });
  }
});

test('handleLaunchMpvEntry forwards configured launchMode and executable path into Windows launcher', async () => {
  const configDir = createTempConfigDir();
  try {
    fs.writeFileSync(
      path.join(configDir, 'config.jsonc'),
      `{
  "mpv": {
    "executablePath": "C:\\\\Configured\\\\mpv.exe",
    "launchMode": "maximized"
  }
}
`,
    );

    const deps: WindowsMpvLaunchDeps = {
      getEnv: () => undefined,
      runWhere: () => ({ status: 1, stdout: '' }),
      fileExists: () => false,
      spawnDetached: async () => undefined,
      showError: () => undefined,
    };
    const calls: unknown[] = [];

    const result = await handleLaunchMpvEntry({
      argv: [
        'SubMiner.exe',
        '--launch-mpv',
        '--input-ipc-server',
        '\\\\.\\pipe\\custom-subminer-socket',
        'C:\\video.mkv',
      ],
      userDataPath: configDir,
      processExecPath: 'C:\\SubMiner\\SubMiner.exe',
      createWindowsMpvLaunchDeps: () => deps,
      launchWindowsMpv: async (
        targets,
        receivedDeps,
        extraArgs,
        binaryPath,
        pluginEntrypointPath,
        configuredMpvPath,
        launchMode,
      ) => {
        calls.push({
          targets,
          receivedDeps,
          extraArgs,
          binaryPath,
          pluginEntrypointPath,
          configuredMpvPath,
          launchMode,
        });
        return { ok: true, mpvPath: 'C:\\Configured\\mpv.exe' };
      },
      resolveBundledWindowsMpvPluginEntrypoint: () => 'C:\\SubMiner\\plugin\\main.lua',
    });

    assert.deepEqual(result, { ok: true, mpvPath: 'C:\\Configured\\mpv.exe' });
    assert.deepEqual(calls, [
      {
        targets: ['C:\\video.mkv'],
        receivedDeps: deps,
        extraArgs: ['--input-ipc-server', '\\\\.\\pipe\\custom-subminer-socket'],
        binaryPath: 'C:\\SubMiner\\SubMiner.exe',
        pluginEntrypointPath: 'C:\\SubMiner\\plugin\\main.lua',
        configuredMpvPath: 'C:\\Configured\\mpv.exe',
        launchMode: 'maximized',
      },
    ]);
  } finally {
    fs.rmSync(configDir, { recursive: true, force: true });
  }
});
