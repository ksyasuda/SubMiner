import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import net from 'node:net';
import os from 'node:os';
import path from 'node:path';
import { spawn, spawnSync } from 'node:child_process';
import {
  createDefaultSetupState,
  getDefaultConfigDir,
  getSetupStatePath,
  readSetupState,
  writeSetupState,
} from '../src/shared/setup-state.js';

type RunResult = {
  status: number | null;
  stdout: string;
  stderr: string;
};

type SmokeCase = {
  root: string;
  artifactsDir: string;
  binDir: string;
  xdgConfigHome: string;
  xdgDataHome: string;
  appDataDir: string;
  localAppDataDir: string;
  homeDir: string;
  socketDir: string;
  socketPath: string;
  videoPath: string;
  fakeAppPath: string;
  fakeMpvPath: string;
  mpvOverlayLogPath: string;
};

const LAUNCHER_RUN_TIMEOUT_MS = 25000;
const LONG_SMOKE_TEST_TIMEOUT_MS = 30000;

function writeExecutable(filePath: string, body: string): void {
  fs.writeFileSync(filePath, body);
  fs.chmodSync(filePath, 0o755);
}

function writeFixtureExecutable(basePath: string, body: string): string {
  if (process.platform !== 'win32') {
    writeExecutable(basePath, body);
    return basePath;
  }

  const scriptPath = `${basePath}.js`;
  const commandPath = `${basePath}.cmd`;
  fs.writeFileSync(scriptPath, body);
  fs.writeFileSync(commandPath, `@echo off\r\n"${process.execPath}" "${scriptPath}" %*\r\n`);
  return commandPath;
}

function createSmokeCase(name: string): SmokeCase {
  const baseDir = path.join(process.cwd(), '.tmp', 'launcher-smoke');
  fs.mkdirSync(baseDir, { recursive: true });

  const root = fs.mkdtempSync(path.join(baseDir, `${name}-`));
  const artifactsDir = path.join(root, 'artifacts');
  const binDir = path.join(root, 'bin');
  const xdgConfigHome = path.join(root, 'xdg');
  const xdgDataHome = path.join(root, 'xdg-data');
  const appDataDir = path.join(root, 'AppData', 'Roaming');
  const localAppDataDir = path.join(root, 'AppData', 'Local');
  const homeDir = path.join(root, 'home');
  const socketDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-smoke-sock-'));
  const socketPath = path.join(socketDir, 'subminer.sock');
  const videoPath = path.join(root, 'video.mkv');
  const fakeAppBasePath = path.join(binDir, 'fake-subminer');
  const fakeMpvBasePath = path.join(binDir, 'mpv');
  const mpvOverlayLogPath = path.join(artifactsDir, 'mpv-overlay.log');

  fs.mkdirSync(artifactsDir, { recursive: true });
  fs.mkdirSync(binDir, { recursive: true });
  fs.writeFileSync(videoPath, 'fake video fixture');

  const configDir = getDefaultConfigDir({ xdgConfigHome, appDataDir, homeDir });
  fs.mkdirSync(configDir, { recursive: true });
  fs.writeFileSync(path.join(configDir, 'config.jsonc'), JSON.stringify({ mpv: { socketPath } }));
  const setupState = createDefaultSetupState();
  setupState.status = 'completed';
  setupState.completedAt = '2026-03-07T00:00:00.000Z';
  setupState.completionSource = 'user';
  writeSetupState(getSetupStatePath(configDir), setupState);

  const fakeMpvLogPath = path.join(artifactsDir, 'fake-mpv.log');
  const fakeAppLogPath = path.join(artifactsDir, 'fake-app.log');
  const fakeAppStartLogPath = path.join(artifactsDir, 'fake-app-start.log');
  const fakeAppStopLogPath = path.join(artifactsDir, 'fake-app-stop.log');

  const fakeMpvPath = writeFixtureExecutable(
    fakeMpvBasePath,
    `#!/usr/bin/env bun
const fs = require('node:fs');
const net = require('node:net');
const path = require('node:path');

const logPath = ${JSON.stringify(fakeMpvLogPath)};
const args = process.argv.slice(2);
const socketArg = args.find((arg) => arg.startsWith('--input-ipc-server='));
const socketPath = socketArg ? socketArg.slice('--input-ipc-server='.length) : '';
fs.appendFileSync(logPath, JSON.stringify({ argv: args, socketPath }) + '\\n');

if (!socketPath) {
  process.exit(2);
}

try {
  fs.rmSync(socketPath, { force: true });
} catch {}

fs.mkdirSync(path.dirname(socketPath), { recursive: true });

const server = net.createServer((socket) => socket.end());
server.on('error', (error) => {
  fs.appendFileSync(logPath, JSON.stringify({ error: String(error) }) + '\\n');
  process.exit(3);
});
server.listen(socketPath);

const closeAndExit = () => {
  server.close(() => process.exit(0));
};

setTimeout(closeAndExit, 3000);
process.on('SIGTERM', closeAndExit);
`,
  );

  const fakeAppPath = writeFixtureExecutable(
    fakeAppBasePath,
    `#!/usr/bin/env bun
const fs = require('node:fs');
const path = require('node:path');

const logPath = ${JSON.stringify(fakeAppLogPath)};
const startPath = ${JSON.stringify(fakeAppStartLogPath)};
const stopPath = ${JSON.stringify(fakeAppStopLogPath)};
const entry = {
  argv: process.argv.slice(2),
  subminerMpvLog: process.env.SUBMINER_MPV_LOG || '',
};
fs.appendFileSync(logPath, JSON.stringify(entry) + '\\n');

if (entry.argv.includes('--start')) {
  fs.appendFileSync(startPath, JSON.stringify(entry) + '\\n');
}
if (entry.argv.includes('--stop')) {
  fs.appendFileSync(stopPath, JSON.stringify(entry) + '\\n');
}
if (entry.argv.includes('--app-ping')) {
  process.exit(process.env.SUBMINER_FAKE_APP_RUNNING === '1' ? 0 : 1);
}
if (entry.argv.includes('--ensure-linux-runtime-plugin-assets')) {
  const responseFlagIndex = entry.argv.indexOf('--ensure-linux-runtime-plugin-assets-response-path');
  const responsePath = responseFlagIndex >= 0 ? entry.argv[responseFlagIndex + 1] : '';
  const xdgDataHome = process.env.XDG_DATA_HOME || path.join(process.env.HOME || '', '.local', 'share');
  const dataDir = path.join(xdgDataHome, 'SubMiner');
  const pluginDir = path.join(dataDir, 'plugin', 'subminer');
  const pluginConfigPath = path.join(dataDir, 'plugin', 'subminer.conf');
  const themePath = path.join(dataDir, 'themes', 'subminer.rasi');
  fs.mkdirSync(pluginDir, { recursive: true });
  fs.mkdirSync(path.dirname(themePath), { recursive: true });
  fs.writeFileSync(path.join(pluginDir, 'main.lua'), '-- smoke plugin\\n');
  fs.writeFileSync(pluginConfigPath, 'smoke=true\\n');
  fs.writeFileSync(themePath, '/* smoke theme */\\n');
  if (responsePath) {
    fs.mkdirSync(path.dirname(responsePath), { recursive: true });
    fs.writeFileSync(responsePath, JSON.stringify({ ok: true, status: 'installed', path: path.join(pluginDir, 'main.lua') }));
  }
  process.exit(0);
}

process.exit(0);
`,
  );

  return {
    root,
    artifactsDir,
    binDir,
    xdgConfigHome,
    xdgDataHome,
    appDataDir,
    localAppDataDir,
    homeDir,
    socketDir,
    socketPath,
    videoPath,
    fakeAppPath,
    fakeMpvPath,
    mpvOverlayLogPath,
  };
}

function makeTestEnv(smokeCase: SmokeCase): NodeJS.ProcessEnv {
  const env: NodeJS.ProcessEnv = {
    ...process.env,
    HOME: smokeCase.homeDir,
    XDG_CONFIG_HOME: smokeCase.xdgConfigHome,
    XDG_DATA_HOME: smokeCase.xdgDataHome,
    APPDATA: smokeCase.appDataDir,
    LOCALAPPDATA: smokeCase.localAppDataDir,
    SUBMINER_APPIMAGE_PATH: smokeCase.fakeAppPath,
    SUBMINER_MPV_LOG: smokeCase.mpvOverlayLogPath,
  };
  const pathKey = Object.keys(env).find((key) => key.toLowerCase() === 'path') ?? 'PATH';
  env[pathKey] = `${smokeCase.binDir}${path.delimiter}${env[pathKey] || ''}`;
  for (const key of Object.keys(env)) {
    if (key !== pathKey && key.toLowerCase() === 'path') {
      delete env[key];
    }
  }
  return env;
}

function runLauncher(
  smokeCase: SmokeCase,
  argv: string[],
  env: NodeJS.ProcessEnv,
  label: string,
): RunResult {
  const result = spawnSync(
    process.execPath,
    ['run', path.join(process.cwd(), 'launcher/main.ts'), ...argv],
    {
      env,
      encoding: 'utf8',
      timeout: LAUNCHER_RUN_TIMEOUT_MS,
    },
  );

  const stdout = result.stdout || '';
  const stderr = result.stderr || '';
  fs.writeFileSync(path.join(smokeCase.artifactsDir, `${label}.stdout.log`), stdout);
  fs.writeFileSync(path.join(smokeCase.artifactsDir, `${label}.stderr.log`), stderr);

  return {
    status: result.status,
    stdout,
    stderr,
  };
}

async function withSmokeCase(
  name: string,
  fn: (smokeCase: SmokeCase) => Promise<void>,
): Promise<void> {
  const smokeCase = createSmokeCase(name);
  let completed = false;
  try {
    await fn(smokeCase);
    completed = true;
  } catch (error) {
    process.stderr.write(`[launcher-smoke] preserved artifacts: ${smokeCase.root}\n`);
    throw error;
  } finally {
    if (completed) {
      fs.rmSync(smokeCase.root, { recursive: true, force: true });
    }
    fs.rmSync(smokeCase.socketDir, { recursive: true, force: true });
  }
}

function readJsonLines(filePath: string): Array<Record<string, unknown>> {
  if (!fs.existsSync(filePath)) return [];
  return fs
    .readFileSync(filePath, 'utf8')
    .split(/\r?\n/)
    .filter((line) => line.trim().length > 0)
    .map((line) => JSON.parse(line) as Record<string, unknown>);
}

async function waitForJsonLines(
  filePath: string,
  minCount: number,
  timeoutMs = 1500,
): Promise<void> {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (readJsonLines(filePath).length >= minCount) {
      return;
    }
    await new Promise<void>((resolve) => setTimeout(resolve, 50));
  }
}

async function waitForFile(filePath: string, timeoutMs = 1500): Promise<void> {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (fs.existsSync(filePath)) return;
    await new Promise<void>((resolve) => setTimeout(resolve, 50));
  }
  throw new Error(`Timed out waiting for file ${filePath} after ${timeoutMs}ms`);
}

async function waitForSocketReady(socketPath: string, timeoutMs = 1500): Promise<boolean> {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (!fs.existsSync(socketPath)) {
      await new Promise<void>((resolve) => setTimeout(resolve, 50));
      continue;
    }
    const ready = await new Promise<boolean>((resolve) => {
      const socket = new net.Socket();
      socket.once('connect', () => {
        socket.end();
        resolve(true);
      });
      socket.once('error', () => {
        socket.destroy();
        resolve(false);
      });
      socket.connect(socketPath);
    });
    if (ready) return true;
    await new Promise<void>((resolve) => setTimeout(resolve, 50));
  }
  return false;
}

async function startFakeControlServer(
  smokeCase: SmokeCase,
): Promise<{ socketPath: string; logPath: string; stop: () => Promise<void> }> {
  const socketPath = path.join(smokeCase.socketDir, 'app-control.sock');
  const logPath = path.join(smokeCase.artifactsDir, 'fake-control.log');
  const readyPath = path.join(smokeCase.artifactsDir, 'fake-control.ready');
  const scriptPath = path.join(smokeCase.artifactsDir, 'fake-control-server.js');

  fs.writeFileSync(
    scriptPath,
    `const fs = require('node:fs');
const net = require('node:net');
const path = require('node:path');

const socketPath = ${JSON.stringify(socketPath)};
const logPath = ${JSON.stringify(logPath)};
const readyPath = ${JSON.stringify(readyPath)};
try { fs.rmSync(socketPath, { force: true }); } catch {}
fs.mkdirSync(path.dirname(socketPath), { recursive: true });

const server = net.createServer((socket) => {
  let buffer = '';
  socket.on('data', (chunk) => {
    buffer += chunk.toString('utf8');
    let handledLine = false;
    while (true) {
      const newlineMatch = buffer.match(/\\r?\\n/);
      if (!newlineMatch || newlineMatch.index === undefined) break;
      const line = buffer.slice(0, newlineMatch.index).trim();
      buffer = buffer.slice(newlineMatch.index + newlineMatch[0].length);
      if (!line) continue;
      fs.appendFileSync(logPath, line + '\\n');
      handledLine = true;
    }
    if (handledLine) {
      socket.end(JSON.stringify({ ok: true }) + '\\n');
    }
  });
});

server.listen(socketPath, () => {
  fs.writeFileSync(readyPath, 'ready');
});

const shutdown = () => {
  server.close(() => {
    try { fs.rmSync(socketPath, { force: true }); } catch {}
    process.exit(0);
  });
};
process.on('SIGTERM', shutdown);
process.on('SIGINT', shutdown);
setInterval(() => {}, 1000);
`,
  );

  const proc = spawn(process.execPath, [scriptPath], { stdio: 'ignore' });
  await waitForFile(readyPath);

  return {
    socketPath,
    logPath,
    stop: async () => {
      if (proc.exitCode !== null || proc.signalCode !== null) return;
      proc.kill('SIGTERM');
      await new Promise<void>((resolve) => {
        const timer = setTimeout(() => {
          proc.kill('SIGKILL');
          resolve();
        }, 1000);
        proc.once('close', () => {
          clearTimeout(timer);
          resolve();
        });
      });
    },
  };
}

test('launcher smoke fixture seeds completed setup state', () => {
  const smokeCase = createSmokeCase('setup-state');
  try {
    const configDir = getDefaultConfigDir({
      xdgConfigHome: smokeCase.xdgConfigHome,
      homeDir: smokeCase.homeDir,
    });
    const statePath = getSetupStatePath(configDir);

    assert.equal(readSetupState(statePath)?.status, 'completed');
  } finally {
    fs.rmSync(smokeCase.root, { recursive: true, force: true });
    fs.rmSync(smokeCase.socketDir, { recursive: true, force: true });
  }
});

test('launcher mpv status returns ready when socket is connectable', async () => {
  await withSmokeCase('mpv-status', async (smokeCase) => {
    const env = makeTestEnv(smokeCase);
    const fakeMpv = spawn(smokeCase.fakeMpvPath, [`--input-ipc-server=${smokeCase.socketPath}`], {
      env,
      stdio: 'ignore',
    });

    try {
      await waitForSocketReady(smokeCase.socketPath);
      const result = runLauncher(
        smokeCase,
        ['mpv', 'status', '--log-level', 'debug'],
        env,
        'mpv-status',
      );
      const fakeMpvEntries = readJsonLines(path.join(smokeCase.artifactsDir, 'fake-mpv.log'));
      const fakeMpvError = fakeMpvEntries.find(
        (entry): entry is { error: string } => typeof entry.error === 'string',
      )?.error;
      const unixSocketDenied =
        typeof fakeMpvError === 'string' && /eperm|operation not permitted/i.test(fakeMpvError);

      if (unixSocketDenied) {
        assert.equal(result.status, 1);
        assert.match(result.stdout, /socket not ready/i);
      } else {
        assert.equal(result.status, 0);
        assert.match(result.stdout, /socket ready/i);
      }
    } finally {
      if (fakeMpv.exitCode === null) {
        await new Promise<void>((resolve) => {
          fakeMpv.once('close', () => resolve());
        });
      }
    }
  });
});

test(
  'launcher start-overlay run forwards socket/backend and stops owned background app after mpv exits',
  { timeout: LONG_SMOKE_TEST_TIMEOUT_MS },
  async () => {
    await withSmokeCase('overlay-start-stop', async (smokeCase) => {
      const env = makeTestEnv(smokeCase);
      const result = runLauncher(
        smokeCase,
        ['--backend', 'x11', '--log-level', 'info', '--start-overlay', smokeCase.videoPath],
        env,
        'overlay-start-stop',
      );

      const appStartPath = path.join(smokeCase.artifactsDir, 'fake-app-start.log');
      const appStopPath = path.join(smokeCase.artifactsDir, 'fake-app-stop.log');
      await waitForJsonLines(appStartPath, 1);
      await waitForJsonLines(appStopPath, 1);

      const appStartEntries = readJsonLines(appStartPath);
      const appStopEntries = readJsonLines(appStopPath);
      const mpvEntries = readJsonLines(path.join(smokeCase.artifactsDir, 'fake-mpv.log'));
      const mpvError = mpvEntries.find(
        (entry): entry is { error: string } => typeof entry.error === 'string',
      )?.error;
      const unixSocketDenied =
        typeof mpvError === 'string' && /eperm|operation not permitted/i.test(mpvError);

      assert.equal(result.status, unixSocketDenied ? 3 : 0);
      assert.match(result.stdout, /Starting SubMiner overlay/i);

      assert.equal(appStartEntries.length, 1);
      assert.equal(appStopEntries.length, 1);
      assert.equal(mpvEntries.length >= 1, true);

      const appStartArgs = appStartEntries[0]?.argv;
      assert.equal(Array.isArray(appStartArgs), true);
      assert.equal((appStartArgs as string[]).includes('--background'), true);
      assert.equal((appStartArgs as string[]).includes('--start'), true);
      assert.equal((appStartArgs as string[]).includes('--managed-playback'), true);
      assert.equal((appStartArgs as string[]).includes('--backend'), true);
      assert.equal((appStartArgs as string[]).includes('x11'), true);
      assert.equal((appStartArgs as string[]).includes('--socket'), true);
      assert.equal((appStartArgs as string[]).includes(smokeCase.socketPath), true);
      assert.equal(appStartEntries[0]?.subminerMpvLog, smokeCase.mpvOverlayLogPath);

      const mpvFirstArgs = mpvEntries[0]?.argv;
      assert.equal(Array.isArray(mpvFirstArgs), true);
      assert.equal(
        (mpvFirstArgs as string[]).some(
          (arg) => arg === `--input-ipc-server=${smokeCase.socketPath}`,
        ),
        true,
      );
      assert.equal((mpvFirstArgs as string[]).includes(smokeCase.videoPath), true);
    });
  },
);

test(
  'launcher start-overlay attaches to a running background app without spawning another app start command',
  { timeout: LONG_SMOKE_TEST_TIMEOUT_MS },
  async () => {
    await withSmokeCase('overlay-borrow-background', async (smokeCase) => {
      const controlServer = await startFakeControlServer(smokeCase);
      const env = {
        ...makeTestEnv(smokeCase),
        SUBMINER_FAKE_APP_RUNNING: '1',
        SUBMINER_APP_CONTROL_SOCKET: controlServer.socketPath,
      };
      try {
        const result = runLauncher(
          smokeCase,
          ['--backend', 'x11', '--start-overlay', smokeCase.videoPath],
          env,
          'overlay-borrow-background',
        );

        const appLogPath = path.join(smokeCase.artifactsDir, 'fake-app.log');
        const appStartPath = path.join(smokeCase.artifactsDir, 'fake-app-start.log');
        const appStopPath = path.join(smokeCase.artifactsDir, 'fake-app-stop.log');
        await waitForJsonLines(controlServer.logPath, 1);

        const appEntries = readJsonLines(appLogPath);
        const appStartEntries = readJsonLines(appStartPath);
        const appStopEntries = readJsonLines(appStopPath);
        const controlEntries = readJsonLines(controlServer.logPath);
        const mpvEntries = readJsonLines(path.join(smokeCase.artifactsDir, 'fake-mpv.log'));
        const mpvError = mpvEntries.find(
          (entry): entry is { error: string } => typeof entry.error === 'string',
        )?.error;
        const unixSocketDenied =
          typeof mpvError === 'string' && /eperm|operation not permitted/i.test(mpvError);

        assert.equal(result.status, unixSocketDenied ? 3 : 0);
        assert.equal(appEntries.length > 0, true);
        assert.equal(
          appEntries.every((entry) =>
            (entry.argv as string[]).includes('--ensure-linux-runtime-plugin-assets'),
          ),
          true,
        );
        assert.equal(appStartEntries.length, 0);
        assert.equal(appStopEntries.length, 0);
        assert.equal(controlEntries.length, 1);
        const controlArgs = controlEntries[0]?.argv;
        assert.equal(Array.isArray(controlArgs), true);
        assert.equal((controlArgs as string[]).includes('--background'), false);
        assert.equal((controlArgs as string[]).includes('--start'), true);
        assert.equal((controlArgs as string[]).includes('--managed-playback'), true);
      } finally {
        await controlServer.stop();
      }
    });
  },
);

test(
  'launcher starts mpv paused when plugin auto-start visible overlay gate is enabled',
  { timeout: LONG_SMOKE_TEST_TIMEOUT_MS },
  async () => {
    await withSmokeCase('autoplay-ready-gate', async (smokeCase) => {
      fs.writeFileSync(
        path.join(getDefaultConfigDir(smokeCase), 'config.jsonc'),
        JSON.stringify({
          auto_start_overlay: true,
          mpv: {
            socketPath: smokeCase.socketPath,
            autoStartSubMiner: true,
            pauseUntilOverlayReady: true,
          },
        }),
      );

      const env = makeTestEnv(smokeCase);
      const result = runLauncher(
        smokeCase,
        [smokeCase.videoPath, '--log-level', 'debug'],
        env,
        'autoplay-ready-gate',
      );

      const mpvEntries = readJsonLines(path.join(smokeCase.artifactsDir, 'fake-mpv.log'));
      const mpvError = mpvEntries.find(
        (entry): entry is { error: string } => typeof entry.error === 'string',
      )?.error;
      const unixSocketDenied =
        typeof mpvError === 'string' && /eperm|operation not permitted/i.test(mpvError);
      const mpvFirstArgs = mpvEntries[0]?.argv;

      assert.equal(result.status, unixSocketDenied ? 3 : 0);
      assert.equal(Array.isArray(mpvFirstArgs), true);
      assert.equal((mpvFirstArgs as string[]).includes('--pause=yes'), true);
      assert.match(
        (mpvFirstArgs as string[]).find((arg) => arg.startsWith('--script-opts=')) ?? '',
        /subminer-auto_start_pause_until_ready_owns_initial_pause=yes/,
      );
      assert.match(result.stdout, /pause mpv until overlay and tokenization are ready/i);
      assert.match(result.stdout, /managed plugin\/theme assets/i);
      assert.equal(
        fs.existsSync(path.join(smokeCase.xdgDataHome, 'SubMiner', 'themes', 'subminer.rasi')),
        true,
      );
    });
  },
);
