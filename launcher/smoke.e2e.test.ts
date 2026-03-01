import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawn, spawnSync } from 'node:child_process';

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
  homeDir: string;
  socketDir: string;
  socketPath: string;
  videoPath: string;
  fakeAppPath: string;
  fakeMpvPath: string;
  mpvOverlayLogPath: string;
};

function writeExecutable(filePath: string, body: string): void {
  fs.writeFileSync(filePath, body);
  fs.chmodSync(filePath, 0o755);
}

function createSmokeCase(name: string): SmokeCase {
  const baseDir = path.join(process.cwd(), '.tmp', 'launcher-smoke');
  fs.mkdirSync(baseDir, { recursive: true });

  const root = fs.mkdtempSync(path.join(baseDir, `${name}-`));
  const artifactsDir = path.join(root, 'artifacts');
  const binDir = path.join(root, 'bin');
  const xdgConfigHome = path.join(root, 'xdg');
  const homeDir = path.join(root, 'home');
  const socketDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-smoke-sock-'));
  const socketPath = path.join(socketDir, 'subminer.sock');
  const videoPath = path.join(root, 'video.mkv');
  const fakeAppPath = path.join(binDir, 'fake-subminer');
  const fakeMpvPath = path.join(binDir, 'mpv');
  const mpvOverlayLogPath = path.join(artifactsDir, 'mpv-overlay.log');

  fs.mkdirSync(artifactsDir, { recursive: true });
  fs.mkdirSync(binDir, { recursive: true });
  fs.mkdirSync(path.join(xdgConfigHome, 'mpv', 'script-opts'), { recursive: true });
  fs.writeFileSync(videoPath, 'fake video fixture');
  fs.writeFileSync(
    path.join(xdgConfigHome, 'mpv', 'script-opts', 'subminer.conf'),
    `socket_path=${socketPath}\n`,
  );

  const fakeMpvLogPath = path.join(artifactsDir, 'fake-mpv.log');
  const fakeAppLogPath = path.join(artifactsDir, 'fake-app.log');
  const fakeAppStartLogPath = path.join(artifactsDir, 'fake-app-start.log');
  const fakeAppStopLogPath = path.join(artifactsDir, 'fake-app-stop.log');

  writeExecutable(
    fakeMpvPath,
    `#!/usr/bin/env node
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

  writeExecutable(
    fakeAppPath,
    `#!/usr/bin/env node
const fs = require('node:fs');

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

process.exit(0);
`,
  );

  return {
    root,
    artifactsDir,
    binDir,
    xdgConfigHome,
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
  return {
    ...process.env,
    HOME: smokeCase.homeDir,
    XDG_CONFIG_HOME: smokeCase.xdgConfigHome,
    SUBMINER_APPIMAGE_PATH: smokeCase.fakeAppPath,
    SUBMINER_MPV_LOG: smokeCase.mpvOverlayLogPath,
    PATH: `${smokeCase.binDir}${path.delimiter}${process.env.PATH || ''}`,
  };
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
      timeout: 15000,
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

test('launcher mpv status returns ready when socket is connectable', async () => {
  await withSmokeCase('mpv-status', async (smokeCase) => {
    const env = makeTestEnv(smokeCase);
    const fakeMpv = spawn(smokeCase.fakeMpvPath, [`--input-ipc-server=${smokeCase.socketPath}`], {
      env,
      stdio: 'ignore',
    });

    try {
      await new Promise<void>((resolve) => setTimeout(resolve, 120));
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
  'launcher start-overlay run forwards socket/backend and stops overlay after mpv exits',
  { timeout: 20000 },
  async () => {
    await withSmokeCase('overlay-start-stop', async (smokeCase) => {
      const env = makeTestEnv(smokeCase);
      const result = runLauncher(
        smokeCase,
        ['--backend', 'x11', '--start-overlay', smokeCase.videoPath],
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
      assert.equal((appStartArgs as string[]).includes('--start'), true);
      assert.equal((appStartArgs as string[]).includes('--backend'), true);
      assert.equal((appStartArgs as string[]).includes('x11'), true);
      assert.equal((appStartArgs as string[]).includes('--socket'), true);
      assert.equal((appStartArgs as string[]).includes(smokeCase.socketPath), true);
      assert.equal(appStartEntries[0]?.subminerMpvLog, smokeCase.mpvOverlayLogPath);

      const appStopArgs = appStopEntries[0]?.argv;
      assert.deepEqual(appStopArgs, ['--stop']);

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
  'launcher starts mpv paused when plugin auto-start visible overlay gate is enabled',
  { timeout: 20000 },
  async () => {
    await withSmokeCase('autoplay-ready-gate', async (smokeCase) => {
      fs.writeFileSync(
        path.join(smokeCase.xdgConfigHome, 'mpv', 'script-opts', 'subminer.conf'),
        [
          `socket_path=${smokeCase.socketPath}`,
          'auto_start=yes',
          'auto_start_visible_overlay=yes',
          'auto_start_pause_until_ready=yes',
          '',
        ].join('\n'),
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
      assert.match(result.stdout, /pause mpv until overlay and tokenization are ready/i);
    });
  },
);
