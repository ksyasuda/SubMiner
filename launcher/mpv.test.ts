import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import net from 'node:net';
import { EventEmitter } from 'node:events';
import type { Args } from './types';
import {
  cleanupPlaybackSession,
  findAppBinary,
  runAppCommandCaptureOutput,
  shouldResolveAniSkipMetadata,
  startOverlay,
  state,
  waitForUnixSocketReady,
} from './mpv';
import * as mpvModule from './mpv';

function createTempSocketPath(): { dir: string; socketPath: string } {
  const baseDir = path.join(process.cwd(), '.tmp', 'launcher-mpv-tests');
  fs.mkdirSync(baseDir, { recursive: true });
  const dir = fs.mkdtempSync(path.join(baseDir, 'case-'));
  return { dir, socketPath: path.join(dir, 'mpv.sock') };
}

test('mpv module exposes only canonical socket readiness helper', () => {
  assert.equal('waitForSocket' in mpvModule, false);
});

test('runAppCommandCaptureOutput captures status and stdio', () => {
  const result = runAppCommandCaptureOutput(process.execPath, [
    '-e',
    'process.stdout.write("stdout-line"); process.stderr.write("stderr-line");',
  ]);

  assert.equal(result.status, 0);
  assert.equal(result.stdout, 'stdout-line');
  assert.equal(result.stderr, 'stderr-line');
  assert.equal(result.error, undefined);
});

test('waitForUnixSocketReady returns false when socket never appears', async () => {
  const { dir, socketPath } = createTempSocketPath();
  try {
    const ready = await waitForUnixSocketReady(socketPath, 120);
    assert.equal(ready, false);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('waitForUnixSocketReady returns false when path exists but is not socket', async () => {
  const { dir, socketPath } = createTempSocketPath();
  try {
    fs.writeFileSync(socketPath, 'not-a-socket');
    const ready = await waitForUnixSocketReady(socketPath, 200);
    assert.equal(ready, false);
  } finally {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('waitForUnixSocketReady returns true when socket becomes connectable before timeout', async () => {
  const { dir, socketPath } = createTempSocketPath();
  fs.writeFileSync(socketPath, '');
  const originalCreateConnection = net.createConnection;
  try {
    net.createConnection = (() => {
      const socket = new EventEmitter() as net.Socket;
      socket.destroy = (() => socket) as net.Socket['destroy'];
      socket.setTimeout = (() => socket) as net.Socket['setTimeout'];
      setTimeout(() => socket.emit('connect'), 25);
      return socket;
    }) as typeof net.createConnection;

    const ready = await waitForUnixSocketReady(socketPath, 400);
    assert.equal(ready, true);
  } finally {
    net.createConnection = originalCreateConnection;
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('shouldResolveAniSkipMetadata skips URL and YouTube-preloaded playback', () => {
  assert.equal(shouldResolveAniSkipMetadata('/media/show.mkv', 'file'), true);
  assert.equal(
    shouldResolveAniSkipMetadata('https://www.youtube.com/watch?v=test123', 'url'),
    false,
  );
  assert.equal(
    shouldResolveAniSkipMetadata('/tmp/video123.webm', 'file', {
      primaryPath: '/tmp/video123.ja.srt',
    }),
    false,
  );
});

function makeArgs(overrides: Partial<Args> = {}): Args {
  return {
    backend: 'x11',
    directory: '.',
    recursive: false,
    profile: '',
    startOverlay: false,
    whisperBin: '',
    whisperModel: '',
    whisperVadModel: '',
    whisperThreads: 4,
    youtubeSubgenOutDir: '',
    youtubeSubgenAudioFormat: 'wav',
    youtubeSubgenKeepTemp: false,
    youtubeFixWithAi: false,
    youtubePrimarySubLangs: [],
    youtubeSecondarySubLangs: [],
    youtubeAudioLangs: [],
    youtubeWhisperSourceLanguage: 'ja',
    aiConfig: {},
    useTexthooker: false,
    autoStartOverlay: false,
    texthookerOnly: false,
    useRofi: false,
    logLevel: 'error',
    passwordStore: '',
    target: '',
    targetKind: '',
    jimakuApiKey: '',
    jimakuApiKeyCommand: '',
    jimakuApiBaseUrl: '',
    jimakuLanguagePreference: 'none',
    jimakuMaxEntryResults: 10,
    jellyfin: false,
    jellyfinLogin: false,
    jellyfinLogout: false,
    jellyfinPlay: false,
    jellyfinDiscovery: false,
    dictionary: false,
    stats: false,
    doctor: false,
    configPath: false,
    configShow: false,
    mpvIdle: false,
    mpvSocket: false,
    mpvStatus: false,
    mpvArgs: '',
    appPassthrough: false,
    appArgs: [],
    jellyfinServer: '',
    jellyfinUsername: '',
    jellyfinPassword: '',
    ...overrides,
  };
}

test('startOverlay resolves without fixed 2s sleep when readiness signals arrive quickly', async () => {
  const { dir, socketPath } = createTempSocketPath();
  const appPath = path.join(dir, 'fake-subminer.sh');
  fs.writeFileSync(appPath, '#!/bin/sh\nexit 0\n');
  fs.chmodSync(appPath, 0o755);
  fs.writeFileSync(socketPath, '');
  const originalCreateConnection = net.createConnection;
  try {
    net.createConnection = (() => {
      const socket = new EventEmitter() as net.Socket;
      socket.destroy = (() => socket) as net.Socket['destroy'];
      socket.setTimeout = (() => socket) as net.Socket['setTimeout'];
      setTimeout(() => socket.emit('connect'), 10);
      return socket;
    }) as typeof net.createConnection;

    const startedAt = Date.now();
    await startOverlay(appPath, makeArgs(), socketPath);
    const elapsedMs = Date.now() - startedAt;

    assert.ok(elapsedMs < 1200, `expected startOverlay <1200ms, got ${elapsedMs}ms`);
  } finally {
    net.createConnection = originalCreateConnection;
    state.overlayProc = null;
    state.overlayManagedByLauncher = false;
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('cleanupPlaybackSession preserves background app while stopping mpv-owned children', async () => {
  const { dir } = createTempSocketPath();
  const appPath = path.join(dir, 'fake-subminer.sh');
  const appInvocationsPath = path.join(dir, 'app-invocations.log');
  fs.writeFileSync(
    appPath,
    `#!/bin/sh\necho \"$@\" >> ${JSON.stringify(appInvocationsPath)}\nexit 0\n`,
  );
  fs.chmodSync(appPath, 0o755);

  const calls: string[] = [];
  const overlayProc = {
    killed: false,
    kill: () => {
      calls.push('overlay-kill');
      return true;
    },
  } as unknown as NonNullable<typeof state.overlayProc>;
  const mpvProc = {
    killed: false,
    kill: () => {
      calls.push('mpv-kill');
      return true;
    },
  } as unknown as NonNullable<typeof state.mpvProc>;
  const helperProc = {
    killed: false,
    kill: () => {
      calls.push('helper-kill');
      return true;
    },
  } as unknown as NonNullable<typeof state.overlayProc>;

  state.stopRequested = false;
  state.appPath = appPath;
  state.overlayManagedByLauncher = true;
  state.overlayProc = overlayProc;
  state.mpvProc = mpvProc;
  state.youtubeSubgenChildren.add(helperProc);

  try {
    await cleanupPlaybackSession(makeArgs());

    assert.deepEqual(calls, ['mpv-kill', 'helper-kill']);
    assert.equal(fs.existsSync(appInvocationsPath), false);
  } finally {
    state.overlayProc = null;
    state.mpvProc = null;
    state.youtubeSubgenChildren.clear();
    state.overlayManagedByLauncher = false;
    state.appPath = '';
    state.stopRequested = false;
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

// ── findAppBinary: Linux packaged path discovery ──────────────────────────────

function makeExecutable(filePath: string): void {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, '#!/bin/sh\nexit 0\n');
  fs.chmodSync(filePath, 0o755);
}

function withFindAppBinaryEnvSandbox(run: () => void): void {
  const originalAppImagePath = process.env.SUBMINER_APPIMAGE_PATH;
  const originalBinaryPath = process.env.SUBMINER_BINARY_PATH;
  try {
    delete process.env.SUBMINER_APPIMAGE_PATH;
    delete process.env.SUBMINER_BINARY_PATH;
    run();
  } finally {
    if (originalAppImagePath === undefined) {
      delete process.env.SUBMINER_APPIMAGE_PATH;
    } else {
      process.env.SUBMINER_APPIMAGE_PATH = originalAppImagePath;
    }
    if (originalBinaryPath === undefined) {
      delete process.env.SUBMINER_BINARY_PATH;
    } else {
      process.env.SUBMINER_BINARY_PATH = originalBinaryPath;
    }
  }
}

function withAccessSyncStub(isExecutablePath: (filePath: string) => boolean, run: () => void): void {
  const originalAccessSync = fs.accessSync;
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (fs as any).accessSync = (filePath: string): void => {
      if (isExecutablePath(filePath)) {
        return;
      }
      throw Object.assign(new Error(`EACCES: ${filePath}`), { code: 'EACCES' });
    };
    run();
  } finally {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (fs as any).accessSync = originalAccessSync;
  }
}

test('findAppBinary resolves ~/.local/bin/SubMiner.AppImage when it exists', () => {
  const baseDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-test-home-'));
  const originalHomedir = os.homedir;
  try {
    os.homedir = () => baseDir;
    const appImage = path.join(baseDir, '.local/bin/SubMiner.AppImage');
    makeExecutable(appImage);

    withFindAppBinaryEnvSandbox(() => {
      const result = findAppBinary('/some/other/path/subminer');
      assert.equal(result, appImage);
    });
  } finally {
    os.homedir = originalHomedir;
    fs.rmSync(baseDir, { recursive: true, force: true });
  }
});

test('findAppBinary resolves /opt/SubMiner/SubMiner.AppImage when ~/.local/bin candidate does not exist', () => {
  const baseDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-test-home-'));
  const originalHomedir = os.homedir;
  try {
    os.homedir = () => baseDir;
    withFindAppBinaryEnvSandbox(() => {
      withAccessSyncStub((filePath) => filePath === '/opt/SubMiner/SubMiner.AppImage', () => {
        const result = findAppBinary('/some/other/path/subminer');
        assert.equal(result, '/opt/SubMiner/SubMiner.AppImage');
      });
    });
  } finally {
    os.homedir = originalHomedir;
    fs.rmSync(baseDir, { recursive: true, force: true });
  }
});

test('findAppBinary finds subminer on PATH when AppImage candidates do not exist', () => {
  const baseDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-test-path-'));
  const originalHomedir = os.homedir;
  const originalPath = process.env.PATH;
  try {
    os.homedir = () => baseDir;
    // No AppImage candidates in empty home dir; place subminer wrapper on PATH
    const binDir = path.join(baseDir, 'bin');
    const wrapperPath = path.join(binDir, 'subminer');
    makeExecutable(wrapperPath);
    process.env.PATH = `${binDir}${path.delimiter}${originalPath ?? ''}`;

    withFindAppBinaryEnvSandbox(() => {
      withAccessSyncStub((filePath) => filePath === wrapperPath, () => {
        // selfPath must differ from wrapperPath so the self-check does not exclude it
        const result = findAppBinary(path.join(baseDir, 'launcher', 'subminer'));
        assert.equal(result, wrapperPath);
      });
    });
  } finally {
    os.homedir = originalHomedir;
    process.env.PATH = originalPath;
    fs.rmSync(baseDir, { recursive: true, force: true });
  }
});
