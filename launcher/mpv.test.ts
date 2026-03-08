import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import net from 'node:net';
import { EventEmitter } from 'node:events';
import type { Args } from './types';
import { runAppCommandCaptureOutput, startOverlay, state, waitForUnixSocketReady } from './mpv';
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

function makeArgs(overrides: Partial<Args> = {}): Args {
  return {
    backend: 'x11',
    directory: '.',
    recursive: false,
    profile: '',
    startOverlay: false,
    youtubeSubgenMode: 'off',
    whisperBin: '',
    whisperModel: '',
    youtubeSubgenOutDir: '',
    youtubeSubgenAudioFormat: 'wav',
    youtubeSubgenKeepTemp: false,
    youtubePrimarySubLangs: [],
    youtubeSecondarySubLangs: [],
    youtubeAudioLangs: [],
    youtubeWhisperSourceLanguage: 'ja',
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
    doctor: false,
    configPath: false,
    configShow: false,
    mpvIdle: false,
    mpvSocket: false,
    mpvStatus: false,
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
