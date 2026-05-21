import assert from 'node:assert/strict';
import { EventEmitter } from 'node:events';
import fs from 'node:fs';
import net from 'node:net';
import os from 'node:os';
import path from 'node:path';
import test from 'node:test';
import { sendAppControlCommand } from '../../shared/app-control-client';
import { startAppControlServer } from './app-control-server';

async function waitForSocketPath(socketPath: string): Promise<void> {
  const timeoutMs = 1000;
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (fs.existsSync(socketPath)) return;
    await new Promise<void>((resolve) => setTimeout(resolve, 10));
  }
  throw new Error(`Timed out waiting for control socket ${socketPath} after ${timeoutMs}ms`);
}

test('app control server dispatches argv requests and replies ok', async () => {
  if (process.platform === 'win32') return;

  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-control-test-'));
  const socketPath = path.join(dir, 'control.sock');
  const received: string[][] = [];
  const server = startAppControlServer({
    socketPath,
    platform: 'linux',
    handleArgv: (argv) => {
      received.push(argv);
    },
  });

  try {
    await waitForSocketPath(socketPath);
    const result = await sendAppControlCommand(['--start', '--socket', '/tmp/mpv.sock'], {
      socketPath,
    });

    assert.deepEqual(result, { ok: true });
    assert.deepEqual(received, [['--start', '--socket', '/tmp/mpv.sock']]);
  } finally {
    server.close();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('app control server rejects requests larger than 64KB by UTF-8 byte length', async () => {
  if (process.platform === 'win32') return;

  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-control-test-'));
  const socketPath = path.join(dir, 'control.sock');
  const received: string[][] = [];
  const server = startAppControlServer({
    socketPath,
    platform: 'linux',
    handleArgv: (argv) => {
      received.push(argv);
    },
  });

  try {
    await waitForSocketPath(socketPath);
    const result = await sendAppControlCommand(
      Array.from({ length: 4 }, () => 'あ'.repeat(6000)),
      {
        socketPath,
      },
    );

    assert.deepEqual(result, { ok: false, error: 'App control request too large' });
    assert.deepEqual(received, []);
  } finally {
    server.close();
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

test('app control server logs and closes errored client sockets', () => {
  const originalCreateServer = net.createServer;
  let socketHandler: ((socket: net.Socket) => void) | null = null;
  const fakeServer = new EventEmitter() as net.Server;
  fakeServer.listen = (() => fakeServer) as net.Server['listen'];
  fakeServer.close = ((callback?: (err?: Error) => void) => {
    callback?.();
    return fakeServer;
  }) as net.Server['close'];
  const received: string[][] = [];
  const warnings: Array<{ message: string; error?: unknown }> = [];

  try {
    net.createServer = ((handler?: (socket: net.Socket) => void) => {
      socketHandler = handler ?? null;
      return fakeServer;
    }) as typeof net.createServer;

    const server = startAppControlServer({
      socketPath: '\\\\.\\pipe\\subminer-test-control',
      platform: 'win32',
      handleArgv: (argv) => {
        received.push(argv);
      },
      logWarn: (message, error) => {
        warnings.push({ message, error });
      },
    });

    const error = new Error('client reset');
    let destroyed = false;
    const socket = new EventEmitter() as net.Socket;
    socket.destroy = (() => {
      destroyed = true;
      return socket;
    }) as net.Socket['destroy'];

    const handler = socketHandler as ((socket: net.Socket) => void) | null;
    assert.ok(handler);
    handler(socket);
    socket.emit('error', error);
    socket.emit('data', Buffer.from('{"argv":["--start"]}\n'));

    assert.equal(destroyed, true);
    assert.deepEqual(received, []);
    assert.deepEqual(warnings, [{ message: 'App control client socket error.', error }]);

    server.close();
  } finally {
    net.createServer = originalCreateServer;
  }
});
