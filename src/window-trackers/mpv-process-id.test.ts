import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import net from 'node:net';
import os from 'node:os';
import path from 'node:path';
import { queryMpvProcessId } from './mpv-process-id';

// Runs a stand-in mpv IPC server that hands each request line to `onRequest`.
async function withFakeMpv(
  onRequest: (socket: net.Socket, request: string) => void,
  run: (socketPath: string) => Promise<void>,
): Promise<void> {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'mpv-pid-'));
  const socketPath =
    process.platform === 'win32'
      ? path.win32.join(String.raw`\\.\pipe`, path.basename(dir))
      : path.join(dir, 'mpv.sock');
  const server = net.createServer((socket) => {
    socket.once('data', (data) => onRequest(socket, data.toString()));
  });
  try {
    await new Promise<void>((resolve, reject) => {
      server.once('error', reject);
      server.listen(socketPath, resolve);
    });
    await run(socketPath);
  } finally {
    await new Promise((resolve) => server.close(resolve));
    fs.rmSync(dir, { recursive: true, force: true });
  }
}

test('queryMpvProcessId asks mpv for pid and returns the answer', async () => {
  let request = '';
  await withFakeMpv(
    (socket, line) => {
      request = line;
      socket.write('{"data":4242,"request_id":1,"error":"success"}\n');
    },
    async (socketPath) => {
      assert.equal(await queryMpvProcessId(socketPath), 4242);
    },
  );
  assert.deepEqual(JSON.parse(request), { command: ['get_property', 'pid'], request_id: 1 });
});

test('queryMpvProcessId skips events and malformed lines, across split chunks', async () => {
  await withFakeMpv(
    (socket) => {
      socket.write('{"event":"idle"}\nnot json\n{"data":42');
      setTimeout(() => socket.write('42,"request_id":1,"error":"success"}\n'), 10);
    },
    async (socketPath) => {
      assert.equal(await queryMpvProcessId(socketPath), 4242);
    },
  );
});

test('queryMpvProcessId rejects when mpv refuses the property', async () => {
  await withFakeMpv(
    (socket) => socket.write('{"request_id":1,"error":"property unavailable"}\n'),
    async (socketPath) => {
      await assert.rejects(queryMpvProcessId(socketPath), /no process ID/);
    },
  );
});

test('queryMpvProcessId rejects when nothing is listening', async () => {
  const missing =
    process.platform === 'win32'
      ? path.win32.join(String.raw`\\.\pipe`, `mpv-pid-missing-${process.pid}`)
      : path.join(os.tmpdir(), `mpv-pid-missing-${process.pid}.sock`);
  await assert.rejects(queryMpvProcessId(missing));
});

test('queryMpvProcessId rejects when mpv hangs up or never answers', async () => {
  await withFakeMpv(
    (socket) => socket.end(),
    async (socketPath) => {
      await assert.rejects(queryMpvProcessId(socketPath), /closed before mpv answered/);
    },
  );

  const held: net.Socket[] = [];
  await withFakeMpv(
    (socket) => held.push(socket),
    async (socketPath) => {
      await assert.rejects(queryMpvProcessId(socketPath, 20), /no answer within 20ms/);
      held.forEach((socket) => socket.destroy());
    },
  );
});
