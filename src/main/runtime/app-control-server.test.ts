import assert from 'node:assert/strict';
import fs from 'node:fs';
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
  throw new Error(
    `Timed out waiting for control socket ${socketPath} after ${timeoutMs}ms`,
  );
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
