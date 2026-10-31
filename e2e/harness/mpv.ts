import { spawn, type ChildProcess } from 'node:child_process';
import { once } from 'node:events';
import fs from 'node:fs';
import net from 'node:net';
import type { E2eDisplay } from './display';
import { E2E_SCREEN } from './display';
import { applyEnvDelta } from './process';
import { waitUntil } from './wait';

export type MpvClient = {
  /** Sends one mpv JSON IPC command and resolves with its `data` field. */
  command: (...args: unknown[]) => Promise<unknown>;
  close: () => void;
};

type MpvReply = { request_id?: number; error?: string; data?: unknown };

export async function startMpv(options: {
  socketPath: string;
  mediaPath: string;
  subtitlePath: string;
  /** mpv's own verbose log (`--log-file`); `--no-terminal` silences its console output. */
  logPath: string;
  /** Process stdout/stderr, which still catches crash output from mpv's libraries. */
  stdioLogPath: string;
  display: E2eDisplay;
}): Promise<ChildProcess> {
  const log = fs.openSync(options.stdioLogPath, 'w');
  const child = spawn(
    'mpv',
    [
      '--no-config',
      '--no-terminal',
      `--log-file=${options.logPath}`,
      '--pause',
      '--keep-open=yes',
      '--ao=null',
      `--geometry=${E2E_SCREEN.width}x${E2E_SCREEN.height}+0+0`,
      // Xvfb has no GPU; the software X11 output always works there.
      ...(options.display.kind === 'xvfb' ? ['--vo=x11'] : []),
      `--input-ipc-server=${options.socketPath}`,
      `--sub-file=${options.subtitlePath}`,
      options.mediaPath,
    ],
    { env: applyEnvDelta(process.env, options.display.env), stdio: ['ignore', log, log] },
  );
  fs.closeSync(log);
  await Promise.race([
    once(child, 'spawn'),
    once(child, 'error').then(([error]) => {
      throw new Error(`Could not start mpv (${String(error)}). Is it on PATH?`);
    }),
  ]);
  return child;
}

function tryConnect(socketPath: string): Promise<net.Socket | null> {
  return new Promise((resolve) => {
    const socket = net.connect(socketPath);
    socket.once('connect', () => resolve(socket));
    socket.once('error', () => resolve(null));
  });
}

/**
 * Connects to mpv's JSON IPC socket. Once the socket closes (mpv exited or
 * crashed), pending and later commands reject instead of hanging until the test
 * timeout; `describeClose` adds context such as mpv's exit status.
 */
export async function connectMpv(
  socketPath: string,
  describeClose: () => string = () => 'mpv IPC connection closed',
): Promise<MpvClient> {
  const socket = await waitUntil(() => tryConnect(socketPath), {
    description: `mpv IPC socket ${socketPath}`,
  });
  const pending = new Map<number, (reply: MpvReply) => void>();
  let nextRequestId = 1;
  let buffer = '';
  let closedError: Error | null = null;

  const failPending = () => {
    closedError ??= new Error(describeClose());
    for (const settle of pending.values()) settle({ error: closedError.message });
    pending.clear();
  };
  socket.on('close', failPending);
  // 'close' always follows 'error'; listening here keeps the error from being thrown.
  socket.on('error', () => {});

  socket.setEncoding('utf8');
  socket.on('data', (chunk: string) => {
    buffer += chunk;
    const lines = buffer.split('\n');
    buffer = lines.pop() ?? '';
    for (const line of lines) {
      if (!line.trim()) continue;
      const reply: MpvReply = JSON.parse(line);
      if (reply.request_id === undefined) continue;
      pending.get(reply.request_id)?.(reply);
      pending.delete(reply.request_id);
    }
  });

  return {
    command: (...args) =>
      new Promise((resolve, reject) => {
        if (closedError) {
          reject(new Error(`mpv ${JSON.stringify(args)} failed: ${closedError.message}`));
          return;
        }
        const requestId = nextRequestId++;
        pending.set(requestId, (reply) => {
          if (reply.error === 'success') resolve(reply.data);
          else reject(new Error(`mpv ${JSON.stringify(args)} failed: ${reply.error}`));
        });
        socket.write(`${JSON.stringify({ command: args, request_id: requestId })}\n`);
      }),
    close: () => socket.destroy(),
  };
}
