import { spawn, type ChildProcess } from 'child_process';
import fs from 'fs';
import net, { type Socket } from 'net';
import os from 'os';
import path from 'path';
import { randomUUID } from 'crypto';

const CONNECT_TIMEOUT_MS = 5_000;
const CONNECT_ATTEMPT_TIMEOUT_MS = 500;
const CONNECT_RETRY_MS = 40;
/**
 * mpv flips eof-reached as soon as the decoder passes `end`, while its audio buffer is still
 * draining; keep-open then pauses once the buffer has played out. A preview has ended when
 * both have happened.
 */
const EOF_OBSERVER_ID = 1;
const PAUSE_OBSERVER_ID = 2;

export interface MediaTimingPreviewStartOptions {
  mediaPath: string;
  executablePath?: string;
  audioTrackId?: number;
  volume?: number;
  /** The file keeps source timestamps (a cached remote window); seek with the original times. */
  absoluteTimestamps?: boolean;
}

type PreviewProcess = Pick<ChildProcess, 'kill' | 'once'>;

interface MediaTimingPreviewDeps {
  platform: NodeJS.Platform;
  spawnProcess: (command: string, args: string[]) => PreviewProcess;
  connectSocket: (socketPath: string) => Socket;
  now: () => number;
  schedule: (callback: () => void, delayMs: number) => ReturnType<typeof setTimeout>;
  cancelSchedule: (timeout: ReturnType<typeof setTimeout>) => void;
  removeSocketFile: (socketPath: string) => void;
  createSocketPath: () => string;
}

export function buildMediaTimingPreviewArgs(
  socketPath: string,
  options: MediaTimingPreviewStartOptions,
): string[] {
  const args = [
    '--no-config',
    '--no-video',
    '--audio-display=no',
    '--force-window=no',
    '--idle=yes',
    '--keep-open=yes',
    '--pause=yes',
    '--terminal=no',
    '--msg-level=all=warn',
    `--input-ipc-server=${socketPath}`,
  ];
  if (typeof options.audioTrackId === 'number' && Number.isInteger(options.audioTrackId)) {
    args.push(`--aid=${options.audioTrackId}`);
  }
  if (typeof options.volume === 'number' && Number.isFinite(options.volume)) {
    args.push(`--volume=${Math.max(0, options.volume)}`);
  }
  if (options.absoluteTimestamps) {
    args.push('--rebase-start-time=no');
  }
  args.push('--', options.mediaPath);
  return args;
}

function createDefaultSocketPath(): string {
  const suffix = `${process.pid}-${randomUUID()}`;
  return process.platform === 'win32'
    ? `\\\\.\\pipe\\subminer-timing-preview-${suffix}`
    : path.join(
        // macOS limits Unix socket paths to 104 bytes, while its temp directory can be long.
        process.platform === 'darwin' ? '/tmp' : os.tmpdir(),
        `subminer-timing-preview-${suffix}.sock`,
      );
}

function removePosixSocketFile(socketPath: string): void {
  if (process.platform === 'win32') return;
  try {
    fs.unlinkSync(socketPath);
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code !== 'ENOENT') {
      throw error;
    }
  }
}

export class MediaTimingPreviewSession {
  private readonly deps: MediaTimingPreviewDeps;
  private socketPath: string | null = null;
  private socket: Socket | null = null;
  private process: PreviewProcess | null = null;
  private startupError: Error | null = null;
  private startPromise: Promise<void> | null = null;
  private retryWait: {
    timeout: ReturnType<typeof setTimeout>;
    resolve: () => void;
  } | null = null;
  private disposed = false;
  private readBuffer = '';
  private playing = false;
  private eofReached = false;
  private paused = true;
  private readonly endedListeners = new Set<() => void>();

  constructor(deps: Partial<MediaTimingPreviewDeps> = {}) {
    this.deps = {
      platform: process.platform,
      spawnProcess: (command, args) => spawn(command, args, { stdio: 'ignore' }),
      connectSocket: (socketPath) => net.createConnection(socketPath),
      now: Date.now,
      schedule: (callback, delayMs) => setTimeout(callback, delayMs),
      cancelSchedule: (timeout) => clearTimeout(timeout),
      removeSocketFile: removePosixSocketFile,
      createSocketPath: createDefaultSocketPath,
      ...deps,
    };
  }

  async start(options: MediaTimingPreviewStartOptions): Promise<void> {
    if (this.disposed) throw new Error('Preview session is closed');
    if (this.socket) return;
    if (this.startPromise) return await this.startPromise;

    const startPromise = this.startOnce(options);
    this.startPromise = startPromise;
    try {
      await startPromise;
    } catch (error) {
      this.releaseResources();
      throw error;
    } finally {
      if (this.startPromise === startPromise) this.startPromise = null;
    }
  }

  private async startOnce(options: MediaTimingPreviewStartOptions): Promise<void> {
    const mediaPath = options.mediaPath.trim();
    if (!mediaPath) throw new Error('No media source is available for preview');

    const socketPath = this.deps.createSocketPath();
    this.socketPath = socketPath;
    if (this.deps.platform !== 'win32') {
      this.deps.removeSocketFile(socketPath);
    }

    const command = options.executablePath?.trim() || 'mpv';
    this.startupError = null;
    const child = this.deps.spawnProcess(
      command,
      buildMediaTimingPreviewArgs(socketPath, { ...options, mediaPath }),
    );
    this.process = child;
    child.once('error', (error) => {
      if (this.process !== child) return;
      this.startupError = error;
    });
    child.once('exit', () => {
      if (this.process !== child) return;
      if (!this.socket && !this.disposed && !this.startupError) {
        this.startupError = new Error('The hidden mpv preview player exited during startup');
      }
      this.socket?.destroy();
      this.socket = null;
      this.process = null;
    });

    await this.connectWithRetry(socketPath);
  }

  /**
   * Plays [startTime, endTime) once. mpv stops itself at `end` and, thanks to keep-open,
   * pauses after draining the audio device, so the listener hears the whole clip even on
   * high-latency outputs. onPlaybackEnded fires when mpv reports the end was reached.
   */
  async play(startTime: number, endTime: number): Promise<void> {
    if (!this.socket || this.socket.destroyed) {
      throw new Error('Preview player is not ready');
    }
    if (!Number.isFinite(startTime) || !Number.isFinite(endTime) || endTime <= startTime) {
      throw new Error('Preview timing is invalid');
    }

    this.playing = false;
    this.send(['set_property', 'pause', true]);
    this.send(['seek', startTime, 'absolute+exact']);
    // The option parser wants a time string; a raw JSON number is not accepted for `end`.
    this.send(['set_property', 'end', endTime.toFixed(3)]);
    this.send(['set_property', 'pause', false]);
    // Only the seek's eof-reached=false and the later keep-open pause count for this play.
    this.eofReached = false;
    this.paused = false;
    this.playing = true;
  }

  async stop(): Promise<void> {
    this.playing = false;
    if (!this.socket || this.socket.destroyed) return;
    this.send(['set_property', 'pause', true]);
  }

  onPlaybackEnded(listener: () => void): void {
    this.endedListeners.add(listener);
  }

  private finishPlayback(): void {
    if (!this.playing) return;
    this.playing = false;
    for (const listener of this.endedListeners) listener();
  }

  private handleSocketData(chunk: Buffer | string): void {
    this.readBuffer += chunk.toString();
    let newline = this.readBuffer.indexOf('\n');
    while (newline !== -1) {
      const line = this.readBuffer.slice(0, newline).trim();
      this.readBuffer = this.readBuffer.slice(newline + 1);
      newline = this.readBuffer.indexOf('\n');
      if (!line) continue;
      let message: unknown;
      try {
        message = JSON.parse(line);
      } catch {
        continue;
      }
      if (
        typeof message === 'object' &&
        message !== null &&
        'event' in message &&
        message.event === 'property-change' &&
        'name' in message &&
        'data' in message
      ) {
        this.handlePropertyChange(message.name, message.data);
      }
    }
  }

  private handlePropertyChange(name: unknown, data: unknown): void {
    if (name === 'eof-reached') this.eofReached = data === true;
    else if (name === 'pause') this.paused = data === true;
    else return;
    if (this.playing && this.eofReached && this.paused) this.finishPlayback();
  }

  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    this.releaseResources();
  }

  private releaseResources(): void {
    this.cancelRetryWait();
    try {
      this.send(['quit']);
    } catch {
      // The process may already have exited.
    }
    this.socket?.end();
    this.socket?.destroy();
    this.socket = null;
    const child = this.process;
    this.process = null;
    child?.kill();
    if (this.socketPath && this.deps.platform !== 'win32') {
      try {
        this.deps.removeSocketFile(this.socketPath);
      } catch {
        // mpv may still be releasing the socket. The OS temp directory owns cleanup.
      }
    }
    this.socketPath = null;
  }

  private send(command: Array<string | number | boolean>): void {
    if (!this.socket || this.socket.destroyed) {
      throw new Error('Preview player is not connected');
    }
    this.socket.write(`${JSON.stringify({ command })}\n`);
  }

  private async connectWithRetry(socketPath: string): Promise<void> {
    const deadline = this.deps.now() + CONNECT_TIMEOUT_MS;
    while (!this.disposed && this.deps.now() < deadline) {
      if (this.startupError) {
        throw this.startupError;
      }
      try {
        const remainingMs = deadline - this.deps.now();
        if (remainingMs <= 0) break;
        const socket = await this.connectOnce(
          socketPath,
          Math.min(CONNECT_ATTEMPT_TIMEOUT_MS, remainingMs),
        );
        if (this.disposed) {
          socket.destroy();
          throw new Error('Preview session is closed');
        }
        this.socket = socket;
        this.readBuffer = '';
        socket.on('data', (chunk: Buffer | string) => {
          if (this.socket === socket) this.handleSocketData(chunk);
        });
        socket.once('close', () => this.finishPlayback());
        this.send(['observe_property', EOF_OBSERVER_ID, 'eof-reached']);
        this.send(['observe_property', PAUSE_OBSERVER_ID, 'pause']);
        return;
      } catch {
        if (this.disposed) {
          throw new Error('Preview session is closed');
        }
        const remainingMs = deadline - this.deps.now();
        if (remainingMs <= 0) break;
        await this.waitForRetry(Math.min(CONNECT_RETRY_MS, remainingMs));
      }
    }
    if (this.startupError) {
      throw this.startupError;
    }
    if (this.disposed) {
      throw new Error('Preview session is closed');
    }
    throw new Error('Timed out starting the hidden mpv preview player');
  }

  private waitForRetry(delayMs: number): Promise<void> {
    return new Promise<void>((resolve) => {
      const timeout = this.deps.schedule(() => {
        if (this.retryWait?.timeout === timeout) this.retryWait = null;
        resolve();
      }, delayMs);
      this.retryWait = { timeout, resolve };
    });
  }

  private cancelRetryWait(): void {
    const pending = this.retryWait;
    this.retryWait = null;
    if (!pending) return;
    this.deps.cancelSchedule(pending.timeout);
    pending.resolve();
  }

  private connectOnce(socketPath: string, timeoutMs: number): Promise<Socket> {
    return new Promise<Socket>((resolve, reject) => {
      let timeout: ReturnType<typeof setTimeout> | null = null;
      let settled = false;
      const clearAttemptTimeout = (): void => {
        if (timeout !== null) this.deps.cancelSchedule(timeout);
        timeout = null;
      };
      const socket = this.deps.connectSocket(socketPath);
      const onConnect = (): void => {
        if (settled) return;
        settled = true;
        clearAttemptTimeout();
        socket.off('error', onError);
        socket.on('error', () => {
          socket.destroy();
          if (this.socket === socket) this.socket = null;
        });
        socket.once('close', () => {
          if (this.socket === socket) this.socket = null;
        });
        resolve(socket);
      };
      const onError = (error: Error): void => {
        if (settled) return;
        settled = true;
        clearAttemptTimeout();
        socket.off('connect', onConnect);
        socket.on('error', () => {});
        socket.destroy();
        reject(error);
      };
      socket.once('connect', onConnect);
      socket.once('error', onError);
      timeout = this.deps.schedule(() => {
        if (settled) return;
        settled = true;
        timeout = null;
        socket.off('connect', onConnect);
        socket.off('error', onError);
        socket.on('error', () => {});
        socket.destroy();
        reject(new Error('Timed out connecting to the hidden mpv preview player'));
      }, timeoutMs);
    });
  }
}
