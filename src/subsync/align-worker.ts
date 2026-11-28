import fs from 'node:fs';
import path from 'node:path';
import type { Worker } from 'node:worker_threads';
import { createLogger } from '../logger';
import {
  alignSubtitleFileToReferenceSpans,
  evaluateReferenceAlignmentToAudio,
  evaluateReferenceAlignmentToSampledAudio,
} from './alass-align';

/**
 * Off-main-thread alignment. The alass-style aligner is pure CPU work (up to ~1s per call on
 * a long movie, several calls per auto run), so subsync runs it in a reusable worker thread to
 * keep mpv IPC and the overlay responsive. Aborting a call terminates the worker (stopping the
 * computation mid-way); the next call starts a fresh one. If the worker cannot start, calls run
 * on the current thread instead.
 */

type AlignTasks = {
  alignSubtitle: typeof alignSubtitleFileToReferenceSpans;
  evaluateToAudio: typeof evaluateReferenceAlignmentToAudio;
  evaluateToSampledAudio: typeof evaluateReferenceAlignmentToSampledAudio;
};

export type AlignTaskKind = keyof AlignTasks;
type AlignTaskArgs<K extends AlignTaskKind> = Parameters<AlignTasks[K]>;
type AlignTaskResult<K extends AlignTaskKind> = ReturnType<AlignTasks[K]>;

export interface AlignWorkerRequest {
  id: number;
  kind: AlignTaskKind;
  args: unknown[];
}

export type AlignWorkerResponse =
  | { id: number; ok: true; result: unknown }
  | { id: number; ok: false; error: string };

/**
 * Run one alignment task synchronously. Untyped because it sits on the worker message
 * boundary (shared by the worker thread and the fallback path); the exported `*OffThread`
 * wrappers keep callers type-checked.
 */
export function runAlignTask(kind: AlignTaskKind, args: unknown[]): unknown {
  switch (kind) {
    case 'alignSubtitle':
      return alignSubtitleFileToReferenceSpans(...(args as AlignTaskArgs<'alignSubtitle'>));
    case 'evaluateToAudio':
      return evaluateReferenceAlignmentToAudio(...(args as AlignTaskArgs<'evaluateToAudio'>));
    case 'evaluateToSampledAudio':
      return evaluateReferenceAlignmentToSampledAudio(
        ...(args as AlignTaskArgs<'evaluateToSampledAudio'>),
      );
  }
}

const logger = createLogger('main:subsync:align-worker');
const DISABLE_WORKER_ENV = 'SUBMINER_DISABLE_SUBSYNC_ALIGN_WORKER';
const IDLE_TERMINATE_MS = 30_000;

function resolveAlignWorkerPath(): string | null {
  const fileName = __filename.endsWith('.ts') ? 'align-worker-thread.ts' : 'align-worker-thread.js';
  const workerPath = path.join(__dirname, fileName);
  return fs.existsSync(workerPath) ? workerPath : null;
}

type Pending = {
  kind: AlignTaskKind;
  args: unknown[];
  resolve: (value: unknown) => void;
  reject: (reason: Error) => void;
};

class AlignWorkerRuntime {
  private worker: Worker | null = null;
  private workerUnavailable = false;
  private nextId = 1;
  private readonly pending = new Map<number, Pending>();
  private idleTimer: ReturnType<typeof setTimeout> | null = null;

  async run<K extends AlignTaskKind>(
    kind: K,
    args: AlignTaskArgs<K>,
    signal?: AbortSignal,
  ): Promise<AlignTaskResult<K>> {
    signal?.throwIfAborted();
    const worker = await this.getWorker();
    signal?.throwIfAborted();
    if (!worker) return runAlignTask(kind, args) as AlignTaskResult<K>;

    const id = this.nextId++;
    this.clearIdleTimer();
    const result = new Promise<unknown>((resolve, reject) => {
      this.pending.set(id, { kind, args, resolve, reject });
    });
    const onAbort = (): void => this.reset(signal?.reason ?? new Error('aborted'));
    signal?.addEventListener('abort', onAbort, { once: true });
    const request: AlignWorkerRequest = { id, kind, args };
    worker.postMessage(request);
    try {
      return (await result) as AlignTaskResult<K>;
    } finally {
      signal?.removeEventListener('abort', onAbort);
      if (this.pending.size === 0) this.scheduleIdleTerminate();
    }
  }

  private async getWorker(): Promise<Worker | null> {
    if (this.worker) return this.worker;
    if (this.workerUnavailable || process.env[DISABLE_WORKER_ENV] === '1') return null;
    const workerPath = resolveAlignWorkerPath();
    if (!workerPath) {
      this.markUnavailable(new Error('align worker module not found'));
      return null;
    }
    try {
      const { Worker: WorkerCtor } = await import('node:worker_threads');
      const worker = new WorkerCtor(workerPath);
      let online = false;
      worker.unref();
      worker.once('online', () => {
        online = true;
      });
      worker.on('message', (message: AlignWorkerResponse) => this.handleMessage(message));
      worker.on('error', (error: Error) => {
        // Failing to start (e.g. the module cannot load) is not the task's fault: finish the
        // queued calls on this thread and stop using the worker.
        if (!online) {
          this.markUnavailable(error);
          this.runPendingInThread();
          return;
        }
        this.reset(error);
      });
      worker.on('exit', (code) => {
        if (this.worker !== worker) return;
        this.worker = null;
        if (this.pending.size > 0) {
          this.reset(new Error(`subsync align worker exited with code ${code}`));
        }
      });
      this.worker = worker;
      return worker;
    } catch (error) {
      this.markUnavailable(error as Error);
      return null;
    }
  }

  private handleMessage(message: AlignWorkerResponse): void {
    const pending = this.pending.get(message.id);
    if (!pending) return;
    this.pending.delete(message.id);
    if (message.ok) pending.resolve(message.result);
    else pending.reject(new Error(message.error));
  }

  private runPendingInThread(): void {
    const worker = this.worker;
    this.worker = null;
    if (worker) void worker.terminate().catch(() => undefined);
    const queued = [...this.pending.values()];
    this.pending.clear();
    for (const pending of queued) {
      try {
        pending.resolve(runAlignTask(pending.kind, pending.args));
      } catch (error) {
        pending.reject(error instanceof Error ? error : new Error(String(error)));
      }
    }
  }

  /** Stop the worker (and any computation in it) and fail every in-flight call. */
  private reset(reason: unknown): void {
    const error = reason instanceof Error ? reason : new Error(String(reason));
    const worker = this.worker;
    this.worker = null;
    this.clearIdleTimer();
    if (worker) void worker.terminate().catch(() => undefined);
    for (const pending of this.pending.values()) pending.reject(error);
    this.pending.clear();
  }

  private markUnavailable(error: Error): void {
    if (!this.workerUnavailable) {
      logger.warn(
        `Subsync align worker unavailable, aligning on the main thread: ${error.message}`,
      );
    }
    this.workerUnavailable = true;
  }

  private scheduleIdleTerminate(): void {
    this.clearIdleTimer();
    this.idleTimer = setTimeout(() => {
      if (this.pending.size === 0) this.reset(new Error('idle'));
    }, IDLE_TERMINATE_MS);
    this.idleTimer.unref?.();
  }

  private clearIdleTimer(): void {
    if (this.idleTimer) clearTimeout(this.idleTimer);
    this.idleTimer = null;
  }
}

const runtime = new AlignWorkerRuntime();

/** `alignSubtitleFileToReferenceSpans` off the main thread; rejects promptly when `signal` aborts. */
export function alignSubtitleOffThread(
  signal: AbortSignal | undefined,
  ...args: AlignTaskArgs<'alignSubtitle'>
): Promise<AlignTaskResult<'alignSubtitle'>> {
  return runtime.run('alignSubtitle', args, signal);
}

/** `evaluateReferenceAlignmentToAudio` off the main thread. */
export function evaluateToAudioOffThread(
  signal: AbortSignal | undefined,
  ...args: AlignTaskArgs<'evaluateToAudio'>
): Promise<AlignTaskResult<'evaluateToAudio'>> {
  return runtime.run('evaluateToAudio', args, signal);
}

/** `evaluateReferenceAlignmentToSampledAudio` off the main thread. */
export function evaluateToSampledAudioOffThread(
  signal: AbortSignal | undefined,
  ...args: AlignTaskArgs<'evaluateToSampledAudio'>
): Promise<AlignTaskResult<'evaluateToSampledAudio'>> {
  return runtime.run('evaluateToSampledAudio', args, signal);
}
