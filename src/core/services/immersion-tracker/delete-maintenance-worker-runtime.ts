import fs from 'node:fs';
import path from 'node:path';
import { createLogger } from '../../../logger';
import { executeDeleteMaintenanceTask, type DeleteMaintenanceTask } from './delete-maintenance';

interface DeleteMaintenanceWorkerResponse {
  ok?: unknown;
  error?: unknown;
}

export type RunDeleteMaintenanceTask = (
  dbPath: string,
  task: DeleteMaintenanceTask,
) => Promise<void>;

interface DeleteMaintenanceWorkerHandle {
  once(event: 'message', listener: (message: DeleteMaintenanceWorkerResponse) => void): this;
  once(event: 'error', listener: (error: Error) => void): this;
  once(event: 'exit', listener: (code: number) => void): this;
  terminate(): Promise<number>;
}

interface DeleteMaintenanceWorkerRuntimeOptions {
  resolveWorkerPath?: () => string | null;
  createWorker?: (
    workerPath: string,
    workerData: { dbPath: string; task: DeleteMaintenanceTask },
  ) => Promise<DeleteMaintenanceWorkerHandle>;
  executeFallback?: typeof executeDeleteMaintenanceTask;
  warn?: (message: string, ...meta: unknown[]) => void;
}

export function resolveDeleteMaintenanceWorkerPath(): string | null {
  // When the process runs TypeScript directly (Bun from source), the emitted
  // .js sibling doesn't exist — spawn the .ts module instead, which such
  // runtimes transpile for workers too. Compiled layouts keep using the .js.
  const fileName = __filename.endsWith('.ts')
    ? 'delete-maintenance-worker-thread.ts'
    : 'delete-maintenance-worker-thread.js';
  const workerPath = path.join(__dirname, fileName);
  return fs.existsSync(workerPath) ? workerPath : null;
}

const logger = createLogger('main:immersion-tracker:delete-worker');

export class DeleteMaintenanceWorkerRuntime {
  private readonly activeWorkers = new Set<DeleteMaintenanceWorkerHandle>();
  private destroyed = false;

  constructor(private readonly options: DeleteMaintenanceWorkerRuntimeOptions = {}) {}

  async run(dbPath: string, task: DeleteMaintenanceTask): Promise<void> {
    if (this.destroyed) {
      throw new Error('Delete maintenance worker is shut down');
    }

    let worker: DeleteMaintenanceWorkerHandle;
    try {
      const workerPath = (this.options.resolveWorkerPath ?? resolveDeleteMaintenanceWorkerPath)();
      if (!workerPath) throw new Error('Emitted delete-maintenance worker module was not found');
      const createWorker =
        this.options.createWorker ??
        (async (resolvedPath, workerData) => {
          const { Worker } = await import('node:worker_threads');
          return new Worker(resolvedPath, { workerData });
        });
      worker = await createWorker(workerPath, { dbPath, task });
    } catch (error) {
      if (this.destroyed) {
        throw new Error('Delete maintenance worker is shut down');
      }
      (this.options.warn ?? logger.warn)(
        'Delete maintenance worker unavailable; running maintenance on the current thread',
        error,
      );
      (this.options.executeFallback ?? executeDeleteMaintenanceTask)(dbPath, task);
      return;
    }

    if (this.destroyed) {
      await worker.terminate().catch(() => undefined);
      throw new Error('Delete maintenance worker is shut down');
    }

    type WorkerOutcome =
      | { kind: 'ok' }
      | { kind: 'task-error'; detail: string }
      | { kind: 'worker-failure'; error: Error };

    const outcome = await new Promise<WorkerOutcome>((resolve) => {
      let settled = false;
      this.activeWorkers.add(worker);

      const settle = (result: WorkerOutcome) => {
        if (settled) return;
        settled = true;
        this.activeWorkers.delete(worker);
        resolve(result);
        void worker.terminate();
      };

      worker.once('message', (message: DeleteMaintenanceWorkerResponse) => {
        if (message.ok === true) {
          settle({ kind: 'ok' });
          return;
        }
        const detail = typeof message.error === 'string' ? message.error : 'unknown worker error';
        settle({ kind: 'task-error', detail });
      });
      worker.once('error', (error) => settle({ kind: 'worker-failure', error }));
      worker.once('exit', (code) => {
        settle({
          kind: 'worker-failure',
          error: new Error(
            code === 0
              ? 'Delete maintenance worker exited without a response'
              : `Delete maintenance worker exited with code ${code}`,
          ),
        });
      });
    });

    if (outcome.kind === 'ok') return;
    // The maintenance itself failed inside the worker — rerunning it on this
    // thread would hit the same error, so surface it instead.
    if (outcome.kind === 'task-error') {
      throw new Error(`Delete maintenance failed: ${outcome.detail}`);
    }
    if (this.destroyed) {
      throw new Error('Delete maintenance worker is shut down');
    }
    // The worker died without reporting a result (failed to load, crashed).
    // Its transaction rolled back with its connection, and a rerun re-plans
    // against the current rows, so falling back on this thread is safe.
    (this.options.warn ?? logger.warn)(
      'Delete maintenance worker failed; running maintenance on the current thread',
      outcome.error,
    );
    (this.options.executeFallback ?? executeDeleteMaintenanceTask)(dbPath, task);
  }

  destroy(): void {
    if (this.destroyed) return;
    this.destroyed = true;
    for (const worker of this.activeWorkers) {
      void worker.terminate();
    }
    this.activeWorkers.clear();
  }
}
