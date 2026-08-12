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
  const workerPath = path.join(__dirname, 'delete-maintenance-worker-thread.js');
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
      (this.options.warn ?? logger.warn)(
        'Delete maintenance worker unavailable; running maintenance on the current thread',
        error,
      );
      (this.options.executeFallback ?? executeDeleteMaintenanceTask)(dbPath, task);
      return;
    }

    await new Promise<void>((resolve, reject) => {
      let settled = false;
      this.activeWorkers.add(worker);

      const settle = (error?: Error) => {
        if (settled) return;
        settled = true;
        this.activeWorkers.delete(worker);
        if (error) reject(error);
        else resolve();
        void worker.terminate();
      };

      worker.once('message', (message: DeleteMaintenanceWorkerResponse) => {
        if (message.ok === true) {
          settle();
          return;
        }
        const detail = typeof message.error === 'string' ? message.error : 'unknown worker error';
        settle(new Error(`Delete maintenance failed: ${detail}`));
      });
      worker.once('error', (error) => settle(error));
      worker.once('exit', (code) => {
        settle(
          new Error(
            code === 0
              ? 'Delete maintenance worker exited without a response'
              : `Delete maintenance worker exited with code ${code}`,
          ),
        );
      });
    });
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
