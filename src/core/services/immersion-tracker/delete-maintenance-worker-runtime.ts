import { executeDeleteMaintenanceTask, type DeleteMaintenanceTask } from './delete-maintenance';

interface DeleteMaintenanceWorkerResponse {
  ok?: unknown;
  error?: unknown;
}

export type RunDeleteMaintenanceTask = (
  dbPath: string,
  task: DeleteMaintenanceTask,
) => Promise<void>;

export class DeleteMaintenanceWorkerRuntime {
  private readonly activeWorkers = new Set<import('node:worker_threads').Worker>();
  private destroyed = false;

  async run(dbPath: string, task: DeleteMaintenanceTask): Promise<void> {
    if (this.destroyed) {
      throw new Error('Delete maintenance worker is shut down');
    }

    let workerThreads: typeof import('node:worker_threads');
    let workerPath: string;
    try {
      workerThreads = await import('node:worker_threads');
      workerPath = require.resolve('./delete-maintenance-worker-thread.js');
    } catch {
      executeDeleteMaintenanceTask(dbPath, task);
      return;
    }

    await new Promise<void>((resolve, reject) => {
      let settled = false;
      const worker = new workerThreads.Worker(workerPath, { workerData: { dbPath, task } });
      this.activeWorkers.add(worker);

      const settle = (error?: Error) => {
        if (settled) return;
        settled = true;
        this.activeWorkers.delete(worker);
        if (error) reject(error);
        else resolve();
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
