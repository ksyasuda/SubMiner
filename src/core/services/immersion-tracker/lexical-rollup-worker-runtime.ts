import fs from 'node:fs';
import path from 'node:path';
import { createLogger } from '../../../logger';
import { executeLexicalRollupBackfillTask } from './lexical-rollup-worker';

interface WorkerResponse {
  ok?: boolean;
  error?: unknown;
}

interface WorkerHandle {
  once(event: 'message', listener: (message: WorkerResponse) => void): this;
  once(event: 'error', listener: (error: Error) => void): this;
  once(event: 'exit', listener: (code: number) => void): this;
  terminate(): Promise<number>;
}

const logger = createLogger('main:immersion-tracker:lexical-rollup-worker');

export function resolveLexicalRollupWorkerPath(): string | null {
  const fileName = __filename.endsWith('.ts')
    ? 'lexical-rollup-worker-thread.ts'
    : 'lexical-rollup-worker-thread.js';
  const workerPath = path.join(__dirname, fileName);
  return fs.existsSync(workerPath) ? workerPath : null;
}

export class LexicalRollupWorkerRuntime {
  private readonly activeWorkers = new Set<WorkerHandle>();
  private destroyed = false;

  async run(dbPath: string): Promise<void> {
    if (this.destroyed) throw new Error('Lexical rollup worker is shut down');
    let worker: WorkerHandle;
    try {
      const workerPath = resolveLexicalRollupWorkerPath();
      if (!workerPath) throw new Error('Emitted lexical rollup worker module was not found');
      const { Worker } = await import('node:worker_threads');
      worker = new Worker(workerPath, { workerData: { dbPath } });
    } catch (error) {
      if (this.destroyed) throw new Error('Lexical rollup worker is shut down');
      logger.warn(
        'Lexical rollup worker unavailable; running backfill on the current thread',
        error,
      );
      executeLexicalRollupBackfillTask(dbPath);
      return;
    }

    if (this.destroyed) {
      await worker.terminate().catch(() => undefined);
      throw new Error('Lexical rollup worker is shut down');
    }

    return new Promise<void>((resolve, reject) => {
      let settled = false;
      this.activeWorkers.add(worker);
      const settle = (error?: Error) => {
        if (settled) return;
        settled = true;
        this.activeWorkers.delete(worker);
        void worker.terminate();
        if (error) reject(error);
        else resolve();
      };
      worker.once('message', (message) => {
        if (message.ok) settle();
        else
          settle(
            new Error(
              `Lexical rollup backfill failed: ${String(message.error ?? 'unknown error')}`,
            ),
          );
      });
      worker.once('error', (error) => settle(error));
      worker.once('exit', (code) => {
        if (!settled) {
          settle(
            new Error(
              code === 0
                ? 'Lexical rollup worker exited without a response'
                : `Lexical rollup worker exited with code ${code}`,
            ),
          );
        }
      });
    });
  }

  destroy(): void {
    if (this.destroyed) return;
    this.destroyed = true;
    for (const worker of this.activeWorkers) void worker.terminate();
    this.activeWorkers.clear();
  }
}
