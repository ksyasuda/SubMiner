import fs from 'node:fs';
import path from 'node:path';
import { createLogger } from '../../../logger';
import type { VocabularyStatsSummary } from './types';

interface VocabularySummaryWorkerResponse {
  summary?: VocabularyStatsSummary;
  error?: unknown;
}

interface VocabularySummaryWorkerHandle {
  once(event: 'message', listener: (message: VocabularySummaryWorkerResponse) => void): this;
  once(event: 'error', listener: (error: Error) => void): this;
  once(event: 'exit', listener: (code: number) => void): this;
  terminate(): Promise<number>;
}

interface VocabularySummaryWorkerRuntimeOptions {
  resolveWorkerPath?: () => string | null;
  createWorker?: (
    workerPath: string,
    workerData: { dbPath: string; knownWords: string[] | null },
  ) => Promise<VocabularySummaryWorkerHandle>;
  timeoutMs?: number;
  warn?: (message: string, ...meta: unknown[]) => void;
}

export type RunVocabularySummaryTask = (
  dbPath: string,
  knownWords: ReadonlySet<string> | null,
) => Promise<VocabularyStatsSummary>;

export function resolveVocabularySummaryWorkerPath(): string | null {
  const fileName = __filename.endsWith('.ts')
    ? 'vocabulary-summary-worker-thread.ts'
    : 'vocabulary-summary-worker-thread.js';
  const workerPath = path.join(__dirname, fileName);
  return fs.existsSync(workerPath) ? workerPath : null;
}

const logger = createLogger('main:immersion-tracker:vocabulary-summary-worker');
const DEFAULT_WORKER_TIMEOUT_MS = 5 * 60 * 1_000;

export class VocabularySummaryWorkerRuntime {
  private readonly activeWorkers = new Set<VocabularySummaryWorkerHandle>();
  private destroyed = false;

  constructor(private readonly options: VocabularySummaryWorkerRuntimeOptions = {}) {}

  async run(
    dbPath: string,
    knownWords: ReadonlySet<string> | null,
  ): Promise<VocabularyStatsSummary> {
    if (this.destroyed) throw new Error('Vocabulary summary worker is shut down');
    const workerData = { dbPath, knownWords: knownWords ? [...knownWords] : null };
    let worker: VocabularySummaryWorkerHandle;
    try {
      const workerPath = (this.options.resolveWorkerPath ?? resolveVocabularySummaryWorkerPath)();
      if (!workerPath) throw new Error('Emitted vocabulary summary worker module was not found');
      const createWorker =
        this.options.createWorker ??
        (async (resolvedPath, data) => {
          const { Worker } = await import('node:worker_threads');
          return new Worker(resolvedPath, { workerData: data });
        });
      worker = await createWorker(workerPath, workerData);
    } catch (error) {
      if (this.destroyed) throw new Error('Vocabulary summary worker is shut down');
      (this.options.warn ?? logger.warn)(
        'Vocabulary summary worker unavailable; refusing to scan vocabulary on the current thread',
        error,
      );
      throw new Error('Vocabulary summary worker unavailable');
    }

    if (this.destroyed) {
      await worker.terminate().catch(() => undefined);
      throw new Error('Vocabulary summary worker is shut down');
    }

    return new Promise<VocabularyStatsSummary>((resolve, reject) => {
      let settled = false;
      let timeout: ReturnType<typeof setTimeout> | null = null;
      this.activeWorkers.add(worker);
      const settle = (result: VocabularyStatsSummary | Error) => {
        if (settled) return;
        settled = true;
        if (timeout) clearTimeout(timeout);
        this.activeWorkers.delete(worker);
        void worker.terminate().catch(() => undefined);
        if (result instanceof Error) reject(result);
        else resolve(result);
      };
      timeout = setTimeout(
        () => settle(new Error('Vocabulary summary worker timed out')),
        this.options.timeoutMs ?? DEFAULT_WORKER_TIMEOUT_MS,
      );

      worker.once('message', (message) => {
        if (message.summary) {
          settle(message.summary);
          return;
        }
        settle(
          new Error(
            `Vocabulary summary failed: ${String(message.error ?? 'unknown worker error')}`,
          ),
        );
      });
      worker.once('error', (error) => settle(error));
      worker.once('exit', (code) => {
        if (!settled) {
          settle(
            new Error(
              code === 0
                ? 'Vocabulary summary worker exited without a response'
                : `Vocabulary summary worker exited with code ${code}`,
            ),
          );
        }
      });
    });
  }

  destroy(): void {
    if (this.destroyed) return;
    this.destroyed = true;
    for (const worker of this.activeWorkers) {
      void worker.terminate().catch(() => undefined);
    }
    this.activeWorkers.clear();
  }
}
