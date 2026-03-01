import type { MergedToken } from '../../../types';
import { createLogger } from '../../../logger';
import { enrichTokensWithMecabPos1 } from './parser-enrichment-stage';

const logger = createLogger('main:tokenizer');
const DISABLE_WORKER_ENV = 'SUBMINER_DISABLE_MECAB_ENRICHMENT_WORKER';

interface WorkerRequest {
  id: number;
  tokens: MergedToken[];
  mecabTokens: MergedToken[] | null;
}

interface WorkerResponse {
  id?: unknown;
  result?: unknown;
  error?: unknown;
}

type PendingRequest = {
  resolve: (value: MergedToken[]) => void;
  reject: (reason?: unknown) => void;
};

class ParserEnrichmentWorkerRuntime {
  private worker: import('node:worker_threads').Worker | null = null;
  private nextRequestId = 1;
  private pending = new Map<number, PendingRequest>();
  private initAttempted = false;

  async enrichTokens(
    tokens: MergedToken[],
    mecabTokens: MergedToken[] | null,
  ): Promise<MergedToken[]> {
    const worker = await this.getWorker();
    if (!worker) {
      return enrichTokensWithMecabPos1(tokens, mecabTokens);
    }

    return new Promise<MergedToken[]>((resolve, reject) => {
      const id = this.nextRequestId++;
      this.pending.set(id, { resolve, reject });
      const request: WorkerRequest = { id, tokens, mecabTokens };
      worker.postMessage(request);
    });
  }

  private async getWorker(): Promise<import('node:worker_threads').Worker | null> {
    if (process.env[DISABLE_WORKER_ENV] === '1') {
      return null;
    }
    if (this.worker) {
      return this.worker;
    }
    if (this.initAttempted) {
      return null;
    }

    this.initAttempted = true;

    let workerThreads: typeof import('node:worker_threads');
    try {
      workerThreads = await import('node:worker_threads');
    } catch {
      return null;
    }

    let workerPath = '';
    try {
      workerPath = require.resolve('./parser-enrichment-worker-thread.js');
    } catch {
      return null;
    }

    try {
      const worker = new workerThreads.Worker(workerPath);
      worker.on('message', (message: WorkerResponse) => this.handleWorkerMessage(message));
      worker.on('error', (error: Error) => this.handleWorkerFailure(error));
      worker.on('exit', (code: number) => {
        if (code !== 0) {
          this.handleWorkerFailure(new Error(`parser enrichment worker exited with code ${code}`));
        } else {
          this.worker = null;
        }
      });
      this.worker = worker;
      return worker;
    } catch (error) {
      logger.debug(`Failed to start parser enrichment worker: ${(error as Error).message}`);
      return null;
    }
  }

  private handleWorkerMessage(message: WorkerResponse): void {
    if (typeof message.id !== 'number') {
      return;
    }

    const request = this.pending.get(message.id);
    if (!request) {
      return;
    }
    this.pending.delete(message.id);

    if (typeof message.error === 'string' && message.error.length > 0) {
      request.reject(new Error(message.error));
      return;
    }

    if (!Array.isArray(message.result)) {
      request.reject(new Error('Parser enrichment worker returned invalid payload'));
      return;
    }

    request.resolve(message.result as MergedToken[]);
  }

  private handleWorkerFailure(error: Error): void {
    logger.debug(
      `Parser enrichment worker unavailable, falling back to main thread: ${error.message}`,
    );
    for (const pending of this.pending.values()) {
      pending.reject(error);
    }
    this.pending.clear();

    if (this.worker) {
      this.worker.removeAllListeners();
      this.worker = null;
    }
  }
}

let runtime: ParserEnrichmentWorkerRuntime | null = null;

export async function enrichTokensWithMecabPos1Async(
  tokens: MergedToken[],
  mecabTokens: MergedToken[] | null,
): Promise<MergedToken[]> {
  if (!runtime) {
    runtime = new ParserEnrichmentWorkerRuntime();
  }

  try {
    return await runtime.enrichTokens(tokens, mecabTokens);
  } catch {
    return enrichTokensWithMecabPos1(tokens, mecabTokens);
  }
}
