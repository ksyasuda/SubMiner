import { parentPort, workerData } from 'node:worker_threads';
import { executeLexicalRollupBackfillTask } from './lexical-rollup-worker';

if (!parentPort) throw new Error('lexical rollup worker missing parent port');

try {
  executeLexicalRollupBackfillTask((workerData as { dbPath: string }).dbPath);
  parentPort.postMessage({ ok: true });
} catch (error) {
  parentPort.postMessage({ error: error instanceof Error ? error.message : String(error) });
}
