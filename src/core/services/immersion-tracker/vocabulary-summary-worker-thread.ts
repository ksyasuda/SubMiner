import { parentPort, workerData } from 'node:worker_threads';
import { executeVocabularySummaryTask } from './vocabulary-summary-worker';

interface VocabularySummaryWorkerData {
  dbPath: string;
  knownWords: string[] | null;
}

if (!parentPort) throw new Error('vocabulary summary worker missing parent port');

const request = workerData as VocabularySummaryWorkerData;

try {
  parentPort.postMessage({
    summary: executeVocabularySummaryTask(request.dbPath, request.knownWords),
  });
} catch (error) {
  parentPort.postMessage({ error: error instanceof Error ? error.message : String(error) });
}
