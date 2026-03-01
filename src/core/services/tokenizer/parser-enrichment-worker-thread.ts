import { parentPort } from 'node:worker_threads';
import type { MergedToken } from '../../../types';
import { enrichTokensWithMecabPos1 } from './parser-enrichment-stage';

interface WorkerRequest {
  id: number;
  tokens: MergedToken[];
  mecabTokens: MergedToken[] | null;
}

if (!parentPort) {
  throw new Error('parser-enrichment worker missing parent port');
}

const port = parentPort;

port.on('message', (message: WorkerRequest) => {
  try {
    const result = enrichTokensWithMecabPos1(message.tokens, message.mecabTokens);
    port.postMessage({ id: message.id, result });
  } catch (error) {
    const messageText = error instanceof Error ? error.message : String(error);
    port.postMessage({ id: message.id, error: messageText });
  }
});
