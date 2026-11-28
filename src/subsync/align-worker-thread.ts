import { parentPort } from 'node:worker_threads';
import { runAlignTask, type AlignWorkerRequest, type AlignWorkerResponse } from './align-worker';

// Worker entry for align-worker.ts: runs one alignment task per message.
if (!parentPort) {
  throw new Error('subsync align worker missing parent port');
}

const port = parentPort;

port.on('message', (message: AlignWorkerRequest) => {
  let response: AlignWorkerResponse;
  try {
    const result = runAlignTask(message.kind, message.args);
    response = { id: message.id, ok: true, result };
  } catch (error) {
    response = {
      id: message.id,
      ok: false,
      error: error instanceof Error ? error.message : String(error),
    };
  }
  port.postMessage(response);
});
