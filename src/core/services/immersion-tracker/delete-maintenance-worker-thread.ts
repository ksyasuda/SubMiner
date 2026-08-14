import { parentPort, workerData } from 'node:worker_threads';
import { executeDeleteMaintenanceTask, type DeleteMaintenanceTask } from './delete-maintenance';

interface DeleteMaintenanceWorkerData {
  dbPath: string;
  task: DeleteMaintenanceTask;
}

if (!parentPort) {
  throw new Error('delete maintenance worker missing parent port');
}

const port = parentPort;
const request = workerData as DeleteMaintenanceWorkerData;

try {
  executeDeleteMaintenanceTask(request.dbPath, request.task);
  port.postMessage({ ok: true });
} catch (error) {
  const message = error instanceof Error ? error.message : String(error);
  port.postMessage({ ok: false, error: message });
}
