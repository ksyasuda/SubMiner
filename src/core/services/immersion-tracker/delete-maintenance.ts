import { Database } from './sqlite';
import { applyPragmas } from './storage';
import {
  deleteMaintenanceBatch,
  type DeleteMaintenanceOperation,
} from './query-delete-maintenance';

export type { DeleteMaintenanceOperation } from './query-delete-maintenance';

export type DeleteMaintenanceTask =
  | DeleteMaintenanceOperation
  | { kind: 'batch'; tasks: DeleteMaintenanceOperation[] };

export function executeDeleteMaintenanceTask(dbPath: string, task: DeleteMaintenanceTask): void {
  const db = new Database(dbPath);
  try {
    applyPragmas(db);
    deleteMaintenanceBatch(db, task.kind === 'batch' ? task.tasks : [task]);
  } finally {
    db.close();
  }
}
