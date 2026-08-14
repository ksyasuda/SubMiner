import { Database } from './sqlite';
import { applyPragmas } from './storage';
import { deleteAnime, deleteSession, deleteSessions, deleteVideo } from './query-maintenance';
import {
  deleteMaintenanceBatch,
  type DeleteMaintenanceOperation,
} from './query-delete-maintenance';

export type { DeleteMaintenanceOperation } from './query-delete-maintenance';

export type DeleteMaintenanceTask =
  | DeleteMaintenanceOperation
  | { kind: 'batch'; tasks: DeleteMaintenanceOperation[] };

function executeDeleteMaintenanceOperation(
  db: InstanceType<typeof Database>,
  task: DeleteMaintenanceOperation,
): void {
  switch (task.kind) {
    case 'session':
      deleteSession(db, task.sessionId);
      return;
    case 'sessions':
      deleteSessions(db, task.sessionIds);
      return;
    case 'video':
      deleteVideo(db, task.videoId);
      return;
    case 'anime':
      deleteAnime(db, task.animeId);
      return;
  }
}

export function executeDeleteMaintenanceTask(dbPath: string, task: DeleteMaintenanceTask): void {
  const db = new Database(dbPath);
  try {
    applyPragmas(db);
    if (task.kind === 'batch') {
      deleteMaintenanceBatch(db, task.tasks);
      return;
    }
    executeDeleteMaintenanceOperation(db, task);
  } finally {
    db.close();
  }
}
