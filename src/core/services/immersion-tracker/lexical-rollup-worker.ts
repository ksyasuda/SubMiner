import { areLexicalDailyRollupsReady, rebuildLexicalDailyRollups } from './lexical-rollups';
import { Database } from './sqlite';
import { applyPragmas } from './storage';

export function executeLexicalRollupBackfillTask(dbPath: string): void {
  const db = new Database(dbPath);
  try {
    applyPragmas(db);
    if (!areLexicalDailyRollupsReady(db)) {
      rebuildLexicalDailyRollups(db);
    }
  } finally {
    db.close();
  }
}
