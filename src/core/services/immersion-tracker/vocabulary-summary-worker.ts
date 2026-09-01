import { getVocabularySummary } from './query-lexical';
import { Database } from './sqlite';
import { applyPragmas } from './storage';
import type { VocabularyStatsSummary } from './types';

export function executeVocabularySummaryTask(
  dbPath: string,
  knownWords: string[] | null,
): VocabularyStatsSummary {
  const db = new Database(dbPath);
  try {
    applyPragmas(db);
    return getVocabularySummary(db, knownWords ? new Set(knownWords) : null);
  } finally {
    db.close();
  }
}
