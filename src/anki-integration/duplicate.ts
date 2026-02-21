export interface NoteField {
  value: string;
}

export interface NoteInfo {
  noteId: number;
  fields: Record<string, NoteField>;
}

export interface DuplicateDetectionDeps {
  findNotes: (query: string, options?: { maxRetries?: number }) => Promise<unknown>;
  notesInfo: (noteIds: number[]) => Promise<unknown>;
  getDeck: () => string | null | undefined;
  resolveFieldName: (noteInfo: NoteInfo, preferredName: string) => string | null;
  logInfo?: (message: string) => void;
  logDebug?: (message: string) => void;
  logWarn: (message: string, error: unknown) => void;
}

export async function findDuplicateNote(
  expression: string,
  excludeNoteId: number,
  noteInfo: NoteInfo,
  deps: DuplicateDetectionDeps,
): Promise<number | null> {
  const sourceCandidates = getDuplicateSourceCandidates(noteInfo, expression);
  if (sourceCandidates.length === 0) return null;
  deps.logInfo?.(
    `[duplicate] start expr="${expression}" sourceCandidates=${sourceCandidates
      .map((entry) => `${entry.fieldName}:${entry.value}`)
      .join('|')}`,
  );

  const deckValue = deps.getDeck();
  const queryPrefixes = deckValue
    ? [`"deck:${escapeAnkiSearchValue(deckValue)}" `, '']
    : [''];

  try {
    const noteIds = new Set<number>();
    const executedQueries = new Set<string>();
    for (const queryPrefix of queryPrefixes) {
      for (const sourceCandidate of sourceCandidates) {
        const escapedExpression = escapeAnkiSearchValue(sourceCandidate.value);
        const queryFieldNames = getDuplicateCandidateFieldNames(sourceCandidate.fieldName);
        for (const queryFieldName of queryFieldNames) {
          const escapedFieldName = escapeAnkiSearchValue(queryFieldName);
          const query = `${queryPrefix}"${escapedFieldName}:${escapedExpression}"`;
          if (executedQueries.has(query)) continue;
          executedQueries.add(query);
          const results = (await deps.findNotes(query)) as number[];
          deps.logDebug?.(
            `[duplicate] query(field)="${query}" hits=${Array.isArray(results) ? results.length : 0}`,
          );
          for (const noteId of results) {
            noteIds.add(noteId);
          }
        }
      }
      if (noteIds.size > 0) break;
    }

    if (noteIds.size === 0) {
      for (const queryPrefix of queryPrefixes) {
        for (const sourceCandidate of sourceCandidates) {
          const escapedExpression = escapeAnkiSearchValue(sourceCandidate.value);
          const query = `${queryPrefix}"${escapedExpression}"`;
          if (executedQueries.has(query)) continue;
          executedQueries.add(query);
          const results = (await deps.findNotes(query)) as number[];
          deps.logDebug?.(
            `[duplicate] query(text)="${query}" hits=${Array.isArray(results) ? results.length : 0}`,
          );
          for (const noteId of results) {
            noteIds.add(noteId);
          }
        }
        if (noteIds.size > 0) break;
      }
    }

    return await findFirstExactDuplicateNoteId(
      noteIds,
      excludeNoteId,
      sourceCandidates.map((candidate) => candidate.value),
      deps,
    );
  } catch (error) {
    deps.logWarn('Duplicate search failed:', error);
    return null;
  }
}

function findFirstExactDuplicateNoteId(
  candidateNoteIds: Iterable<number>,
  excludeNoteId: number,
  sourceValues: string[],
  deps: DuplicateDetectionDeps,
): Promise<number | null> {
  const candidates = Array.from(candidateNoteIds).filter((id) => id !== excludeNoteId);
  deps.logDebug?.(`[duplicate] candidateIds=${candidates.length} exclude=${excludeNoteId}`);
  if (candidates.length === 0) {
    deps.logInfo?.('[duplicate] no candidates after query + exclude');
    return Promise.resolve(null);
  }

  const normalizedValues = new Set(
    sourceValues.map((value) => normalizeDuplicateValue(value)).filter((value) => value.length > 0),
  );
  if (normalizedValues.size === 0) {
    return Promise.resolve(null);
  }

  const chunkSize = 50;
  return (async () => {
    for (let i = 0; i < candidates.length; i += chunkSize) {
      const chunk = candidates.slice(i, i + chunkSize);
      const notesInfoResult = (await deps.notesInfo(chunk)) as unknown[];
      const notesInfo = notesInfoResult as NoteInfo[];
      for (const noteInfo of notesInfo) {
        const candidateFieldNames = ['word', 'expression'];
        for (const candidateFieldName of candidateFieldNames) {
          const resolvedField = deps.resolveFieldName(noteInfo, candidateFieldName);
          if (!resolvedField) continue;
          const candidateValue = noteInfo.fields[resolvedField]?.value || '';
          if (normalizedValues.has(normalizeDuplicateValue(candidateValue))) {
            deps.logDebug?.(
              `[duplicate] exact-match noteId=${noteInfo.noteId} field=${resolvedField}`,
            );
            deps.logInfo?.(`[duplicate] matched noteId=${noteInfo.noteId} field=${resolvedField}`);
            return noteInfo.noteId;
          }
        }
      }
    }
    deps.logInfo?.('[duplicate] no exact match in candidate notes');
    return null;
  })();
}

function getDuplicateCandidateFieldNames(fieldName: string): string[] {
  const candidates = [fieldName];
  const lower = fieldName.toLowerCase();
  if (lower === 'word') {
    candidates.push('expression');
  } else if (lower === 'expression') {
    candidates.push('word');
  }
  return candidates;
}

function getDuplicateSourceCandidates(
  noteInfo: NoteInfo,
  fallbackExpression: string,
): Array<{ fieldName: string; value: string }> {
  const candidates: Array<{ fieldName: string; value: string }> = [];
  const dedupeKey = new Set<string>();

  for (const fieldName of Object.keys(noteInfo.fields)) {
    const lower = fieldName.toLowerCase();
    if (lower !== 'word' && lower !== 'expression') continue;
    const value = noteInfo.fields[fieldName]?.value?.trim() ?? '';
    if (!value) continue;
    const key = `${lower}:${normalizeDuplicateValue(value)}`;
    if (dedupeKey.has(key)) continue;
    dedupeKey.add(key);
    candidates.push({ fieldName, value });
  }

  const trimmedFallback = fallbackExpression.trim();
  if (trimmedFallback.length > 0) {
    const fallbackKey = `expression:${normalizeDuplicateValue(trimmedFallback)}`;
    if (!dedupeKey.has(fallbackKey)) {
      candidates.push({ fieldName: 'expression', value: trimmedFallback });
    }
  }

  return candidates;
}

function normalizeDuplicateValue(value: string): string {
  return value
    .replace(/<[^>]*>/g, '')
    .replace(/([^\s\[\]]+)\[[^\]]*\]/g, '$1')
    .replace(/\s+/g, ' ')
    .trim();
}

function escapeAnkiSearchValue(value: string): string {
  return value
    .replace(/\\/g, '\\\\')
    .replace(/"/g, '\\"')
    .replace(/([:*?()[\]{}])/g, '\\$1');
}
