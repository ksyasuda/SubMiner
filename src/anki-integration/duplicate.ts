export interface NoteField {
  value: string;
}

export interface NoteInfo {
  noteId: number;
  fields: Record<string, NoteField>;
}

export interface DuplicateDetectionDeps {
  findNotes: (
    query: string,
    options?: { maxRetries?: number },
  ) => Promise<unknown>;
  notesInfo: (noteIds: number[]) => Promise<unknown>;
  getDeck: () => string | null | undefined;
  resolveFieldName: (noteInfo: NoteInfo, preferredName: string) => string | null;
  logWarn: (message: string, error: unknown) => void;
}

export async function findDuplicateNote(
  expression: string,
  excludeNoteId: number,
  noteInfo: NoteInfo,
  deps: DuplicateDetectionDeps,
): Promise<number | null> {
  let fieldName = "";
  for (const name of Object.keys(noteInfo.fields)) {
    if (
      ["word", "expression"].includes(name.toLowerCase()) &&
      noteInfo.fields[name].value
    ) {
      fieldName = name;
      break;
    }
  }
  if (!fieldName) return null;

  const escapedFieldName = escapeAnkiSearchValue(fieldName);
  const escapedExpression = escapeAnkiSearchValue(expression);
  const deckPrefix = deps.getDeck()
    ? `"deck:${escapeAnkiSearchValue(deps.getDeck()!)}" `
    : "";
  const query = `${deckPrefix}"${escapedFieldName}:${escapedExpression}"`;

  try {
    const noteIds = (await deps.findNotes(query, { maxRetries: 0 }) as number[]);
    return await findFirstExactDuplicateNoteId(
      noteIds,
      excludeNoteId,
      fieldName,
      expression,
      deps,
    );
  } catch (error) {
    deps.logWarn("Duplicate search failed:", error);
    return null;
  }
}

function findFirstExactDuplicateNoteId(
  candidateNoteIds: number[],
  excludeNoteId: number,
  fieldName: string,
  expression: string,
  deps: DuplicateDetectionDeps,
): Promise<number | null> {
  const candidates = candidateNoteIds.filter((id) => id !== excludeNoteId);
  if (candidates.length === 0) {
    return Promise.resolve(null);
  }

  const normalizedExpression = normalizeDuplicateValue(expression);
  const chunkSize = 50;
  return (async () => {
    for (let i = 0; i < candidates.length; i += chunkSize) {
      const chunk = candidates.slice(i, i + chunkSize);
      const notesInfoResult = (await deps.notesInfo(chunk)) as unknown[];
      const notesInfo = notesInfoResult as NoteInfo[];
      for (const noteInfo of notesInfo) {
        const resolvedField = deps.resolveFieldName(noteInfo, fieldName);
        if (!resolvedField) continue;
        const candidateValue = noteInfo.fields[resolvedField]?.value || "";
        if (normalizeDuplicateValue(candidateValue) === normalizedExpression) {
          return noteInfo.noteId;
        }
      }
    }
    return null;
  })();
}

function normalizeDuplicateValue(value: string): string {
  return value.replace(/\s+/g, " ").trim();
}

function escapeAnkiSearchValue(value: string): string {
  return value
    .replace(/\\/g, "\\\\")
    .replace(/"/g, '\\"')
    .replace(/([:*?()[\]{}])/g, "\\$1");
}
