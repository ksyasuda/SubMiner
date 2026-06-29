export interface NoteFieldValueInfo {
  fields: Record<string, { value: string }>;
}

export function getNoteFieldValue(
  noteInfo: NoteFieldValueInfo,
  preferredName: string,
): string | null {
  const resolvedFieldName = Object.keys(noteInfo.fields).find(
    (fieldName) => fieldName.toLowerCase() === preferredName.toLowerCase(),
  );
  return resolvedFieldName ? (noteInfo.fields[resolvedFieldName]?.value ?? '') : null;
}

export function hasNoteFieldValue(noteInfo: NoteFieldValueInfo, preferredName: string): boolean {
  return (getNoteFieldValue(noteInfo, preferredName) ?? '').trim().length > 0;
}

export function shouldMarkWordAndSentenceCard(
  noteInfo: NoteFieldValueInfo,
  sentenceCardConfig: { lapisEnabled: boolean; kikuEnabled: boolean },
): boolean {
  if (!sentenceCardConfig.lapisEnabled && !sentenceCardConfig.kikuEnabled) {
    return false;
  }

  const wordAndSentenceValue = getNoteFieldValue(noteInfo, 'IsWordAndSentenceCard');
  if (wordAndSentenceValue === null) {
    return false;
  }
  if (wordAndSentenceValue.trim().length > 0) {
    return true;
  }
  return (
    !hasNoteFieldValue(noteInfo, 'IsSentenceCard') && !hasNoteFieldValue(noteInfo, 'IsAudioCard')
  );
}
