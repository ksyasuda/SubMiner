import type { AnkiConnectConfig } from './types/anki';

type NoteFieldValue = { value?: string } | string | null | undefined;

function normalizeFieldName(value: string | null | undefined): string | null {
  if (typeof value !== 'string') return null;
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : null;
}

export function getConfiguredWordFieldName(
  config?: Pick<AnkiConnectConfig, 'fields'> | null,
): string {
  return normalizeFieldName(config?.fields?.word) ?? 'Expression';
}

export function getConfiguredSentenceFieldName(
  config?: Pick<AnkiConnectConfig, 'fields'> | null,
): string {
  return normalizeFieldName(config?.fields?.sentence) ?? 'Sentence';
}

export function getConfiguredTranslationFieldName(
  config?: Pick<AnkiConnectConfig, 'fields'> | null,
): string {
  return normalizeFieldName(config?.fields?.translation) ?? 'SelectionText';
}

export function getConfiguredWordFieldCandidates(
  config?: Pick<AnkiConnectConfig, 'fields'> | null,
): string[] {
  const preferred = getConfiguredWordFieldName(config);
  const candidates = [preferred, 'Expression', 'Word'];
  const seen = new Set<string>();
  return candidates.filter((candidate) => {
    const key = candidate.toLowerCase();
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function coerceFieldValue(value: NoteFieldValue): string {
  if (typeof value === 'string') return value;
  if (value && typeof value === 'object' && typeof value.value === 'string') {
    return value.value;
  }
  return '';
}

export function stripAnkiFieldHtml(value: string): string {
  return value
    .replace(/\[sound:[^\]]+\]/gi, ' ')
    .replace(/<br\s*\/?>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

export function getPreferredNoteFieldValue(
  fields: Record<string, NoteFieldValue> | null | undefined,
  preferredNames: string[],
): string {
  if (!fields) return '';
  const entries = Object.entries(fields);
  for (const preferredName of preferredNames) {
    const preferredKey = preferredName.trim().toLowerCase();
    if (!preferredKey) continue;
    const entry = entries.find(([fieldName]) => fieldName.trim().toLowerCase() === preferredKey);
    if (!entry) continue;
    const cleaned = stripAnkiFieldHtml(coerceFieldValue(entry[1]));
    if (cleaned) return cleaned;
  }
  return '';
}

export function getPreferredWordValueFromExtractedFields(
  fields: Record<string, string>,
  config?: Pick<AnkiConnectConfig, 'fields'> | null,
): string {
  for (const candidate of getConfiguredWordFieldCandidates(config)) {
    const value = fields[candidate.toLowerCase()]?.trim();
    if (value) return value;
  }
  return '';
}
