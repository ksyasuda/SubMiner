const MIN_FLATTENED_DUPLICATE_LENGTH = 16;
const TERMINAL_SENTENCE_PUNCTUATION = /[.!?。！？…⋯]+$/gu;

/**
 * Identifies long lines that become duplicates when positioned ASS events are
 * flattened into the secondary subtitle bar. Short dialogue stays distinct.
 */
export function flattenedSecondarySubtitleLineIdentity(text: string): string | null {
  const identity = text
    .normalize('NFKC')
    .replace(/\s+/gu, '')
    .replace(TERMINAL_SENTENCE_PUNCTUATION, '');

  return identity.length >= MIN_FLATTENED_DUPLICATE_LENGTH ? identity : null;
}
