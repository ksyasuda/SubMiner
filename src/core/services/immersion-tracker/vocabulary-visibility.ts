import { PartOfSpeech, type MergedToken } from '../../../types';
import { shouldExcludeTokenFromVocabularyPersistence } from '../tokenizer/annotation-stage';

export interface VocabularyVisibilityRow {
  word: string | null;
  headword: string | null;
  reading?: string | null;
  partOfSpeech?: string | null;
  pos1?: string | null;
  pos2?: string | null;
  pos3?: string | null;
  frequencyRank?: number | null;
}

function toVocabularyToken(row: VocabularyVisibilityRow): MergedToken {
  const word = row.word ?? '';
  const headword = row.headword ?? word;
  const partOfSpeech =
    row.partOfSpeech && Object.values(PartOfSpeech).includes(row.partOfSpeech as PartOfSpeech)
      ? (row.partOfSpeech as PartOfSpeech)
      : PartOfSpeech.other;

  return {
    surface: word,
    reading: row.reading ?? '',
    headword,
    startPos: 0,
    endPos: word.length,
    partOfSpeech,
    pos1: row.pos1 ?? '',
    pos2: row.pos2 ?? '',
    pos3: row.pos3 ?? '',
    frequencyRank: row.frequencyRank ?? undefined,
    isMerged: false,
    isKnown: false,
    isNPlusOneTarget: false,
  };
}

export function isVocabularyStatsRowVisible(row: VocabularyVisibilityRow): boolean {
  if (!(row.word?.trim() || row.headword?.trim())) return false;
  return !shouldExcludeTokenFromVocabularyPersistence(toVocabularyToken(row));
}
