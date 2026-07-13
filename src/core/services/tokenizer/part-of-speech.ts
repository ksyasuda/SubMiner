import { PartOfSpeech } from '../../../types';
import { normalizePosTag, splitPosTag } from './token-classification';

export function isPartOfSpeechValue(value: unknown): value is PartOfSpeech {
  return typeof value === 'string' && Object.values(PartOfSpeech).includes(value as PartOfSpeech);
}

export function mapMecabPos1ToPartOfSpeech(pos1: string | null | undefined): PartOfSpeech {
  switch (normalizePosTag(pos1)) {
    case '名詞':
      return PartOfSpeech.noun;
    case '動詞':
      return PartOfSpeech.verb;
    case '形容詞':
      return PartOfSpeech.i_adjective;
    case '形状詞':
    case '形容動詞':
      return PartOfSpeech.na_adjective;
    case '助詞':
      return PartOfSpeech.particle;
    case '助動詞':
      return PartOfSpeech.bound_auxiliary;
    case '記号':
    case '補助記号':
      return PartOfSpeech.symbol;
    default:
      return PartOfSpeech.other;
  }
}

export function deriveStoredPartOfSpeech(input: {
  partOfSpeech?: string | null;
  pos1?: string | null;
}): PartOfSpeech {
  const pos1Parts = splitPosTag(input.pos1);

  if (pos1Parts.length > 0) {
    const derivedParts = [...new Set(pos1Parts.map((part) => mapMecabPos1ToPartOfSpeech(part)))];
    if (derivedParts.length === 1) {
      return derivedParts[0]!;
    }
    return PartOfSpeech.other;
  }

  if (isPartOfSpeechValue(input.partOfSpeech)) {
    return input.partOfSpeech;
  }

  return PartOfSpeech.other;
}
