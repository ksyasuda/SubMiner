import { buildReading, buildReadingFromHint } from './name-reading';
import { expandRawNameVariants, isJapaneseNameSplitCandidate } from './term-building';
import type {
  CharacterRecord,
  NameSplitToken,
  NameSplitTokenizer,
  ResolvedNameSplit,
} from './types';

const NAME_SEPARATOR_PATTERN = /[\s　・･·•]/;

function joinSurfaces(tokens: NameSplitToken[]): string {
  return tokens.map((token) => token.word).join('');
}

// MeCab tags dictionary-known person names as 名詞,固有名詞,人名,姓|名. A split is
// trusted only when every leading token is 姓 and every remaining token is 名.
function splitIndexFromPersonNamePos(tokens: NameSplitToken[]): number | null {
  let familyEnd = 0;
  while (
    familyEnd < tokens.length &&
    tokens[familyEnd]!.pos3 === '人名' &&
    tokens[familyEnd]!.pos4 === '姓'
  ) {
    familyEnd += 1;
  }
  if (familyEnd === 0 || familyEnd >= tokens.length) {
    return null;
  }
  for (let index = familyEnd; index < tokens.length; index += 1) {
    if (tokens[index]!.pos3 !== '人名' || tokens[index]!.pos4 !== '名') {
      return null;
    }
  }
  return familyEnd;
}

function splitIndexFromHintReadings(
  tokens: NameSplitToken[],
  familyHintReading: string,
  givenHintReading: string,
): number | null {
  if (!familyHintReading && !givenHintReading) {
    return null;
  }
  const readings = tokens.map((token) => buildReading(token.katakanaReading || ''));
  let matchedIndex: number | null = null;
  for (let index = 1; index < tokens.length; index += 1) {
    const familyReadings = readings.slice(0, index);
    const givenReadings = readings.slice(index);
    const familyMatches =
      !!familyHintReading &&
      familyReadings.every((reading) => reading.length > 0) &&
      familyReadings.join('') === familyHintReading;
    const givenMatches =
      !!givenHintReading &&
      givenReadings.every((reading) => reading.length > 0) &&
      givenReadings.join('') === givenHintReading;
    if (!familyMatches && !givenMatches) {
      continue;
    }
    if (matchedIndex !== null && matchedIndex !== index) {
      return null;
    }
    matchedIndex = index;
  }
  return matchedIndex;
}

function collectSplitCandidateNames(character: CharacterRecord): string[] {
  const candidates = new Set<string>();
  const rawNames = [character.nativeName, character.fullName, ...character.alternativeNames];
  for (const rawName of rawNames) {
    for (const name of expandRawNameVariants(rawName)) {
      const trimmed = name.trim();
      if (!trimmed || NAME_SEPARATOR_PATTERN.test(trimmed)) continue;
      if (!isJapaneseNameSplitCandidate(trimmed)) continue;
      if ([...trimmed].length < 2) continue;
      candidates.add(trimmed);
    }
  }
  return [...candidates];
}

export async function resolveJapaneseNameSplits(
  characters: CharacterRecord[],
  tokenize: NameSplitTokenizer,
  logWarn?: (message: string) => void,
  onCharacterResolved?: (completed: number, total: number) => void,
): Promise<Map<string, ResolvedNameSplit>> {
  const splits = new Map<string, ResolvedNameSplit>();
  let resolvedCharacters = 0;
  for (const character of characters) {
    const familyHintReading = buildReadingFromHint(character.lastNameHint?.trim() || '');
    const givenHintReading = buildReadingFromHint(character.firstNameHint?.trim() || '');
    for (const name of collectSplitCandidateNames(character)) {
      if (splits.has(name)) continue;
      let tokens: NameSplitToken[] | null = null;
      try {
        tokens = await tokenize(name);
      } catch (err) {
        logWarn?.(
          `[dictionary] name split tokenization failed for "${name}": ${(err as Error).message}`,
        );
        continue;
      }
      if (!tokens || tokens.length < 2 || joinSurfaces(tokens) !== name) continue;
      const splitIndex =
        splitIndexFromPersonNamePos(tokens) ??
        splitIndexFromHintReadings(tokens, familyHintReading, givenHintReading);
      if (splitIndex === null) continue;
      const family = joinSurfaces(tokens.slice(0, splitIndex));
      const given = joinSurfaces(tokens.slice(splitIndex));
      if (family && given) {
        splits.set(name, { family, given });
      }
    }
    resolvedCharacters += 1;
    onCharacterResolved?.(resolvedCharacters, characters.length);
  }
  return splits;
}
