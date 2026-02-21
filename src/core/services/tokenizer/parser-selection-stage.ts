import { MergedToken, NPlusOneMatchMode, PartOfSpeech } from '../../../types';

interface YomitanParseHeadword {
  term?: unknown;
}

interface YomitanParseSegment {
  text?: string;
  reading?: string;
  headwords?: unknown;
}

interface YomitanParseResultItem {
  source?: unknown;
  index?: unknown;
  content?: unknown;
}

type YomitanParseLine = YomitanParseSegment[];

export interface YomitanParseCandidate {
  source: string;
  index: number;
  tokens: MergedToken[];
}

function isObject(value: unknown): value is Record<string, unknown> {
  return Boolean(value && typeof value === 'object');
}

function isString(value: unknown): value is string {
  return typeof value === 'string';
}

function resolveKnownWordText(
  surface: string,
  headword: string,
  matchMode: NPlusOneMatchMode,
): string {
  return matchMode === 'surface' ? surface : headword;
}

function isKanaChar(char: string): boolean {
  const code = char.codePointAt(0);
  if (code === undefined) {
    return false;
  }

  return (
    (code >= 0x3041 && code <= 0x3096) ||
    (code >= 0x309b && code <= 0x309f) ||
    (code >= 0x30a0 && code <= 0x30fa) ||
    (code >= 0x30fd && code <= 0x30ff)
  );
}

function isYomitanParseLine(value: unknown): value is YomitanParseLine {
  if (!Array.isArray(value)) {
    return false;
  }

  return value.every((segment) => {
    if (!isObject(segment)) {
      return false;
    }

    const candidate = segment as YomitanParseSegment;
    return isString(candidate.text);
  });
}

export function isYomitanParseResultItem(value: unknown): value is YomitanParseResultItem {
  if (!isObject(value)) {
    return false;
  }
  if (!isString((value as YomitanParseResultItem).source)) {
    return false;
  }
  if (!Array.isArray((value as YomitanParseResultItem).content)) {
    return false;
  }
  return true;
}

function isYomitanHeadwordRows(value: unknown): value is YomitanParseHeadword[][] {
  return (
    Array.isArray(value) &&
    value.every(
      (group) =>
        Array.isArray(group) &&
        group.every((item) => isObject(item) && isString((item as YomitanParseHeadword).term)),
    )
  );
}

function extractYomitanHeadword(segment: YomitanParseSegment): string {
  const headwords = segment.headwords;
  if (!isYomitanHeadwordRows(headwords)) {
    return '';
  }

  for (const group of headwords) {
    if (group.length > 0) {
      const firstHeadword = group[0] as YomitanParseHeadword;
      if (isString(firstHeadword?.term)) {
        return firstHeadword.term;
      }
    }
  }

  return '';
}

export function mapYomitanParseResultItemToMergedTokens(
  parseResult: YomitanParseResultItem,
  isKnownWord: (text: string) => boolean,
  knownWordMatchMode: NPlusOneMatchMode,
): YomitanParseCandidate | null {
  const content = parseResult.content;
  if (!Array.isArray(content) || content.length === 0) {
    return null;
  }

  const source = String(parseResult.source ?? '');
  const index =
    typeof parseResult.index === 'number' && Number.isInteger(parseResult.index)
      ? parseResult.index
      : 0;

  const tokens: MergedToken[] = [];
  let charOffset = 0;
  let validLineCount = 0;

  for (const line of content) {
    if (!isYomitanParseLine(line)) {
      continue;
    }
    validLineCount += 1;

    let combinedSurface = '';
    let combinedReading = '';
    let combinedHeadword = '';

    for (const segment of line) {
      const segmentText = segment.text;
      if (!segmentText || segmentText.length === 0) {
        continue;
      }

      combinedSurface += segmentText;
      if (typeof segment.reading === 'string') {
        combinedReading += segment.reading;
      }
      if (!combinedHeadword) {
        combinedHeadword = extractYomitanHeadword(segment);
      }
    }

    if (!combinedSurface) {
      continue;
    }

    const start = charOffset;
    const end = start + combinedSurface.length;
    charOffset = end;
    const headword = combinedHeadword || combinedSurface;

    tokens.push({
      surface: combinedSurface,
      reading: combinedReading,
      headword,
      startPos: start,
      endPos: end,
      partOfSpeech: PartOfSpeech.other,
      pos1: '',
      isMerged: true,
      isNPlusOneTarget: false,
      isKnown: (() => {
        const matchText = resolveKnownWordText(combinedSurface, headword, knownWordMatchMode);
        return matchText ? isKnownWord(matchText) : false;
      })(),
    });
  }

  if (validLineCount === 0 || tokens.length === 0) {
    return null;
  }

  return { source, index, tokens };
}

export function selectBestYomitanParseCandidate(
  candidates: YomitanParseCandidate[],
): MergedToken[] | null {
  if (candidates.length === 0) {
    return null;
  }

  const scanningCandidates = candidates.filter(
    (candidate) => candidate.source === 'scanning-parser',
  );
  const mecabCandidates = candidates.filter((candidate) => candidate.source === 'mecab');

  const getBestByTokenCount = (items: YomitanParseCandidate[]): YomitanParseCandidate | null =>
    items.length === 0
      ? null
      : items.reduce((best, current) =>
          current.tokens.length > best.tokens.length ? current : best,
        );

  const getCandidateScore = (candidate: YomitanParseCandidate): number => {
    const readableTokenCount = candidate.tokens.filter(
      (token) => token.reading.trim().length > 0,
    ).length;
    const suspiciousKanaFragmentCount = candidate.tokens.filter(
      (token) =>
        token.reading.trim().length === 0 &&
        token.surface.length >= 2 &&
        Array.from(token.surface).every((char) => isKanaChar(char)),
    ).length;

    return readableTokenCount * 100 - suspiciousKanaFragmentCount * 50 - candidate.tokens.length;
  };

  const chooseBestCandidate = (items: YomitanParseCandidate[]): YomitanParseCandidate | null => {
    if (items.length === 0) {
      return null;
    }

    return items.reduce((best, current) => {
      const bestScore = getCandidateScore(best);
      const currentScore = getCandidateScore(current);
      if (currentScore !== bestScore) {
        return currentScore > bestScore ? current : best;
      }

      if (current.tokens.length !== best.tokens.length) {
        return current.tokens.length < best.tokens.length ? current : best;
      }

      return best;
    });
  };

  if (scanningCandidates.length > 0) {
    const bestScanning = getBestByTokenCount(scanningCandidates);
    if (bestScanning && bestScanning.tokens.length > 1) {
      return bestScanning.tokens;
    }

    const bestMecab = chooseBestCandidate(mecabCandidates);
    if (bestMecab && bestMecab.tokens.length > (bestScanning?.tokens.length ?? 0)) {
      return bestMecab.tokens;
    }

    return bestScanning ? bestScanning.tokens : null;
  }

  const multiTokenCandidates = candidates.filter((candidate) => candidate.tokens.length > 1);
  const pool = multiTokenCandidates.length > 0 ? multiTokenCandidates : candidates;
  const bestCandidate = chooseBestCandidate(pool);
  return bestCandidate ? bestCandidate.tokens : null;
}

export function selectYomitanParseTokens(
  parseResults: unknown,
  isKnownWord: (text: string) => boolean,
  knownWordMatchMode: NPlusOneMatchMode,
): MergedToken[] | null {
  if (!Array.isArray(parseResults) || parseResults.length === 0) {
    return null;
  }

  const candidates = parseResults
    .filter((item): item is YomitanParseResultItem => isYomitanParseResultItem(item))
    .map((item) => mapYomitanParseResultItemToMergedTokens(item, isKnownWord, knownWordMatchMode))
    .filter((candidate): candidate is YomitanParseCandidate => candidate !== null);

  const bestCandidate = selectBestYomitanParseCandidate(candidates);
  return bestCandidate;
}
