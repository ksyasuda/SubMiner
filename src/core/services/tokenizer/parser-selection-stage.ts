import { MergedToken, NPlusOneMatchMode, PartOfSpeech } from '../../../types';
import { isStandaloneGrammarEndingText } from './grammar-ending';
import { isKanaChar, isKanaOnlyText } from './token-classification';

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

function selectMergedHeadword(
  firstHeadword: string,
  expandedHeadwords: string[],
  surface: string,
): string {
  if (expandedHeadwords.length > 0) {
    const exactSurfaceMatch = expandedHeadwords.find((headword) => headword === surface);
    if (exactSurfaceMatch) {
      return exactSurfaceMatch;
    }

    return expandedHeadwords.reduce((best, current) => {
      if (current.length !== best.length) {
        return current.length > best.length ? current : best;
      }
      return best;
    });
  }

  if (!firstHeadword) {
    return '';
  }
  return firstHeadword;
}

function isStandaloneGrammarEndingSegment(segment: YomitanParseSegment): boolean {
  const surface = segment.text?.trim() ?? '';
  const headword = extractYomitanHeadword(segment).trim();
  return (
    headword.length > 0 &&
    (isStandaloneGrammarEndingText(surface) || isStandaloneGrammarEndingText(headword))
  );
}

// Caption-style asides (SFX/speaker labels) start with a bracket and should stay
// non-interactive; dropping punctuation-only runs also keeps the source-text gap that
// sentence-boundary detection relies on (e.g. a dropped 「…」).
const CAPTION_OPENING_BRACKETS = new Set([
  '(',
  '（',
  '[',
  '［',
  '{',
  '｛',
  '「',
  '『',
  '【',
  '〈',
  '《',
  '≪',
  '＜',
  '<',
]);

function hasLookupWorthyText(text: string): boolean {
  return /[\p{L}\p{N}]/u.test(text);
}

function isCaptionLikeUnparsedText(text: string): boolean {
  const firstChar = Array.from(text.trim())[0];
  return firstChar !== undefined && CAPTION_OPENING_BRACKETS.has(firstChar);
}

function shouldEmitUnparsedRunAsToken(text: string): boolean {
  return hasLookupWorthyText(text) && !isCaptionLikeUnparsedText(text);
}

function shouldMergeKanaContinuation(
  previousToken: MergedToken | undefined,
  continuationSurface: string,
): previousToken is MergedToken {
  if (!previousToken || !continuationSurface || !isKanaOnlyText(continuationSurface)) {
    return false;
  }

  if (!previousToken.headword || previousToken.headword.length <= previousToken.surface.length) {
    return false;
  }

  const appendedSurface = previousToken.surface + continuationSurface;
  return previousToken.headword.startsWith(appendedSurface);
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
  let hasDictionaryMatch = false;

  for (const line of content) {
    if (!isYomitanParseLine(line)) {
      continue;
    }
    validLineCount += 1;

    let combinedSurface = '';
    let combinedReading = '';
    let combinedStart = charOffset;
    let firstHeadword = '';
    const expandedHeadwords: string[] = [];

    const pushToken = (
      surface: string,
      reading: string,
      headword: string,
      start: number,
      end: number,
      isUnparsedRun = false,
    ): void => {
      tokens.push({
        surface,
        reading,
        headword,
        startPos: start,
        endPos: end,
        partOfSpeech: PartOfSpeech.other,
        pos1: '',
        isMerged: true,
        isNPlusOneTarget: false,
        isKnown: (() => {
          const matchText = resolveKnownWordText(surface, headword, knownWordMatchMode);
          return matchText ? isKnownWord(matchText) : false;
        })(),
        ...(isUnparsedRun ? { isUnparsedRun: true } : {}),
      });
    };

    const flushCombinedToken = (end: number): void => {
      if (!combinedSurface) {
        combinedStart = end;
        return;
      }

      const combinedHeadword = selectMergedHeadword(
        firstHeadword,
        expandedHeadwords,
        combinedSurface,
      );
      if (!combinedHeadword) {
        const previousToken = tokens[tokens.length - 1];
        if (shouldMergeKanaContinuation(previousToken, combinedSurface)) {
          previousToken.surface += combinedSurface;
          previousToken.reading += combinedReading;
          previousToken.endPos = end;
        } else if (shouldEmitUnparsedRunAsToken(combinedSurface)) {
          // Yomitan couldn't parse this run (e.g. 戻ろ… truncated volitional, ぅ～
          // elongations). Keep it as a token with its surface as headword so it stays
          // hoverable, but flag it so annotation/N+1/vocab logic ignores it — there is
          // no dictionary entry behind it.
          pushToken(combinedSurface, combinedReading, combinedSurface, combinedStart, end, true);
        }
      } else {
        hasDictionaryMatch = true;
        pushToken(combinedSurface, combinedReading, combinedHeadword, combinedStart, end);
      }

      combinedSurface = '';
      combinedReading = '';
      firstHeadword = '';
      expandedHeadwords.length = 0;
      combinedStart = end;
    };

    for (const segment of line) {
      const segmentText = segment.text;
      if (!segmentText || segmentText.length === 0) {
        continue;
      }

      const segmentStart = charOffset;
      const segmentEnd = segmentStart + segmentText.length;
      charOffset = segmentEnd;
      combinedSurface += segmentText;
      if (typeof segment.reading === 'string') {
        combinedReading += segment.reading;
      }
      const segmentHeadword = extractYomitanHeadword(segment);
      if (isStandaloneGrammarEndingSegment(segment)) {
        combinedSurface = combinedSurface.slice(0, -segmentText.length);
        if (typeof segment.reading === 'string' && segment.reading.length > 0) {
          combinedReading = combinedReading.slice(0, -segment.reading.length);
        }
        flushCombinedToken(segmentStart);
        const grammarHeadword = segmentHeadword || segmentText;
        hasDictionaryMatch = true;
        pushToken(
          segmentText,
          typeof segment.reading === 'string' ? segment.reading : '',
          grammarHeadword,
          segmentStart,
          segmentEnd,
        );
        combinedStart = segmentEnd;
        continue;
      }

      if (segmentHeadword) {
        if (!firstHeadword) {
          firstHeadword = segmentHeadword;
        }
        if (segmentHeadword.length > segmentText.length) {
          expandedHeadwords.push(segmentHeadword);
        }
      }
    }

    flushCombinedToken(charOffset);
  }

  if (validLineCount === 0 || tokens.length === 0 || !hasDictionaryMatch) {
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
  if (scanningCandidates.length === 0) {
    return null;
  }

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

  const multiTokenCandidates = scanningCandidates.filter(
    (candidate) => candidate.tokens.length > 1,
  );
  const pool = multiTokenCandidates.length > 0 ? multiTokenCandidates : scanningCandidates;
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
