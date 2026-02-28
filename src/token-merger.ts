/*
 * SubMiner - All-in-one sentence mining overlay
 * Copyright (C) 2024 sudacode
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <https://www.gnu.org/licenses/>.
 */

import { PartOfSpeech, Token, MergedToken } from './types';

export function isNoun(tok: Token): boolean {
  return tok.partOfSpeech === PartOfSpeech.noun;
}

export function isProperNoun(tok: Token): boolean {
  return tok.partOfSpeech === PartOfSpeech.noun && tok.pos2 === '固有名詞';
}

export function ignoreReading(tok: Token): boolean {
  return tok.partOfSpeech === PartOfSpeech.symbol && tok.pos2 === '文字';
}

export function isCopula(tok: Token): boolean {
  const raw = tok.inflectionType;
  if (!raw) {
    return false;
  }
  return ['特殊・ダ', '特殊・デス', '特殊|だ', '特殊|デス'].includes(raw);
}

export function isAuxVerb(tok: Token): boolean {
  return tok.partOfSpeech === PartOfSpeech.bound_auxiliary && !isCopula(tok);
}

export function isContinuativeForm(tok: Token): boolean {
  if (!tok.inflectionForm) {
    return false;
  }
  const inflectionForm = tok.inflectionForm;
  const isContinuative =
    inflectionForm === '連用デ接続' ||
    inflectionForm === '連用タ接続' ||
    inflectionForm.startsWith('連用形');

  if (!isContinuative) {
    return false;
  }
  return tok.headword !== 'ない';
}

export function isVerbSuffix(tok: Token): boolean {
  return tok.partOfSpeech === PartOfSpeech.verb && (tok.pos2 === '非自立' || tok.pos2 === '接尾');
}

export function isTatteParticle(tok: Token): boolean {
  return (
    tok.partOfSpeech === PartOfSpeech.particle &&
    tok.pos2 === '接続助詞' &&
    tok.headword === 'たって'
  );
}

export function isBaParticle(tok: Token): boolean {
  return tok.partOfSpeech === PartOfSpeech.particle && tok.pos2 === '接続助詞' && tok.word === 'ば';
}

export function isTeDeParticle(tok: Token): boolean {
  return (
    tok.partOfSpeech === PartOfSpeech.particle &&
    tok.pos2 === '接続助詞' &&
    ['て', 'で', 'ちゃ'].includes(tok.word)
  );
}

export function isTaDaParticle(tok: Token): boolean {
  return isAuxVerb(tok) && ['た', 'だ'].includes(tok.word);
}

export function isVerb(tok: Token): boolean {
  return [PartOfSpeech.verb, PartOfSpeech.bound_auxiliary].includes(tok.partOfSpeech);
}

export function isVerbNonIndependent(): boolean {
  return true;
}

export function canReceiveAuxiliary(tok: Token): boolean {
  return [PartOfSpeech.verb, PartOfSpeech.bound_auxiliary, PartOfSpeech.i_adjective].includes(
    tok.partOfSpeech,
  );
}

export function isNounSuffix(tok: Token): boolean {
  return tok.partOfSpeech === PartOfSpeech.verb && tok.pos2 === '接尾';
}

export function isCounter(tok: Token): boolean {
  return (
    tok.partOfSpeech === PartOfSpeech.noun &&
    tok.pos3 !== undefined &&
    tok.pos3.startsWith('助数詞')
  );
}

export function isNumeral(tok: Token): boolean {
  return (
    tok.partOfSpeech === PartOfSpeech.noun && tok.pos2 !== undefined && tok.pos2.startsWith('数')
  );
}

export function shouldMerge(lastStandaloneToken: Token, token: Token): boolean {
  if (isVerb(lastStandaloneToken)) {
    if (isAuxVerb(token)) {
      return true;
    }
    if (isContinuativeForm(lastStandaloneToken) && isVerbSuffix(token)) {
      return true;
    }
    if (isVerbSuffix(token) && isVerbNonIndependent()) {
      return true;
    }
  }

  if (isNoun(lastStandaloneToken) && !isProperNoun(lastStandaloneToken) && isNounSuffix(token)) {
    return true;
  }

  if (isCounter(token) && isNumeral(lastStandaloneToken)) {
    return true;
  }

  if (isBaParticle(token) && canReceiveAuxiliary(lastStandaloneToken)) {
    return true;
  }

  if (isTatteParticle(token) && canReceiveAuxiliary(lastStandaloneToken)) {
    return true;
  }

  if (isTeDeParticle(token) && isContinuativeForm(lastStandaloneToken)) {
    return true;
  }

  if (isTaDaParticle(token) && canReceiveAuxiliary(lastStandaloneToken)) {
    return true;
  }

  if (isTeDeParticle(lastStandaloneToken) && isVerbSuffix(token)) {
    return true;
  }

  return false;
}

export function mergeTokens(
  tokens: Token[],
  isKnownWord: (text: string) => boolean = () => false,
  knownWordMatchMode: 'headword' | 'surface' = 'headword',
): MergedToken[] {
  if (!tokens || tokens.length === 0) {
    return [];
  }

  const result: MergedToken[] = [];
  let charOffset = 0;
  let lastStandaloneToken: Token | null = null;

  for (const token of tokens) {
    const start = charOffset;
    const end = charOffset + token.word.length;
    charOffset = end;

    let shouldMergeToken = false;

    if (result.length > 0 && lastStandaloneToken !== null) {
      shouldMergeToken = shouldMerge(lastStandaloneToken, token);
    }

    const tokenReading = ignoreReading(token) ? '' : token.katakanaReading || token.word;

    if (shouldMergeToken && result.length > 0) {
      const prev = result.pop()!;
      const mergedHeadword = prev.headword;
      const headwordForKnownMatch = (() => {
        if (knownWordMatchMode === 'surface') {
          return prev.surface;
        }
        return mergedHeadword;
      })();
      result.push({
        surface: prev.surface + token.word,
        reading: prev.reading + tokenReading,
        headword: prev.headword,
        startPos: prev.startPos,
        endPos: end,
        partOfSpeech: prev.partOfSpeech,
        pos1: prev.pos1 ?? token.pos1,
        pos2: prev.pos2 ?? token.pos2,
        pos3: prev.pos3 ?? token.pos3,
        isMerged: true,
        isKnown: headwordForKnownMatch ? isKnownWord(headwordForKnownMatch) : false,
        isNPlusOneTarget: false,
      });
    } else {
      const headwordForKnownMatch = (() => {
        if (knownWordMatchMode === 'surface') {
          return token.word;
        }
        return token.headword;
      })();
      result.push({
        surface: token.word,
        reading: tokenReading,
        headword: token.headword,
        startPos: start,
        endPos: end,
        partOfSpeech: token.partOfSpeech,
        pos1: token.pos1,
        pos2: token.pos2,
        pos3: token.pos3,
        isMerged: false,
        isKnown: headwordForKnownMatch ? isKnownWord(headwordForKnownMatch) : false,
        isNPlusOneTarget: false,
      });
    }

    lastStandaloneToken = token;
  }

  return result;
}

const SENTENCE_BOUNDARY_SURFACES = new Set(['。', '？', '！', '?', '!', '…', '\u2026']);
const N_PLUS_ONE_IGNORED_POS1 = new Set(['助詞', '助動詞', '記号', '補助記号']);

export function isNPlusOneCandidateToken(token: MergedToken): boolean {
  if (token.isKnown) {
    return false;
  }
  return isNPlusOneWordCountToken(token);
}

function isNPlusOneWordCountToken(token: MergedToken): boolean {
  if (token.partOfSpeech === PartOfSpeech.particle) {
    return false;
  }

  if (token.partOfSpeech === PartOfSpeech.bound_auxiliary) {
    return false;
  }

  if (token.partOfSpeech === PartOfSpeech.symbol) {
    return false;
  }

  if (token.partOfSpeech === PartOfSpeech.noun && token.pos2 === '固有名詞') {
    return false;
  }

  if (token.pos3 && token.pos3.startsWith('助数詞')) {
    return false;
  }

  if (token.pos1 && N_PLUS_ONE_IGNORED_POS1.has(token.pos1)) {
    return false;
  }

  if (token.surface.trim().length === 0) {
    return false;
  }

  return true;
}

function isSentenceBoundaryToken(token: MergedToken): boolean {
  if (token.partOfSpeech !== PartOfSpeech.symbol) {
    return false;
  }

  return SENTENCE_BOUNDARY_SURFACES.has(token.surface);
}

export function markNPlusOneTargets(tokens: MergedToken[], minSentenceWords = 3): MergedToken[] {
  if (tokens.length === 0) {
    return [];
  }

  const markedTokens = tokens.map((token) => ({
    ...token,
    isNPlusOneTarget: false,
  }));

  let sentenceStart = 0;
  const minimumSentenceWords = Number.isInteger(minSentenceWords)
    ? Math.max(1, minSentenceWords)
    : 3;

  const markSentence = (start: number, endExclusive: number): void => {
    const sentenceCandidates: number[] = [];
    let sentenceWordCount = 0;
    for (let i = start; i < endExclusive; i++) {
      const token = markedTokens[i];
      if (!token) continue;
      if (isNPlusOneWordCountToken(token)) {
        sentenceWordCount += 1;
      }

      if (isNPlusOneCandidateToken(token)) {
        sentenceCandidates.push(i);
      }
    }

    if (sentenceWordCount >= minimumSentenceWords && sentenceCandidates.length === 1) {
      markedTokens[sentenceCandidates[0]!] = {
        ...markedTokens[sentenceCandidates[0]!]!,
        isNPlusOneTarget: true,
      };
    }
  };

  for (let i = 0; i < markedTokens.length; i++) {
    const token = markedTokens[i];
    if (!token) continue;
    if (isSentenceBoundaryToken(token)) {
      markSentence(sentenceStart, i);
      sentenceStart = i + 1;
    }
  }

  if (sentenceStart < markedTokens.length) {
    markSentence(sentenceStart, markedTokens.length);
  }

  return markedTokens;
}
