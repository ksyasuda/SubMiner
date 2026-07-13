import { MergedToken } from '../../../types';
import { isKanaChar, isKanaOnlyText, normalizeKana, splitPosTag } from './token-classification';

export function isTrailingSmallTsuKanaSfx(text: string): boolean {
  const chars = [...normalizeKana(text)];
  return (
    chars.length >= 2 &&
    chars.length <= 4 &&
    chars.every(isKanaChar) &&
    chars[chars.length - 1] === 'っ'
  );
}

function isReduplicatedKanaSfx(text: string): boolean {
  const chars = [...normalizeKana(text)];
  if (chars.length < 4 || chars.length % 2 !== 0 || !chars.every(isKanaChar)) {
    return false;
  }
  const half = chars.length / 2;
  return chars.slice(0, half).join('') === chars.slice(half).join('');
}

export function isReduplicatedKanaSfxWithOptionalTrailingTo(text: string): boolean {
  const normalized = normalizeKana(text);
  if (!normalized) {
    return false;
  }
  if (isReduplicatedKanaSfx(normalized)) {
    return true;
  }
  return normalized.length > 1 && normalized.endsWith('と')
    ? isReduplicatedKanaSfx(normalized.slice(0, -1))
    : false;
}

export function isExcludedTrailingParticleMergedToken(
  token: MergedToken,
  suffixes: readonly string[],
  leadingPos1Exclusions: ReadonlySet<string>,
): boolean {
  const surface = normalizeKana(token.surface);
  const headword = normalizeKana(token.headword);
  if (!surface || !headword || !surface.startsWith(headword)) {
    return false;
  }
  if (!suffixes.includes(surface.slice(headword.length))) {
    return false;
  }

  const [leadingPos1, ...trailingPos1] = splitPosTag(token.pos1);
  if (!leadingPos1 || leadingPos1Exclusions.has(leadingPos1)) {
    return false;
  }
  return trailingPos1.length > 0 && trailingPos1.every((part) => part === '助詞');
}

export function isAuxiliaryStemGrammarTailToken(
  token: MergedToken,
  allowedPos1: readonly string[],
): boolean {
  const pos1Parts = splitPosTag(token.pos1);
  if (pos1Parts.length === 0 || !pos1Parts.every((part) => allowedPos1.includes(part))) {
    return false;
  }
  return splitPosTag(token.pos3).includes('助動詞語幹');
}

export function isKanaOnlyNonIndependentNounHelperMerge(
  token: MergedToken,
  tailPos1: readonly string[],
): boolean {
  const surface = normalizeKana(token.surface);
  const headword = normalizeKana(token.headword);
  if (!surface || !headword || surface === headword || ![...surface].every(isKanaChar)) {
    return false;
  }

  const pos1Parts = splitPosTag(token.pos1);
  if (pos1Parts.length < 2 || pos1Parts[0] !== '名詞') {
    return false;
  }
  const pos2Parts = splitPosTag(token.pos2);
  return pos2Parts[0] === '非自立' && pos1Parts.slice(1).every((part) => tailPos1.includes(part));
}

export function isStandaloneAuxiliaryInflectionFragment(
  token: MergedToken,
  trailingPos1: readonly string[],
): boolean {
  if (!isKanaOnlyText(token.surface)) {
    return false;
  }
  const pos1Parts = splitPosTag(token.pos1);
  if (pos1Parts.length === 0) {
    return false;
  }
  if (pos1Parts.every((part) => part === '助動詞')) {
    return true;
  }
  const pos2Parts = splitPosTag(token.pos2);
  return (
    pos1Parts[0] === '動詞' &&
    pos2Parts[0] === '接尾' &&
    pos1Parts.slice(1).every((part) => trailingPos1.includes(part))
  );
}

export function isAuxiliaryOnlyHelperSpan(
  token: MergedToken,
  allowedPos1: readonly string[],
  lexicalVerbPos2: readonly string[],
): boolean {
  if (!isKanaOnlyText(token.surface) || !isKanaOnlyText(token.headword)) {
    return false;
  }
  const pos1Parts = splitPosTag(token.pos1);
  if (
    pos1Parts.length === 0 ||
    !pos1Parts.every((part) => allowedPos1.includes(part)) ||
    !pos1Parts.includes('助詞') ||
    !pos1Parts.includes('動詞')
  ) {
    return false;
  }
  return !splitPosTag(token.pos2).some((part) => lexicalVerbPos2.includes(part));
}

export function isStandaloneSuruTeGrammarHelper(token: MergedToken): boolean {
  const surface = normalizeKana(token.surface);
  const headword = normalizeKana(token.headword);
  if (!surface.startsWith('して') || headword !== 'する') {
    return false;
  }
  const pos1Parts = splitPosTag(token.pos1);
  return isKanaOnlyText(surface) && (pos1Parts.length === 0 || pos1Parts.includes('動詞'));
}

export function isStandaloneGrammarParticle(
  token: MergedToken,
  surfaces: readonly string[],
  phrases: readonly string[],
): boolean {
  const surface = normalizeKana(token.surface);
  return (
    surface === normalizeKana(token.headword) &&
    (surfaces.includes(surface) || phrases.includes(surface))
  );
}

export function isSingleKanaSurfaceFragment(token: MergedToken): boolean {
  const chars = [...normalizeKana(token.surface)];
  return chars.length === 1 && chars.every(isKanaChar);
}
