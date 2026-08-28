import type {
  MergedToken,
  PrimarySubMode,
  SecondarySubMode,
  SubtitleData,
  SubtitleRendererStyleConfig,
} from '../types';
import { assToPlainText, normalizePlainSubtitleText } from '../core/services/ass-text.js';
import { flattenedSecondarySubtitleLineIdentity } from '../core/services/secondary-subtitle-line-identity.js';
import type { RendererContext } from './context';
import { PRIMARY_SUB_VISIBLE_ON_YOMITAN_POPUP_CLASS } from './yomitan-popup.js';

type FrequencyRenderSettings = {
  enabled: boolean;
  topX: number;
  mode: 'single' | 'banded';
  singleColor: string;
  bandedColors: [string, string, string, string, string];
};

type TokenRenderSettings = FrequencyRenderSettings & {
  nameMatchEnabled: boolean;
};

export type SubtitleTokenHoverRange = {
  start: number;
  end: number;
  tokenIndex: number;
};

let _spanTemplate: HTMLSpanElement | null = null;
function getSpanTemplate(): HTMLSpanElement {
  if (!_spanTemplate) {
    _spanTemplate = document.createElement('span');
  }
  return _spanTemplate;
}

export function shouldRenderTokenizedSubtitle(tokenCount: number): boolean {
  return tokenCount > 0;
}

function isWhitespaceOnly(value: string): boolean {
  return value.trim().length === 0;
}

// Text reaching the overlay has already been decoded from ASS -- by mpv for live lines,
// by the cue parser for prefetched ones -- so this only settles line breaks.
export function normalizeSubtitle(text: string, trim = true, collapseLineBreaks = false): string {
  return normalizePlainSubtitleText(text, { trim, collapseLineBreaks });
}

/**
 * Display form of a resolved subtitle. `preserveLineBreaks` governs wrapping inside one
 * utterance, which is what a typesetter's `\N` means. The blank line the resolver puts
 * between two simultaneous cues is a different thing and always breaks, so a sign or a
 * second speaker never runs into the line beside it.
 */
export function normalizeSubtitleForDisplay(text: string, preserveLineBreaks: boolean): string {
  return text
    .split(/\n{2,}/)
    .map((cueText) => normalizeSubtitle(cueText, true, !preserveLineBreaks))
    .filter((cueText) => cueText.length > 0)
    .join('\n');
}

const HEX_COLOR_PATTERN = /^#(?:[0-9a-fA-F]{3}|[0-9a-fA-F]{4}|[0-9a-fA-F]{6}|[0-9a-fA-F]{8})$/;
const SAFE_CSS_COLOR_PATTERN =
  /^(?:#(?:[0-9a-fA-F]{3}|[0-9a-fA-F]{4}|[0-9a-fA-F]{6}|[0-9a-fA-F]{8})|(?:rgba?|hsla?)\([^)]*\)|var\([^)]*\)|[a-zA-Z]+)$/;

function sanitizeHexColor(value: unknown, fallback: string): string {
  return typeof value === 'string' && HEX_COLOR_PATTERN.test(value.trim())
    ? value.trim()
    : fallback;
}

export function sanitizeSubtitleHoverTokenColor(value: unknown): string {
  const sanitized = sanitizeHexColor(value, '#f4dbd6');
  const normalized = sanitized.replace(/^#/, '').toLowerCase();
  if (
    normalized === '000' ||
    normalized === '0000' ||
    normalized === '000000' ||
    normalized === '00000000'
  ) {
    return '#f4dbd6';
  }
  return sanitized;
}

function sanitizeSubtitleHoverTokenBackgroundColor(value: unknown): string {
  if (typeof value !== 'string') {
    return 'transparent';
  }
  const trimmed = value.trim();
  return trimmed.length > 0 && SAFE_CSS_COLOR_PATTERN.test(trimmed) ? trimmed : 'transparent';
}

const DEFAULT_FREQUENCY_RENDER_SETTINGS: FrequencyRenderSettings = {
  enabled: false,
  topX: 10000,
  mode: 'single',
  singleColor: '#f5a97f',
  bandedColors: ['#ed8796', '#f5a97f', '#f9e2af', '#8bd5ca', '#8aadf4'],
};
const DEFAULT_NAME_MATCH_ENABLED = false;

function hasPrioritizedNameMatch(
  token: MergedToken,
  tokenRenderSettings?: Partial<Pick<TokenRenderSettings, 'nameMatchEnabled'>>,
): boolean {
  return (
    (tokenRenderSettings?.nameMatchEnabled ?? DEFAULT_NAME_MATCH_ENABLED) &&
    token.isNameMatch === true
  );
}

function hasTokenCharacterImage(token: MergedToken): boolean {
  return (
    typeof token.characterImage?.src === 'string' && token.characterImage.src.trim().length > 0
  );
}

function shouldRenderTokenCharacterImage(
  token: MergedToken,
  tokenRenderSettings: Partial<Pick<TokenRenderSettings, 'nameMatchEnabled'>>,
): boolean {
  return hasPrioritizedNameMatch(token, tokenRenderSettings) && hasTokenCharacterImage(token);
}

function appendTokenSurface(
  span: HTMLSpanElement,
  token: MergedToken,
  surface: string,
  tokenRenderSettings: Partial<Pick<TokenRenderSettings, 'nameMatchEnabled'>>,
): void {
  if (!shouldRenderTokenCharacterImage(token, tokenRenderSettings)) {
    span.textContent = surface;
    return;
  }

  const image = document.createElement('img');
  image.className = 'word-character-image';
  image.src = token.characterImage!.src;
  image.alt = token.characterImage!.alt || token.headword || surface;
  image.decoding = 'async';
  image.loading = 'eager';
  span.appendChild(image);
  span.appendChild(document.createTextNode(surface));
}

function sanitizeFrequencyTopX(value: unknown, fallback: number): number {
  if (typeof value !== 'number' || !Number.isFinite(value) || value <= 0) {
    return fallback;
  }
  return Math.max(1, Math.floor(value));
}

function sanitizeFrequencyBandedColors(
  value: unknown,
  fallback: FrequencyRenderSettings['bandedColors'],
): FrequencyRenderSettings['bandedColors'] {
  if (!Array.isArray(value) || value.length !== 5) {
    return fallback;
  }

  return [
    sanitizeHexColor(value[0], fallback[0]),
    sanitizeHexColor(value[1], fallback[1]),
    sanitizeHexColor(value[2], fallback[2]),
    sanitizeHexColor(value[3], fallback[3]),
    sanitizeHexColor(value[4], fallback[4]),
  ];
}

function applyInlineStyleDeclarations(
  target: HTMLElement,
  declarations: Record<string, unknown>,
  excludedKeys: ReadonlySet<string> = new Set<string>(),
): void {
  for (const [key, value] of Object.entries(declarations)) {
    if (excludedKeys.has(key)) {
      continue;
    }
    if (value === null || value === undefined || typeof value === 'object') {
      continue;
    }

    const cssValue = String(value);
    if (key.includes('-')) {
      target.style.setProperty(key, cssValue);
      if (key === '--webkit-text-stroke') {
        target.style.setProperty('-webkit-text-stroke', cssValue);
      }
      continue;
    }

    const styleTarget = target.style as unknown as Record<string, string>;
    styleTarget[key] = cssValue;
  }
}

const appliedCssKeys = new WeakMap<HTMLElement, Set<string>>();

function inlineStyleDeclarationKeys(
  declarations: Record<string, unknown>,
  excludedKeys: ReadonlySet<string>,
): Set<string> {
  const keys = new Set<string>();
  for (const [key, value] of Object.entries(declarations)) {
    if (excludedKeys.has(key)) continue;
    if (value === null || value === undefined || typeof value === 'object') continue;
    keys.add(key);
  }
  return keys;
}

function clearInlineStyleDeclaration(target: HTMLElement, key: string): void {
  if (key.includes('-')) {
    target.style.removeProperty(key);
    if (key === '--webkit-text-stroke') {
      target.style.removeProperty('-webkit-text-stroke');
    }
    return;
  }

  (target.style as unknown as Record<string, string>)[key] = '';
}

function replaceInlineStyleDeclarations(
  target: HTMLElement,
  declarations: Record<string, unknown>,
  excludedKeys: ReadonlySet<string> = new Set<string>(),
): void {
  const nextKeys = inlineStyleDeclarationKeys(declarations, excludedKeys);
  const previousKeys = appliedCssKeys.get(target) ?? new Set<string>();
  for (const key of previousKeys) {
    if (!nextKeys.has(key)) {
      clearInlineStyleDeclaration(target, key);
    }
  }
  applyInlineStyleDeclarations(target, declarations, excludedKeys);
  appliedCssKeys.set(target, nextKeys);
}

function normalizeCssDeclarationObject(value: unknown): Record<string, string> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    return {};
  }

  const declarations: Record<string, string> = {};
  for (const [key, rawValue] of Object.entries(value)) {
    if (typeof rawValue !== 'string') continue;
    const cssValue = rawValue.trim();
    if (cssValue.length > 0) declarations[key] = cssValue;
  }
  return declarations;
}

function applySubtitleCssDeclarations(
  root: HTMLElement,
  container: HTMLElement,
  declarations: Record<string, string>,
): void {
  replaceInlineStyleDeclarations(root, declarations, CONTAINER_STYLE_KEYS);
  replaceInlineStyleDeclarations(
    container,
    pickInlineStyleDeclarations(declarations, CONTAINER_STYLE_KEYS),
  );
}

function syncPrimaryVisibleOnYomitanPopupClass(ctx: RendererContext): void {
  document.body?.classList?.toggle(
    PRIMARY_SUB_VISIBLE_ON_YOMITAN_POPUP_CLASS,
    ctx.state.yomitanPopupVisible && ctx.state.primaryVisibleOnYomitanPopup,
  );
}

function pickInlineStyleDeclarations(
  declarations: Record<string, unknown>,
  includedKeys: ReadonlySet<string>,
): Record<string, unknown> {
  const picked: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(declarations)) {
    if (!includedKeys.has(key)) continue;
    picked[key] = value;
  }
  return picked;
}

const CONTAINER_STYLE_KEYS = new Set<string>([
  'background',
  'background-color',
  'backgroundColor',
  'backdrop-filter',
  'backdropFilter',
  'WebkitBackdropFilter',
  'webkitBackdropFilter',
  '-webkit-backdrop-filter',
]);

function resolveSecondaryBackgroundColor(declarations: Record<string, unknown>): string {
  for (const key of ['backgroundColor', 'background-color', 'background']) {
    const value = declarations[key];
    if (typeof value === 'string' && value.trim().length > 0) {
      return value.trim();
    }
  }

  return 'transparent';
}

function resolveSecondaryBackdropFilter(declarations: Record<string, unknown>): string {
  for (const key of [
    'backdropFilter',
    'backdrop-filter',
    'WebkitBackdropFilter',
    'webkitBackdropFilter',
    '-webkit-backdrop-filter',
  ]) {
    const value = declarations[key];
    if (typeof value === 'string' && value.trim().length > 0) {
      return value.trim();
    }
  }

  return 'none';
}

function getFrequencyDictionaryClass(
  token: MergedToken,
  settings: FrequencyRenderSettings,
): string {
  if (!settings.enabled) {
    return '';
  }

  if (typeof token.frequencyRank !== 'number' || !Number.isFinite(token.frequencyRank)) {
    return '';
  }

  const rank = Math.max(1, Math.floor(token.frequencyRank));
  const topX = sanitizeFrequencyTopX(settings.topX, DEFAULT_FREQUENCY_RENDER_SETTINGS.topX);
  if (rank > topX) {
    return '';
  }

  if (settings.mode === 'banded') {
    const bandCount = settings.bandedColors.length;
    const normalizedBand = Math.ceil((rank / topX) * bandCount);
    const band = Math.min(bandCount, Math.max(1, normalizedBand));
    return `word-frequency-band-${band}`;
  }

  return 'word-frequency-single';
}

function getNormalizedFrequencyRank(token: MergedToken): number | null {
  if (typeof token.frequencyRank !== 'number' || !Number.isFinite(token.frequencyRank)) {
    return null;
  }
  return Math.max(1, Math.floor(token.frequencyRank));
}

export function getFrequencyRankLabelForToken(
  token: MergedToken,
  frequencySettings?: Partial<TokenRenderSettings>,
): string | null {
  if (hasPrioritizedNameMatch(token, frequencySettings)) {
    return null;
  }

  const resolvedFrequencySettings = {
    ...DEFAULT_FREQUENCY_RENDER_SETTINGS,
    ...frequencySettings,
    bandedColors: sanitizeFrequencyBandedColors(
      frequencySettings?.bandedColors,
      DEFAULT_FREQUENCY_RENDER_SETTINGS.bandedColors,
    ),
    topX: sanitizeFrequencyTopX(frequencySettings?.topX, DEFAULT_FREQUENCY_RENDER_SETTINGS.topX),
    singleColor: sanitizeHexColor(
      frequencySettings?.singleColor,
      DEFAULT_FREQUENCY_RENDER_SETTINGS.singleColor,
    ),
  };

  if (!getFrequencyDictionaryClass(token, resolvedFrequencySettings)) {
    return null;
  }

  const rank = getNormalizedFrequencyRank(token);
  return rank === null ? null : String(rank);
}

export function getJlptLevelLabelForToken(
  token: MergedToken,
  tokenRenderSettings?: Partial<Pick<TokenRenderSettings, 'nameMatchEnabled'>>,
): string | null {
  if (hasPrioritizedNameMatch(token, tokenRenderSettings)) {
    return null;
  }

  return token.jlptLevel ?? null;
}

function renderWithTokens(
  root: HTMLElement,
  tokens: MergedToken[],
  tokenRenderSettings?: Partial<TokenRenderSettings>,
  sourceText?: string,
  preserveLineBreaks = false,
): void {
  const resolvedTokenRenderSettings = {
    ...DEFAULT_FREQUENCY_RENDER_SETTINGS,
    ...tokenRenderSettings,
    bandedColors: sanitizeFrequencyBandedColors(
      tokenRenderSettings?.bandedColors,
      DEFAULT_FREQUENCY_RENDER_SETTINGS.bandedColors,
    ),
    topX: sanitizeFrequencyTopX(tokenRenderSettings?.topX, DEFAULT_FREQUENCY_RENDER_SETTINGS.topX),
    singleColor: sanitizeHexColor(
      tokenRenderSettings?.singleColor,
      DEFAULT_FREQUENCY_RENDER_SETTINGS.singleColor,
    ),
    nameMatchEnabled: tokenRenderSettings?.nameMatchEnabled ?? DEFAULT_NAME_MATCH_ENABLED,
  };

  const fragment = document.createDocumentFragment();

  if (sourceText) {
    const normalizedSource = normalizeSubtitleForDisplay(sourceText, preserveLineBreaks);
    const segments = alignTokensToSourceText(tokens, normalizedSource);

    for (const segment of segments) {
      if (segment.kind === 'text') {
        // Normalization already resolved which breaks survive; every one left is real.
        renderPlainTextPreserveLineBreaks(fragment, segment.text);
        continue;
      }

      const token = segment.token;
      const span = getSpanTemplate().cloneNode(false) as HTMLSpanElement;
      span.className = computeWordClass(token, resolvedTokenRenderSettings);
      appendTokenSurface(span, token, token.surface, resolvedTokenRenderSettings);
      span.dataset.tokenIndex = String(segment.tokenIndex);
      if (token.reading) span.dataset.reading = token.reading;
      if (token.headword) span.dataset.headword = token.headword;
      const frequencyRankLabel = getFrequencyRankLabelForToken(token, resolvedTokenRenderSettings);
      if (frequencyRankLabel) {
        span.dataset.frequencyRank = frequencyRankLabel;
      }
      const jlptLevelLabel = getJlptLevelLabelForToken(token, resolvedTokenRenderSettings);
      if (jlptLevelLabel) {
        span.dataset.jlptLevel = jlptLevelLabel;
      }
      fragment.appendChild(span);
    }

    root.appendChild(fragment);
    return;
  }

  for (let index = 0; index < tokens.length; index += 1) {
    const token = tokens[index];
    if (!token) {
      continue;
    }
    const surface = token.surface.replace(/\n/g, ' ');
    if (!surface) {
      continue;
    }

    if (isWhitespaceOnly(surface)) {
      fragment.appendChild(document.createTextNode(surface));
      continue;
    }

    const span = getSpanTemplate().cloneNode(false) as HTMLSpanElement;
    span.className = computeWordClass(token, resolvedTokenRenderSettings);
    appendTokenSurface(span, token, surface, resolvedTokenRenderSettings);
    span.dataset.tokenIndex = String(index);
    if (token.reading) span.dataset.reading = token.reading;
    if (token.headword) span.dataset.headword = token.headword;
    const frequencyRankLabel = getFrequencyRankLabelForToken(token, resolvedTokenRenderSettings);
    if (frequencyRankLabel) {
      span.dataset.frequencyRank = frequencyRankLabel;
    }
    const jlptLevelLabel = getJlptLevelLabelForToken(token, resolvedTokenRenderSettings);
    if (jlptLevelLabel) {
      span.dataset.jlptLevel = jlptLevelLabel;
    }
    fragment.appendChild(span);
  }

  root.appendChild(fragment);
}

type SubtitleRenderSegment =
  | { kind: 'text'; text: string }
  | { kind: 'token'; token: MergedToken; tokenIndex: number };

export function alignTokensToSourceText(
  tokens: MergedToken[],
  sourceText: string,
): SubtitleRenderSegment[] {
  if (tokens.length === 0) {
    return sourceText ? [{ kind: 'text', text: sourceText }] : [];
  }

  const segments: SubtitleRenderSegment[] = [];
  let cursor = 0;

  for (let tokenIndex = 0; tokenIndex < tokens.length; tokenIndex += 1) {
    const token = tokens[tokenIndex];
    if (!token) {
      continue;
    }
    const surface = token.surface;
    if (!surface || isWhitespaceOnly(surface)) {
      continue;
    }

    const foundIndex = sourceText.indexOf(surface, cursor);
    if (foundIndex < 0) {
      // Token text can diverge from source normalization (e.g., half/full-width forms).
      // Skip unmatched token to avoid duplicating visible tail text in preserve-line-break mode.
      continue;
    }

    if (foundIndex > cursor) {
      segments.push({ kind: 'text', text: sourceText.slice(cursor, foundIndex) });
    }

    segments.push({ kind: 'token', token, tokenIndex });
    cursor = foundIndex + surface.length;
  }

  if (cursor < sourceText.length) {
    segments.push({ kind: 'text', text: sourceText.slice(cursor) });
  }

  return segments;
}

export function buildSubtitleTokenHoverRanges(
  tokens: MergedToken[],
  sourceText: string,
): SubtitleTokenHoverRange[] {
  if (tokens.length === 0 || sourceText.length === 0) {
    return [];
  }

  const segments = alignTokensToSourceText(tokens, sourceText);
  const ranges: SubtitleTokenHoverRange[] = [];
  let cursor = 0;

  for (const segment of segments) {
    if (segment.kind === 'text') {
      cursor += segment.text.length;
      continue;
    }

    const tokenLength = segment.token.surface.length;
    if (tokenLength <= 0) {
      continue;
    }

    ranges.push({
      start: cursor,
      end: cursor + tokenLength,
      tokenIndex: segment.tokenIndex,
    });
    cursor += tokenLength;
  }

  return ranges;
}

export function computeWordClass(
  token: MergedToken,
  tokenRenderSettings?: Partial<TokenRenderSettings>,
): string {
  const resolvedTokenRenderSettings = {
    ...DEFAULT_FREQUENCY_RENDER_SETTINGS,
    ...tokenRenderSettings,
    bandedColors: sanitizeFrequencyBandedColors(
      tokenRenderSettings?.bandedColors,
      DEFAULT_FREQUENCY_RENDER_SETTINGS.bandedColors,
    ),
    topX: sanitizeFrequencyTopX(tokenRenderSettings?.topX, DEFAULT_FREQUENCY_RENDER_SETTINGS.topX),
    singleColor: sanitizeHexColor(
      tokenRenderSettings?.singleColor,
      DEFAULT_FREQUENCY_RENDER_SETTINGS.singleColor,
    ),
    nameMatchEnabled: tokenRenderSettings?.nameMatchEnabled ?? DEFAULT_NAME_MATCH_ENABLED,
  };

  const classes = ['word'];

  if (hasPrioritizedNameMatch(token, resolvedTokenRenderSettings)) {
    classes.push('word-name-match');
  } else if (token.isNPlusOneTarget) {
    classes.push('word-n-plus-one');
  } else if (token.isKnown) {
    classes.push('word-known');
    // The maturity class rides on word-known so hover/selection rules keyed
    // on word-known keep applying; it only overrides the color.
    if (token.knownMaturity) {
      classes.push(`word-maturity-${token.knownMaturity}`);
    }
  }

  if (!hasPrioritizedNameMatch(token, resolvedTokenRenderSettings) && token.jlptLevel) {
    classes.push(`word-jlpt-${token.jlptLevel.toLowerCase()}`);
  }

  if (
    !token.isKnown &&
    !token.isNPlusOneTarget &&
    !hasPrioritizedNameMatch(token, resolvedTokenRenderSettings)
  ) {
    const frequencyClass = getFrequencyDictionaryClass(token, resolvedTokenRenderSettings);
    if (frequencyClass) {
      classes.push(frequencyClass);
    }
  }

  if (shouldRenderTokenCharacterImage(token, resolvedTokenRenderSettings)) {
    classes.push('word-character-image-token');
  }

  return classes.join(' ');
}

function renderCharacterLevel(root: HTMLElement, text: string): void {
  const fragment = document.createDocumentFragment();

  for (const char of text) {
    if (char === '\n') {
      fragment.appendChild(document.createElement('br'));
      continue;
    }
    const span = getSpanTemplate().cloneNode(false) as HTMLSpanElement;
    span.className = 'c';
    span.textContent = char;
    fragment.appendChild(span);
  }

  root.appendChild(fragment);
}

function renderPlainTextPreserveLineBreaks(root: ParentNode, text: string): void {
  const lines = text.split('\n');
  const fragment = document.createDocumentFragment();

  for (let i = 0; i < lines.length; i += 1) {
    fragment.appendChild(document.createTextNode(lines[i] ?? ''));
    if (i < lines.length - 1) {
      fragment.appendChild(document.createElement('br'));
    }
  }

  root.appendChild(fragment);
}

// Karaoke-typeset OP/ED tracks emit one ASS event per syllable (duplicated across
// layers), and mpv's secondary-sub-text joins every active event with newlines —
// rendering those verbatim stacks dozens of tiny lines down the screen and turns the
// hover-pause band into a full-screen trap.
const KARAOKE_MIN_LINE_COUNT = 8;
const KARAOKE_MAX_MEDIAN_LINE_LENGTH = 4;

function isKaraokeLikeLineSet(lines: string[]): boolean {
  if (lines.length < KARAOKE_MIN_LINE_COUNT) return false;
  const lengths = lines.map((line) => line.length).sort((a, b) => a - b);
  const median = lengths[Math.floor(lengths.length / 2)] ?? 0;
  return median <= KARAOKE_MAX_MEDIAN_LINE_LENGTH;
}

function collapseFullLineFallbackCopies(lines: string[]): string[] {
  const seenExact = new Set<string>();
  const seenFlattened = new Set<string>();
  return lines.filter((line) => {
    const exactIdentity = line.normalize('NFKC');
    if (seenExact.has(exactIdentity)) return false;
    seenExact.add(exactIdentity);

    const flattenedIdentity = flattenedSecondarySubtitleLineIdentity(line);
    if (!flattenedIdentity) return true;
    if (seenFlattened.has(flattenedIdentity)) return false;
    seenFlattened.add(flattenedIdentity);
    return true;
  });
}

export function prepareSecondarySubtitleLines(text: string): string[] {
  // The one display-side ASS decode: secondary text also reaches the overlay from
  // websocket clients that forward their source line untouched, so unlike the primary
  // path it cannot assume mpv already decoded it.
  const normalized = assToPlainText(text).trim();

  if (!normalized) return [];

  const lines = normalized
    .split('\n')
    .map((line) => line.trim())
    .filter((line) => line.length > 0);
  if (!isKaraokeLikeLineSet(lines)) {
    return collapseFullLineFallbackCopies(lines);
  }

  const seen = new Set<string>();
  const unique: string[] = [];
  for (const line of lines) {
    if (seen.has(line)) continue;
    seen.add(line);
    unique.push(line);
  }
  return [unique.join(' ')];
}

export function createSubtitleRenderer(ctx: RendererContext) {
  let lastPrimarySubtitleRenderKey: string | null = null;
  let lastPrimarySubtitleNormalizedText: string | null = null;
  let lastPrimarySubtitleRenderedTokenized = false;

  function getPrimarySubtitleRenderKey(
    text: string,
    normalized: string,
    tokens: MergedToken[] | null,
  ): string {
    if (!shouldRenderTokenizedSubtitle(tokens?.length ?? 0) || !tokens) {
      return JSON.stringify({
        mode: 'plain',
        text: normalized,
      });
    }

    return JSON.stringify({
      mode: 'tokens',
      text,
      tokens,
      settings: getTokenRenderSettings(),
      preserveSubtitleLineBreaks: ctx.state.preserveSubtitleLineBreaks,
    });
  }

  function renderSubtitle(data: SubtitleData | string): void {
    let text: string;
    let tokens: MergedToken[] | null;

    if (typeof data === 'string') {
      text = data;
      tokens = null;
    } else if (data && typeof data === 'object') {
      text = data.text;
      tokens = data.tokens;
    } else {
      return;
    }

    const normalized = normalizeSubtitleForDisplay(text, ctx.state.preserveSubtitleLineBreaks);
    const hasRenderableTokens =
      shouldRenderTokenizedSubtitle(tokens?.length ?? 0) && Boolean(tokens);
    if (
      lastPrimarySubtitleRenderKey !== null &&
      !hasRenderableTokens &&
      lastPrimarySubtitleRenderedTokenized &&
      normalized === lastPrimarySubtitleNormalizedText
    ) {
      return;
    }

    const renderKey = getPrimarySubtitleRenderKey(text, normalized, tokens);
    if (renderKey === lastPrimarySubtitleRenderKey) {
      return;
    }
    lastPrimarySubtitleRenderKey = renderKey;
    lastPrimarySubtitleNormalizedText = normalized;
    lastPrimarySubtitleRenderedTokenized = hasRenderableTokens;

    ctx.dom.subtitleRoot.replaceChildren();

    if (!text) return;

    if (shouldRenderTokenizedSubtitle(tokens?.length ?? 0) && tokens) {
      renderWithTokens(
        ctx.dom.subtitleRoot,
        tokens,
        getTokenRenderSettings(),
        text,
        ctx.state.preserveSubtitleLineBreaks,
      );
      return;
    }
    renderCharacterLevel(ctx.dom.subtitleRoot, normalized);
  }

  function getTokenRenderSettings(): Partial<TokenRenderSettings> {
    return {
      nameMatchEnabled: ctx.state.nameMatchEnabled,
      enabled: ctx.state.frequencyDictionaryEnabled,
      topX: ctx.state.frequencyDictionaryTopX,
      mode: ctx.state.frequencyDictionaryMode,
      singleColor: ctx.state.frequencyDictionarySingleColor,
      bandedColors: [
        ctx.state.frequencyDictionaryBand1Color,
        ctx.state.frequencyDictionaryBand2Color,
        ctx.state.frequencyDictionaryBand3Color,
        ctx.state.frequencyDictionaryBand4Color,
        ctx.state.frequencyDictionaryBand5Color,
      ] as [string, string, string, string, string],
    };
  }

  function renderSecondarySub(text: string): void {
    ctx.dom.secondarySubRoot.replaceChildren();
    if (!text) return;

    const lines = prepareSecondarySubtitleLines(text);
    for (let i = 0; i < lines.length; i += 1) {
      const line = lines[i];
      if (line) {
        ctx.dom.secondarySubRoot.appendChild(document.createTextNode(line));
      }
      if (i < lines.length - 1) {
        ctx.dom.secondarySubRoot.appendChild(document.createElement('br'));
      }
    }
  }

  function updateSecondarySubMode(mode: SecondarySubMode): void {
    ctx.dom.secondarySubContainer.classList.remove(
      'secondary-sub-hidden',
      'secondary-sub-visible',
      'secondary-sub-hover',
    );
    ctx.dom.secondarySubContainer.classList.add(`secondary-sub-${mode}`);
  }

  function updatePrimarySubMode(mode: PrimarySubMode): void {
    ctx.state.primarySubtitleMode = mode;
    ctx.dom.subtitleContainer.classList.remove(
      'primary-sub-hidden',
      'primary-sub-visible',
      'primary-sub-hover',
    );
    ctx.dom.subtitleContainer.classList.add(`primary-sub-${mode}`);
  }

  function applySubtitleFontSize(fontSize: number): void {
    const clampedSize = Math.max(10, fontSize);
    ctx.dom.subtitleRoot.style.fontSize = `${clampedSize}px`;
    document.documentElement.style.setProperty('--subtitle-font-size', `${clampedSize}px`);
  }

  function applySubtitleStyle(style: SubtitleRendererStyleConfig | null): void {
    if (!style) return;
    lastPrimarySubtitleRenderKey = null;

    const styleDeclarations = style as Record<string, unknown>;
    applyInlineStyleDeclarations(ctx.dom.subtitleRoot, styleDeclarations, CONTAINER_STYLE_KEYS);
    applyInlineStyleDeclarations(
      ctx.dom.subtitleContainer,
      pickInlineStyleDeclarations(styleDeclarations, CONTAINER_STYLE_KEYS),
    );

    if (style.fontFamily) ctx.dom.subtitleRoot.style.fontFamily = style.fontFamily;
    if (style.fontSize) ctx.dom.subtitleRoot.style.fontSize = `${style.fontSize}px`;
    if (style.fontColor) {
      ctx.dom.subtitleRoot.style.color = style.fontColor;
    }
    if (style.fontWeight) ctx.dom.subtitleRoot.style.fontWeight = String(style.fontWeight);
    if (style.fontStyle) ctx.dom.subtitleRoot.style.fontStyle = style.fontStyle;
    const knownWordColor = style.knownWordColor ?? ctx.state.knownWordColor ?? '#a6da95';
    const nPlusOneColor = style.nPlusOneColor ?? ctx.state.nPlusOneColor ?? '#c6a0f6';
    const nameMatchEnabled = style.nameMatchEnabled ?? ctx.state.nameMatchEnabled ?? false;
    const nameMatchColor = style.nameMatchColor ?? ctx.state.nameMatchColor ?? '#f5bde6';
    const hoverTokenColor = sanitizeSubtitleHoverTokenColor(style.hoverTokenColor);
    const hoverTokenBackgroundColor = sanitizeSubtitleHoverTokenBackgroundColor(
      style.hoverTokenBackgroundColor,
    );
    const jlptColors = {
      N1: ctx.state.jlptN1Color ?? '#ed8796',
      N2: ctx.state.jlptN2Color ?? '#f5a97f',
      N3: ctx.state.jlptN3Color ?? '#f9e2af',
      N4: ctx.state.jlptN4Color ?? '#a6e3a1',
      N5: ctx.state.jlptN5Color ?? '#8aadf4',
      ...(style.jlptColors
        ? {
            N1: sanitizeHexColor(style.jlptColors?.N1, ctx.state.jlptN1Color),
            N2: sanitizeHexColor(style.jlptColors?.N2, ctx.state.jlptN2Color),
            N3: sanitizeHexColor(style.jlptColors?.N3, ctx.state.jlptN3Color),
            N4: sanitizeHexColor(style.jlptColors?.N4, ctx.state.jlptN4Color),
            N5: sanitizeHexColor(style.jlptColors?.N5, ctx.state.jlptN5Color),
          }
        : {}),
    };

    const maturityColorOverrides = style.knownWordMaturityColors;
    const maturityColors = {
      new: sanitizeHexColor(maturityColorOverrides?.new, ctx.state.knownWordMaturityNewColor),
      learning: sanitizeHexColor(
        maturityColorOverrides?.learning,
        ctx.state.knownWordMaturityLearningColor,
      ),
      young: sanitizeHexColor(maturityColorOverrides?.young, ctx.state.knownWordMaturityYoungColor),
      mature: sanitizeHexColor(
        maturityColorOverrides?.mature,
        ctx.state.knownWordMaturityMatureColor,
      ),
    };

    ctx.state.knownWordColor = knownWordColor;
    ctx.state.knownWordMaturityNewColor = maturityColors.new;
    ctx.state.knownWordMaturityLearningColor = maturityColors.learning;
    ctx.state.knownWordMaturityYoungColor = maturityColors.young;
    ctx.state.knownWordMaturityMatureColor = maturityColors.mature;
    ctx.state.nPlusOneColor = nPlusOneColor;
    ctx.state.nameMatchEnabled = nameMatchEnabled;
    ctx.state.nameMatchColor = nameMatchColor;
    ctx.dom.subtitleRoot.style.setProperty('--subtitle-known-word-color', knownWordColor);
    ctx.dom.subtitleRoot.style.setProperty('--subtitle-maturity-new-color', maturityColors.new);
    ctx.dom.subtitleRoot.style.setProperty(
      '--subtitle-maturity-learning-color',
      maturityColors.learning,
    );
    ctx.dom.subtitleRoot.style.setProperty('--subtitle-maturity-young-color', maturityColors.young);
    ctx.dom.subtitleRoot.style.setProperty(
      '--subtitle-maturity-mature-color',
      maturityColors.mature,
    );
    ctx.dom.subtitleRoot.style.setProperty('--subtitle-n-plus-one-color', nPlusOneColor);
    ctx.dom.subtitleRoot.style.setProperty('--subtitle-name-match-color', nameMatchColor);
    ctx.dom.subtitleRoot.style.setProperty('--subtitle-hover-token-color', hoverTokenColor);
    ctx.dom.subtitleRoot.style.setProperty(
      '--subtitle-hover-token-background-color',
      hoverTokenBackgroundColor,
    );
    ctx.state.jlptN1Color = jlptColors.N1;
    ctx.state.jlptN2Color = jlptColors.N2;
    ctx.state.jlptN3Color = jlptColors.N3;
    ctx.state.jlptN4Color = jlptColors.N4;
    ctx.state.jlptN5Color = jlptColors.N5;
    ctx.state.preserveSubtitleLineBreaks = style.preserveLineBreaks ?? false;
    ctx.state.autoPauseVideoOnSubtitleHover = style.autoPauseVideoOnHover ?? false;
    ctx.state.autoPauseVideoOnYomitanPopup = style.autoPauseVideoOnYomitanPopup ?? false;
    ctx.state.primaryVisibleOnYomitanPopup = style.primaryVisibleOnYomitanPopup ?? true;
    syncPrimaryVisibleOnYomitanPopupClass(ctx);
    ctx.dom.subtitleRoot.style.setProperty('--subtitle-jlpt-n1-color', jlptColors.N1);
    ctx.dom.subtitleRoot.style.setProperty('--subtitle-jlpt-n2-color', jlptColors.N2);
    ctx.dom.subtitleRoot.style.setProperty('--subtitle-jlpt-n3-color', jlptColors.N3);
    ctx.dom.subtitleRoot.style.setProperty('--subtitle-jlpt-n4-color', jlptColors.N4);
    ctx.dom.subtitleRoot.style.setProperty('--subtitle-jlpt-n5-color', jlptColors.N5);
    const frequencyDictionarySettings = style.frequencyDictionary ?? {};
    const frequencyEnabled =
      frequencyDictionarySettings.enabled ?? ctx.state.frequencyDictionaryEnabled;
    const frequencyTopX = sanitizeFrequencyTopX(
      frequencyDictionarySettings.topX,
      ctx.state.frequencyDictionaryTopX,
    );
    const frequencyMode = frequencyDictionarySettings.mode
      ? frequencyDictionarySettings.mode
      : ctx.state.frequencyDictionaryMode;
    const frequencySingleColor = sanitizeHexColor(
      frequencyDictionarySettings.singleColor,
      ctx.state.frequencyDictionarySingleColor,
    );
    const frequencyBandedColors = sanitizeFrequencyBandedColors(
      frequencyDictionarySettings.bandedColors,
      [
        ctx.state.frequencyDictionaryBand1Color,
        ctx.state.frequencyDictionaryBand2Color,
        ctx.state.frequencyDictionaryBand3Color,
        ctx.state.frequencyDictionaryBand4Color,
        ctx.state.frequencyDictionaryBand5Color,
      ] as [string, string, string, string, string],
    );

    ctx.state.frequencyDictionaryEnabled = frequencyEnabled;
    ctx.state.frequencyDictionaryTopX = frequencyTopX;
    ctx.state.frequencyDictionaryMode = frequencyMode;
    ctx.state.frequencyDictionarySingleColor = frequencySingleColor;
    [
      ctx.state.frequencyDictionaryBand1Color,
      ctx.state.frequencyDictionaryBand2Color,
      ctx.state.frequencyDictionaryBand3Color,
      ctx.state.frequencyDictionaryBand4Color,
      ctx.state.frequencyDictionaryBand5Color,
    ] = frequencyBandedColors;
    ctx.dom.subtitleRoot.style.setProperty(
      '--subtitle-frequency-single-color',
      frequencySingleColor,
    );
    ctx.dom.subtitleRoot.style.setProperty(
      '--subtitle-frequency-band-1-color',
      frequencyBandedColors[0],
    );
    ctx.dom.subtitleRoot.style.setProperty(
      '--subtitle-frequency-band-2-color',
      frequencyBandedColors[1],
    );
    ctx.dom.subtitleRoot.style.setProperty(
      '--subtitle-frequency-band-3-color',
      frequencyBandedColors[2],
    );
    ctx.dom.subtitleRoot.style.setProperty(
      '--subtitle-frequency-band-4-color',
      frequencyBandedColors[3],
    );
    ctx.dom.subtitleRoot.style.setProperty(
      '--subtitle-frequency-band-5-color',
      frequencyBandedColors[4],
    );
    applySubtitleCssDeclarations(
      ctx.dom.subtitleRoot,
      ctx.dom.subtitleContainer,
      normalizeCssDeclarationObject(style.css),
    );

    const secondaryStyle = style.secondary;
    if (!secondaryStyle) return;

    const secondaryStyleDeclarations = secondaryStyle as Record<string, unknown>;
    const secondaryCssDeclarations = normalizeCssDeclarationObject(secondaryStyle.css);
    applyInlineStyleDeclarations(
      ctx.dom.secondarySubRoot,
      secondaryStyleDeclarations,
      CONTAINER_STYLE_KEYS,
    );
    const secondaryContainerStyleDeclarations = {
      ...pickInlineStyleDeclarations(secondaryStyleDeclarations, CONTAINER_STYLE_KEYS),
      ...pickInlineStyleDeclarations(secondaryCssDeclarations, CONTAINER_STYLE_KEYS),
    };
    ctx.dom.secondarySubContainer.style.setProperty(
      '--secondary-sub-background-color',
      resolveSecondaryBackgroundColor(secondaryContainerStyleDeclarations),
    );
    ctx.dom.secondarySubContainer.style.setProperty(
      '--secondary-sub-backdrop-filter',
      resolveSecondaryBackdropFilter(secondaryContainerStyleDeclarations),
    );

    if (secondaryStyle.fontFamily) {
      ctx.dom.secondarySubRoot.style.fontFamily = secondaryStyle.fontFamily;
    }
    if (secondaryStyle.fontSize) {
      ctx.dom.secondarySubRoot.style.fontSize = `${secondaryStyle.fontSize}px`;
    }
    if (secondaryStyle.fontColor) {
      ctx.dom.secondarySubRoot.style.color = secondaryStyle.fontColor;
    }
    if (secondaryStyle.fontWeight) {
      ctx.dom.secondarySubRoot.style.fontWeight = String(secondaryStyle.fontWeight);
    }
    if (secondaryStyle.fontStyle) {
      ctx.dom.secondarySubRoot.style.fontStyle = secondaryStyle.fontStyle;
    }
    applySubtitleCssDeclarations(
      ctx.dom.secondarySubRoot,
      ctx.dom.secondarySubContainer,
      secondaryCssDeclarations,
    );
  }

  return {
    applySubtitleFontSize,
    applySubtitleStyle,
    renderSecondarySub,
    renderSubtitle,
    updatePrimarySubMode,
    updateSecondarySubMode,
  };
}
