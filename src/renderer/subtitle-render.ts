import type { MergedToken, SecondarySubMode, SubtitleData, SubtitleStyleConfig } from '../types';
import type { RendererContext } from './context';

type FrequencyRenderSettings = {
  enabled: boolean;
  topX: number;
  mode: 'single' | 'banded';
  singleColor: string;
  bandedColors: [string, string, string, string, string];
};

export type SubtitleTokenHoverRange = {
  start: number;
  end: number;
  tokenIndex: number;
};

export function shouldRenderTokenizedSubtitle(tokenCount: number): boolean {
  return tokenCount > 0;
}

function isWhitespaceOnly(value: string): boolean {
  return value.trim().length === 0;
}

export function normalizeSubtitle(text: string, trim = true, collapseLineBreaks = false): string {
  if (!text) return '';

  let normalized = text.replace(/\\N/g, '\n').replace(/\\n/g, '\n');
  normalized = normalized.replace(/\{[^}]*\}/g, '');
  if (collapseLineBreaks) {
    normalized = normalized.replace(/\n/g, ' ');
    normalized = normalized.replace(/\s+/g, ' ');
  }

  return trim ? normalized.trim() : normalized;
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
    return 'rgba(54, 58, 79, 0.84)';
  }
  const trimmed = value.trim();
  return trimmed.length > 0 && SAFE_CSS_COLOR_PATTERN.test(trimmed)
    ? trimmed
    : 'rgba(54, 58, 79, 0.84)';
}

const DEFAULT_FREQUENCY_RENDER_SETTINGS: FrequencyRenderSettings = {
  enabled: false,
  topX: 1000,
  mode: 'single',
  singleColor: '#f5a97f',
  bandedColors: ['#ed8796', '#f5a97f', '#f9e2af', '#a6e3a1', '#8aadf4'],
};

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
  'backgroundColor',
  'backdropFilter',
  'WebkitBackdropFilter',
  'webkitBackdropFilter',
  '-webkit-backdrop-filter',
]);

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
  frequencySettings?: Partial<FrequencyRenderSettings>,
): string | null {
  if (token.isNPlusOneTarget) {
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

export function getJlptLevelLabelForToken(token: MergedToken): string | null {
  return token.jlptLevel ?? null;
}

function renderWithTokens(
  root: HTMLElement,
  tokens: MergedToken[],
  frequencyRenderSettings?: Partial<FrequencyRenderSettings>,
  sourceText?: string,
  preserveLineBreaks = false,
): void {
  const resolvedFrequencyRenderSettings = {
    ...DEFAULT_FREQUENCY_RENDER_SETTINGS,
    ...frequencyRenderSettings,
    bandedColors: sanitizeFrequencyBandedColors(
      frequencyRenderSettings?.bandedColors,
      DEFAULT_FREQUENCY_RENDER_SETTINGS.bandedColors,
    ),
    topX: sanitizeFrequencyTopX(
      frequencyRenderSettings?.topX,
      DEFAULT_FREQUENCY_RENDER_SETTINGS.topX,
    ),
    singleColor: sanitizeHexColor(
      frequencyRenderSettings?.singleColor,
      DEFAULT_FREQUENCY_RENDER_SETTINGS.singleColor,
    ),
  };

  const fragment = document.createDocumentFragment();

  if (preserveLineBreaks && sourceText) {
    const normalizedSource = normalizeSubtitle(sourceText, true, false);
    const segments = alignTokensToSourceText(tokens, normalizedSource);

    for (const segment of segments) {
      if (segment.kind === 'text') {
        renderPlainTextPreserveLineBreaks(fragment, segment.text);
        continue;
      }

      const token = segment.token;
      const span = document.createElement('span');
      span.className = computeWordClass(token, resolvedFrequencyRenderSettings);
      span.textContent = token.surface;
      span.dataset.tokenIndex = String(segment.tokenIndex);
      if (token.reading) span.dataset.reading = token.reading;
      if (token.headword) span.dataset.headword = token.headword;
      const frequencyRankLabel = getFrequencyRankLabelForToken(
        token,
        resolvedFrequencyRenderSettings,
      );
      if (frequencyRankLabel) {
        span.dataset.frequencyRank = frequencyRankLabel;
      }
      const jlptLevelLabel = getJlptLevelLabelForToken(token);
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

    const span = document.createElement('span');
    span.className = computeWordClass(token, resolvedFrequencyRenderSettings);
    span.textContent = surface;
    span.dataset.tokenIndex = String(index);
    if (token.reading) span.dataset.reading = token.reading;
    if (token.headword) span.dataset.headword = token.headword;
    const frequencyRankLabel = getFrequencyRankLabelForToken(
      token,
      resolvedFrequencyRenderSettings,
    );
    if (frequencyRankLabel) {
      span.dataset.frequencyRank = frequencyRankLabel;
    }
    const jlptLevelLabel = getJlptLevelLabelForToken(token);
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
  frequencySettings?: Partial<FrequencyRenderSettings>,
): string {
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

  const classes = ['word'];

  if (token.isNPlusOneTarget) {
    classes.push('word-n-plus-one');
  } else if (token.isKnown) {
    classes.push('word-known');
  }

  if (token.jlptLevel) {
    classes.push(`word-jlpt-${token.jlptLevel.toLowerCase()}`);
  }

  if (!token.isNPlusOneTarget) {
    const frequencyClass = getFrequencyDictionaryClass(token, resolvedFrequencySettings);
    if (frequencyClass) {
      classes.push(frequencyClass);
    }
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
    const span = document.createElement('span');
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

export function createSubtitleRenderer(ctx: RendererContext) {
  function renderSubtitle(data: SubtitleData | string): void {
    ctx.dom.subtitleRoot.innerHTML = '';

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

    if (!text) return;

    const normalized = normalizeSubtitle(text, true, !ctx.state.preserveSubtitleLineBreaks);
    if (shouldRenderTokenizedSubtitle(tokens?.length ?? 0) && tokens) {
      renderWithTokens(
        ctx.dom.subtitleRoot,
        tokens,
        getFrequencyRenderSettings(),
        text,
        ctx.state.preserveSubtitleLineBreaks,
      );
      return;
    }
    renderCharacterLevel(ctx.dom.subtitleRoot, normalized);
  }

  function getFrequencyRenderSettings(): Partial<FrequencyRenderSettings> {
    return {
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
    ctx.dom.secondarySubRoot.innerHTML = '';
    if (!text) return;

    const normalized = text
      .replace(/\\N/g, '\n')
      .replace(/\\n/g, '\n')
      .replace(/\{[^}]*\}/g, '')
      .trim();

    if (!normalized) return;

    const lines = normalized.split('\n');
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

  function applySubtitleFontSize(fontSize: number): void {
    const clampedSize = Math.max(10, fontSize);
    ctx.dom.subtitleRoot.style.fontSize = `${clampedSize}px`;
    document.documentElement.style.setProperty('--subtitle-font-size', `${clampedSize}px`);
  }

  function applySubtitleStyle(style: SubtitleStyleConfig | null): void {
    if (!style) return;

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

    ctx.state.knownWordColor = knownWordColor;
    ctx.state.nPlusOneColor = nPlusOneColor;
    ctx.dom.subtitleRoot.style.setProperty('--subtitle-known-word-color', knownWordColor);
    ctx.dom.subtitleRoot.style.setProperty('--subtitle-n-plus-one-color', nPlusOneColor);
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

    const secondaryStyle = style.secondary;
    if (!secondaryStyle) return;

    const secondaryStyleDeclarations = secondaryStyle as Record<string, unknown>;
    applyInlineStyleDeclarations(
      ctx.dom.secondarySubRoot,
      secondaryStyleDeclarations,
      CONTAINER_STYLE_KEYS,
    );
    applyInlineStyleDeclarations(
      ctx.dom.secondarySubContainer,
      pickInlineStyleDeclarations(secondaryStyleDeclarations, CONTAINER_STYLE_KEYS),
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
  }

  return {
    applySubtitleFontSize,
    applySubtitleStyle,
    renderSecondarySub,
    renderSubtitle,
    updateSecondarySubMode,
  };
}
