import type { MergedToken, SecondarySubMode, SubtitleData, SubtitleStyleConfig } from '../types';
import type { RendererContext } from './context';

type FrequencyRenderSettings = {
  enabled: boolean;
  topX: number;
  mode: 'single' | 'banded';
  singleColor: string;
  bandedColors: [string, string, string, string, string];
};

function normalizeSubtitle(text: string, trim = true, collapseLineBreaks = false): string {
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

function sanitizeHexColor(value: unknown, fallback: string): string {
  return typeof value === 'string' && HEX_COLOR_PATTERN.test(value.trim())
    ? value.trim()
    : fallback;
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
      if (token.reading) span.dataset.reading = token.reading;
      if (token.headword) span.dataset.headword = token.headword;
      fragment.appendChild(span);
    }

    root.appendChild(fragment);
    return;
  }

  for (const token of tokens) {
    const surface = token.surface;

    if (surface.includes('\n')) {
      const parts = surface.split('\n');
      for (let i = 0; i < parts.length; i += 1) {
        const part = parts[i];
        if (part) {
          const span = document.createElement('span');
          span.className = computeWordClass(token, resolvedFrequencyRenderSettings);
          span.textContent = part;
          if (token.reading) span.dataset.reading = token.reading;
          if (token.headword) span.dataset.headword = token.headword;
          fragment.appendChild(span);
        }
        if (i < parts.length - 1) {
          fragment.appendChild(document.createElement('br'));
        }
      }
      continue;
    }

    const span = document.createElement('span');
    span.className = computeWordClass(token, resolvedFrequencyRenderSettings);
    span.textContent = surface;
    if (token.reading) span.dataset.reading = token.reading;
    if (token.headword) span.dataset.headword = token.headword;
    fragment.appendChild(span);
  }

  root.appendChild(fragment);
}

type SubtitleRenderSegment = { kind: 'text'; text: string } | { kind: 'token'; token: MergedToken };

export function alignTokensToSourceText(
  tokens: MergedToken[],
  sourceText: string,
): SubtitleRenderSegment[] {
  if (tokens.length === 0) {
    return sourceText ? [{ kind: 'text', text: sourceText }] : [];
  }

  const segments: SubtitleRenderSegment[] = [];
  let cursor = 0;

  for (const token of tokens) {
    const surface = token.surface;
    if (!surface) {
      continue;
    }

    const foundIndex = sourceText.indexOf(surface, cursor);
    if (foundIndex < 0) {
      if (cursor < sourceText.length) {
        segments.push({ kind: 'text', text: sourceText.slice(cursor) });
      }
      segments.push({ kind: 'token', token });
      cursor = sourceText.length;
      continue;
    }

    if (foundIndex > cursor) {
      segments.push({ kind: 'text', text: sourceText.slice(cursor, foundIndex) });
    }

    segments.push({ kind: 'token', token });
    cursor = foundIndex + surface.length;
  }

  if (cursor < sourceText.length) {
    segments.push({ kind: 'text', text: sourceText.slice(cursor) });
  }

  return segments;
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

  if (!token.isKnown && !token.isNPlusOneTarget) {
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
    ctx.state.lastHoverSelectionKey = '';
    ctx.state.lastHoverSelectionNode = null;

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

    if (ctx.platform.isInvisibleLayer) {
      const normalizedInvisible = normalizeSubtitle(text, false);
      ctx.state.currentInvisibleSubtitleLineCount = Math.max(
        1,
        normalizedInvisible.split('\n').length,
      );
      renderPlainTextPreserveLineBreaks(ctx.dom.subtitleRoot, normalizedInvisible);
      return;
    }

    const normalized = normalizeSubtitle(text);
    if (tokens && tokens.length > 0) {
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

    if (style.fontFamily) ctx.dom.subtitleRoot.style.fontFamily = style.fontFamily;
    if (style.fontSize) ctx.dom.subtitleRoot.style.fontSize = `${style.fontSize}px`;
    if (style.fontColor) ctx.dom.subtitleRoot.style.color = style.fontColor;
    if (style.fontWeight) ctx.dom.subtitleRoot.style.fontWeight = style.fontWeight;
    if (style.fontStyle) ctx.dom.subtitleRoot.style.fontStyle = style.fontStyle;
    if (style.backgroundColor) {
      ctx.dom.subtitleContainer.style.background = style.backgroundColor;
    }

    const knownWordColor = style.knownWordColor ?? ctx.state.knownWordColor ?? '#a6da95';
    const nPlusOneColor = style.nPlusOneColor ?? ctx.state.nPlusOneColor ?? '#c6a0f6';
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
    ctx.state.jlptN1Color = jlptColors.N1;
    ctx.state.jlptN2Color = jlptColors.N2;
    ctx.state.jlptN3Color = jlptColors.N3;
    ctx.state.jlptN4Color = jlptColors.N4;
    ctx.state.jlptN5Color = jlptColors.N5;
    ctx.state.preserveSubtitleLineBreaks = style.preserveLineBreaks ?? false;
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
      ctx.dom.secondarySubRoot.style.fontWeight = secondaryStyle.fontWeight;
    }
    if (secondaryStyle.fontStyle) {
      ctx.dom.secondarySubRoot.style.fontStyle = secondaryStyle.fontStyle;
    }
    if (secondaryStyle.backgroundColor) {
      ctx.dom.secondarySubContainer.style.background = secondaryStyle.backgroundColor;
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
