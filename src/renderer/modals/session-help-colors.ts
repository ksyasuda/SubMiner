import type { SessionHelpSection } from './session-help-sections';

export type SessionHelpSubtitleStyle = {
  knownWordColor?: unknown;
  knownWordMaturityColors?: {
    new?: unknown;
    learning?: unknown;
    young?: unknown;
    mature?: unknown;
  };
  nPlusOneColor?: unknown;
  nameMatchColor?: unknown;
  jlptColors?: {
    N1?: unknown;
    N2?: unknown;
    N3?: unknown;
    N4?: unknown;
    N5?: unknown;
  };
};

export type SessionHelpColorOptions = {
  /** When true, known words are colored per Anki card maturity instead of one flat color. */
  knownWordMaturityEnabled?: boolean;
};

const HEX_COLOR_RE = /^#(?:[0-9a-fA-F]{3}|[0-9a-fA-F]{4}|[0-9a-fA-F]{6}|[0-9a-fA-F]{8})$/;

const FALLBACK_COLORS = {
  knownWordColor: '#a6da95',
  knownWordMaturityNewColor: '#ee99a0',
  knownWordMaturityLearningColor: '#b7bdf8',
  knownWordMaturityYoungColor: '#91d7e3',
  knownWordMaturityMatureColor: '#a6da95',
  nPlusOneColor: '#c6a0f6',
  nameMatchColor: '#f5bde6',
  jlptN1Color: '#ed8796',
  jlptN2Color: '#f5a97f',
  jlptN3Color: '#f9e2af',
  jlptN4Color: '#a6e3a1',
  jlptN5Color: '#8aadf4',
};

function normalizeColor(value: unknown, fallback: string): string {
  if (typeof value !== 'string') return fallback;
  const next = value.trim();
  return HEX_COLOR_RE.test(next) ? next : fallback;
}

function buildKnownWordRows(
  style: SessionHelpSubtitleStyle,
  options: SessionHelpColorOptions,
): SessionHelpSection['rows'] {
  if (!options.knownWordMaturityEnabled) {
    const knownWordColor = normalizeColor(style.knownWordColor, FALLBACK_COLORS.knownWordColor);
    return [{ shortcut: 'Known words', action: knownWordColor, color: knownWordColor }];
  }

  const maturityColors = style.knownWordMaturityColors;
  const tiers: Array<{ label: string; value: unknown; fallback: string }> = [
    {
      label: 'Known words (new)',
      value: maturityColors?.new,
      fallback: FALLBACK_COLORS.knownWordMaturityNewColor,
    },
    {
      label: 'Known words (learning)',
      value: maturityColors?.learning,
      fallback: FALLBACK_COLORS.knownWordMaturityLearningColor,
    },
    {
      label: 'Known words (young)',
      value: maturityColors?.young,
      fallback: FALLBACK_COLORS.knownWordMaturityYoungColor,
    },
    {
      label: 'Known words (mature)',
      value: maturityColors?.mature,
      fallback: FALLBACK_COLORS.knownWordMaturityMatureColor,
    },
  ];

  return tiers.map((tier) => {
    const color = normalizeColor(tier.value, tier.fallback);
    return { shortcut: tier.label, action: color, color };
  });
}

export function buildColorSection(
  style: SessionHelpSubtitleStyle,
  options: SessionHelpColorOptions = {},
): SessionHelpSection {
  return {
    title: 'Color legend',
    rows: [
      ...buildKnownWordRows(style, options),
      {
        shortcut: 'N+1 words',
        action: normalizeColor(style.nPlusOneColor, FALLBACK_COLORS.nPlusOneColor),
        color: normalizeColor(style.nPlusOneColor, FALLBACK_COLORS.nPlusOneColor),
      },
      {
        shortcut: 'Character names',
        action: normalizeColor(style.nameMatchColor, FALLBACK_COLORS.nameMatchColor),
        color: normalizeColor(style.nameMatchColor, FALLBACK_COLORS.nameMatchColor),
      },
      {
        shortcut: 'JLPT N1',
        action: normalizeColor(style.jlptColors?.N1, FALLBACK_COLORS.jlptN1Color),
        color: normalizeColor(style.jlptColors?.N1, FALLBACK_COLORS.jlptN1Color),
      },
      {
        shortcut: 'JLPT N2',
        action: normalizeColor(style.jlptColors?.N2, FALLBACK_COLORS.jlptN2Color),
        color: normalizeColor(style.jlptColors?.N2, FALLBACK_COLORS.jlptN2Color),
      },
      {
        shortcut: 'JLPT N3',
        action: normalizeColor(style.jlptColors?.N3, FALLBACK_COLORS.jlptN3Color),
        color: normalizeColor(style.jlptColors?.N3, FALLBACK_COLORS.jlptN3Color),
      },
      {
        shortcut: 'JLPT N4',
        action: normalizeColor(style.jlptColors?.N4, FALLBACK_COLORS.jlptN4Color),
        color: normalizeColor(style.jlptColors?.N4, FALLBACK_COLORS.jlptN4Color),
      },
      {
        shortcut: 'JLPT N5',
        action: normalizeColor(style.jlptColors?.N5, FALLBACK_COLORS.jlptN5Color),
        color: normalizeColor(style.jlptColors?.N5, FALLBACK_COLORS.jlptN5Color),
      },
    ],
  };
}
