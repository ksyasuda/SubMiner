import type { SessionHelpSection } from './session-help-sections';
import { i18n } from '../../i18n/index.js';

export type SessionHelpSubtitleStyle = {
  knownWordColor?: unknown;
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

const HEX_COLOR_RE = /^#(?:[0-9a-fA-F]{3}|[0-9a-fA-F]{4}|[0-9a-fA-F]{6}|[0-9a-fA-F]{8})$/;

const FALLBACK_COLORS = {
  knownWordColor: '#a6da95',
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

export function buildColorSection(style: SessionHelpSubtitleStyle): SessionHelpSection {
  return {
    title: i18n.t('sessionHelp.colorLegend'),
    rows: [
      {
        shortcut: i18n.t('sessionHelp.colors.knownWords'),
        action: normalizeColor(style.knownWordColor, FALLBACK_COLORS.knownWordColor),
        color: normalizeColor(style.knownWordColor, FALLBACK_COLORS.knownWordColor),
      },
      {
        shortcut: i18n.t('sessionHelp.colors.nPlusOne'),
        action: normalizeColor(style.nPlusOneColor, FALLBACK_COLORS.nPlusOneColor),
        color: normalizeColor(style.nPlusOneColor, FALLBACK_COLORS.nPlusOneColor),
      },
      {
        shortcut: i18n.t('sessionHelp.colors.characterNames'),
        action: normalizeColor(style.nameMatchColor, FALLBACK_COLORS.nameMatchColor),
        color: normalizeColor(style.nameMatchColor, FALLBACK_COLORS.nameMatchColor),
      },
      {
        shortcut: i18n.t('sessionHelp.colors.jlptN1'),
        action: normalizeColor(style.jlptColors?.N1, FALLBACK_COLORS.jlptN1Color),
        color: normalizeColor(style.jlptColors?.N1, FALLBACK_COLORS.jlptN1Color),
      },
      {
        shortcut: i18n.t('sessionHelp.colors.jlptN2'),
        action: normalizeColor(style.jlptColors?.N2, FALLBACK_COLORS.jlptN2Color),
        color: normalizeColor(style.jlptColors?.N2, FALLBACK_COLORS.jlptN2Color),
      },
      {
        shortcut: i18n.t('sessionHelp.colors.jlptN3'),
        action: normalizeColor(style.jlptColors?.N3, FALLBACK_COLORS.jlptN3Color),
        color: normalizeColor(style.jlptColors?.N3, FALLBACK_COLORS.jlptN3Color),
      },
      {
        shortcut: i18n.t('sessionHelp.colors.jlptN4'),
        action: normalizeColor(style.jlptColors?.N4, FALLBACK_COLORS.jlptN4Color),
        color: normalizeColor(style.jlptColors?.N4, FALLBACK_COLORS.jlptN4Color),
      },
      {
        shortcut: i18n.t('sessionHelp.colors.jlptN5'),
        action: normalizeColor(style.jlptColors?.N5, FALLBACK_COLORS.jlptN5Color),
        color: normalizeColor(style.jlptColors?.N5, FALLBACK_COLORS.jlptN5Color),
      },
    ],
  };
}
