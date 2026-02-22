import type { SubtitleData } from '../../types';

export const HOVER_SCRIPT_NAME = 'subminer';
export const HOVER_TOKEN_MESSAGE = 'subminer-hover-token';

const DEFAULT_HOVER_TOKEN_COLOR = 'C6A0F6';
const DEFAULT_TOKEN_COLOR = 'FFFFFF';

export type HoverPayloadToken = {
  text: string;
  index: number;
  startPos: number | null;
  endPos: number | null;
};

export type HoverTokenPayload = {
  revision: number;
  subtitle: string | null;
  hoveredTokenIndex: number | null;
  tokens: HoverPayloadToken[];
  colors: {
    base: string;
    hover: string;
  };
};

type HoverTokenInput = {
  subtitle: SubtitleData | null;
  hoveredTokenIndex: number | null;
  revision: number;
  hoverColor?: string | null;
};

function normalizeHexColor(color: string | null | undefined, fallback: string): string {
  if (typeof color !== 'string') {
    return fallback;
  }
  const normalized = color.trim().replace(/^#/, '').toUpperCase();
  return /^[0-9A-F]{6}$/.test(normalized) ? normalized : fallback;
}

function sanitizeSubtitleText(text: string): string {
  return text
    .replace(/\\N/g, '\n')
    .replace(/\\n/g, '\n')
    .replace(/\{[^}]*\}/g, '')
    .trim();
}

function sanitizeTokenSurface(surface: unknown): string {
  return typeof surface === 'string' ? surface : '';
}

function hasHoveredToken(subtitle: SubtitleData | null, hoveredTokenIndex: number | null): boolean {
  if (!subtitle || hoveredTokenIndex === null || hoveredTokenIndex < 0) {
    return false;
  }

  return subtitle.tokens?.some((token, index) => index === hoveredTokenIndex) ?? false;
}

export function buildHoveredTokenPayload(input: HoverTokenInput): HoverTokenPayload {
  const { subtitle, hoveredTokenIndex, revision, hoverColor } = input;

  const tokens: HoverPayloadToken[] = [];

  if (subtitle?.tokens && subtitle.tokens.length > 0) {
    for (let tokenIndex = 0; tokenIndex < subtitle.tokens.length; tokenIndex += 1) {
      const token = subtitle.tokens[tokenIndex];
      if (!token) {
        continue;
      }
      const surface = sanitizeTokenSurface(token?.surface);
      if (!surface || surface.trim().length === 0) {
        continue;
      }

      tokens.push({
        text: surface,
        index: tokenIndex,
        startPos: Number.isFinite(token.startPos) ? token.startPos : null,
        endPos: Number.isFinite(token.endPos) ? token.endPos : null,
      });
    }
  }

  return {
    revision,
    subtitle: subtitle ? sanitizeSubtitleText(subtitle.text) : null,
    hoveredTokenIndex:
      hoveredTokenIndex !== null && hoveredTokenIndex >= 0 ? hoveredTokenIndex : null,
    tokens,
    colors: {
      base: DEFAULT_TOKEN_COLOR,
      hover: normalizeHexColor(hoverColor, DEFAULT_HOVER_TOKEN_COLOR),
    },
  };
}

export function buildHoveredTokenMessageCommand(payload: HoverTokenPayload): (string | number)[] {
  return [
    'script-message-to',
    HOVER_SCRIPT_NAME,
    HOVER_TOKEN_MESSAGE,
    JSON.stringify(payload),
  ];
}

export function createApplyHoveredTokenOverlayHandler(deps: {
  getMpvClient: () => {
    connected: boolean;
    send: (payload: { command: (string | number)[] }) => boolean;
  } | null;
  getCurrentSubtitleData: () => SubtitleData | null;
  getHoveredTokenIndex: () => number | null;
  getHoveredSubtitleRevision: () => number;
  getHoverTokenColor: () => string | null;
}) {
  return (): void => {
    const mpvClient = deps.getMpvClient();
    if (!mpvClient || !mpvClient.connected) {
      return;
    }

    const subtitle = deps.getCurrentSubtitleData();
    const hoveredTokenIndex = deps.getHoveredTokenIndex();
    const revision = deps.getHoveredSubtitleRevision();
    const hoverColor = deps.getHoverTokenColor();
    const payload = buildHoveredTokenPayload({
      subtitle: subtitle && hasHoveredToken(subtitle, hoveredTokenIndex) ? subtitle : null,
      hoveredTokenIndex: hoveredTokenIndex,
      revision,
      hoverColor,
    });

    mpvClient.send({ command: buildHoveredTokenMessageCommand(payload) });
  };
}
