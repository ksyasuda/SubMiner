import type {
  MergedToken,
  SecondarySubMode,
  SubtitleData,
  SubtitleStyleConfig,
} from "../types";
import type { RendererContext } from "./context";

function normalizeSubtitle(text: string, trim = true): string {
  if (!text) return "";

  let normalized = text.replace(/\\N/g, "\n").replace(/\\n/g, "\n");
  normalized = normalized.replace(/\{[^}]*\}/g, "");

  return trim ? normalized.trim() : normalized;
}

const HEX_COLOR_PATTERN =
  /^#(?:[0-9a-fA-F]{3}|[0-9a-fA-F]{4}|[0-9a-fA-F]{6}|[0-9a-fA-F]{8})$/;

function sanitizeHexColor(value: unknown, fallback: string): string {
  return typeof value === "string" && HEX_COLOR_PATTERN.test(value.trim())
    ? value.trim()
    : fallback;
}

function renderWithTokens(root: HTMLElement, tokens: MergedToken[]): void {
  const fragment = document.createDocumentFragment();

  for (const token of tokens) {
    const surface = token.surface;

    if (surface.includes("\n")) {
      const parts = surface.split("\n");
        for (let i = 0; i < parts.length; i += 1) {
          if (parts[i]) {
            const span = document.createElement("span");
            span.className = computeWordClass(token);
            span.textContent = parts[i];
            if (token.reading) span.dataset.reading = token.reading;
            if (token.headword) span.dataset.headword = token.headword;
          fragment.appendChild(span);
        }
        if (i < parts.length - 1) {
          fragment.appendChild(document.createElement("br"));
        }
      }
      continue;
    }

    const span = document.createElement("span");
    span.className = computeWordClass(token);
    span.textContent = surface;
    if (token.reading) span.dataset.reading = token.reading;
    if (token.headword) span.dataset.headword = token.headword;
    fragment.appendChild(span);
  }

  root.appendChild(fragment);
}

export function computeWordClass(token: MergedToken): string {
  const classes = ["word"];

  if (token.isNPlusOneTarget) {
    classes.push("word-n-plus-one");
  } else if (token.isKnown) {
    classes.push("word-known");
  }

  if (token.jlptLevel) {
    classes.push(`word-jlpt-${token.jlptLevel.toLowerCase()}`);
  }

  return classes.join(" ");
}

function renderCharacterLevel(root: HTMLElement, text: string): void {
  const fragment = document.createDocumentFragment();

  for (const char of text) {
    if (char === "\n") {
      fragment.appendChild(document.createElement("br"));
      continue;
    }
    const span = document.createElement("span");
    span.className = "c";
    span.textContent = char;
    fragment.appendChild(span);
  }

  root.appendChild(fragment);
}

function renderPlainTextPreserveLineBreaks(root: HTMLElement, text: string): void {
  const lines = text.split("\n");
  const fragment = document.createDocumentFragment();

  for (let i = 0; i < lines.length; i += 1) {
    fragment.appendChild(document.createTextNode(lines[i]));
    if (i < lines.length - 1) {
      fragment.appendChild(document.createElement("br"));
    }
  }

  root.appendChild(fragment);
}

export function createSubtitleRenderer(ctx: RendererContext) {
  function renderSubtitle(data: SubtitleData | string): void {
    ctx.dom.subtitleRoot.innerHTML = "";
    ctx.state.lastHoverSelectionKey = "";
    ctx.state.lastHoverSelectionNode = null;

    let text: string;
    let tokens: MergedToken[] | null;

    if (typeof data === "string") {
      text = data;
      tokens = null;
    } else if (data && typeof data === "object") {
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
        normalizedInvisible.split("\n").length,
      );
      renderPlainTextPreserveLineBreaks(ctx.dom.subtitleRoot, normalizedInvisible);
      return;
    }

    const normalized = normalizeSubtitle(text);
    if (tokens && tokens.length > 0) {
      renderWithTokens(ctx.dom.subtitleRoot, tokens);
      return;
    }
    renderCharacterLevel(ctx.dom.subtitleRoot, normalized);
  }

  function renderSecondarySub(text: string): void {
    ctx.dom.secondarySubRoot.innerHTML = "";
    if (!text) return;

    const normalized = text
      .replace(/\\N/g, "\n")
      .replace(/\\n/g, "\n")
      .replace(/\{[^}]*\}/g, "")
      .trim();

    if (!normalized) return;

    const lines = normalized.split("\n");
    for (let i = 0; i < lines.length; i += 1) {
      if (lines[i]) {
        ctx.dom.secondarySubRoot.appendChild(document.createTextNode(lines[i]));
      }
      if (i < lines.length - 1) {
        ctx.dom.secondarySubRoot.appendChild(document.createElement("br"));
      }
    }
  }

  function updateSecondarySubMode(mode: SecondarySubMode): void {
    ctx.dom.secondarySubContainer.classList.remove(
      "secondary-sub-hidden",
      "secondary-sub-visible",
      "secondary-sub-hover",
    );
    ctx.dom.secondarySubContainer.classList.add(`secondary-sub-${mode}`);
  }

  function applySubtitleFontSize(fontSize: number): void {
    const clampedSize = Math.max(10, fontSize);
    ctx.dom.subtitleRoot.style.fontSize = `${clampedSize}px`;
    document.documentElement.style.setProperty(
      "--subtitle-font-size",
      `${clampedSize}px`,
    );
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

    const knownWordColor =
      style.knownWordColor ?? ctx.state.knownWordColor ?? "#a6da95";
    const nPlusOneColor =
      style.nPlusOneColor ?? ctx.state.nPlusOneColor ?? "#c6a0f6";
    const jlptColors = {
      N1: ctx.state.jlptN1Color ?? "#ed8796",
      N2: ctx.state.jlptN2Color ?? "#f5a97f",
      N3: ctx.state.jlptN3Color ?? "#f9e2af",
      N4: ctx.state.jlptN4Color ?? "#a6e3a1",
      N5: ctx.state.jlptN5Color ?? "#8aadf4",
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
    ctx.dom.subtitleRoot.style.setProperty(
      "--subtitle-known-word-color",
      knownWordColor,
    );
    ctx.dom.subtitleRoot.style.setProperty("--subtitle-n-plus-one-color", nPlusOneColor);
    ctx.state.jlptN1Color = jlptColors.N1;
    ctx.state.jlptN2Color = jlptColors.N2;
    ctx.state.jlptN3Color = jlptColors.N3;
    ctx.state.jlptN4Color = jlptColors.N4;
    ctx.state.jlptN5Color = jlptColors.N5;
    ctx.dom.subtitleRoot.style.setProperty("--subtitle-jlpt-n1-color", jlptColors.N1);
    ctx.dom.subtitleRoot.style.setProperty("--subtitle-jlpt-n2-color", jlptColors.N2);
    ctx.dom.subtitleRoot.style.setProperty("--subtitle-jlpt-n3-color", jlptColors.N3);
    ctx.dom.subtitleRoot.style.setProperty("--subtitle-jlpt-n4-color", jlptColors.N4);
    ctx.dom.subtitleRoot.style.setProperty("--subtitle-jlpt-n5-color", jlptColors.N5);

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
