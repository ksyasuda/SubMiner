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

function renderWithTokens(root: HTMLElement, tokens: MergedToken[]): void {
  const fragment = document.createDocumentFragment();

  for (const token of tokens) {
    const surface = token.surface;

    if (surface.includes("\n")) {
      const parts = surface.split("\n");
      for (let i = 0; i < parts.length; i += 1) {
        if (parts[i]) {
          const span = document.createElement("span");
          span.className = token.isKnown ? "word word-known" : "word";
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
    span.className = token.isKnown ? "word word-known" : "word";
    span.textContent = surface;
    if (token.reading) span.dataset.reading = token.reading;
    if (token.headword) span.dataset.headword = token.headword;
    fragment.appendChild(span);
  }

  root.appendChild(fragment);
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
