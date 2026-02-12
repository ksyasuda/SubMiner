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

interface MergedToken {
  surface: string;
  reading: string;
  headword: string;
  startPos: number;
  endPos: number;
  partOfSpeech: string;
  isMerged: boolean;
}

interface SubtitleData {
  text: string;
  tokens: MergedToken[] | null;
}

interface MpvSubtitleRenderMetrics {
  subPos: number;
  subFontSize: number;
  subScale: number;
  subMarginY: number;
  subMarginX: number;
  subFont: string;
  subSpacing: number;
  subBold: boolean;
  subItalic: boolean;
  subBorderSize: number;
  subShadowOffset: number;
  subAssOverride: string;
  subScaleByWindow: boolean;
  subUseMargins: boolean;
  osdHeight: number;
  osdDimensions: {
    w: number;
    h: number;
    ml: number;
    mr: number;
    mt: number;
    mb: number;
  } | null;
}

interface Keybinding {
  key: string;
  command: (string | number)[] | null;
}

interface SubtitlePosition {
  yPercent: number;
  invisibleOffsetXPx?: number;
  invisibleOffsetYPx?: number;
}

type SecondarySubMode = "hidden" | "visible" | "hover";

interface SubtitleStyleConfig {
  fontFamily?: string;
  fontSize?: number;
  fontColor?: string;
  fontWeight?: string;
  fontStyle?: string;
  backgroundColor?: string;
  secondary?: {
    fontFamily?: string;
    fontSize?: number;
    fontColor?: string;
    fontWeight?: string;
    fontStyle?: string;
    backgroundColor?: string;
  };
}

type JimakuConfidence = "high" | "medium" | "low";

interface JimakuMediaInfo {
  title: string;
  season: number | null;
  episode: number | null;
  confidence: JimakuConfidence;
  filename: string;
  rawTitle: string;
}

interface JimakuEntryFlags {
  anime?: boolean;
  movie?: boolean;
  adult?: boolean;
  external?: boolean;
  unverified?: boolean;
}

interface JimakuEntry {
  id: number;
  name: string;
  english_name?: string | null;
  japanese_name?: string | null;
  flags?: JimakuEntryFlags;
  last_modified?: string;
}

interface JimakuFileEntry {
  name: string;
  url: string;
  size: number;
  last_modified: string;
}

interface JimakuApiError {
  error: string;
  code?: number;
  retryAfter?: number;
}

type JimakuApiResponse<T> =
  | { ok: true; data: T }
  | { ok: false; error: JimakuApiError };

type JimakuDownloadResult =
  | { ok: true; path: string }
  | { ok: false; error: JimakuApiError };

interface KikuDuplicateCardInfo {
  noteId: number;
  expression: string;
  sentencePreview: string;
  hasAudio: boolean;
  hasImage: boolean;
  isOriginal: boolean;
}

interface KikuFieldGroupingChoice {
  keepNoteId: number;
  deleteNoteId: number;
  deleteDuplicate: boolean;
  cancelled: boolean;
}

interface KikuMergePreviewResponse {
  ok: boolean;
  compact?: Record<string, unknown>;
  full?: Record<string, unknown>;
  error?: string;
}

type RuntimeOptionId = "anki.autoUpdateNewCards" | "anki.kikuFieldGrouping";
type RuntimeOptionValue = boolean | string;
type RuntimeOptionValueType = "boolean" | "enum";

interface RuntimeOptionState {
  id: RuntimeOptionId;
  label: string;
  scope: "ankiConnect";
  valueType: RuntimeOptionValueType;
  value: RuntimeOptionValue;
  allowedValues: RuntimeOptionValue[];
  requiresRestart: boolean;
}

interface RuntimeOptionApplyResult {
  ok: boolean;
  option?: RuntimeOptionState;
  osdMessage?: string;
  requiresRestart?: boolean;
  error?: string;
}

interface SubsyncSourceTrack {
  id: number;
  label: string;
}

interface SubsyncManualPayload {
  sourceTracks: SubsyncSourceTrack[];
}

type KikuModalStep = "select" | "preview";
type KikuPreviewMode = "compact" | "full";

const subtitleRoot = document.getElementById("subtitleRoot")!;
const subtitleContainer = document.getElementById("subtitleContainer")!;
const overlay = document.getElementById("overlay")!;
const secondarySubContainer = document.getElementById("secondarySubContainer")!;
const secondarySubRoot = document.getElementById("secondarySubRoot")!;
const jimakuModal = document.getElementById("jimakuModal") as HTMLDivElement;
const jimakuTitleInput = document.getElementById(
  "jimakuTitle",
) as HTMLInputElement;
const jimakuSeasonInput = document.getElementById(
  "jimakuSeason",
) as HTMLInputElement;
const jimakuEpisodeInput = document.getElementById(
  "jimakuEpisode",
) as HTMLInputElement;
const jimakuSearchButton = document.getElementById(
  "jimakuSearch",
) as HTMLButtonElement;
const jimakuCloseButton = document.getElementById(
  "jimakuClose",
) as HTMLButtonElement;
const jimakuStatus = document.getElementById("jimakuStatus") as HTMLDivElement;
const jimakuEntriesSection = document.getElementById(
  "jimakuEntriesSection",
) as HTMLDivElement;
const jimakuEntriesList = document.getElementById(
  "jimakuEntries",
) as HTMLUListElement;
const jimakuFilesSection = document.getElementById(
  "jimakuFilesSection",
) as HTMLDivElement;
const jimakuFilesList = document.getElementById(
  "jimakuFiles",
) as HTMLUListElement;
const jimakuBroadenButton = document.getElementById(
  "jimakuBroaden",
) as HTMLButtonElement;

const kikuModal = document.getElementById(
  "kikuFieldGroupingModal",
) as HTMLDivElement;
const kikuCard1 = document.getElementById("kikuCard1") as HTMLDivElement;
const kikuCard2 = document.getElementById("kikuCard2") as HTMLDivElement;
const kikuCard1Expression = document.getElementById(
  "kikuCard1Expression",
) as HTMLDivElement;
const kikuCard2Expression = document.getElementById(
  "kikuCard2Expression",
) as HTMLDivElement;
const kikuCard1Sentence = document.getElementById(
  "kikuCard1Sentence",
) as HTMLDivElement;
const kikuCard2Sentence = document.getElementById(
  "kikuCard2Sentence",
) as HTMLDivElement;
const kikuCard1Meta = document.getElementById(
  "kikuCard1Meta",
) as HTMLDivElement;
const kikuCard2Meta = document.getElementById(
  "kikuCard2Meta",
) as HTMLDivElement;
const kikuConfirmButton = document.getElementById(
  "kikuConfirmButton",
) as HTMLButtonElement;
const kikuCancelButton = document.getElementById(
  "kikuCancelButton",
) as HTMLButtonElement;
const kikuDeleteDuplicateCheckbox = document.getElementById(
  "kikuDeleteDuplicate",
) as HTMLInputElement;
const kikuSelectionStep = document.getElementById(
  "kikuSelectionStep",
) as HTMLDivElement;
const kikuPreviewStep = document.getElementById(
  "kikuPreviewStep",
) as HTMLDivElement;
const kikuPreviewJson = document.getElementById(
  "kikuPreviewJson",
) as HTMLPreElement;
const kikuPreviewCompactButton = document.getElementById(
  "kikuPreviewCompact",
) as HTMLButtonElement;
const kikuPreviewFullButton = document.getElementById(
  "kikuPreviewFull",
) as HTMLButtonElement;
const kikuPreviewError = document.getElementById(
  "kikuPreviewError",
) as HTMLDivElement;
const kikuBackButton = document.getElementById(
  "kikuBackButton",
) as HTMLButtonElement;
const kikuFinalConfirmButton = document.getElementById(
  "kikuFinalConfirmButton",
) as HTMLButtonElement;
const kikuFinalCancelButton = document.getElementById(
  "kikuFinalCancelButton",
) as HTMLButtonElement;
const kikuHint = document.getElementById("kikuHint") as HTMLDivElement;
const runtimeOptionsModal = document.getElementById(
  "runtimeOptionsModal",
) as HTMLDivElement;
const runtimeOptionsClose = document.getElementById(
  "runtimeOptionsClose",
) as HTMLButtonElement;
const runtimeOptionsList = document.getElementById(
  "runtimeOptionsList",
) as HTMLUListElement;
const runtimeOptionsStatus = document.getElementById(
  "runtimeOptionsStatus",
) as HTMLDivElement;
const subsyncModal = document.getElementById("subsyncModal") as HTMLDivElement;
const subsyncCloseButton = document.getElementById(
  "subsyncClose",
) as HTMLButtonElement;
const subsyncEngineAlass = document.getElementById(
  "subsyncEngineAlass",
) as HTMLInputElement;
const subsyncEngineFfsubsync = document.getElementById(
  "subsyncEngineFfsubsync",
) as HTMLInputElement;
const subsyncSourceLabel = document.getElementById(
  "subsyncSourceLabel",
) as HTMLLabelElement;
const subsyncSourceSelect = document.getElementById(
  "subsyncSourceSelect",
) as HTMLSelectElement;
const subsyncRunButton = document.getElementById(
  "subsyncRun",
) as HTMLButtonElement;
const subsyncStatus = document.getElementById(
  "subsyncStatus",
) as HTMLDivElement;

const overlayLayerFromPreload = window.electronAPI.getOverlayLayer();
const overlayLayerFromQuery =
  new URLSearchParams(window.location.search).get("layer") === "invisible"
    ? "invisible"
    : "visible";
const overlayLayer =
  overlayLayerFromPreload === "visible" ||
  overlayLayerFromPreload === "invisible"
    ? overlayLayerFromPreload
    : overlayLayerFromQuery;
const isInvisibleLayer = overlayLayer === "invisible";
const isLinuxPlatform = navigator.platform.toLowerCase().includes("linux");
const isMacOSPlatform =
  navigator.platform.toLowerCase().includes("mac") ||
  /mac/i.test(navigator.userAgent);
// Linux passthrough forwarding is not reliable for this overlay; keep pointer
// routing local so hover lookup, drag-reposition, and key handling remain usable.
const shouldToggleMouseIgnore = !isLinuxPlatform;
const INVISIBLE_POSITION_EDIT_TOGGLE_CODE = "KeyP";
const INVISIBLE_POSITION_STEP_PX = 1;
const INVISIBLE_POSITION_STEP_FAST_PX = 4;
const INVISIBLE_MACOS_VERTICAL_NUDGE_PX = 4;

let isOverSubtitle = false;
let isDragging = false;
let dragStartY = 0;
let startYPercent = 0;
let currentYPercent: number | null = null;
let persistedSubtitlePosition: SubtitlePosition = { yPercent: 10 };
let jimakuModalOpen = false;
let jimakuEntries: JimakuEntry[] = [];
let jimakuFiles: JimakuFileEntry[] = [];
let selectedEntryIndex = 0;
let selectedFileIndex = 0;
let currentEpisodeFilter: number | null = null;
let currentEntryId: number | null = null;

let kikuModalOpen = false;
let kikuSelectedCard: 1 | 2 = 1;
let kikuOriginalData: KikuDuplicateCardInfo | null = null;
let kikuDuplicateData: KikuDuplicateCardInfo | null = null;
let kikuModalStep: KikuModalStep = "select";
let kikuPreviewMode: KikuPreviewMode = "compact";
let kikuPendingChoice: KikuFieldGroupingChoice | null = null;
let kikuPreviewCompactData: Record<string, unknown> | null = null;
let kikuPreviewFullData: Record<string, unknown> | null = null;
let runtimeOptionsModalOpen = false;
let runtimeOptions: RuntimeOptionState[] = [];
let runtimeOptionSelectedIndex = 0;
let runtimeOptionDraftValues = new Map<RuntimeOptionId, RuntimeOptionValue>();
let subsyncModalOpen = false;
let subsyncSourceTracks: SubsyncSourceTrack[] = [];
let subsyncSubmitting = false;
const DEFAULT_MPV_SUBTITLE_RENDER_METRICS: MpvSubtitleRenderMetrics = {
  subPos: 100,
  subFontSize: 38,
  subScale: 1,
  subMarginY: 34,
  subMarginX: 19,
  subFont: "sans-serif",
  subSpacing: 0,
  subBold: false,
  subItalic: false,
  subBorderSize: 2.5,
  subShadowOffset: 0,
  subAssOverride: "yes",
  subScaleByWindow: true,
  subUseMargins: true,
  osdHeight: 720,
  osdDimensions: null,
};
let mpvSubtitleRenderMetrics: MpvSubtitleRenderMetrics = {
  ...DEFAULT_MPV_SUBTITLE_RENDER_METRICS,
};
let invisiblePositionEditMode = false;
let invisiblePositionEditStartX = 0;
let invisiblePositionEditStartY = 0;
let invisibleSubtitleOffsetXPx = 0;
let invisibleSubtitleOffsetYPx = 0;
let invisibleLayoutBaseLeftPx = 0;
let invisibleLayoutBaseBottomPx: number | null = null;
let invisibleLayoutBaseTopPx: number | null = null;
let invisiblePositionEditHud: HTMLDivElement | null = null;
let currentInvisibleSubtitleLineCount = 1;
let lastHoverSelectionKey = "";
let lastHoverSelectionNode: Text | null = null;
const wordSegmenter =
  typeof Intl !== "undefined" && "Segmenter" in Intl
    ? new Intl.Segmenter(undefined, { granularity: "word" })
    : null;

function isAnySettingsModalOpen(): boolean {
  return runtimeOptionsModalOpen || subsyncModalOpen || kikuModalOpen;
}

function syncSettingsModalSubtitleSuppression(): void {
  const suppressSubtitles = isAnySettingsModalOpen();
  document.body.classList.toggle("settings-modal-open", suppressSubtitles);
  if (suppressSubtitles) {
    isOverSubtitle = false;
  }
}

function normalizeSubtitle(text: string, trim = true): string {
  if (!text) return "";

  let normalized = text.replace(/\\N/g, "\n").replace(/\\n/g, "\n");

  normalized = normalized.replace(/\{[^}]*\}/g, "");

  return trim ? normalized.trim() : normalized;
}

function renderWithTokens(tokens: MergedToken[]): void {
  const fragment = document.createDocumentFragment();

  for (const token of tokens) {
    const surface = token.surface;

    if (surface.includes("\n")) {
      const parts = surface.split("\n");
      for (let i = 0; i < parts.length; i++) {
        if (parts[i]) {
          const span = document.createElement("span");
          span.className = "word";
          span.textContent = parts[i];
          if (token.reading) {
            span.dataset.reading = token.reading;
          }
          if (token.headword) {
            span.dataset.headword = token.headword;
          }
          fragment.appendChild(span);
        }
        if (i < parts.length - 1) {
          fragment.appendChild(document.createElement("br"));
        }
      }
    } else {
      const span = document.createElement("span");
      span.className = "word";
      span.textContent = surface;
      if (token.reading) {
        span.dataset.reading = token.reading;
      }
      if (token.headword) {
        span.dataset.headword = token.headword;
      }
      fragment.appendChild(span);
    }
  }

  subtitleRoot.appendChild(fragment);
}

function renderCharacterLevel(text: string): void {
  const fragment = document.createDocumentFragment();

  for (const char of text) {
    if (char === "\n") {
      fragment.appendChild(document.createElement("br"));
    } else {
      const span = document.createElement("span");
      span.className = "c";
      span.textContent = char;
      fragment.appendChild(span);
    }
  }

  subtitleRoot.appendChild(fragment);
}

function renderPlainTextPreserveLineBreaks(text: string): void {
  const lines = text.split("\n");
  const fragment = document.createDocumentFragment();

  for (let i = 0; i < lines.length; i += 1) {
    fragment.appendChild(document.createTextNode(lines[i]));
    if (i < lines.length - 1) {
      fragment.appendChild(document.createElement("br"));
    }
  }

  subtitleRoot.appendChild(fragment);
}

function renderSubtitle(data: SubtitleData | string): void {
  subtitleRoot.innerHTML = "";
  lastHoverSelectionKey = "";
  lastHoverSelectionNode = null;

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

  if (!text) {
    return;
  }

  if (isInvisibleLayer) {
    // Keep natural kerning/shaping for accurate hitbox alignment with mpv.
    const normalizedInvisible = normalizeSubtitle(text, false);
    currentInvisibleSubtitleLineCount = Math.max(
      1,
      normalizedInvisible.split("\n").length,
    );
    renderPlainTextPreserveLineBreaks(normalizedInvisible);
    return;
  }

  const normalized = normalizeSubtitle(text);

  if (tokens && tokens.length > 0) {
    renderWithTokens(tokens);
  } else {
    renderCharacterLevel(normalized);
  }
}

function handleMouseEnter(): void {
  isOverSubtitle = true;
  overlay.classList.add("interactive");
  if (shouldToggleMouseIgnore) {
    window.electronAPI.setIgnoreMouseEvents(false);
  }
}

function handleMouseLeave(): void {
  isOverSubtitle = false;
  const yomitanPopup = document.querySelector('iframe[id^="yomitan-popup"]');
  if (
    !yomitanPopup &&
    !jimakuModalOpen &&
    !kikuModalOpen &&
    !runtimeOptionsModalOpen &&
    !subsyncModalOpen &&
    !invisiblePositionEditMode
  ) {
    overlay.classList.remove("interactive");
    if (shouldToggleMouseIgnore) {
      window.electronAPI.setIgnoreMouseEvents(true, { forward: true });
    }
  }
}

function clampYPercent(yPercent: number): number {
  return Math.max(2, Math.min(80, yPercent));
}

function getCurrentYPercent(): number {
  if (currentYPercent !== null) {
    return currentYPercent;
  }
  const marginBottom = parseFloat(subtitleContainer.style.marginBottom) || 60;
  const windowHeight = window.innerHeight;
  currentYPercent = clampYPercent((marginBottom / windowHeight) * 100);
  return currentYPercent;
}

function applyYPercent(yPercent: number): void {
  const clampedPercent = clampYPercent(yPercent);
  currentYPercent = clampedPercent;
  const marginBottom = (clampedPercent / 100) * window.innerHeight;

  subtitleContainer.style.position = "";
  subtitleContainer.style.left = "";
  subtitleContainer.style.top = "";
  subtitleContainer.style.right = "";
  subtitleContainer.style.transform = "";

  subtitleContainer.style.marginBottom = `${marginBottom}px`;
}

function updatePersistedSubtitlePosition(position: SubtitlePosition | null): void {
  const nextYPercent =
    position && typeof position.yPercent === "number" && Number.isFinite(position.yPercent)
      ? position.yPercent
      : persistedSubtitlePosition.yPercent;
  const nextXOffset =
    position && typeof position.invisibleOffsetXPx === "number" && Number.isFinite(position.invisibleOffsetXPx)
      ? position.invisibleOffsetXPx
      : 0;
  const nextYOffset =
    position && typeof position.invisibleOffsetYPx === "number" && Number.isFinite(position.invisibleOffsetYPx)
      ? position.invisibleOffsetYPx
      : 0;
  persistedSubtitlePosition = {
    yPercent: nextYPercent,
    invisibleOffsetXPx: nextXOffset,
    invisibleOffsetYPx: nextYOffset,
  };
}

function persistSubtitlePositionPatch(patch: Partial<SubtitlePosition>): void {
  const nextPosition: SubtitlePosition = {
    yPercent:
      typeof patch.yPercent === "number" && Number.isFinite(patch.yPercent)
        ? patch.yPercent
        : persistedSubtitlePosition.yPercent,
    invisibleOffsetXPx:
      typeof patch.invisibleOffsetXPx === "number" &&
      Number.isFinite(patch.invisibleOffsetXPx)
        ? patch.invisibleOffsetXPx
        : persistedSubtitlePosition.invisibleOffsetXPx ?? 0,
    invisibleOffsetYPx:
      typeof patch.invisibleOffsetYPx === "number" &&
      Number.isFinite(patch.invisibleOffsetYPx)
        ? patch.invisibleOffsetYPx
        : persistedSubtitlePosition.invisibleOffsetYPx ?? 0,
  };
  persistedSubtitlePosition = nextPosition;
  window.electronAPI.saveSubtitlePosition(nextPosition);
}

function applyStoredSubtitlePosition(
  position: SubtitlePosition | null,
  source: string,
): void {
  updatePersistedSubtitlePosition(position);
  if (position && position.yPercent !== undefined) {
    applyYPercent(position.yPercent);
    console.log(
      "Applied subtitle position from",
      source,
      ":",
      position.yPercent,
      "%",
    );
  } else {
    const defaultMarginBottom = 60;
    const defaultYPercent = (defaultMarginBottom / window.innerHeight) * 100;
    applyYPercent(defaultYPercent);
    console.log("Applied default subtitle position from", source);
  }
}

function applyInvisibleSubtitleOffsetPosition(): void {
  const nextLeft = invisibleLayoutBaseLeftPx + invisibleSubtitleOffsetXPx;
  subtitleContainer.style.left = `${nextLeft}px`;

  if (invisibleLayoutBaseBottomPx !== null) {
    subtitleContainer.style.bottom = `${Math.max(0, invisibleLayoutBaseBottomPx + invisibleSubtitleOffsetYPx)}px`;
    subtitleContainer.style.top = "";
    return;
  }

  if (invisibleLayoutBaseTopPx !== null) {
    subtitleContainer.style.top = `${Math.max(0, invisibleLayoutBaseTopPx - invisibleSubtitleOffsetYPx)}px`;
    subtitleContainer.style.bottom = "";
  }
}

function updateInvisiblePositionEditHud(): void {
  if (!invisiblePositionEditHud) return;
  invisiblePositionEditHud.textContent =
    `Position Edit  Ctrl/Cmd+Shift+P toggle  Arrow keys move  Enter/Ctrl+S save  Esc cancel  x:${Math.round(invisibleSubtitleOffsetXPx)} y:${Math.round(invisibleSubtitleOffsetYPx)}`;
}

function setInvisiblePositionEditMode(enabled: boolean): void {
  if (!isInvisibleLayer) return;
  if (invisiblePositionEditMode === enabled) return;
  invisiblePositionEditMode = enabled;
  document.body.classList.toggle("invisible-position-edit", enabled);
  if (enabled) {
    invisiblePositionEditStartX = invisibleSubtitleOffsetXPx;
    invisiblePositionEditStartY = invisibleSubtitleOffsetYPx;
    overlay.classList.add("interactive");
    if (shouldToggleMouseIgnore) {
      window.electronAPI.setIgnoreMouseEvents(false);
    }
  } else if (!isOverSubtitle && !isAnySettingsModalOpen()) {
    overlay.classList.remove("interactive");
    if (shouldToggleMouseIgnore) {
      window.electronAPI.setIgnoreMouseEvents(true, { forward: true });
    }
  }
  updateInvisiblePositionEditHud();
}

function applyInvisibleStoredSubtitlePosition(
  position: SubtitlePosition | null,
  source: string,
): void {
  updatePersistedSubtitlePosition(position);
  invisibleSubtitleOffsetXPx = persistedSubtitlePosition.invisibleOffsetXPx ?? 0;
  invisibleSubtitleOffsetYPx = persistedSubtitlePosition.invisibleOffsetYPx ?? 0;
  applyInvisibleSubtitleOffsetPosition();
  console.log(
    "[invisible-overlay] Applied subtitle offset from",
    source,
    `${invisibleSubtitleOffsetXPx}px`,
    `${invisibleSubtitleOffsetYPx}px`,
  );
  updateInvisiblePositionEditHud();
}

function applySubtitleFontSize(fontSize: number): void {
  const clampedSize = Math.max(10, fontSize);
  subtitleRoot.style.fontSize = `${clampedSize}px`;
  document.documentElement.style.setProperty(
    "--subtitle-font-size",
    `${clampedSize}px`,
  );
}

function coerceFiniteNumber(
  value: unknown,
  fallback: number,
  min?: number,
  max?: number,
): number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    return fallback;
  }
  let next = value;
  if (typeof min === "number") next = Math.max(min, next);
  if (typeof max === "number") next = Math.min(max, next);
  return next;
}

function sanitizeMpvSubtitleRenderMetrics(
  metrics: Partial<MpvSubtitleRenderMetrics> | null | undefined,
): MpvSubtitleRenderMetrics {
  const dims = metrics?.osdDimensions;
  const nextOsdDimensions =
    dims &&
    typeof dims.w === "number" &&
    typeof dims.h === "number" &&
    typeof dims.ml === "number" &&
    typeof dims.mr === "number" &&
    typeof dims.mt === "number" &&
    typeof dims.mb === "number"
      ? {
          w: coerceFiniteNumber(dims.w, 0, 1, 100000),
          h: coerceFiniteNumber(dims.h, 0, 1, 100000),
          ml: coerceFiniteNumber(dims.ml, 0, 0, 100000),
          mr: coerceFiniteNumber(dims.mr, 0, 0, 100000),
          mt: coerceFiniteNumber(dims.mt, 0, 0, 100000),
          mb: coerceFiniteNumber(dims.mb, 0, 0, 100000),
        }
      : dims === null
        ? null
        : mpvSubtitleRenderMetrics.osdDimensions;

  return {
    subPos: coerceFiniteNumber(
      metrics?.subPos,
      mpvSubtitleRenderMetrics.subPos,
      0,
      150,
    ),
    subFontSize: coerceFiniteNumber(
      metrics?.subFontSize,
      mpvSubtitleRenderMetrics.subFontSize,
      1,
      200,
    ),
    subScale: coerceFiniteNumber(
      metrics?.subScale,
      mpvSubtitleRenderMetrics.subScale,
      0.1,
      10,
    ),
    subMarginY: coerceFiniteNumber(
      metrics?.subMarginY,
      mpvSubtitleRenderMetrics.subMarginY,
      0,
      200,
    ),
    subMarginX: coerceFiniteNumber(
      metrics?.subMarginX,
      mpvSubtitleRenderMetrics.subMarginX,
      0,
      200,
    ),
    subFont:
      typeof metrics?.subFont === "string" && metrics.subFont.trim().length > 0
        ? metrics.subFont.trim()
        : mpvSubtitleRenderMetrics.subFont,
    subSpacing: coerceFiniteNumber(
      metrics?.subSpacing,
      mpvSubtitleRenderMetrics.subSpacing,
      -100,
      100,
    ),
    subBold:
      typeof metrics?.subBold === "boolean"
        ? metrics.subBold
        : mpvSubtitleRenderMetrics.subBold,
    subItalic:
      typeof metrics?.subItalic === "boolean"
        ? metrics.subItalic
        : mpvSubtitleRenderMetrics.subItalic,
    subBorderSize: coerceFiniteNumber(
      metrics?.subBorderSize,
      mpvSubtitleRenderMetrics.subBorderSize,
      0,
      100,
    ),
    subShadowOffset: coerceFiniteNumber(
      metrics?.subShadowOffset,
      mpvSubtitleRenderMetrics.subShadowOffset,
      0,
      100,
    ),
    subAssOverride:
      typeof metrics?.subAssOverride === "string" &&
      metrics.subAssOverride.trim().length > 0
        ? metrics.subAssOverride.trim()
        : mpvSubtitleRenderMetrics.subAssOverride,
    subScaleByWindow:
      typeof metrics?.subScaleByWindow === "boolean"
        ? metrics.subScaleByWindow
        : mpvSubtitleRenderMetrics.subScaleByWindow,
    subUseMargins:
      typeof metrics?.subUseMargins === "boolean"
        ? metrics.subUseMargins
        : mpvSubtitleRenderMetrics.subUseMargins,
    osdHeight: coerceFiniteNumber(
      metrics?.osdHeight,
      mpvSubtitleRenderMetrics.osdHeight,
      1,
      10000,
    ),
    osdDimensions: nextOsdDimensions,
  };
}

function applyInvisibleSubtitleLayoutFromMpvMetrics(
  metrics: Partial<MpvSubtitleRenderMetrics> | null | undefined,
  source: string,
): void {
  mpvSubtitleRenderMetrics = sanitizeMpvSubtitleRenderMetrics(metrics);

  const dims = mpvSubtitleRenderMetrics.osdDimensions;
  const dpr = window.devicePixelRatio || 1;
  // On macOS, mpv osd-dimensions can be reported in either physical pixels or
  // point-like units. Infer the osd->CSS scale from current viewport ratio.
  const osdToCssScale =
    isMacOSPlatform && dims
      ? (() => {
          const ratios = [
            dims.w / Math.max(1, window.innerWidth),
            dims.h / Math.max(1, window.innerHeight),
          ].filter((value) => Number.isFinite(value) && value > 0);
          const avgRatio =
            ratios.length > 0
              ? ratios.reduce((sum, value) => sum + value, 0) / ratios.length
              : dpr;
          return avgRatio > 1.25 ? avgRatio : 1;
        })()
      : dpr;
  const renderAreaHeight = dims ? dims.h / osdToCssScale : window.innerHeight;
  const renderAreaWidth = dims ? dims.w / osdToCssScale : window.innerWidth;
  const videoLeftInset = dims ? dims.ml / osdToCssScale : 0;
  const videoRightInset = dims ? dims.mr / osdToCssScale : 0;
  const videoTopInset = dims ? dims.mt / osdToCssScale : 0;
  const videoBottomInset = dims ? dims.mb / osdToCssScale : 0;
  const anchorToVideoArea = !mpvSubtitleRenderMetrics.subUseMargins;
  const leftInset = anchorToVideoArea ? videoLeftInset : 0;
  const rightInset = anchorToVideoArea ? videoRightInset : 0;
  const topInset = anchorToVideoArea ? videoTopInset : 0;
  const bottomInset = anchorToVideoArea ? videoBottomInset : 0;

  const videoHeight = renderAreaHeight - videoTopInset - videoBottomInset;
  const scaleRefHeight = mpvSubtitleRenderMetrics.subScaleByWindow
    ? renderAreaHeight
    : videoHeight;
  const pxPerScaledPixel = Math.max(0.1, scaleRefHeight / 720);
  const computedFontSize =
    mpvSubtitleRenderMetrics.subFontSize *
    mpvSubtitleRenderMetrics.subScale *
    (isLinuxPlatform ? 1 : pxPerScaledPixel);

  const macOsFontCompensation = isMacOSPlatform ? 0.87 : 1;
  const effectiveFontSize = computedFontSize * macOsFontCompensation;
  applySubtitleFontSize(effectiveFontSize);

  const alignment = 2;
  const hAlign = ((alignment - 1) % 3) as 0 | 1 | 2; // 0=left, 1=center, 2=right
  const vAlign = Math.floor((alignment - 1) / 3) as 0 | 1 | 2; // 0=bottom, 1=middle, 2=top

  const marginY = mpvSubtitleRenderMetrics.subMarginY * pxPerScaledPixel;
  const marginX = Math.max(
    0,
    mpvSubtitleRenderMetrics.subMarginX * pxPerScaledPixel,
  );
  const horizontalAvailable = Math.max(
    0,
    renderAreaWidth - leftInset - rightInset - Math.round(marginX * 2),
  );

  const effectiveBorderSize = mpvSubtitleRenderMetrics.subBorderSize * pxPerScaledPixel;
  document.documentElement.style.setProperty(
    "--sub-border-size",
    `${effectiveBorderSize}px`,
  );

  subtitleContainer.style.position = "absolute";
  subtitleContainer.style.maxWidth = `${horizontalAvailable}px`;
  subtitleContainer.style.width = `${horizontalAvailable}px`;
  subtitleContainer.style.padding = "0";
  subtitleContainer.style.background = "transparent";
  subtitleContainer.style.marginBottom = "0";
  subtitleContainer.style.pointerEvents = "none";

  // Horizontal positioning based on \an alignment.
  // All alignments position the container at the left margin with full available
  // width and use text-align for alignment. This avoids translateX(-50%) which
  // can cause subpixel rounding artifacts from GPU rasterization.
  subtitleContainer.style.left = `${leftInset + marginX}px`;
  subtitleContainer.style.right = "";
  subtitleContainer.style.transform = "";
  subtitleContainer.style.textAlign = "";
  if (hAlign === 0) {
    subtitleContainer.style.textAlign = "left";
    subtitleRoot.style.textAlign = "left";
  } else if (hAlign === 2) {
    subtitleContainer.style.textAlign = "right";
    subtitleRoot.style.textAlign = "right";
  } else {
    subtitleContainer.style.textAlign = "center";
    subtitleRoot.style.textAlign = "center";
  }
  subtitleRoot.style.display = "inline-block";
  subtitleRoot.style.maxWidth = "100%";
  subtitleRoot.style.pointerEvents = "auto";

  // Vertical positioning based on \an alignment
  const lineCount = Math.max(1, currentInvisibleSubtitleLineCount);
  const multiline = lineCount > 1;
  const baselineCompensationFactor =
    lineCount >= 3 ? 0.46 : multiline ? 0.58 : 0.7;
  const baselineCompensationPx = Math.max(
    0,
    effectiveFontSize * baselineCompensationFactor,
  );
  if (vAlign === 2) {
    subtitleContainer.style.top = `${Math.max(0, topInset + marginY - baselineCompensationPx)}px`;
    subtitleContainer.style.bottom = "";
  } else if (vAlign === 1) {
    subtitleContainer.style.top = "50%";
    subtitleContainer.style.bottom = "";
    subtitleContainer.style.transform = "translateY(-50%)";
  } else {
    const subPosMargin =
      ((100 - mpvSubtitleRenderMetrics.subPos) / 100) * renderAreaHeight;
    const effectiveMargin = Math.max(marginY, subPosMargin);
    const bottomPx = Math.max(
      0,
      bottomInset + effectiveMargin + baselineCompensationPx,
    );
    subtitleContainer.style.top = "";
    subtitleContainer.style.bottom = `${bottomPx}px`;
  }

  subtitleRoot.style.setProperty(
    "line-height",
    isMacOSPlatform
      ? lineCount >= 3
        ? "1.18"
        : multiline
          ? "1.08"
          : "0.86"
      : "normal",
    isMacOSPlatform ? "important" : "",
  );

  const rawFont = mpvSubtitleRenderMetrics.subFont;
  const strippedFont = rawFont
    .replace(
      /\s+(Regular|Bold|Italic|Light|Medium|Semi\s*Bold|Extra\s*Bold|Extra\s*Light|Thin|Black|Heavy|Demi\s*Bold|Book|Condensed)\s*$/i,
      "",
    )
    .trim();
  subtitleRoot.style.fontFamily =
    strippedFont !== rawFont
      ? `"${rawFont}", "${strippedFont}", sans-serif`
      : `"${rawFont}", sans-serif`;
  const effectiveSpacing = mpvSubtitleRenderMetrics.subSpacing;
  subtitleRoot.style.setProperty(
    "letter-spacing",
    Math.abs(effectiveSpacing) > 0.0001
      ? `${effectiveSpacing * pxPerScaledPixel * (isMacOSPlatform ? 0.7 : 1)}px`
      : isMacOSPlatform
        ? "-0.02em"
        : "0px",
    isMacOSPlatform ? "important" : "",
  );
  subtitleRoot.style.fontKerning = isMacOSPlatform ? "auto" : "none";
  const effectiveBold = mpvSubtitleRenderMetrics.subBold;
  const effectiveItalic = mpvSubtitleRenderMetrics.subItalic;
  subtitleRoot.style.fontWeight = effectiveBold ? "700" : "400";
  subtitleRoot.style.fontStyle = effectiveItalic ? "italic" : "normal";
  const scaleX = 1;
  const scaleY = 1;
  if (Math.abs(scaleX - 1) > 0.0001 || Math.abs(scaleY - 1) > 0.0001) {
    subtitleRoot.style.transform = `scale(${scaleX}, ${scaleY})`;
    subtitleRoot.style.transformOrigin = "50% 100%";
  } else {
    subtitleRoot.style.transform = "";
    subtitleRoot.style.transformOrigin = "";
  }

  // CSS line-height: normal adds "half-leading" — extra space equally distributed
  // above and below text within each line box. This means the text's visual bottom
  // is above the element's bottom edge by halfLeading pixels. We must compensate
  // by shifting the element down by that amount so glyph positions better match mpv.
  const computedLineHeight = parseFloat(
    getComputedStyle(subtitleRoot).lineHeight,
  );
  if (
    Number.isFinite(computedLineHeight) &&
    computedLineHeight > effectiveFontSize
  ) {
    const halfLeading = (computedLineHeight - effectiveFontSize) / 2;
    if (vAlign === 0 && halfLeading > 0.5) {
      const currentBottom = parseFloat(subtitleContainer.style.bottom);
      if (Number.isFinite(currentBottom)) {
        subtitleContainer.style.bottom = `${Math.max(0, currentBottom - halfLeading)}px`;
      }
    } else if (vAlign === 2 && halfLeading > 0.5) {
      const currentTop = parseFloat(subtitleContainer.style.top);
      if (Number.isFinite(currentTop)) {
        subtitleContainer.style.top = `${Math.max(0, currentTop - halfLeading)}px`;
      }
    }
  }

  if (isMacOSPlatform && vAlign === 0) {
    const currentBottom = parseFloat(subtitleContainer.style.bottom);
    if (Number.isFinite(currentBottom)) {
      subtitleContainer.style.bottom = `${Math.max(0, currentBottom + INVISIBLE_MACOS_VERTICAL_NUDGE_PX)}px`;
    }
  }

  invisibleLayoutBaseLeftPx = parseFloat(subtitleContainer.style.left) || 0;
  invisibleLayoutBaseBottomPx = Number.isFinite(parseFloat(subtitleContainer.style.bottom))
    ? parseFloat(subtitleContainer.style.bottom)
    : null;
  invisibleLayoutBaseTopPx = Number.isFinite(parseFloat(subtitleContainer.style.top))
    ? parseFloat(subtitleContainer.style.top)
    : null;
  applyInvisibleSubtitleOffsetPosition();
  updateInvisiblePositionEditHud();
  console.log("[invisible-overlay] Applied mpv subtitle render metrics from", source);
}

function setJimakuStatus(message: string, isError = false): void {
  jimakuStatus.textContent = message;
  jimakuStatus.style.color = isError
    ? "rgba(255, 120, 120, 0.95)"
    : "rgba(255, 255, 255, 0.8)";
}

function resetJimakuLists(): void {
  jimakuEntries = [];
  jimakuFiles = [];
  selectedEntryIndex = 0;
  selectedFileIndex = 0;
  currentEntryId = null;
  jimakuEntriesList.innerHTML = "";
  jimakuFilesList.innerHTML = "";
  jimakuEntriesSection.classList.add("hidden");
  jimakuFilesSection.classList.add("hidden");
  jimakuBroadenButton.classList.add("hidden");
}

function openJimakuModal(): void {
  if (isInvisibleLayer) return;
  if (jimakuModalOpen) return;
  jimakuModalOpen = true;
  overlay.classList.add("interactive");
  jimakuModal.classList.remove("hidden");
  jimakuModal.setAttribute("aria-hidden", "false");
  setJimakuStatus("Loading media info...");
  resetJimakuLists();

  window.electronAPI
    .getJimakuMediaInfo()
    .then((info: JimakuMediaInfo) => {
      jimakuTitleInput.value = info.title || "";
      jimakuSeasonInput.value = info.season ? String(info.season) : "";
      jimakuEpisodeInput.value = info.episode ? String(info.episode) : "";
      currentEpisodeFilter = info.episode ?? null;

      if (info.confidence === "high" && info.title && info.episode) {
        performJimakuSearch();
      } else if (info.title) {
        setJimakuStatus("Check title/season/episode and press Search.");
      } else {
        setJimakuStatus("Enter title/season/episode and press Search.");
      }
    })
    .catch(() => {
      setJimakuStatus("Failed to load media info.", true);
    });
}

function closeJimakuModal(): void {
  if (!jimakuModalOpen) return;
  jimakuModalOpen = false;
  jimakuModal.classList.add("hidden");
  jimakuModal.setAttribute("aria-hidden", "true");
  if (
    !isOverSubtitle &&
    !kikuModalOpen &&
    !runtimeOptionsModalOpen &&
    !subsyncModalOpen
  ) {
    overlay.classList.remove("interactive");
  }
  resetJimakuLists();
}

function formatRuntimeOptionValue(value: RuntimeOptionValue): string {
  if (typeof value === "boolean") {
    return value ? "On" : "Off";
  }
  return value;
}

function setRuntimeOptionsStatus(message: string, isError = false): void {
  runtimeOptionsStatus.textContent = message;
  runtimeOptionsStatus.classList.toggle("error", isError);
}

function getRuntimeOptionDisplayValue(
  option: RuntimeOptionState,
): RuntimeOptionValue {
  return runtimeOptionDraftValues.get(option.id) ?? option.value;
}

function renderRuntimeOptionsList(): void {
  runtimeOptionsList.innerHTML = "";
  runtimeOptions.forEach((option, index) => {
    const li = document.createElement("li");
    li.className = "runtime-options-item";
    li.classList.toggle("active", index === runtimeOptionSelectedIndex);

    const label = document.createElement("div");
    label.className = "runtime-options-label";
    label.textContent = option.label;

    const value = document.createElement("div");
    value.className = "runtime-options-value";
    value.textContent = `Value: ${formatRuntimeOptionValue(getRuntimeOptionDisplayValue(option))}`;
    value.title = "Click to cycle value, right-click to cycle backward";

    const allowed = document.createElement("div");
    allowed.className = "runtime-options-allowed";
    allowed.textContent = `Allowed: ${option.allowedValues
      .map((entry) => formatRuntimeOptionValue(entry))
      .join(" | ")}`;

    li.appendChild(label);
    li.appendChild(value);
    li.appendChild(allowed);
    li.addEventListener("click", () => {
      runtimeOptionSelectedIndex = index;
      renderRuntimeOptionsList();
    });
    li.addEventListener("dblclick", () => {
      runtimeOptionSelectedIndex = index;
      void applySelectedRuntimeOption();
    });

    value.addEventListener("click", (event) => {
      event.stopPropagation();
      runtimeOptionSelectedIndex = index;
      cycleRuntimeDraftValue(1);
    });
    value.addEventListener("contextmenu", (event) => {
      event.preventDefault();
      event.stopPropagation();
      runtimeOptionSelectedIndex = index;
      cycleRuntimeDraftValue(-1);
    });

    runtimeOptionsList.appendChild(li);
  });
}

function updateRuntimeOptions(options: RuntimeOptionState[]): void {
  const previousId =
    runtimeOptions[runtimeOptionSelectedIndex]?.id ?? runtimeOptions[0]?.id;
  runtimeOptions = options;

  runtimeOptionDraftValues.clear();
  for (const option of runtimeOptions) {
    runtimeOptionDraftValues.set(option.id, option.value);
  }

  const nextIndex = runtimeOptions.findIndex(
    (option) => option.id === previousId,
  );
  runtimeOptionSelectedIndex = nextIndex >= 0 ? nextIndex : 0;
  renderRuntimeOptionsList();
}

function closeRuntimeOptionsModal(): void {
  if (!runtimeOptionsModalOpen) return;
  runtimeOptionsModalOpen = false;
  syncSettingsModalSubtitleSuppression();
  runtimeOptionsModal.classList.add("hidden");
  runtimeOptionsModal.setAttribute("aria-hidden", "true");
  window.electronAPI.notifyOverlayModalClosed("runtime-options");
  setRuntimeOptionsStatus("");
  if (
    !isOverSubtitle &&
    !jimakuModalOpen &&
    !kikuModalOpen &&
    !subsyncModalOpen
  ) {
    overlay.classList.remove("interactive");
  }
}

async function openRuntimeOptionsModal(): Promise<void> {
  if (isInvisibleLayer) return;
  const options = await window.electronAPI.getRuntimeOptions();
  updateRuntimeOptions(options);
  runtimeOptionsModalOpen = true;
  syncSettingsModalSubtitleSuppression();
  overlay.classList.add("interactive");
  runtimeOptionsModal.classList.remove("hidden");
  runtimeOptionsModal.setAttribute("aria-hidden", "false");
  setRuntimeOptionsStatus(
    "Use arrow keys. Click value to cycle. Enter or double-click to apply.",
  );
}

function getSelectedRuntimeOption(): RuntimeOptionState | null {
  if (runtimeOptions.length === 0) return null;
  if (runtimeOptionSelectedIndex < 0) return null;
  if (runtimeOptionSelectedIndex >= runtimeOptions.length) return null;
  return runtimeOptions[runtimeOptionSelectedIndex];
}

function cycleRuntimeDraftValue(direction: 1 | -1): void {
  const option = getSelectedRuntimeOption();
  if (!option || option.allowedValues.length === 0) return;
  const currentValue = getRuntimeOptionDisplayValue(option);
  const currentIndex = option.allowedValues.findIndex(
    (value) => value === currentValue,
  );
  const safeIndex = currentIndex >= 0 ? currentIndex : 0;
  const nextIndex =
    direction === 1
      ? (safeIndex + 1) % option.allowedValues.length
      : (safeIndex - 1 + option.allowedValues.length) %
        option.allowedValues.length;
  runtimeOptionDraftValues.set(option.id, option.allowedValues[nextIndex]);
  renderRuntimeOptionsList();
  setRuntimeOptionsStatus(
    `Selected ${option.label}: ${formatRuntimeOptionValue(option.allowedValues[nextIndex])}`,
  );
}

async function applySelectedRuntimeOption(): Promise<void> {
  const option = getSelectedRuntimeOption();
  if (!option) return;
  const nextValue = getRuntimeOptionDisplayValue(option);
  const result: RuntimeOptionApplyResult =
    await window.electronAPI.setRuntimeOptionValue(option.id, nextValue);
  if (!result.ok) {
    setRuntimeOptionsStatus(result.error || "Failed to apply option", true);
    return;
  }

  if (result.option) {
    runtimeOptionDraftValues.set(result.option.id, result.option.value);
  }
  const latest = await window.electronAPI.getRuntimeOptions();
  updateRuntimeOptions(latest);
  setRuntimeOptionsStatus(result.osdMessage || "Option applied.");
}

function handleRuntimeOptionsKeydown(e: KeyboardEvent): boolean {
  if (e.key === "Escape") {
    e.preventDefault();
    closeRuntimeOptionsModal();
    return true;
  }

  if (
    e.key === "ArrowDown" ||
    e.key === "j" ||
    e.key === "J" ||
    (e.ctrlKey && (e.key === "n" || e.key === "N"))
  ) {
    e.preventDefault();
    if (runtimeOptions.length > 0) {
      runtimeOptionSelectedIndex = Math.min(
        runtimeOptions.length - 1,
        runtimeOptionSelectedIndex + 1,
      );
      renderRuntimeOptionsList();
    }
    return true;
  }

  if (
    e.key === "ArrowUp" ||
    e.key === "k" ||
    e.key === "K" ||
    (e.ctrlKey && (e.key === "p" || e.key === "P"))
  ) {
    e.preventDefault();
    if (runtimeOptions.length > 0) {
      runtimeOptionSelectedIndex = Math.max(0, runtimeOptionSelectedIndex - 1);
      renderRuntimeOptionsList();
    }
    return true;
  }

  if (e.key === "ArrowRight" || e.key === "l" || e.key === "L") {
    e.preventDefault();
    cycleRuntimeDraftValue(1);
    return true;
  }

  if (e.key === "ArrowLeft" || e.key === "h" || e.key === "H") {
    e.preventDefault();
    cycleRuntimeDraftValue(-1);
    return true;
  }

  if (e.key === "Enter") {
    e.preventDefault();
    void applySelectedRuntimeOption();
    return true;
  }

  return true;
}

function setSubsyncStatus(message: string, isError = false): void {
  subsyncStatus.textContent = message;
  subsyncStatus.classList.toggle("error", isError);
}

function updateSubsyncSourceVisibility(): void {
  const useAlass = subsyncEngineAlass.checked;
  subsyncSourceLabel.classList.toggle("hidden", !useAlass);
}

function renderSubsyncSourceTracks(): void {
  subsyncSourceSelect.innerHTML = "";
  for (const track of subsyncSourceTracks) {
    const option = document.createElement("option");
    option.value = String(track.id);
    option.textContent = track.label;
    subsyncSourceSelect.appendChild(option);
  }
  subsyncSourceSelect.disabled = subsyncSourceTracks.length === 0;
}

function closeSubsyncModal(): void {
  if (!subsyncModalOpen) return;
  subsyncModalOpen = false;
  syncSettingsModalSubtitleSuppression();
  subsyncModal.classList.add("hidden");
  subsyncModal.setAttribute("aria-hidden", "true");
  window.electronAPI.notifyOverlayModalClosed("subsync");
  if (
    !isOverSubtitle &&
    !jimakuModalOpen &&
    !kikuModalOpen &&
    !runtimeOptionsModalOpen
  ) {
    overlay.classList.remove("interactive");
  }
}

function openSubsyncModal(payload: SubsyncManualPayload): void {
  if (isInvisibleLayer) return;
  subsyncSubmitting = false;
  subsyncRunButton.disabled = false;
  subsyncSourceTracks = payload.sourceTracks;
  const hasSources = subsyncSourceTracks.length > 0;
  subsyncEngineAlass.checked = hasSources;
  subsyncEngineFfsubsync.checked = !hasSources;
  renderSubsyncSourceTracks();
  updateSubsyncSourceVisibility();
  setSubsyncStatus(
    hasSources
      ? "Choose engine and source, then run."
      : "No source subtitles available for alass. Use ffsubsync.",
    false,
  );
  subsyncModalOpen = true;
  syncSettingsModalSubtitleSuppression();
  overlay.classList.add("interactive");
  subsyncModal.classList.remove("hidden");
  subsyncModal.setAttribute("aria-hidden", "false");
}

async function runSubsyncManualFromModal(): Promise<void> {
  if (subsyncSubmitting) return;
  const engine = subsyncEngineAlass.checked ? "alass" : "ffsubsync";
  const sourceTrackId =
    engine === "alass" && subsyncSourceSelect.value
      ? Number.parseInt(subsyncSourceSelect.value, 10)
      : null;

  if (engine === "alass" && !Number.isFinite(sourceTrackId)) {
    setSubsyncStatus("Select a source subtitle track for alass.", true);
    return;
  }

  subsyncSubmitting = true;
  subsyncRunButton.disabled = true;
  closeSubsyncModal();
  try {
    await window.electronAPI.runSubsyncManual({
      engine,
      sourceTrackId,
    });
  } finally {
    subsyncSubmitting = false;
    subsyncRunButton.disabled = false;
  }
}

function handleSubsyncKeydown(e: KeyboardEvent): boolean {
  if (e.key === "Escape") {
    e.preventDefault();
    closeSubsyncModal();
    return true;
  }
  if (e.key === "Enter") {
    e.preventDefault();
    void runSubsyncManualFromModal();
    return true;
  }
  return true;
}

function formatMediaMeta(card: KikuDuplicateCardInfo): string {
  const parts: string[] = [];
  parts.push(card.hasAudio ? "Audio: Yes" : "Audio: No");
  parts.push(card.hasImage ? "Image: Yes" : "Image: No");
  return parts.join(" | ");
}

function updateKikuCardSelection(): void {
  kikuCard1.classList.toggle("active", kikuSelectedCard === 1);
  kikuCard2.classList.toggle("active", kikuSelectedCard === 2);
}

function setKikuModalStep(step: KikuModalStep): void {
  kikuModalStep = step;
  const isSelect = step === "select";
  kikuSelectionStep.classList.toggle("hidden", !isSelect);
  kikuPreviewStep.classList.toggle("hidden", isSelect);
  kikuHint.textContent = isSelect
    ? "Press 1 or 2 to select · Enter to continue · Esc to cancel"
    : "Enter to confirm merge · Backspace to go back · Esc to cancel";
}

function updateKikuPreviewToggle(): void {
  kikuPreviewCompactButton.classList.toggle(
    "active",
    kikuPreviewMode === "compact",
  );
  kikuPreviewFullButton.classList.toggle("active", kikuPreviewMode === "full");
}

function renderKikuPreview(): void {
  const payload =
    kikuPreviewMode === "compact"
      ? kikuPreviewCompactData
      : kikuPreviewFullData;
  kikuPreviewJson.textContent = payload
    ? JSON.stringify(payload, null, 2)
    : "{}";
  updateKikuPreviewToggle();
}

function setKikuPreviewError(message: string | null): void {
  if (!message) {
    kikuPreviewError.textContent = "";
    kikuPreviewError.classList.add("hidden");
    return;
  }
  kikuPreviewError.textContent = message;
  kikuPreviewError.classList.remove("hidden");
}

function openKikuFieldGroupingModal(data: {
  original: KikuDuplicateCardInfo;
  duplicate: KikuDuplicateCardInfo;
}): void {
  if (isInvisibleLayer) return;
  if (kikuModalOpen) return;
  kikuModalOpen = true;
  kikuOriginalData = data.original;
  kikuDuplicateData = data.duplicate;
  kikuSelectedCard = 1;

  kikuCard1Expression.textContent = data.original.expression;
  kikuCard1Sentence.textContent =
    data.original.sentencePreview || "(no sentence)";
  kikuCard1Meta.textContent = formatMediaMeta(data.original);

  kikuCard2Expression.textContent = data.duplicate.expression;
  kikuCard2Sentence.textContent =
    data.duplicate.sentencePreview || "(current subtitle)";
  kikuCard2Meta.textContent = formatMediaMeta(data.duplicate);
  kikuDeleteDuplicateCheckbox.checked = true;
  kikuPendingChoice = null;
  kikuPreviewCompactData = null;
  kikuPreviewFullData = null;
  kikuPreviewMode = "compact";
  renderKikuPreview();
  setKikuPreviewError(null);
  setKikuModalStep("select");

  updateKikuCardSelection();

  syncSettingsModalSubtitleSuppression();
  overlay.classList.add("interactive");
  kikuModal.classList.remove("hidden");
  kikuModal.setAttribute("aria-hidden", "false");
}

function closeKikuFieldGroupingModal(): void {
  if (!kikuModalOpen) return;
  kikuModalOpen = false;
  syncSettingsModalSubtitleSuppression();
  kikuModal.classList.add("hidden");
  kikuModal.setAttribute("aria-hidden", "true");
  setKikuPreviewError(null);
  kikuPreviewJson.textContent = "";
  kikuPendingChoice = null;
  kikuPreviewCompactData = null;
  kikuPreviewFullData = null;
  kikuPreviewMode = "compact";
  setKikuModalStep("select");
  kikuOriginalData = null;
  kikuDuplicateData = null;
  if (
    !isOverSubtitle &&
    !jimakuModalOpen &&
    !runtimeOptionsModalOpen &&
    !subsyncModalOpen
  ) {
    overlay.classList.remove("interactive");
  }
}

async function confirmKikuSelection(): Promise<void> {
  if (!kikuOriginalData || !kikuDuplicateData) return;

  const keepData =
    kikuSelectedCard === 1 ? kikuOriginalData : kikuDuplicateData;
  const deleteData =
    kikuSelectedCard === 1 ? kikuDuplicateData : kikuOriginalData;

  const choice: KikuFieldGroupingChoice = {
    keepNoteId: keepData.noteId,
    deleteNoteId: deleteData.noteId,
    deleteDuplicate: kikuDeleteDuplicateCheckbox.checked,
    cancelled: false,
  };
  kikuPendingChoice = choice;
  setKikuPreviewError(null);
  kikuConfirmButton.disabled = true;

  try {
    const preview: KikuMergePreviewResponse =
      await window.electronAPI.kikuBuildMergePreview({
        keepNoteId: choice.keepNoteId,
        deleteNoteId: choice.deleteNoteId,
        deleteDuplicate: choice.deleteDuplicate,
      });

    if (!preview.ok) {
      setKikuPreviewError(preview.error || "Failed to build merge preview");
      return;
    }

    kikuPreviewCompactData = preview.compact || {};
    kikuPreviewFullData = preview.full || {};
    kikuPreviewMode = "compact";
    renderKikuPreview();
    setKikuModalStep("preview");
  } finally {
    kikuConfirmButton.disabled = false;
  }
}

function confirmKikuMerge(): void {
  if (!kikuPendingChoice) return;
  window.electronAPI.kikuFieldGroupingRespond(kikuPendingChoice);
  closeKikuFieldGroupingModal();
}

function goBackFromKikuPreview(): void {
  setKikuPreviewError(null);
  setKikuModalStep("select");
}

function cancelKikuFieldGrouping(): void {
  const choice: KikuFieldGroupingChoice = {
    keepNoteId: 0,
    deleteNoteId: 0,
    deleteDuplicate: true,
    cancelled: true,
  };

  window.electronAPI.kikuFieldGroupingRespond(choice);
  closeKikuFieldGroupingModal();
}

function handleKikuKeydown(e: KeyboardEvent): boolean {
  if (kikuModalStep === "preview") {
    if (e.key === "Escape") {
      e.preventDefault();
      cancelKikuFieldGrouping();
      return true;
    }

    if (e.key === "Backspace") {
      e.preventDefault();
      goBackFromKikuPreview();
      return true;
    }

    if (e.key === "Enter") {
      e.preventDefault();
      confirmKikuMerge();
      return true;
    }

    return true;
  }

  if (e.key === "Escape") {
    e.preventDefault();
    cancelKikuFieldGrouping();
    return true;
  }

  if (e.key === "1") {
    e.preventDefault();
    kikuSelectedCard = 1;
    updateKikuCardSelection();
    return true;
  }

  if (e.key === "2") {
    e.preventDefault();
    kikuSelectedCard = 2;
    updateKikuCardSelection();
    return true;
  }

  if (e.key === "ArrowLeft" || e.key === "ArrowRight") {
    e.preventDefault();
    kikuSelectedCard = kikuSelectedCard === 1 ? 2 : 1;
    updateKikuCardSelection();
    return true;
  }

  if (e.key === "Enter") {
    e.preventDefault();
    void confirmKikuSelection();
    return true;
  }

  return true;
}

function formatEntryLabel(entry: JimakuEntry): string {
  if (entry.english_name && entry.english_name !== entry.name) {
    return `${entry.name} / ${entry.english_name}`;
  }
  return entry.name;
}

function renderEntries(): void {
  jimakuEntriesList.innerHTML = "";
  if (jimakuEntries.length === 0) {
    jimakuEntriesSection.classList.add("hidden");
    return;
  }
  jimakuEntriesSection.classList.remove("hidden");
  jimakuEntries.forEach((entry, index) => {
    const li = document.createElement("li");
    li.textContent = formatEntryLabel(entry);
    if (entry.japanese_name) {
      const sub = document.createElement("div");
      sub.className = "jimaku-subtext";
      sub.textContent = entry.japanese_name;
      li.appendChild(sub);
    }
    if (index === selectedEntryIndex) {
      li.classList.add("active");
    }
    li.addEventListener("click", () => {
      selectEntry(index);
    });
    jimakuEntriesList.appendChild(li);
  });
}

function formatBytes(size: number): string {
  if (!Number.isFinite(size)) return "";
  const units = ["B", "KB", "MB", "GB"];
  let value = size;
  let idx = 0;
  while (value >= 1024 && idx < units.length - 1) {
    value /= 1024;
    idx += 1;
  }
  return `${value.toFixed(value >= 10 || idx === 0 ? 0 : 1)} ${units[idx]}`;
}

function renderFiles(): void {
  jimakuFilesList.innerHTML = "";
  if (jimakuFiles.length === 0) {
    jimakuFilesSection.classList.add("hidden");
    return;
  }
  jimakuFilesSection.classList.remove("hidden");
  jimakuFiles.forEach((file, index) => {
    const li = document.createElement("li");
    li.textContent = file.name;
    const sub = document.createElement("div");
    sub.className = "jimaku-subtext";
    sub.textContent = `${formatBytes(file.size)} • ${file.last_modified}`;
    li.appendChild(sub);
    if (index === selectedFileIndex) {
      li.classList.add("active");
    }
    li.addEventListener("click", () => {
      selectFile(index);
    });
    jimakuFilesList.appendChild(li);
  });
}

function getSearchQuery(): { query: string; episode: number | null } {
  const title = jimakuTitleInput.value.trim();
  const episode = jimakuEpisodeInput.value
    ? Number.parseInt(jimakuEpisodeInput.value, 10)
    : null;
  const query = title;
  return { query, episode: Number.isFinite(episode) ? episode : null };
}

async function performJimakuSearch(): Promise<void> {
  const { query, episode } = getSearchQuery();
  if (!query) {
    setJimakuStatus("Enter a title before searching.", true);
    return;
  }
  resetJimakuLists();
  setJimakuStatus("Searching Jimaku...");
  currentEpisodeFilter = episode;

  const response: JimakuApiResponse<JimakuEntry[]> =
    await window.electronAPI.jimakuSearchEntries({ query });
  if (!response.ok) {
    const retry = response.error.retryAfter
      ? ` Retry after ${response.error.retryAfter.toFixed(1)}s.`
      : "";
    setJimakuStatus(`${response.error.error}${retry}`, true);
    return;
  }

  jimakuEntries = response.data;
  selectedEntryIndex = 0;
  if (jimakuEntries.length === 0) {
    setJimakuStatus("No entries found.");
    return;
  }
  setJimakuStatus("Select an entry.");
  renderEntries();
  if (jimakuEntries.length === 1) {
    selectEntry(0);
  }
}

async function loadFiles(
  entryId: number,
  episode: number | null,
): Promise<void> {
  setJimakuStatus("Loading files...");
  jimakuFiles = [];
  selectedFileIndex = 0;
  jimakuFilesList.innerHTML = "";
  jimakuFilesSection.classList.add("hidden");

  const response: JimakuApiResponse<JimakuFileEntry[]> =
    await window.electronAPI.jimakuListFiles({
      entryId,
      episode,
    });
  if (!response.ok) {
    const retry = response.error.retryAfter
      ? ` Retry after ${response.error.retryAfter.toFixed(1)}s.`
      : "";
    setJimakuStatus(`${response.error.error}${retry}`, true);
    return;
  }

  jimakuFiles = response.data;
  if (jimakuFiles.length === 0) {
    if (episode !== null) {
      setJimakuStatus("No files found for this episode.");
      jimakuBroadenButton.classList.remove("hidden");
    } else {
      setJimakuStatus("No files found.");
    }
    return;
  }

  jimakuBroadenButton.classList.add("hidden");
  setJimakuStatus("Select a subtitle file.");
  renderFiles();
  if (jimakuFiles.length === 1) {
    selectFile(0);
  }
}

function selectEntry(index: number): void {
  if (index < 0 || index >= jimakuEntries.length) return;
  selectedEntryIndex = index;
  currentEntryId = jimakuEntries[index].id;
  renderEntries();
  if (currentEntryId !== null) {
    loadFiles(currentEntryId, currentEpisodeFilter);
  }
}

async function selectFile(index: number): Promise<void> {
  if (index < 0 || index >= jimakuFiles.length) return;
  selectedFileIndex = index;
  renderFiles();
  if (currentEntryId === null) {
    setJimakuStatus("Select an entry first.", true);
    return;
  }

  const file = jimakuFiles[index];
  setJimakuStatus("Downloading subtitle...");
  const result: JimakuDownloadResult =
    await window.electronAPI.jimakuDownloadFile({
      entryId: currentEntryId,
      url: file.url,
      name: file.name,
    });

  if (result.ok) {
    setJimakuStatus(`Downloaded and loaded: ${result.path}`);
  } else {
    const retry = result.error.retryAfter
      ? ` Retry after ${result.error.retryAfter.toFixed(1)}s.`
      : "";
    setJimakuStatus(`${result.error.error}${retry}`, true);
  }
}

function isTextInputFocused(): boolean {
  const active = document.activeElement;
  if (!active) return false;
  const tag = active.tagName.toLowerCase();
  return tag === "input" || tag === "textarea";
}

function handleJimakuKeydown(e: KeyboardEvent): boolean {
  if (e.key === "Escape") {
    e.preventDefault();
    closeJimakuModal();
    return true;
  }

  if (isTextInputFocused()) {
    if (e.key === "Enter") {
      e.preventDefault();
      performJimakuSearch();
      return true;
    }
    return true;
  }

  if (e.key === "ArrowDown") {
    e.preventDefault();
    if (jimakuFiles.length > 0) {
      selectedFileIndex = Math.min(
        jimakuFiles.length - 1,
        selectedFileIndex + 1,
      );
      renderFiles();
    } else if (jimakuEntries.length > 0) {
      selectedEntryIndex = Math.min(
        jimakuEntries.length - 1,
        selectedEntryIndex + 1,
      );
      renderEntries();
    }
    return true;
  }

  if (e.key === "ArrowUp") {
    e.preventDefault();
    if (jimakuFiles.length > 0) {
      selectedFileIndex = Math.max(0, selectedFileIndex - 1);
      renderFiles();
    } else if (jimakuEntries.length > 0) {
      selectedEntryIndex = Math.max(0, selectedEntryIndex - 1);
      renderEntries();
    }
    return true;
  }

  if (e.key === "Enter") {
    e.preventDefault();
    if (jimakuFiles.length > 0) {
      selectFile(selectedFileIndex);
    } else if (jimakuEntries.length > 0) {
      selectEntry(selectedEntryIndex);
    } else {
      performJimakuSearch();
    }
    return true;
  }

  return true;
}

function setupDragging(): void {
  subtitleContainer.addEventListener("mousedown", (e: MouseEvent) => {
    if (e.button === 2) {
      e.preventDefault();
      isDragging = true;
      dragStartY = e.clientY;
      startYPercent = getCurrentYPercent();
      subtitleContainer.style.cursor = "grabbing";
    }
  });

  document.addEventListener("mousemove", (e: MouseEvent) => {
    if (!isDragging) return;

    const deltaY = dragStartY - e.clientY;
    const deltaPercent = (deltaY / window.innerHeight) * 100;
    const newYPercent = startYPercent + deltaPercent;

    applyYPercent(newYPercent);
  });

  document.addEventListener("mouseup", (e: MouseEvent) => {
    if (isDragging && e.button === 2) {
      isDragging = false;
      subtitleContainer.style.cursor = "";

      const yPercent = getCurrentYPercent();
      persistSubtitlePositionPatch({ yPercent });
    }
  });

  subtitleContainer.addEventListener("contextmenu", (e: Event) => {
    e.preventDefault();
  });
}

function getCaretTextPointRange(clientX: number, clientY: number): Range | null {
  const documentWithCaretApi = document as Document & {
    caretRangeFromPoint?: (x: number, y: number) => Range | null;
    caretPositionFromPoint?: (
      x: number,
      y: number,
    ) => { offsetNode: Node; offset: number } | null;
  };

  if (typeof documentWithCaretApi.caretRangeFromPoint === "function") {
    return documentWithCaretApi.caretRangeFromPoint(clientX, clientY);
  }

  if (typeof documentWithCaretApi.caretPositionFromPoint === "function") {
    const caretPosition = documentWithCaretApi.caretPositionFromPoint(
      clientX,
      clientY,
    );
    if (!caretPosition) return null;
    const range = document.createRange();
    range.setStart(caretPosition.offsetNode, caretPosition.offset);
    range.collapse(true);
    return range;
  }

  return null;
}

function getWordBoundsAtOffset(
  text: string,
  offset: number,
): { start: number; end: number } | null {
  if (!text || text.length === 0) return null;

  const clampedOffset = Math.max(0, Math.min(offset, text.length));
  const probeIndex =
    clampedOffset >= text.length ? Math.max(0, text.length - 1) : clampedOffset;

  if (wordSegmenter) {
    for (const part of wordSegmenter.segment(text)) {
      const start = part.index;
      const end = start + part.segment.length;
      if (probeIndex >= start && probeIndex < end) {
        if (part.isWordLike === false) return null;
        return { start, end };
      }
    }
  }

  const isBoundary = (char: string): boolean =>
    /[\s\u3000.,!?;:()[\]{}"'`~<>/\\|@#$%^&*+=\-、。・「」『』【】〈〉《》]/.test(
      char,
    );

  const probeChar = text[probeIndex];
  if (!probeChar || isBoundary(probeChar)) return null;

  let start = probeIndex;
  while (start > 0 && !isBoundary(text[start - 1])) {
    start -= 1;
  }

  let end = probeIndex + 1;
  while (end < text.length && !isBoundary(text[end])) {
    end += 1;
  }

  if (end <= start) return null;
  return { start, end };
}

function updateHoverWordSelection(event: MouseEvent): void {
  if (!isInvisibleLayer || !isMacOSPlatform) return;
  if (event.buttons !== 0) return;
  if (!(event.target instanceof Node)) return;
  if (!subtitleRoot.contains(event.target)) return;

  const caretRange = getCaretTextPointRange(event.clientX, event.clientY);
  if (!caretRange) return;
  if (caretRange.startContainer.nodeType !== Node.TEXT_NODE) return;
  if (!subtitleRoot.contains(caretRange.startContainer)) return;

  const textNode = caretRange.startContainer as Text;
  const wordBounds = getWordBoundsAtOffset(textNode.data, caretRange.startOffset);
  if (!wordBounds) return;

  const selectionKey = `${wordBounds.start}:${wordBounds.end}:${textNode.data.slice(
    wordBounds.start,
    wordBounds.end,
  )}`;
  if (
    selectionKey === lastHoverSelectionKey &&
    textNode === lastHoverSelectionNode
  ) {
    return;
  }

  const selection = window.getSelection();
  if (!selection) return;

  const range = document.createRange();
  range.setStart(textNode, wordBounds.start);
  range.setEnd(textNode, wordBounds.end);

  selection.removeAllRanges();
  selection.addRange(range);
  lastHoverSelectionKey = selectionKey;
  lastHoverSelectionNode = textNode;
}

function setupInvisibleHoverSelection(): void {
  if (!isInvisibleLayer || !isMacOSPlatform) return;

  subtitleRoot.addEventListener("mousemove", (event: MouseEvent) => {
    updateHoverWordSelection(event);
  });

  subtitleRoot.addEventListener("mouseleave", () => {
    lastHoverSelectionKey = "";
    lastHoverSelectionNode = null;
  });
}

function isInteractiveTarget(target: EventTarget | null): boolean {
  if (!(target instanceof Element)) return false;
  if (target.closest(".modal")) return true;
  if (subtitleContainer.contains(target)) return true;
  if (
    target.tagName === "IFRAME" &&
    target.id &&
    target.id.startsWith("yomitan-popup")
  )
    return true;
  if (target.closest && target.closest('iframe[id^="yomitan-popup"]'))
    return true;
  return false;
}

function keyEventToString(e: KeyboardEvent): string {
  const parts: string[] = [];
  if (e.ctrlKey) parts.push("Ctrl");
  if (e.altKey) parts.push("Alt");
  if (e.shiftKey) parts.push("Shift");
  if (e.metaKey) parts.push("Meta");
  parts.push(e.code);
  return parts.join("+");
}

function isInvisiblePositionToggleShortcut(e: KeyboardEvent): boolean {
  return (
    e.code === INVISIBLE_POSITION_EDIT_TOGGLE_CODE &&
    !e.altKey &&
    e.shiftKey &&
    (e.ctrlKey || e.metaKey)
  );
}

function saveInvisiblePositionEdit(): void {
  persistSubtitlePositionPatch({
    invisibleOffsetXPx: invisibleSubtitleOffsetXPx,
    invisibleOffsetYPx: invisibleSubtitleOffsetYPx,
  });
  setInvisiblePositionEditMode(false);
}

function cancelInvisiblePositionEdit(): void {
  invisibleSubtitleOffsetXPx = invisiblePositionEditStartX;
  invisibleSubtitleOffsetYPx = invisiblePositionEditStartY;
  applyInvisibleSubtitleOffsetPosition();
  setInvisiblePositionEditMode(false);
}

function handleInvisiblePositionEditKeydown(e: KeyboardEvent): boolean {
  if (!isInvisibleLayer) return false;

  if (isInvisiblePositionToggleShortcut(e)) {
    e.preventDefault();
    if (invisiblePositionEditMode) {
      cancelInvisiblePositionEdit();
    } else {
      setInvisiblePositionEditMode(true);
    }
    return true;
  }

  if (!invisiblePositionEditMode) return false;

  const step = e.shiftKey ? INVISIBLE_POSITION_STEP_FAST_PX : INVISIBLE_POSITION_STEP_PX;

  if (e.key === "Escape") {
    e.preventDefault();
    cancelInvisiblePositionEdit();
    return true;
  }

  if (e.key === "Enter" || ((e.ctrlKey || e.metaKey) && e.code === "KeyS")) {
    e.preventDefault();
    saveInvisiblePositionEdit();
    return true;
  }

  if (
    e.key === "ArrowUp" ||
    e.key === "ArrowDown" ||
    e.key === "ArrowLeft" ||
    e.key === "ArrowRight" ||
    e.key === "h" ||
    e.key === "j" ||
    e.key === "k" ||
    e.key === "l" ||
    e.key === "H" ||
    e.key === "J" ||
    e.key === "K" ||
    e.key === "L"
  ) {
    e.preventDefault();
    if (e.key === "ArrowUp" || e.key === "k" || e.key === "K") {
      invisibleSubtitleOffsetYPx += step;
    } else if (e.key === "ArrowDown" || e.key === "j" || e.key === "J") {
      invisibleSubtitleOffsetYPx -= step;
    } else if (e.key === "ArrowLeft" || e.key === "h" || e.key === "H") {
      invisibleSubtitleOffsetXPx -= step;
    } else if (e.key === "ArrowRight" || e.key === "l" || e.key === "L") {
      invisibleSubtitleOffsetXPx += step;
    }
    applyInvisibleSubtitleOffsetPosition();
    updateInvisiblePositionEditHud();
    return true;
  }

  return true;
}

let keybindingsMap = new Map<string, (string | number)[]>();

type ChordAction =
  | { type: "mpv"; command: string[] }
  | { type: "electron"; action: () => void }
  | { type: "noop" };

const CHORD_MAP = new Map<string, ChordAction>([
  ["KeyS", { type: "mpv", command: ["script-message", "subminer-start"] }],
  ["Shift+KeyS", { type: "mpv", command: ["script-message", "subminer-stop"] }],
  ["KeyT", { type: "mpv", command: ["script-message", "subminer-toggle"] }],
  [
    "KeyI",
    { type: "mpv", command: ["script-message", "subminer-toggle-invisible"] },
  ],
  [
    "Shift+KeyI",
    { type: "mpv", command: ["script-message", "subminer-show-invisible"] },
  ],
  [
    "KeyU",
    { type: "mpv", command: ["script-message", "subminer-hide-invisible"] },
  ],
  ["KeyO", { type: "mpv", command: ["script-message", "subminer-options"] }],
  ["KeyR", { type: "mpv", command: ["script-message", "subminer-restart"] }],
  ["KeyC", { type: "mpv", command: ["script-message", "subminer-status"] }],
  ["KeyY", { type: "mpv", command: ["script-message", "subminer-menu"] }],
  [
    "KeyD",
    { type: "electron", action: () => window.electronAPI.toggleDevTools() },
  ],
]);

let chordPending = false;
let chordTimeout: ReturnType<typeof setTimeout> | null = null;

function resetChord(): void {
  chordPending = false;
  if (chordTimeout !== null) {
    clearTimeout(chordTimeout);
    chordTimeout = null;
  }
}

async function setupMpvInputForwarding(): Promise<void> {
  const keybindings: Keybinding[] = await window.electronAPI.getKeybindings();
  keybindingsMap = new Map();
  for (const binding of keybindings) {
    if (binding.command) {
      keybindingsMap.set(binding.key, binding.command);
    }
  }

  document.addEventListener("keydown", (e: KeyboardEvent) => {
    const yomitanPopup = document.querySelector('iframe[id^="yomitan-popup"]');
    if (yomitanPopup) return;
    if (handleInvisiblePositionEditKeydown(e)) return;

    if (runtimeOptionsModalOpen) {
      handleRuntimeOptionsKeydown(e);
      return;
    }

    if (subsyncModalOpen) {
      handleSubsyncKeydown(e);
      return;
    }

    if (kikuModalOpen) {
      handleKikuKeydown(e);
      return;
    }

    if (jimakuModalOpen) {
      handleJimakuKeydown(e);
      return;
    }

    if (chordPending) {
      const modifierKeys = [
        "ShiftLeft",
        "ShiftRight",
        "ControlLeft",
        "ControlRight",
        "AltLeft",
        "AltRight",
        "MetaLeft",
        "MetaRight",
      ];
      if (modifierKeys.includes(e.code)) {
        return;
      }

      e.preventDefault();
      const secondKey = keyEventToString(e);
      const action = CHORD_MAP.get(secondKey);
      resetChord();
      if (action) {
        if (action.type === "mpv") {
          window.electronAPI.sendMpvCommand(action.command);
        } else if (action.type === "electron") {
          action.action();
        }
      }
      return;
    }

    if (
      e.code === "KeyY" &&
      !e.ctrlKey &&
      !e.altKey &&
      !e.shiftKey &&
      !e.metaKey &&
      !e.repeat
    ) {
      e.preventDefault();
      chordPending = true;
      chordTimeout = setTimeout(() => {
        resetChord();
      }, 1000);
      return;
    }

    const keyString = keyEventToString(e);
    const command = keybindingsMap.get(keyString);

    if (command) {
      e.preventDefault();
      window.electronAPI.sendMpvCommand(command);
    }
  });

  document.addEventListener("mousedown", (e: MouseEvent) => {
    if (e.button === 2 && !isInteractiveTarget(e.target)) {
      e.preventDefault();
      window.electronAPI.sendMpvCommand(["cycle", "pause"]);
    }
  });

  document.addEventListener("contextmenu", (e: Event) => {
    if (!isInteractiveTarget(e.target)) {
      e.preventDefault();
    }
  });
}

function setupResizeHandler(): void {
  window.addEventListener("resize", () => {
    if (isInvisibleLayer) {
      applyInvisibleSubtitleLayoutFromMpvMetrics(
        mpvSubtitleRenderMetrics,
        "resize",
      );
      return;
    }
    applyYPercent(getCurrentYPercent());
  });
}

async function restoreSubtitlePosition(): Promise<void> {
  const position = await window.electronAPI.getSubtitlePosition();
  applyStoredSubtitlePosition(position, "startup");
}

async function restoreSubtitleFontSize(): Promise<void> {
  const style = await window.electronAPI.getSubtitleStyle();
  if (style && style.fontSize !== undefined) {
    applySubtitleFontSize(style.fontSize);
    console.log("Applied subtitle font size:", style.fontSize);
  }
}

function setupSelectionObserver(): void {
  document.addEventListener("selectionchange", () => {
    const selection = window.getSelection();
    const hasSelection =
      selection && selection.rangeCount > 0 && !selection.isCollapsed;

    if (hasSelection) {
      subtitleRoot.classList.add("has-selection");
    } else {
      subtitleRoot.classList.remove("has-selection");
    }
  });
}

function setupInvisiblePositionEditHud(): void {
  if (!isInvisibleLayer) return;
  const hud = document.createElement("div");
  hud.id = "invisiblePositionEditHud";
  hud.className = "invisible-position-edit-hud";
  overlay.appendChild(hud);
  invisiblePositionEditHud = hud;
  updateInvisiblePositionEditHud();
}

function setupYomitanObserver(): void {
  const observer = new MutationObserver((mutations: MutationRecord[]) => {
    for (const mutation of mutations) {
      mutation.addedNodes.forEach((node) => {
        if (node.nodeType === Node.ELEMENT_NODE) {
          const element = node as Element;
          if (
            element.tagName === "IFRAME" &&
            element.id &&
            element.id.startsWith("yomitan-popup")
          ) {
            overlay.classList.add("interactive");
            if (shouldToggleMouseIgnore) {
              window.electronAPI.setIgnoreMouseEvents(false);
            }
          }
        }
      });
      mutation.removedNodes.forEach((node) => {
        if (node.nodeType === Node.ELEMENT_NODE) {
          const element = node as Element;
          if (
            element.tagName === "IFRAME" &&
            element.id &&
            element.id.startsWith("yomitan-popup")
          ) {
            if (
              !isOverSubtitle &&
              !jimakuModalOpen &&
              !kikuModalOpen &&
              !runtimeOptionsModalOpen &&
              !subsyncModalOpen
            ) {
              overlay.classList.remove("interactive");
              if (shouldToggleMouseIgnore) {
                window.electronAPI.setIgnoreMouseEvents(true, {
                  forward: true,
                });
              }
            }
          }
        }
      });
    }
  });

  observer.observe(document.body, {
    childList: true,
    subtree: true,
  });
}

function renderSecondarySub(text: string): void {
  secondarySubRoot.innerHTML = "";
  if (!text) return;

  let normalized = text
    .replace(/\\N/g, "\n")
    .replace(/\\n/g, "\n")
    .replace(/\{[^}]*\}/g, "")
    .trim();

  if (!normalized) return;

  const lines = normalized.split("\n");
  for (let i = 0; i < lines.length; i++) {
    if (lines[i]) {
      const textNode = document.createTextNode(lines[i]);
      secondarySubRoot.appendChild(textNode);
    }
    if (i < lines.length - 1) {
      secondarySubRoot.appendChild(document.createElement("br"));
    }
  }
}

function updateSecondarySubMode(mode: SecondarySubMode): void {
  secondarySubContainer.classList.remove(
    "secondary-sub-hidden",
    "secondary-sub-visible",
    "secondary-sub-hover",
  );
  secondarySubContainer.classList.add(`secondary-sub-${mode}`);
}

async function applySubtitleStyle(): Promise<void> {
  const style = await window.electronAPI.getSubtitleStyle();
  if (!style) return;

  if (style.fontFamily) {
    subtitleRoot.style.fontFamily = style.fontFamily;
  }
  if (style.fontSize) {
    subtitleRoot.style.fontSize = `${style.fontSize}px`;
  }
  if (style.fontColor) {
    subtitleRoot.style.color = style.fontColor;
  }
  if (style.fontWeight) {
    subtitleRoot.style.fontWeight = style.fontWeight;
  }
  if (style.fontStyle) {
    subtitleRoot.style.fontStyle = style.fontStyle;
  }
  if (style.backgroundColor) {
    subtitleContainer.style.background = style.backgroundColor;
  }

  const sec = style.secondary;
  if (sec) {
    if (sec.fontFamily) {
      secondarySubRoot.style.fontFamily = sec.fontFamily;
    }
    if (sec.fontSize) {
      secondarySubRoot.style.fontSize = `${sec.fontSize}px`;
    }
    if (sec.fontColor) {
      secondarySubRoot.style.color = sec.fontColor;
    }
    if (sec.fontWeight) {
      secondarySubRoot.style.fontWeight = sec.fontWeight;
    }
    if (sec.fontStyle) {
      secondarySubRoot.style.fontStyle = sec.fontStyle;
    }
    if (sec.backgroundColor) {
      secondarySubContainer.style.background = sec.backgroundColor;
    }
  }
}

async function init(): Promise<void> {
  document.body.classList.add(`layer-${overlayLayer}`);

  window.electronAPI.onSubtitle((data: SubtitleData) => {
    renderSubtitle(data);
  });

  window.electronAPI.onSubtitlePosition(
    (position: SubtitlePosition | null) => {
      if (isInvisibleLayer) {
        applyInvisibleStoredSubtitlePosition(position, "media-change");
      } else {
        applyStoredSubtitlePosition(position, "media-change");
      }
    },
  );

  if (isInvisibleLayer) {
    window.electronAPI.onMpvSubtitleRenderMetrics(
      (metrics: MpvSubtitleRenderMetrics) => {
        applyInvisibleSubtitleLayoutFromMpvMetrics(metrics, "event");
      },
    );
    window.electronAPI.onOverlayDebugVisualization((enabled: boolean) => {
      document.body.classList.toggle("debug-invisible-visualization", enabled);
    });
  }

  const initialSubtitle = await window.electronAPI.getCurrentSubtitle();
  renderSubtitle(initialSubtitle);

  window.electronAPI.onSecondarySub((text: string) => {
    renderSecondarySub(text);
  });

  window.electronAPI.onSecondarySubMode((mode: SecondarySubMode) => {
    updateSecondarySubMode(mode);
  });

  const initialMode = await window.electronAPI.getSecondarySubMode();
  updateSecondarySubMode(initialMode);

  const initialSecondary = await window.electronAPI.getCurrentSecondarySub();
  renderSecondarySub(initialSecondary);

  const hoverTarget = isInvisibleLayer ? subtitleRoot : subtitleContainer;
  hoverTarget.addEventListener("mouseenter", handleMouseEnter);
  hoverTarget.addEventListener("mouseleave", handleMouseLeave);
  setupInvisibleHoverSelection();
  setupInvisiblePositionEditHud();

  secondarySubContainer.addEventListener("mouseenter", handleMouseEnter);
  secondarySubContainer.addEventListener("mouseleave", handleMouseLeave);

  jimakuSearchButton.addEventListener("click", () => {
    performJimakuSearch();
  });

  jimakuCloseButton.addEventListener("click", () => {
    closeJimakuModal();
  });

  jimakuBroadenButton.addEventListener("click", () => {
    if (currentEntryId !== null) {
      jimakuBroadenButton.classList.add("hidden");
      loadFiles(currentEntryId, null);
    }
  });

  kikuCard1.addEventListener("click", () => {
    kikuSelectedCard = 1;
    updateKikuCardSelection();
  });

  kikuCard1.addEventListener("dblclick", () => {
    kikuSelectedCard = 1;
    void confirmKikuSelection();
  });

  kikuCard2.addEventListener("click", () => {
    kikuSelectedCard = 2;
    updateKikuCardSelection();
  });

  kikuCard2.addEventListener("dblclick", () => {
    kikuSelectedCard = 2;
    void confirmKikuSelection();
  });

  kikuConfirmButton.addEventListener("click", () => {
    void confirmKikuSelection();
  });

  kikuCancelButton.addEventListener("click", () => {
    cancelKikuFieldGrouping();
  });

  kikuBackButton.addEventListener("click", () => {
    goBackFromKikuPreview();
  });

  kikuFinalConfirmButton.addEventListener("click", () => {
    confirmKikuMerge();
  });

  kikuFinalCancelButton.addEventListener("click", () => {
    cancelKikuFieldGrouping();
  });

  kikuPreviewCompactButton.addEventListener("click", () => {
    kikuPreviewMode = "compact";
    renderKikuPreview();
  });

  kikuPreviewFullButton.addEventListener("click", () => {
    kikuPreviewMode = "full";
    renderKikuPreview();
  });

  runtimeOptionsClose.addEventListener("click", () => {
    closeRuntimeOptionsModal();
  });
  subsyncCloseButton.addEventListener("click", () => {
    closeSubsyncModal();
  });
  subsyncEngineAlass.addEventListener("change", () => {
    updateSubsyncSourceVisibility();
  });
  subsyncEngineFfsubsync.addEventListener("change", () => {
    updateSubsyncSourceVisibility();
  });
  subsyncRunButton.addEventListener("click", () => {
    void runSubsyncManualFromModal();
  });

  window.electronAPI.onRuntimeOptionsChanged(
    (options: RuntimeOptionState[]) => {
      updateRuntimeOptions(options);
    },
  );

  window.electronAPI.onOpenRuntimeOptions(() => {
    openRuntimeOptionsModal().catch(() => {
      setRuntimeOptionsStatus("Failed to load runtime options", true);
      window.electronAPI.notifyOverlayModalClosed("runtime-options");
      syncSettingsModalSubtitleSuppression();
    });
  });
  window.electronAPI.onOpenJimaku(() => {
    openJimakuModal();
  });
  window.electronAPI.onSubsyncManualOpen((payload: SubsyncManualPayload) => {
    openSubsyncModal(payload);
  });

  window.electronAPI.onKikuFieldGroupingRequest(
    (data: {
      original: KikuDuplicateCardInfo;
      duplicate: KikuDuplicateCardInfo;
    }) => {
      openKikuFieldGroupingModal(data);
    },
  );

  if (!isInvisibleLayer) {
    setupDragging();
  }

  await setupMpvInputForwarding();

  setupResizeHandler();

  if (isInvisibleLayer) {
    const position = await window.electronAPI.getSubtitlePosition();
    applyInvisibleStoredSubtitlePosition(position, "startup");
    const metrics = await window.electronAPI.getMpvSubtitleRenderMetrics();
    applyInvisibleSubtitleLayoutFromMpvMetrics(metrics, "startup");
  } else {
    await restoreSubtitlePosition();
    await restoreSubtitleFontSize();
    await applySubtitleStyle();
  }

  setupYomitanObserver();

  setupSelectionObserver();
  if (shouldToggleMouseIgnore) {
    window.electronAPI.setIgnoreMouseEvents(true, { forward: true });
  }
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", init);
} else {
  init();
}
