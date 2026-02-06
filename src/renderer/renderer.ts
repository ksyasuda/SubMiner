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

interface Keybinding {
  key: string;
  command: (string | number)[] | null;
}

interface SubtitlePosition {
  yPercent: number;
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

let isOverSubtitle = false;
let isDragging = false;
let dragStartY = 0;
let startYPercent = 0;
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

function normalizeSubtitle(text: string): string {
  if (!text) return "";

  let normalized = text.replace(/\\N/g, "\n").replace(/\\n/g, "\n");

  normalized = normalized.replace(/\{[^}]*\}/g, "");

  return normalized.trim();
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

function renderSubtitle(data: SubtitleData | string): void {
  subtitleRoot.innerHTML = "";

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
}

function handleMouseLeave(): void {
  isOverSubtitle = false;
  const yomitanPopup = document.querySelector('iframe[id^="yomitan-popup"]');
  if (!yomitanPopup && !jimakuModalOpen && !kikuModalOpen) {
    overlay.classList.remove("interactive");
  }
}

function getCurrentYPercent(): number {
  const marginBottom = parseFloat(subtitleContainer.style.marginBottom) || 60;
  const windowHeight = window.innerHeight;
  return (marginBottom / windowHeight) * 100;
}

function applyYPercent(yPercent: number): void {
  const clampedPercent = Math.max(2, Math.min(80, yPercent));
  const marginBottom = (clampedPercent / 100) * window.innerHeight;

  subtitleContainer.style.position = "";
  subtitleContainer.style.left = "";
  subtitleContainer.style.top = "";
  subtitleContainer.style.right = "";
  subtitleContainer.style.transform = "";

  subtitleContainer.style.marginBottom = `${marginBottom}px`;
}

function applyStoredSubtitlePosition(
  position: SubtitlePosition | null,
  source: string,
): void {
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

function applySubtitleFontSize(fontSize: number): void {
  const clampedSize = Math.max(10, Math.min(96, fontSize));
  document.documentElement.style.setProperty(
    "--subtitle-font-size",
    `${clampedSize}px`,
  );
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
  if (!isOverSubtitle) {
    overlay.classList.remove("interactive");
  }
  resetJimakuLists();
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

  overlay.classList.add("interactive");
  kikuModal.classList.remove("hidden");
  kikuModal.setAttribute("aria-hidden", "false");
}

function closeKikuFieldGroupingModal(): void {
  if (!kikuModalOpen) return;
  kikuModalOpen = false;
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
  if (!isOverSubtitle && !jimakuModalOpen) {
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
      window.electronAPI.saveSubtitlePosition({ yPercent });
    }
  });

  subtitleContainer.addEventListener("contextmenu", (e: Event) => {
    e.preventDefault();
  });
}

function isInteractiveTarget(target: EventTarget | null): boolean {
  if (!(target instanceof Element)) return false;
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

let keybindingsMap = new Map<string, (string | number)[]>();

type ChordAction =
  | { type: "mpv"; command: string[] }
  | { type: "electron"; action: () => void }
  | { type: "noop" };

const CHORD_MAP = new Map<string, ChordAction>([
  [
    "KeyT",
    { type: "electron", action: () => window.electronAPI.toggleOverlay() },
  ],
  [
    "Shift+KeyS",
    { type: "electron", action: () => window.electronAPI.quitApp() },
  ],
  [
    "KeyO",
    {
      type: "electron",
      action: () => window.electronAPI.openYomitanSettings(),
    },
  ],
  ["KeyR", { type: "mpv", command: ["script-message", "subminer-restart"] }],
  ["KeyC", { type: "mpv", command: ["script-message", "subminer-status"] }],
  ["KeyY", { type: "mpv", command: ["script-message", "subminer-menu"] }],
  ["KeyJ", { type: "electron", action: () => openJimakuModal() }],
  [
    "KeyD",
    { type: "electron", action: () => window.electronAPI.toggleDevTools() },
  ],
  ["KeyS", { type: "noop" }],
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
    const currentYPercent = getCurrentYPercent();
    applyYPercent(currentYPercent);
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
            if (!isOverSubtitle && !jimakuModalOpen && !kikuModalOpen) {
              overlay.classList.remove("interactive");
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
  window.electronAPI.onSubtitle((data: SubtitleData) => {
    renderSubtitle(data);
  });

  window.electronAPI.onSubtitlePosition((position: SubtitlePosition | null) => {
    applyStoredSubtitlePosition(position, "media-change");
  });

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

  subtitleContainer.addEventListener("mouseenter", handleMouseEnter);
  subtitleContainer.addEventListener("mouseleave", handleMouseLeave);

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

  window.electronAPI.onKikuFieldGroupingRequest(
    (data: {
      original: KikuDuplicateCardInfo;
      duplicate: KikuDuplicateCardInfo;
    }) => {
      openKikuFieldGroupingModal(data);
    },
  );

  setupDragging();

  await setupMpvInputForwarding();

  setupResizeHandler();

  await restoreSubtitlePosition();
  await restoreSubtitleFontSize();

  await applySubtitleStyle();

  setupYomitanObserver();

  setupSelectionObserver();
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", init);
} else {
  init();
}
