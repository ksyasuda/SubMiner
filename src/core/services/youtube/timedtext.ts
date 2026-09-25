interface YoutubeTimedTextRow {
  startMs: number;
  durationMs: number;
  text: string;
  isGenerated: boolean;
  rollingWindow: YoutubeRollingWindow | null;
}

interface YoutubeRollingWindow {
  rowCount: number;
  columnCount: number;
}

interface YoutubeTimedTextWindowDefinitions {
  rollingStyleIds: Set<string>;
  positions: Map<string, YoutubeRollingWindow>;
  windows: Map<string, YoutubeRollingWindow>;
}

interface YoutubeTimedTextDocument {
  rows: YoutubeTimedTextRow[];
  // Start times of every <p> event, including empty window-append fillers.
  // Rolling speech rows with a 3000ms placeholder display until the next event.
  eventStartsMs: number[];
  hasRollingWindowEvents: boolean;
}

const YOUTUBE_TIMEDTEXT_EXTENSIONS = new Set(['srv1', 'srv2', 'srv3', 'ytsrv3']);
const YOUTUBE_ROLLING_PLACEHOLDER_DURATION_MS = 3_000;

function decodeNumericEntity(match: string, codePoint: number): string {
  if (
    !Number.isInteger(codePoint) ||
    codePoint < 0 ||
    codePoint > 0x10ffff ||
    (codePoint >= 0xd800 && codePoint <= 0xdfff)
  ) {
    return match;
  }
  return String.fromCodePoint(codePoint);
}

function decodeHtmlEntities(value: string): string {
  return value
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&#(\d+);/g, (match, codePoint) => decodeNumericEntity(match, Number(codePoint)))
    .replace(/&#x([0-9a-f]+);/gi, (match, codePoint) =>
      decodeNumericEntity(match, Number.parseInt(codePoint, 16)),
    );
}

function parseAttributeMap(raw: string): Map<string, string> {
  const attrs = new Map<string, string>();
  for (const match of raw.matchAll(/([a-zA-Z0-9:_-]+)="([^"]*)"/g)) {
    attrs.set(match[1]!, match[2]!);
  }
  return attrs;
}

function parsePositiveInteger(value: string | undefined): number | null {
  if (value === undefined) {
    return null;
  }
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) && parsed > 0 ? parsed : null;
}

function extractYoutubeTimedTextWindowDefinitions(xml: string): YoutubeTimedTextWindowDefinitions {
  const rollingStyleIds = new Set<string>();
  for (const match of xml.matchAll(/<ws\b([^>]*)\/?\s*>/g)) {
    const attrs = parseAttributeMap(match[1] ?? '');
    const id = attrs.get('id');
    if (id !== undefined && attrs.get('mh') === '2') {
      rollingStyleIds.add(id);
    }
  }

  const positions = new Map<string, YoutubeRollingWindow>();
  for (const match of xml.matchAll(/<wp\b([^>]*)\/?\s*>/g)) {
    const attrs = parseAttributeMap(match[1] ?? '');
    const id = attrs.get('id');
    const rowCount = parsePositiveInteger(attrs.get('rc'));
    const columnCount = parsePositiveInteger(attrs.get('cc'));
    if (id !== undefined && rowCount !== null && columnCount !== null) {
      positions.set(id, { rowCount, columnCount });
    }
  }

  const windows = new Map<string, YoutubeRollingWindow>();
  for (const match of xml.matchAll(/<w\b([^>]*)\/?\s*>/g)) {
    const attrs = parseAttributeMap(match[1] ?? '');
    const id = attrs.get('id');
    const styleId = attrs.get('ws');
    const positionId = attrs.get('wp');
    const position = positionId === undefined ? undefined : positions.get(positionId);
    if (
      id !== undefined &&
      styleId !== undefined &&
      rollingStyleIds.has(styleId) &&
      position !== undefined
    ) {
      windows.set(id, position);
    }
  }

  return { rollingStyleIds, positions, windows };
}

function resolveRollingWindow(
  attrs: Map<string, string>,
  definitions: YoutubeTimedTextWindowDefinitions,
): YoutubeRollingWindow | null {
  const windowId = attrs.get('w');
  if (windowId !== undefined) {
    return definitions.windows.get(windowId) ?? null;
  }

  const styleId = attrs.get('ws');
  const positionId = attrs.get('wp');
  if (
    styleId === undefined ||
    positionId === undefined ||
    !definitions.rollingStyleIds.has(styleId)
  ) {
    return null;
  }
  return definitions.positions.get(positionId) ?? null;
}

function extractYoutubeTimedTextDocument(xml: string): YoutubeTimedTextDocument {
  const rows: YoutubeTimedTextRow[] = [];
  const eventStartsMs: number[] = [];
  let hasRollingWindowEvents = false;
  const windowDefinitions = extractYoutubeTimedTextWindowDefinitions(xml);

  for (const match of xml.matchAll(/<p\b([^>]*)>([\s\S]*?)<\/p>/g)) {
    const attrs = parseAttributeMap(match[1] ?? '');
    const startMs = Number(attrs.get('t'));
    if (!Number.isFinite(startMs)) {
      continue;
    }
    eventStartsMs.push(startMs);
    if (attrs.get('a') === '1') {
      hasRollingWindowEvents = true;
    }

    const durationMs = Number(attrs.get('d'));
    if (!Number.isFinite(durationMs)) {
      continue;
    }

    const rawInner = match[2] ?? '';
    const inner = rawInner.replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, '');
    const text = decodeHtmlEntities(inner).trim();
    if (!text) {
      continue;
    }

    rows.push({
      startMs,
      durationMs,
      text,
      isGenerated: /<s\b/.test(rawInner),
      rollingWindow: resolveRollingWindow(attrs, windowDefinitions),
    });
  }

  eventStartsMs.sort((a, b) => a - b);
  return { rows, eventStartsMs, hasRollingWindowEvents };
}

function findNextEventStartMs(eventStartsMs: number[], afterMs: number): number | undefined {
  for (const startMs of eventStartsMs) {
    if (startMs > afterMs) {
      return startMs;
    }
  }
  return undefined;
}

function isGeneratedRollingCue(row: YoutubeTimedTextRow, hasRollingWindowEvents: boolean): boolean {
  return row.isGenerated && (row.rollingWindow !== null || hasRollingWindowEvents);
}

function formatVttTimestamp(ms: number): string {
  const totalMs = Math.max(0, Math.floor(ms));
  const hours = Math.floor(totalMs / 3_600_000);
  const minutes = Math.floor((totalMs % 3_600_000) / 60_000);
  const seconds = Math.floor((totalMs % 60_000) / 1_000);
  const millis = totalMs % 1_000;
  return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}.${String(millis).padStart(3, '0')}`;
}

const ROLLING_PAGE_BREAK_PATTERN = /[\s、。！？!?]/u;

// VTT cannot carry SRV3's row and column limits. Page only roll-up windows so
// the overlay keeps their bounded presentation without changing authored cues.
function splitRollingCaptionIntoPages(text: string, rollingWindow: YoutubeRollingWindow): string[] {
  const pageCapacity = rollingWindow.rowCount * rollingWindow.columnCount;
  const characters = [...text];
  if (
    !Number.isSafeInteger(pageCapacity) ||
    pageCapacity <= 0 ||
    characters.length <= pageCapacity
  ) {
    return [text];
  }

  const pages: string[] = [];
  let pageStart = 0;
  while (pageStart < characters.length) {
    let pageEnd = Math.min(pageStart + pageCapacity, characters.length);
    if (pageEnd < characters.length) {
      const earliestNaturalBreak = pageStart + Math.ceil(pageCapacity * 0.6);
      for (let index = pageEnd - 1; index >= earliestNaturalBreak; index -= 1) {
        if (ROLLING_PAGE_BREAK_PATTERN.test(characters[index]!)) {
          pageEnd = index + 1;
          break;
        }
      }
    }
    pages.push(characters.slice(pageStart, pageEnd).join(''));
    pageStart = pageEnd;
  }
  return pages;
}

interface TimedCaptionPage {
  startMs: number;
  endMs: number;
  text: string;
}

function timeCaptionPages(input: {
  text: string;
  pages: string[];
  startMs: number;
  endMs: number;
}): TimedCaptionPage[] {
  const durationMs = input.endMs - input.startMs;
  if (input.pages.length === 1 || durationMs < input.pages.length) {
    return [{ startMs: input.startMs, endMs: input.endMs, text: input.text }];
  }

  const totalCharacters = [...input.text].length;
  const timedPages: TimedCaptionPage[] = [];
  let consumedCharacters = 0;
  let pageStartMs = input.startMs;
  // Automatic captions often omit span offsets, so distribute the known cue
  // duration by page length while guaranteeing every page at least one ms.
  for (let index = 0; index < input.pages.length; index += 1) {
    const page = input.pages[index]!;
    consumedCharacters += [...page].length;
    const remainingPages = input.pages.length - index - 1;
    const proportionalEndMs =
      input.startMs + Math.round((durationMs * consumedCharacters) / totalCharacters);
    const pageEndMs =
      remainingPages === 0
        ? input.endMs
        : Math.min(Math.max(proportionalEndMs, pageStartMs + 1), input.endMs - remainingPages);
    timedPages.push({ startMs: pageStartMs, endMs: pageEndMs, text: page });
    pageStartMs = pageEndMs;
  }
  return timedPages;
}

export function isYoutubeTimedTextExtension(value: string | undefined): boolean {
  if (!value) {
    return false;
  }
  return YOUTUBE_TIMEDTEXT_EXTENSIONS.has(value.trim().toLowerCase());
}

export function convertYoutubeTimedTextToVtt(xml: string): string {
  const { rows, eventStartsMs, hasRollingWindowEvents } = extractYoutubeTimedTextDocument(xml);
  if (rows.length === 0) {
    return 'WEBVTT\n';
  }

  const blocks: string[] = [];
  let previousText = '';
  for (let index = 0; index < rows.length; index += 1) {
    const row = rows[index]!;
    const nextRow = rows[index + 1];
    const unclampedEnd = row.startMs + row.durationMs;
    // YouTube uses exactly 3000ms as a placeholder for generated rolling speech.
    // Plain-text cues can explicitly use the same duration and must keep it.
    const nextEventStart =
      isGeneratedRollingCue(row, hasRollingWindowEvents) &&
      row.durationMs === YOUTUBE_ROLLING_PLACEHOLDER_DURATION_MS
        ? findNextEventStartMs(eventStartsMs, row.startMs)
        : undefined;
    const clampedEnd =
      nextEventStart !== undefined
        ? nextEventStart
        : nextRow && unclampedEnd > nextRow.startMs
          ? Math.max(row.startMs, nextRow.startMs - 1)
          : unclampedEnd;
    if (clampedEnd <= row.startMs) {
      continue;
    }

    const text =
      previousText && row.text.startsWith(previousText)
        ? row.text.slice(previousText.length).trimStart()
        : row.text;
    previousText = row.text;
    if (!text) {
      continue;
    }
    const pages = row.rollingWindow
      ? splitRollingCaptionIntoPages(text, row.rollingWindow)
      : [text];
    for (const page of timeCaptionPages({
      text,
      pages,
      startMs: row.startMs,
      endMs: clampedEnd,
    })) {
      blocks.push(
        `${formatVttTimestamp(page.startMs)} --> ${formatVttTimestamp(page.endMs)}\n${page.text}`,
      );
    }
  }

  return `WEBVTT\n\n${blocks.join('\n\n')}\n`;
}

function normalizeRollingCaptionText(text: string, previousText: string): string {
  if (!previousText || !text.startsWith(previousText)) {
    return text;
  }
  return text.slice(previousText.length).trimStart();
}

export function normalizeYoutubeAutoVtt(content: string): string {
  const normalizedContent = content.replace(/\r\n?/g, '\n');
  const blocks = normalizedContent.split(/\n{2,}/);
  if (blocks.length === 0) {
    return content;
  }

  let previousText = '';
  let changed = false;
  const normalizedBlocks = blocks.map((block) => {
    if (!block.includes('-->')) {
      return block;
    }

    const lines = block.split('\n');
    const timingLineIndex = lines.findIndex((line) => line.includes('-->'));
    if (timingLineIndex < 0 || timingLineIndex === lines.length - 1) {
      return block;
    }

    const textLines = lines.slice(timingLineIndex + 1);
    const originalText = textLines.join('\n').trim();
    if (!originalText) {
      return block;
    }

    const normalizedText = normalizeRollingCaptionText(originalText, previousText);
    previousText = originalText;
    if (!normalizedText || normalizedText === originalText) {
      return block;
    }

    changed = true;
    return [...lines.slice(0, timingLineIndex + 1), normalizedText].join('\n');
  });

  if (!changed) {
    return content;
  }
  return `${normalizedBlocks.join('\n\n')}\n`;
}
