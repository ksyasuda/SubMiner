interface YoutubeTimedTextRow {
  startMs: number;
  durationMs: number;
  text: string;
}

const YOUTUBE_TIMEDTEXT_EXTENSIONS = new Set(['srv1', 'srv2', 'srv3', 'ytsrv3']);

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

function extractYoutubeTimedTextRows(xml: string): YoutubeTimedTextRow[] {
  const rows: YoutubeTimedTextRow[] = [];

  for (const match of xml.matchAll(/<p\b([^>]*)>([\s\S]*?)<\/p>/g)) {
    const attrs = parseAttributeMap(match[1] ?? '');
    const startMs = Number(attrs.get('t'));
    const durationMs = Number(attrs.get('d'));
    if (!Number.isFinite(startMs) || !Number.isFinite(durationMs)) {
      continue;
    }

    const inner = (match[2] ?? '').replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, '');
    const text = decodeHtmlEntities(inner).trim();
    if (!text) {
      continue;
    }

    rows.push({ startMs, durationMs, text });
  }

  return rows;
}

function formatVttTimestamp(ms: number): string {
  const totalMs = Math.max(0, Math.floor(ms));
  const hours = Math.floor(totalMs / 3_600_000);
  const minutes = Math.floor((totalMs % 3_600_000) / 60_000);
  const seconds = Math.floor((totalMs % 60_000) / 1_000);
  const millis = totalMs % 1_000;
  return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}.${String(millis).padStart(3, '0')}`;
}

export function isYoutubeTimedTextExtension(value: string | undefined): boolean {
  if (!value) {
    return false;
  }
  return YOUTUBE_TIMEDTEXT_EXTENSIONS.has(value.trim().toLowerCase());
}

export function convertYoutubeTimedTextToVtt(xml: string): string {
  const rows = extractYoutubeTimedTextRows(xml);
  if (rows.length === 0) {
    return 'WEBVTT\n';
  }

  const blocks: string[] = [];
  let previousText = '';
  for (let index = 0; index < rows.length; index += 1) {
    const row = rows[index]!;
    const nextRow = rows[index + 1];
    const unclampedEnd = row.startMs + row.durationMs;
    const clampedEnd =
      nextRow && unclampedEnd > nextRow.startMs
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
    blocks.push(
      `${formatVttTimestamp(row.startMs)} --> ${formatVttTimestamp(clampedEnd)}\n${text}`,
    );
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
