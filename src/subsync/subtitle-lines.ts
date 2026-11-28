export type SubtitleFileFormat = 'ass' | 'srt' | 'vtt';

export interface SubtitleTimeSpan {
  start: number;
  end: number;
}

export interface ParsedSubtitleLine {
  /** Line index in the original file's `lines` array (0-based). */
  lineIndex: number;
  start: number;
  end: number;
  text: string;
  sign: boolean;
  /** ASS field positions for format-preserving timestamp rewrite. */
  assStartFieldIndex?: number;
  assEndFieldIndex?: number;
  assFieldCount?: number;
  /** SRT/VTT suffix after the end timestamp (e.g., VTT cue settings). */
  timingSuffix?: string;
  /** SRT/VTT leading indentation before start timestamp. */
  timingPrefix?: string;
}

export interface ParsedSubtitleFile {
  format: SubtitleFileFormat;
  eol: '\r\n' | '\n';
  rawLines: string[];
  events: ParsedSubtitleLine[];
  /** False when byte decoding fell back to lossy replacement (e.g. U+FFFD). */
  lossless?: boolean;
  /**
   * True when the decoded text started with a U+FEFF marker: the source had a BOM or was
   * not UTF-8. Rewrites keep a UTF-8 BOM so ANSI-codepage readers (VSFilter) do not misread.
   */
  bom?: boolean;
}

const SIGN_STYLE_WORDS = new Set([
  'sign',
  'signs',
  'signage',
  'typeset',
  'ts',
  'title',
  'titles',
  'logo',
  'screen',
  'onscreen',
  'insert',
  'song',
  'songs',
  'lyric',
  'lyrics',
  'karaoke',
  'kara',
  'kfx',
  'fx',
  'opening',
  'ending',
  'note',
  'notes',
  'tn',
  'credit',
  'credits',
  'eyecatch',
  'next',
  'preview',
  'banner',
]);

// Only all-caps "OP"/"ED" mark song styles; "Ed" or "Op" are usually character names.
const UPPERCASE_ONLY_SIGN_STYLE_WORDS = new Set(['OP', 'ED']);
const JAPANESE_SIGN_STYLE_PATTERN = /看板|歌|タイトル/;

const TRACK_SIGNS_OR_SONGS_PATTERN =
  /(?:^|[^\p{L}\p{N}])(?:signs?(?:\s*(?:&|and|\+|\/|-|_)\s*songs?)?|songs?|lyrics?|karaoke|forced|s[\s&+_-]*s|op|ed|opening|ending)(?=$|[^\p{L}\p{N}])|看板|歌詞/iu;

// "English (Full + Songs)" or "日本語 (歌詞あり)" are full dialogue tracks that also carry lyrics.
const TRACK_FULL_DIALOGUE_PATTERN =
  /\b(?:full|dialogue|dialog|complete)\b|歌詞(?:あり|付き)|台詞|セリフ/i;

const SRT_VTT_TIMING_LINE_PATTERN =
  /^(\s*)(\d{1,2}:\d{2}:\d{2}[,.]\d{1,3}|\d{1,2}:\d{2}[,.]\d{1,3})(\s*-->\s*)(\d{1,2}:\d{2}:\d{2}[,.]\d{1,3}|\d{1,2}:\d{2}[,.]\d{1,3})(.*)$/;

const ASS_DRAWING_PATTERN = /\{[^}]*\\p[1-9][^}]*\}[^{]*/g;
const ASS_TAG_PATTERN = /\{[^}]*\}/g;
const ASS_KARAOKE_TAG_PATTERN = /\{[^}]*\\(?:k|kf|ko|K)\d+[^}]*\}/;
const HTML_TAG_PATTERN = /<\/?[a-zA-Z][^>]*>/g;
const MUSIC_ONLY_TEXT_PATTERN =
  /^(?:[♪♫♩♬\s~～\-–—()（）[\]【】]*|(?:[\[(（](?:music|song|singing|lyrics?|instrumental|op|ed|opening|ending|bgm|音楽|歌|歌唱)[\])）]\s*)+)$/iu;

function tryDecode(bytes: Uint8Array, encoding: string, fatal: boolean): string | null {
  try {
    return new TextDecoder(encoding, { fatal }).decode(bytes);
  } catch {
    return null;
  }
}

export interface DecodedSubtitleBytes {
  /**
   * Decoded text. Starts with U+FEFF when the source had a BOM or was a non-UTF-8 encoding,
   * so `parseSubtitleFileContent` records that a rewrite (always UTF-8) needs a BOM.
   */
  text: string;
  lossless: boolean;
}

const BOM = '\uFEFF';

export function decodeSubtitleBytesWithStatus(bytes: Uint8Array): DecodedSubtitleBytes {
  const decoded = decodeSubtitleBytesRaw(bytes);
  if (!decoded.lossless || !decoded.needsBom) {
    return { text: decoded.text, lossless: decoded.lossless };
  }
  return { text: `${BOM}${decoded.text}`, lossless: true };
}

function decodeSubtitleBytesRaw(bytes: Uint8Array): DecodedSubtitleBytes & { needsBom: boolean } {
  if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    const rest = bytes.subarray(3);
    const strict = tryDecode(rest, 'utf-8', true);
    if (strict !== null) return { text: strict, lossless: true, needsBom: true };
    return { text: new TextDecoder('utf-8').decode(rest), lossless: false, needsBom: true };
  }
  if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) {
    const strict = tryDecode(bytes, 'utf-16le', true);
    if (strict !== null) return { text: strict, lossless: true, needsBom: true };
    return {
      text: tryDecode(bytes, 'utf-16le', false) ?? new TextDecoder('utf-8').decode(bytes),
      lossless: false,
      needsBom: true,
    };
  }
  if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) {
    const strict = tryDecode(bytes, 'utf-16be', true);
    if (strict !== null) return { text: strict, lossless: true, needsBom: true };
    return {
      text: tryDecode(bytes, 'utf-16be', false) ?? new TextDecoder('utf-8').decode(bytes),
      lossless: false,
      needsBom: true,
    };
  }
  const strictUtf8 = tryDecode(bytes, 'utf-8', true);
  if (strictUtf8 !== null) {
    return { text: strictUtf8, lossless: true, needsBom: false };
  }
  const shiftJis = tryDecode(bytes, 'shift_jis', true);
  const eucJp = tryDecode(bytes, 'euc-jp', true);
  if (shiftJis !== null && eucJp === null) {
    return { text: shiftJis, lossless: true, needsBom: true };
  }
  if (eucJp !== null && shiftJis === null) {
    return { text: eucJp, lossless: true, needsBom: true };
  }
  if (shiftJis !== null && eucJp !== null) {
    if (shiftJis === eucJp) {
      return { text: shiftJis, lossless: true, needsBom: true };
    }
    return { text: eucJp, lossless: false, needsBom: true };
  }
  return { text: new TextDecoder('utf-8').decode(bytes), lossless: false, needsBom: true };
}

export function decodeSubtitleBytes(bytes: Uint8Array): string {
  return decodeSubtitleBytesRaw(bytes).text;
}

/**
 * Split a style or effect name into lowercase word tokens so camelCase / snake_case /
 * kebab-case names like "OPRomaji", "TLNote", "sign_0423_EpTitle" match cleanly while
 * dialogue styles like "Subtitle", "Notebook", or "Edward" do not.
 */
export function splitStyleWords(name: string): string[] {
  return splitStyleTokens(name).map((word) => word.toLowerCase());
}

/** Same split as `splitStyleWords`, keeping the original letter case. */
function splitStyleTokens(name: string): string[] {
  const words: string[] = [];
  let word = '';
  for (let i = 0; i < name.length; i += 1) {
    const ch = name[i]!;
    const isAlphaNum = /[A-Za-z0-9]/.test(ch);
    if (!isAlphaNum) {
      if (word) {
        words.push(word);
        word = '';
      }
      continue;
    }
    if (word) {
      const prev = name[i - 1]!;
      const next = name[i + 1] ?? '';
      const prevLower = /[a-z]/.test(prev);
      const prevUpper = /[A-Z]/.test(prev);
      const prevDigit = /[0-9]/.test(prev);
      const currUpper = /[A-Z]/.test(ch);
      const currDigit = /[0-9]/.test(ch);
      const nextLower = /[a-z]/.test(next);
      if (
        (prevLower && currUpper) ||
        (prevUpper && currUpper && nextLower) ||
        prevDigit !== currDigit
      ) {
        words.push(word);
        word = '';
      }
    }
    word += ch;
  }
  if (word) {
    words.push(word);
  }
  return words;
}

export function isSignStyleOrEffect(value: string): boolean {
  if (!value.trim()) return false;
  if (JAPANESE_SIGN_STYLE_PATTERN.test(value)) return true;
  return splitStyleTokens(value).some(
    (token) =>
      SIGN_STYLE_WORDS.has(token.toLowerCase()) || UPPERCASE_ONLY_SIGN_STYLE_WORDS.has(token),
  );
}

export function isSignsOrSongsTrackMeta(track: {
  forced?: boolean;
  title?: string;
  'external-filename'?: string;
}): boolean {
  if (track.forced === true) return true;
  const filename = track['external-filename']?.split(/[\\/]/).pop() ?? '';
  // Judge title and filename separately so a full-dialogue hint only overrides its own field.
  return [track.title ?? '', filename].some(
    (field) => TRACK_SIGNS_OR_SONGS_PATTERN.test(field) && !TRACK_FULL_DIALOGUE_PATTERN.test(field),
  );
}

export function parseSubtitleTimestamp(raw: string): number | null {
  const normalized = raw.trim().replace(',', '.');
  const parts = normalized.split(':');
  if (parts.length === 3) {
    const h = Number(parts[0]);
    const m = Number(parts[1]);
    const s = Number(parts[2]);
    if (!Number.isFinite(h) || !Number.isFinite(m) || !Number.isFinite(s)) return null;
    return h * 3600 + m * 60 + s;
  }
  if (parts.length === 2) {
    const m = Number(parts[0]);
    const s = Number(parts[1]);
    if (!Number.isFinite(m) || !Number.isFinite(s)) return null;
    return m * 60 + s;
  }
  return null;
}

function cleanCuePlainText(raw: string): string {
  return raw
    .replace(ASS_DRAWING_PATTERN, '')
    .replace(ASS_TAG_PATTERN, '')
    .replace(HTML_TAG_PATTERN, '')
    .replace(/\\N|\\n/g, '\n')
    .replace(/\\h/g, ' ')
    .replace(/\u200B/g, '')
    .split('\n')
    .map((line) => line.trim())
    .filter(Boolean)
    .join('\n');
}

function isMusicOrLyricText(plainText: string, rawText: string): boolean {
  if (!plainText) return true;
  // Lyrics open with a music mark ("♪ 夢の中へ ♪"); a trailing one is just tone ("ありがとう♪").
  if (/^[\s(（[【~～]*[♪♫♩♬]/u.test(plainText)) return true;
  if (MUSIC_ONLY_TEXT_PATTERN.test(plainText)) return true;
  if (ASS_KARAOKE_TAG_PATTERN.test(rawText)) return true;
  return false;
}

function detectFormat(text: string, filename?: string): SubtitleFileFormat {
  const lowerName = (filename ?? '').toLowerCase();
  if (
    /^\[events\]/im.test(text) ||
    /^dialogue:/im.test(text) ||
    lowerName.endsWith('.ass') ||
    lowerName.endsWith('.ssa')
  ) {
    return 'ass';
  }
  if (/^\s*WEBVTT\b/m.test(text) || lowerName.endsWith('.vtt')) {
    return 'vtt';
  }
  return 'srt';
}

function splitAssDialogueFields(rest: string, fieldCount: number): string[] {
  const fields: string[] = [];
  let start = 0;
  for (let i = 0; i < fieldCount - 1; i += 1) {
    const comma = rest.indexOf(',', start);
    if (comma === -1) {
      fields.push(rest.slice(start));
      return fields;
    }
    fields.push(rest.slice(start, comma));
    start = comma + 1;
  }
  fields.push(rest.slice(start));
  return fields;
}

function parseAssSubtitle(rawLines: string[], eol: '\r\n' | '\n'): ParsedSubtitleFile {
  interface RawAssEntry {
    lineIndex: number;
    start: number;
    end: number;
    text: string;
    namedSign: boolean;
    placed: boolean;
    musicOrEmpty: boolean;
    startFieldIndex: number;
    endFieldIndex: number;
    fieldCount: number;
  }

  let inEvents = false;
  let format: string[] = [];
  const rawEntries: RawAssEntry[] = [];

  for (let i = 0; i < rawLines.length; i += 1) {
    const trimmed = rawLines[i]!.trim();
    if (trimmed.startsWith('[')) {
      inEvents = trimmed.toLowerCase() === '[events]';
      continue;
    }
    if (!inEvents) continue;

    if (/^format:/i.test(trimmed)) {
      format = trimmed
        .slice(trimmed.indexOf(':') + 1)
        .split(',')
        .map((field) => field.trim().toLowerCase());
      continue;
    }

    if (!/^dialogue:/i.test(trimmed)) continue;

    if (format.length === 0) {
      format = [
        'layer',
        'start',
        'end',
        'style',
        'name',
        'marginl',
        'marginr',
        'marginv',
        'effect',
        'text',
      ];
    }

    const colonIndex = rawLines[i]!.indexOf(':');
    const rest = rawLines[i]!.slice(colonIndex + 1);
    const fields = splitAssDialogueFields(rest, format.length);

    const startIdx = format.indexOf('start');
    const endIdx = format.indexOf('end');
    const styleIdx = format.indexOf('style');
    const effectIdx = format.indexOf('effect');
    const textIdx = format.indexOf('text');
    if (startIdx < 0 || endIdx < 0 || textIdx < 0) continue;

    const start = parseSubtitleTimestamp(fields[startIdx] ?? '');
    const end = parseSubtitleTimestamp(fields[endIdx] ?? '');
    if (start === null || end === null || end <= start) continue;

    const style = (fields[styleIdx] ?? '').trim();
    const effect = (fields[effectIdx] ?? '').trim();
    const rawBody = fields[textIdx] ?? '';
    const plainText = cleanCuePlainText(rawBody);
    const namedSign = isSignStyleOrEffect(style) || isSignStyleOrEffect(effect);
    const placed = /\\(?:pos|move)\(/.test(rawBody);
    const musicOrEmpty = isMusicOrLyricText(plainText, rawBody);

    rawEntries.push({
      lineIndex: i,
      start,
      end,
      text: plainText,
      namedSign,
      placed,
      musicOrEmpty,
      startFieldIndex: startIdx,
      endFieldIndex: endIdx,
      fieldCount: format.length,
    });
  }

  // Only treat \pos / \move as sign indicators when fewer than half of the
  // non-named lines are positioned (some conversion scripts position every line).
  let unnamedCount = 0;
  let placedUnnamedCount = 0;
  for (const entry of rawEntries) {
    if (!entry.namedSign && !entry.musicOrEmpty) {
      unnamedCount += 1;
      if (entry.placed) placedUnnamedCount += 1;
    }
  }
  const placingMarksSigns = placedUnnamedCount * 2 <= unnamedCount;

  const events: ParsedSubtitleLine[] = rawEntries.map((entry) => ({
    lineIndex: entry.lineIndex,
    start: entry.start,
    end: entry.end,
    text: entry.text,
    sign: entry.namedSign || entry.musicOrEmpty || (entry.placed && placingMarksSigns),
    assStartFieldIndex: entry.startFieldIndex,
    assEndFieldIndex: entry.endFieldIndex,
    assFieldCount: entry.fieldCount,
  }));

  return { format: 'ass', eol, rawLines, events };
}

function parseSrtOrVttSubtitle(
  rawLines: string[],
  eol: '\r\n' | '\n',
  format: 'srt' | 'vtt',
): ParsedSubtitleFile {
  const events: ParsedSubtitleLine[] = [];

  for (let i = 0; i < rawLines.length; i += 1) {
    const match = SRT_VTT_TIMING_LINE_PATTERN.exec(rawLines[i]!);
    if (!match) continue;

    const start = parseSubtitleTimestamp(match[2]!);
    const end = parseSubtitleTimestamp(match[4]!);
    if (start === null || end === null || end <= start) continue;

    const bodyLines: string[] = [];
    let j = i + 1;
    while (j < rawLines.length && rawLines[j]!.trim() !== '') {
      if (SRT_VTT_TIMING_LINE_PATTERN.test(rawLines[j]!)) break;
      bodyLines.push(rawLines[j]!);
      j += 1;
    }

    const rawBody = bodyLines.join('\n');
    const plainText = cleanCuePlainText(rawBody);
    events.push({
      lineIndex: i,
      start,
      end,
      text: plainText,
      sign: isMusicOrLyricText(plainText, rawBody),
      timingPrefix: match[1] ?? '',
      timingSuffix: match[5] ?? '',
    });

    i = j - 1;
  }

  return { format, eol, rawLines, events };
}

export function parseSubtitleFileContent(content: string, filename?: string): ParsedSubtitleFile {
  const bom = content.startsWith(BOM);
  const cleanContent = bom ? content.slice(1) : content;
  const eol: '\r\n' | '\n' = cleanContent.includes('\r\n') ? '\r\n' : '\n';
  const rawLines = cleanContent.split(/\r?\n/);
  const format = detectFormat(cleanContent, filename);
  const parsed =
    format === 'ass'
      ? parseAssSubtitle(rawLines, eol)
      : parseSrtOrVttSubtitle(rawLines, eol, format);
  return bom ? { ...parsed, bom } : parsed;
}

export function extractSpansFromParsedSubtitle(
  parsed: ParsedSubtitleFile,
  options: { stripSigns?: boolean } = {},
): {
  spans: SubtitleTimeSpan[];
  dialogueIndices: number[];
  dialogueCount: number;
  signCount: number;
} {
  const stripSigns = options.stripSigns !== false;
  const dialogueIndices: number[] = [];
  let signCount = 0;

  for (let i = 0; i < parsed.events.length; i += 1) {
    const event = parsed.events[i]!;
    if (event.sign) {
      signCount += 1;
      if (stripSigns) continue;
    }
    dialogueIndices.push(i);
  }

  // Fallback: if stripping signs left zero lines in an unusual file, keep all lines.
  const activeIndices =
    dialogueIndices.length > 0 ? dialogueIndices : parsed.events.map((_, idx) => idx);

  const spans = activeIndices.map((idx) => {
    const ev = parsed.events[idx]!;
    return { start: ev.start, end: ev.end };
  });

  return {
    spans,
    dialogueIndices: activeIndices,
    dialogueCount: parsed.events.length - signCount,
    signCount,
  };
}

export function formatAssTimestamp(seconds: number): string {
  const clamped = Math.max(0, seconds);
  const totalCentiseconds = Math.round(clamped * 100);
  const cs = totalCentiseconds % 100;
  const totalSeconds = Math.floor(totalCentiseconds / 100);
  const sec = totalSeconds % 60;
  const totalMinutes = Math.floor(totalSeconds / 60);
  const min = totalMinutes % 60;
  const hours = Math.floor(totalMinutes / 60);
  return `${hours}:${String(min).padStart(2, '0')}:${String(sec).padStart(2, '0')}.${String(cs).padStart(2, '0')}`;
}

export function formatSrtTimestamp(seconds: number): string {
  const clamped = Math.max(0, seconds);
  const totalMs = Math.round(clamped * 1000);
  const ms = totalMs % 1000;
  const totalSeconds = Math.floor(totalMs / 1000);
  const sec = totalSeconds % 60;
  const totalMinutes = Math.floor(totalSeconds / 60);
  const min = totalMinutes % 60;
  const hours = Math.floor(totalMinutes / 60);
  return `${String(hours).padStart(2, '0')}:${String(min).padStart(2, '0')}:${String(sec).padStart(2, '0')},${String(ms).padStart(3, '0')}`;
}

export function formatVttTimestamp(seconds: number): string {
  return formatSrtTimestamp(seconds).replace(',', '.');
}

interface DialogueMoveAnchor {
  start: number;
  move: number;
}

/** Dialogue (start, move) pairs sorted by start time, for time-based sign/song shifting. */
function buildDialogueMoveAnchors(
  events: readonly ParsedSubtitleLine[],
  dialogueIndices: readonly number[],
  movesSec: readonly number[],
): DialogueMoveAnchor[] {
  return dialogueIndices
    .map((eventIndex, k) => ({ start: events[eventIndex]!.start, move: movesSec[k] ?? 0 }))
    .sort((a, b) => a.start - b.start);
}

/**
 * Shift of the dialogue line nearest `start` in time. ASS files often keep typesetting
 * or OP/ED karaoke at the top or bottom of the file, so file position is meaningless.
 */
function nearestDialogueMove(anchors: readonly DialogueMoveAnchor[], start: number): number {
  if (anchors.length === 0) return 0;
  let lo = 0;
  let hi = anchors.length;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (anchors[mid]!.start < start) {
      lo = mid + 1;
    } else {
      hi = mid;
    }
  }
  if (lo === 0) return anchors[0]!.move;
  if (lo >= anchors.length) return anchors[anchors.length - 1]!.move;
  const prev = anchors[lo - 1]!;
  const next = anchors[lo]!;
  return start - prev.start <= next.start - start ? prev.move : next.move;
}

/**
 * Rewrite only the timestamps of the original subtitle file, preserving all ASS
 * styles/headers/tags/comments, SRT cue numbers, and VTT settings. Non-dialogue
 * signs/songs follow the shift of the dialogue line nearest them in time.
 */
export function applyTimingToSubtitleFile(
  parsed: ParsedSubtitleFile,
  dialogueIndices: readonly number[],
  movesSec: readonly number[],
  ratio: number,
): string {
  if (parsed.events.length === 0 || dialogueIndices.length === 0 || movesSec.length === 0) {
    return serializeSubtitleLines(parsed, parsed.rawLines);
  }

  const outLines = [...parsed.rawLines];
  const dialogueMoveByEventIndex = new Map<number, number>();
  for (let k = 0; k < dialogueIndices.length; k += 1) {
    dialogueMoveByEventIndex.set(dialogueIndices[k]!, movesSec[k] ?? 0);
  }
  const anchors = buildDialogueMoveAnchors(parsed.events, dialogueIndices, movesSec);

  for (let i = 0; i < parsed.events.length; i += 1) {
    const ev = parsed.events[i]!;
    const directMove = dialogueMoveByEventIndex.get(i);
    const shift = directMove ?? nearestDialogueMove(anchors, ev.start);

    const newStart = Math.max(0, ev.start * ratio + shift);
    const newEnd = Math.max(newStart + 0.05, ev.end * ratio + shift);

    if (parsed.format === 'ass') {
      const rawLine = outLines[ev.lineIndex]!;
      const colonIndex = rawLine.indexOf(':');
      const prefix = rawLine.slice(0, colonIndex + 1);
      const rest = rawLine.slice(colonIndex + 1);
      const fieldCount = ev.assFieldCount ?? 10;
      const fields = splitAssDialogueFields(rest, fieldCount);
      if (
        ev.assStartFieldIndex !== undefined &&
        ev.assEndFieldIndex !== undefined &&
        ev.assStartFieldIndex < fields.length &&
        ev.assEndFieldIndex < fields.length
      ) {
        fields[ev.assStartFieldIndex] = formatAssTimestamp(newStart);
        fields[ev.assEndFieldIndex] = formatAssTimestamp(newEnd);
        outLines[ev.lineIndex] = `${prefix}${fields.join(',')}`;
      }
    } else if (parsed.format === 'vtt') {
      outLines[ev.lineIndex] =
        `${ev.timingPrefix ?? ''}${formatVttTimestamp(newStart)} --> ${formatVttTimestamp(newEnd)}${ev.timingSuffix ?? ''}`;
    } else {
      outLines[ev.lineIndex] =
        `${ev.timingPrefix ?? ''}${formatSrtTimestamp(newStart)} --> ${formatSrtTimestamp(newEnd)}${ev.timingSuffix ?? ''}`;
    }
  }

  return serializeSubtitleLines(parsed, outLines);
}

/** Join rewritten lines with the source line ending, restoring a BOM when the source needs one. */
export function serializeSubtitleLines(
  parsed: Pick<ParsedSubtitleFile, 'eol' | 'bom'>,
  lines: readonly string[],
): string {
  const joined = lines.join(parsed.eol);
  return parsed.bom ? `${BOM}${joined}` : joined;
}

/**
 * Emit a clean dialogue-only SRT file from spans, useful when feeding a cleaned
 * reference subtitle into external `alass`.
 */
export function formatSpansAsSrt(spans: readonly SubtitleTimeSpan[]): string {
  const valid = spans
    .filter(
      (span) => Number.isFinite(span.start) && Number.isFinite(span.end) && span.end > span.start,
    )
    .sort((a, b) => a.start - b.start || a.end - b.end);

  return valid
    .map(
      (span, index) =>
        `${index + 1}\n${formatSrtTimestamp(span.start)} --> ${formatSrtTimestamp(span.end)}\nDialogue ${index + 1}\n`,
    )
    .join('\n');
}
