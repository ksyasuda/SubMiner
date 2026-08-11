import {
  assOverrideSignature,
  assToPlainText,
  collectAssOverrideCommands,
  parseAssEffectField,
  type AssEffectKind,
  type AssOverrideCommand,
} from './ass-text';
import { mergeDuplicateCues } from './subtitle-cue-dedup';

export interface SubtitleCue {
  startTime: number;
  endTime: number;
  text: string;
}

/**
 * Everything the parser knows about a source event, shared only with the dedup engine.
 * Deduplication needs the authoring context -- which style the line belongs to, which
 * override commands it carries, whether the `Effect` column was set -- to tell a karaoke
 * burst apart from two characters saying the same word in turn. None of it is meaningful
 * outside the parser, so the public API stays `{startTime, endTime, text}`.
 */
export interface AnnotatedSubtitleCue extends SubtitleCue {
  /** Text exactly as authored, override blocks and all. */
  rawText: string;
  style: string;
  layer: number;
  /** ASS `Name`/`Actor` column. */
  name: string;
  /** ASS `Effect` column, verbatim. */
  effect: string;
  effectKind: AssEffectKind;
  /** Override commands found in `{...}` blocks, with their arguments. */
  overrides: readonly AssOverrideCommand[];
  /** Canonical form of `overrides`, for spotting values that change across a run. */
  overrideSignature: string;
  /** Position in the source file, so sorting by time stays deterministic across layers. */
  order: number;
}

export type SubtitleSourceFormat = 'ass' | 'srt';

const HTML_SUBTITLE_TAG_PATTERN = /<\/?[A-Za-z][^>\n]*>/g;

const SRT_TIMING_PATTERN =
  /^\s*(?:(\d{1,2}):)?(\d{2}):(\d{2})[,.](\d{1,3})\s*-->\s*(?:(\d{1,2}):)?(\d{2}):(\d{2})[,.](\d{1,3})/;

function parseTimestamp(
  hours: string | undefined,
  minutes: string,
  seconds: string,
  millis: string,
): number {
  return (
    Number(hours || 0) * 3600 +
    Number(minutes) * 60 +
    Number(seconds) +
    Number(millis.padEnd(3, '0')) / 1000
  );
}

/**
 * The single ASS decode for the file path: cues leave the parser as plain text with real
 * line breaks, matching what mpv hands over for the same line played live. No layer
 * downstream decodes ASS again.
 */
function sanitizeSubtitleCueText(text: string): string {
  return assToPlainText(text, '\n').replace(HTML_SUBTITLE_TAG_PATTERN, '').trim();
}

function toPublicCues(cues: AnnotatedSubtitleCue[]): SubtitleCue[] {
  return cues.map(({ startTime, endTime, text }) => ({ startTime, endTime, text }));
}

function parseAnnotatedSrtCues(content: string): AnnotatedSubtitleCue[] {
  const cues: AnnotatedSubtitleCue[] = [];
  const lines = content.split(/\r?\n/);
  let i = 0;

  while (i < lines.length) {
    const line = lines[i]!;
    const timingMatch = SRT_TIMING_PATTERN.exec(line);
    if (!timingMatch) {
      i += 1;
      continue;
    }

    const startTime = parseTimestamp(
      timingMatch[1],
      timingMatch[2]!,
      timingMatch[3]!,
      timingMatch[4]!,
    );
    const endTime = parseTimestamp(
      timingMatch[5],
      timingMatch[6]!,
      timingMatch[7]!,
      timingMatch[8]!,
    );

    i += 1;
    const textLines: string[] = [];
    while (i < lines.length && lines[i]!.trim() !== '') {
      textLines.push(lines[i]!);
      i += 1;
    }

    const rawText = textLines.join('\n');
    const text = sanitizeSubtitleCueText(rawText);
    if (text) {
      cues.push({
        startTime,
        endTime,
        text,
        rawText,
        style: '',
        layer: 0,
        name: '',
        effect: '',
        effectKind: 'none',
        // SRT and VTT carry no authoring metadata, and the dedup engine never reads
        // overrides for those formats -- collecting them would be parsing for nobody.
        overrides: [],
        overrideSignature: '',
        order: cues.length,
      });
    }
  }

  return cues;
}

export function parseSrtCues(content: string): SubtitleCue[] {
  return toPublicCues(parseAnnotatedSrtCues(content));
}

const ASS_TIMING_PATTERN = /^(\d+):(\d{2}):(\d{2})\.(\d{1,2})$/;
const ASS_FORMAT_PREFIX = 'Format:';
const ASS_DIALOGUE_PREFIX = 'Dialogue:';
const ASS_NAME_FIELD_ALIASES = ['name', 'actor'];

function parseAssTimestamp(raw: string): number | null {
  const match = ASS_TIMING_PATTERN.exec(raw.trim());
  if (!match) {
    return null;
  }
  const hours = Number(match[1]);
  const minutes = Number(match[2]);
  const seconds = Number(match[3]);
  const centiseconds = Number(match[4]!.padEnd(2, '0'));
  return hours * 3600 + minutes * 60 + seconds + centiseconds / 100;
}

function readField(fields: string[], index: number): string {
  return index >= 0 && index < fields.length ? fields[index]!.trim() : '';
}

function findFieldIndex(formatFields: string[], aliases: string[]): number {
  for (const alias of aliases) {
    const index = formatFields.indexOf(alias);
    if (index >= 0) {
      return index;
    }
  }
  return -1;
}

function parseAnnotatedAssCues(content: string): AnnotatedSubtitleCue[] {
  const cues: AnnotatedSubtitleCue[] = [];
  const lines = content.split(/\r?\n/);
  let inEventsSection = false;
  const fieldIndex = {
    start: -1,
    end: -1,
    text: -1,
    style: -1,
    layer: -1,
    name: -1,
    effect: -1,
  };

  const resetFieldIndex = () => {
    fieldIndex.start = -1;
    fieldIndex.end = -1;
    fieldIndex.text = -1;
    fieldIndex.style = -1;
    fieldIndex.layer = -1;
    fieldIndex.name = -1;
    fieldIndex.effect = -1;
  };

  for (const line of lines) {
    const trimmed = line.trim();

    if (trimmed.startsWith('[') && trimmed.endsWith(']')) {
      inEventsSection = trimmed.toLowerCase() === '[events]';
      if (!inEventsSection) {
        resetFieldIndex();
      }
      continue;
    }

    if (!inEventsSection) {
      continue;
    }

    if (trimmed.startsWith(ASS_FORMAT_PREFIX)) {
      const formatFields = trimmed
        .slice(ASS_FORMAT_PREFIX.length)
        .split(',')
        .map((field) => field.trim().toLowerCase());
      fieldIndex.start = formatFields.indexOf('start');
      fieldIndex.end = formatFields.indexOf('end');
      fieldIndex.text = formatFields.indexOf('text');
      fieldIndex.style = formatFields.indexOf('style');
      fieldIndex.layer = formatFields.indexOf('layer');
      // Aegisub writes the speaker column as `Actor`; the v4+ spec calls it `Name`.
      // Missing it costs the burst check its speaker guard, so both spellings count.
      fieldIndex.name = findFieldIndex(formatFields, ASS_NAME_FIELD_ALIASES);
      fieldIndex.effect = formatFields.indexOf('effect');
      continue;
    }

    if (!trimmed.startsWith(ASS_DIALOGUE_PREFIX)) {
      continue;
    }

    if (fieldIndex.start < 0 || fieldIndex.end < 0 || fieldIndex.text < 0) {
      continue;
    }

    const fields = trimmed.slice(ASS_DIALOGUE_PREFIX.length).split(',');
    if (
      fieldIndex.start >= fields.length ||
      fieldIndex.end >= fields.length ||
      fieldIndex.text >= fields.length
    ) {
      continue;
    }

    const startTime = parseAssTimestamp(fields[fieldIndex.start]!);
    const endTime = parseAssTimestamp(fields[fieldIndex.end]!);
    if (startTime === null || endTime === null) {
      continue;
    }

    const rawText = fields.slice(fieldIndex.text).join(',');
    const text = sanitizeSubtitleCueText(rawText);
    if (!text) {
      continue;
    }

    const effect = readField(fields, fieldIndex.effect);
    const layer = Number(readField(fields, fieldIndex.layer));
    const overrides = collectAssOverrideCommands(rawText);
    cues.push({
      startTime,
      endTime,
      text,
      rawText,
      style: readField(fields, fieldIndex.style),
      layer: Number.isFinite(layer) ? layer : 0,
      name: readField(fields, fieldIndex.name),
      effect,
      effectKind: parseAssEffectField(effect),
      overrides,
      overrideSignature: assOverrideSignature(overrides),
      order: cues.length,
    });
  }

  return cues;
}

export function parseAssCues(content: string): SubtitleCue[] {
  return toPublicCues(parseAnnotatedAssCues(content));
}

function detectSubtitleFormat(source: string): 'srt' | 'vtt' | 'ass' | 'ssa' | null {
  const [normalizedSource = source] =
    (() => {
      try {
        return /^[a-z]+:\/\//i.test(source) ? new URL(source).pathname : source;
      } catch {
        return source;
      }
    })().split(/[?#]/, 1)[0] ?? '';
  const ext = normalizedSource.split('.').pop()?.toLowerCase() ?? '';
  if (ext === 'srt') return 'srt';
  if (ext === 'vtt') return 'vtt';
  if (ext === 'ass' || ext === 'ssa') return 'ass';
  return null;
}

export function parseSubtitleCues(content: string, filename: string): SubtitleCue[] {
  const format = detectSubtitleFormat(filename);
  let cues: AnnotatedSubtitleCue[];
  let sourceFormat: SubtitleSourceFormat = 'srt';

  switch (format) {
    case 'srt':
    case 'vtt':
      cues = parseAnnotatedSrtCues(content);
      break;
    case 'ass':
    case 'ssa':
      cues = parseAnnotatedAssCues(content);
      sourceFormat = 'ass';
      break;
    default:
      cues = [];
  }

  if (cues.length === 0) {
    const assCues = parseAnnotatedAssCues(content);
    const srtCues = parseAnnotatedSrtCues(content);
    const preferAss = assCues.length >= srtCues.length;
    cues = preferAss ? assCues : srtCues;
    sourceFormat = preferAss && assCues.length > 0 ? 'ass' : 'srt';
  }

  cues.sort((a, b) => a.startTime - b.startTime || a.endTime - b.endTime || a.order - b.order);
  return toPublicCues(mergeDuplicateCues(cues, sourceFormat));
}
