import {
  assOverrideSignature,
  assToPlainText,
  collectAssOverrideCommands,
  hasAssTemporalOverride,
  isAnimatedAssEffectKind,
  parseAssEffectField,
  type AssEffectKind,
  type AssOverrideCommand,
} from './ass-text';

export interface SubtitleCue {
  startTime: number;
  endTime: number;
  text: string;
}

/**
 * Everything the parser knows about a source event, kept private to this module.
 * Deduplication needs the authoring context -- which style the line belongs to, which
 * override commands it carries, whether the `Effect` column was set -- to tell a karaoke
 * burst apart from two characters saying the same word in turn. None of it is meaningful
 * outside the parser, so the public API stays `{startTime, endTime, text}`.
 */
interface AnnotatedSubtitleCue extends SubtitleCue {
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

type SubtitleSourceFormat = 'ass' | 'srt';

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
      const overrides = collectAssOverrideCommands(rawText);
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
        overrides,
        overrideSignature: assOverrideSignature(overrides),
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

// Back-to-back frames of the same animation are authored flush against each other; a
// tiny tolerance absorbs the centisecond rounding of the ASS timestamp format.
const DUPLICATE_CUE_GAP_TOLERANCE_SECONDS = 0.05;
// A burst is a *sequence*. Two adjacent events are two events, not an animation --
// characters do repeat each other, and a repeated line can legitimately be short.
const MIN_BURST_EVENTS = 3;
// Real dialogue holds on screen for about a second, so a run with a couple of much
// shorter events among them looks like frames. Used only alongside authoring evidence.
const ANIMATION_FRAME_MAX_SECONDS = 0.3;
// A karaoke run usually ends on a long "hold" frame, so not every event is short.
const MIN_TAGGED_BURST_FRAMES = 2;
// SRT and VTT carry no authoring metadata at all, so timing is the only signal available
// -- which makes it the easiest one to get wrong. ASS->SRT conversion leaves frames at
// ~0.04s, well under any real utterance, and a burst leaves many of them behind. Both
// bounds are deliberately far stricter than the ASS path: a run of ordinary short lines
// (`えっ` traded between characters) must not clear them.
const TIMING_ONLY_FRAME_MAX_SECONDS = 0.1;
const MIN_TIMING_ONLY_FRAMES = 5;

function cueKey(cue: SubtitleCue): string {
  return `${cue.startTime}|${cue.endTime}|${cue.text}`;
}

/**
 * Identical text over an identical span is redundant however it was authored -- most
 * often a layered ASS event stacking a shadow copy under the visible one.
 */
function collapseExactDuplicates(cues: AnnotatedSubtitleCue[]): AnnotatedSubtitleCue[] {
  const seen = new Set<string>();
  return cues.filter((cue) => {
    const key = cueKey(cue);
    if (seen.has(key)) {
      return false;
    }
    seen.add(key);
    return true;
  });
}

function countFramesShorterThan(run: AnnotatedSubtitleCue[], maxSeconds: number): number {
  return run.filter((cue) => cue.endTime - cue.startTime < maxSeconds).length;
}

/**
 * Evidence that a run of ASS events is one animation rather than several authored lines.
 * A static tag says nothing on its own -- three events sharing one `\clip(...)` are three
 * signs -- so the tag has to be temporal by nature (`\t`, `\move`, karaoke timing, or
 * anything wrapped in `\t(...)`), an animated `Effect` column, or a value that actually
 * changes from event to event, which is how per-frame typesetting is authored.
 */
function hasAssAnimationEvidence(run: AnnotatedSubtitleCue[]): boolean {
  if (run.every((cue) => hasAssTemporalOverride(cue.overrides))) {
    return true;
  }
  if (run.every((cue) => isAnimatedAssEffectKind(cue.effectKind))) {
    return true;
  }

  const [first] = run;
  const everyEventTypeset = run.every((cue) => cue.overrides.length > 0);
  const signatureChanges = run.some((cue) => cue.overrideSignature !== first!.overrideSignature);
  return everyEventTypeset && signatureChanges;
}

function isAnimationBurst(run: AnnotatedSubtitleCue[], format: SubtitleSourceFormat): boolean {
  if (run.length < MIN_BURST_EVENTS) {
    return false;
  }

  if (format === 'srt') {
    return (
      run.length >= MIN_TIMING_ONLY_FRAMES &&
      countFramesShorterThan(run, TIMING_ONLY_FRAME_MAX_SECONDS) === run.length
    );
  }

  if (countFramesShorterThan(run, ANIMATION_FRAME_MAX_SECONDS) < MIN_TAGGED_BURST_FRAMES) {
    return false;
  }

  // One animation belongs to one styled, one named source line. Two characters trading
  // the same short word are two styles or two actors, and never merge.
  const [first] = run;
  if (run.some((cue) => cue.style !== first!.style || cue.name !== first!.name)) {
    return false;
  }

  return hasAssAnimationEvidence(run);
}

/**
 * Karaoke and sign typesetting emits one Dialogue event per animation frame, all carrying
 * the same visible text over a contiguous span. Collapse each such run into a single cue.
 *
 * Only runs that look like animation collapse. Two ordinary lines that happen to repeat
 * -- several characters each saying `おはよう` in turn, a positioned sign redrawn with a
 * different fade -- stay separate, because merging them would destroy real mineable lines.
 */
function collapseAnimationBursts(
  cues: AnnotatedSubtitleCue[],
  format: SubtitleSourceFormat,
): AnnotatedSubtitleCue[] {
  const indicesByText = new Map<string, number[]>();
  cues.forEach((cue, index) => {
    const bucket = indicesByText.get(cue.text);
    if (bucket) {
      bucket.push(index);
    } else {
      indicesByText.set(cue.text, [index]);
    }
  });

  const dropped = new Set<number>();
  const extendedEnd = new Map<number, number>();

  for (const indices of indicesByText.values()) {
    if (indices.length < MIN_BURST_EVENTS) {
      continue;
    }

    let runStart = 0;
    while (runStart < indices.length) {
      let runEnd = runStart;
      let chainEnd = cues[indices[runStart]!]!.endTime;

      while (runEnd + 1 < indices.length) {
        const next = cues[indices[runEnd + 1]!]!;
        if (next.startTime > chainEnd + DUPLICATE_CUE_GAP_TOLERANCE_SECONDS) {
          break;
        }
        chainEnd = Math.max(chainEnd, next.endTime);
        runEnd += 1;
      }

      const run = indices.slice(runStart, runEnd + 1).map((index) => cues[index]!);
      if (isAnimationBurst(run, format)) {
        for (let i = runStart + 1; i <= runEnd; i += 1) {
          dropped.add(indices[i]!);
        }
        extendedEnd.set(indices[runStart]!, chainEnd);
      }

      runStart = runEnd + 1;
    }
  }

  if (dropped.size === 0) {
    return cues;
  }

  const merged: AnnotatedSubtitleCue[] = [];
  cues.forEach((cue, index) => {
    if (dropped.has(index)) {
      return;
    }
    const end = extendedEnd.get(index);
    merged.push(end !== undefined && end > cue.endTime ? { ...cue, endTime: end } : cue);
  });
  return merged;
}

function mergeDuplicateCues(
  cues: AnnotatedSubtitleCue[],
  format: SubtitleSourceFormat,
): AnnotatedSubtitleCue[] {
  return collapseAnimationBursts(collapseExactDuplicates(cues), format);
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
