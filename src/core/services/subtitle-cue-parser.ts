import {
  assOverrideSignature,
  assToPlainText,
  collectAssOverrideCommands,
  parseAssEffectField,
  removeAssControlDebrisLines,
  type AssEffectKind,
  type AssOverrideCommand,
} from './ass-text';
import { hasAssAnimationEvidence, mergeDuplicateCues } from './subtitle-cue-dedup';

export type AssCueLayout =
  | { kind: 'positioned'; sourceOrder: number; y: number }
  | { kind: 'fragment-grid'; sourceOrder: number }
  | { kind: 'source-order'; sourceOrder: number };

export interface SubtitleCue {
  startTime: number;
  endTime: number;
  text: string;
  /** How a complete line was recovered from generated ASS animation events. */
  source?: 'canonical-ass' | 'reconstructed-ass';
  /**
   * Full span of the generated animation events a recovered cue replaced. Entrance and
   * exit frames can run past canonical authored timing.
   */
  animationStartTime?: number;
  animationEndTime?: number;
  /** ASS style retained only for fragment-reconstructed lines. */
  assStyle?: string;
  /** Authored ASS ordering metadata used when flattening simultaneous positioned cues. */
  assLayout?: AssCueLayout;
}

/**
 * Everything the parser knows about a source event, shared only with the dedup engine.
 * Deduplication needs the authoring context -- which style the line belongs to, which
 * override commands it carries, whether the `Effect` column was set -- to tell a karaoke
 * burst apart from two characters saying the same word in turn. None of it is meaningful
 * outside the parser, so the public API exposes only timing, text, and the optional
 * recovery marker used by live subtitle consumers.
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
function decodeSubtitleCueText(text: string): string {
  return assToPlainText(text, '\n').replace(HTML_SUBTITLE_TAG_PATTERN, '');
}

function sanitizeSubtitleCueText(text: string): string {
  return decodeSubtitleCueText(text).trim();
}

function sanitizeAssCueText(text: string): string {
  return removeAssControlDebrisLines(decodeSubtitleCueText(text)).trim();
}

function attachAssLayout<T extends SubtitleCue>(cue: T, assLayout: AssCueLayout | undefined): T {
  if (assLayout) {
    Object.defineProperty(cue, 'assLayout', { value: assLayout, enumerable: false });
  }
  return cue;
}

function toPublicCues(cues: AnnotatedSubtitleCue[]): SubtitleCue[] {
  return cues.map(
    ({
      startTime,
      endTime,
      text,
      source,
      animationStartTime,
      animationEndTime,
      style,
      assLayout,
    }) => {
      const common = {
        startTime,
        endTime,
        text,
      };
      if (source === 'reconstructed-ass') {
        return attachAssLayout(
          {
            ...common,
            source,
            animationStartTime,
            animationEndTime,
            assStyle: style,
          },
          assLayout,
        );
      }
      return attachAssLayout(
        source ? { ...common, source, animationStartTime, animationEndTime } : common,
        assLayout,
      );
    },
  );
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
const ASS_COMMENT_PREFIX = 'Comment:';
const ASS_NAME_FIELD_ALIASES = ['name', 'actor'];
const CANONICAL_MATCH_MARGIN_SECONDS = 1;
const MIN_CANONICAL_ANIMATION_EVENTS = 3;
// A tiny animated fragment can itself be composed from still smaller glyph events. It is
// not enough evidence that the fragment represents an authored line boundary.
const MIN_CANONICAL_DIALOGUE_TEXT_LENGTH = 4;
const MIN_FRAGMENT_LINE_EVENTS = 8;
const MIN_FRAGMENT_LINE_PARTS = 4;
const MAX_FRAGMENT_MEDIAN_LENGTH = 4;
const MAX_FRAGMENT_LINE_TIMING_VARIANCE_SECONDS = 2;
const MAX_FRAGMENT_LINE_VERTICAL_SPAN = 48;

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

interface ParsedAssEvents {
  dialogue: AnnotatedSubtitleCue[];
  comments: AnnotatedSubtitleCue[];
}

// Every candidate line re-reads the compacted text of each event in its window, so on
// fragment-heavy scripts the same event compacts thousands of times without this cache.
const compactMatchTextCache = new WeakMap<AnnotatedSubtitleCue, string>();

function compactAssMatchText(text: string): string {
  return text.replace(/\s+/gu, '');
}

function compactCueMatchText(cue: AnnotatedSubtitleCue): string {
  let compact = compactMatchTextCache.get(cue);
  if (compact === undefined) {
    compact = compactAssMatchText(cue.text);
    compactMatchTextCache.set(cue, compact);
  }
  return compact;
}

function assEventGroupKey(cue: AnnotatedSubtitleCue): string {
  return `${cue.style}\0${cue.name}`;
}

/**
 * Windowed lookup over one style/name group. Every candidate line queries its time
 * neighborhood, and fragment-heavy scripts put thousands of candidates in one group, so
 * a linear rescan per candidate is quadratic in practice. Events are sorted by start
 * once; `prefixMaxEnd` lets the backward walk stop as soon as no earlier event can still
 * reach the window.
 */
interface AssEventGroupIndex {
  byStart: AnnotatedSubtitleCue[];
  prefixMaxEnd: number[];
}

function buildAssEventGroupIndex(events: readonly AnnotatedSubtitleCue[]): AssEventGroupIndex {
  const byStart = [...events].sort((a, b) => a.startTime - b.startTime || a.order - b.order);
  const prefixMaxEnd: number[] = [];
  let maxEnd = -Infinity;
  for (const event of byStart) {
    maxEnd = Math.max(maxEnd, event.endTime);
    prefixMaxEnd.push(maxEnd);
  }
  return { byStart, prefixMaxEnd };
}

/** Group events overlapping `[startTime, endTime]`, returned in source order. */
function eventsOverlappingWindow(
  index: AssEventGroupIndex,
  startTime: number,
  endTime: number,
): AnnotatedSubtitleCue[] {
  const { byStart, prefixMaxEnd } = index;
  let low = 0;
  let high = byStart.length;
  while (low < high) {
    const mid = (low + high) >>> 1;
    if (byStart[mid]!.startTime <= endTime) {
      low = mid + 1;
    } else {
      high = mid;
    }
  }
  const matches: AnnotatedSubtitleCue[] = [];
  for (let i = low - 1; i >= 0 && prefixMaxEnd[i]! >= startTime; i -= 1) {
    if (byStart[i]!.endTime >= startTime) {
      matches.push(byStart[i]!);
    }
  }
  return matches.sort((a, b) => a.order - b.order);
}

interface FragmentGroup {
  text: string;
  events: AnnotatedSubtitleCue[];
}

function fragmentPlacementAnchors(event: AnnotatedSubtitleCue): Set<string> {
  const anchors = new Set<string>();
  for (const command of event.overrides) {
    const name = command.name.toLowerCase();
    const args = command.args.split(',').map((value) => value.trim());
    if (name === 'pos' && args.length >= 2) {
      anchors.add(`pos:${args[0]},${args[1]}`);
    } else if (name === 'move' && args.length >= 4) {
      anchors.add(`move:${args[0]},${args[1]}`);
      anchors.add(`move:${args[2]},${args[3]}`);
    }
  }
  return anchors;
}

function isRepeatedFragmentCopy(
  previous: AnnotatedSubtitleCue,
  current: AnnotatedSubtitleCue,
): boolean {
  const previousAnchors = fragmentPlacementAnchors(previous);
  if ([...fragmentPlacementAnchors(current)].some((anchor) => previousAnchors.has(anchor))) {
    return true;
  }
  return (
    previous.startTime === current.startTime &&
    previous.endTime === current.endTime &&
    previous.overrideSignature === current.overrideSignature
  );
}

function hasRelaxedAssFragmentEvidence(events: readonly AnnotatedSubtitleCue[]): boolean {
  if (events.length < 2 || !events.every((event) => fragmentPlacementAnchors(event).size > 0)) {
    return false;
  }

  const latestStart = events.reduce(
    (latest, event) => Math.max(latest, event.startTime),
    -Infinity,
  );
  const earliestEnd = events.reduce(
    (earliest, event) => Math.min(earliest, event.endTime),
    Infinity,
  );
  if (latestStart >= earliestEnd) {
    return false;
  }

  const first = events[0]!;
  const hasChangingOverrides = events.some(
    (event) => event.overrideSignature !== first.overrideSignature,
  );
  const hasPositionedLayerCopy = events.some((event, index) =>
    events
      .slice(0, index)
      .some(
        (previous) =>
          compactCueMatchText(previous) === compactCueMatchText(event) &&
          isRepeatedFragmentCopy(previous, event),
      ),
  );
  return hasChangingOverrides || hasPositionedLayerCopy;
}

interface AssFragmentPart {
  cue: AnnotatedSubtitleCue;
  text: string;
}

interface AssFragmentPosition {
  x: number;
  y: number;
}

const MIN_LATIN_POSITION_GAP_SAMPLES = 4;
const LATIN_WORD_GAP_RATIO = 1.2;

function fragmentPosition(cue: AnnotatedSubtitleCue): AssFragmentPosition | null {
  for (const command of cue.overrides) {
    if (command.animated) continue;
    const name = command.name.toLowerCase();
    const args = command.args.split(',').map((value) => Number(value.trim()));
    if (
      name === 'pos' &&
      args.length >= 2 &&
      Number.isFinite(args[0]) &&
      Number.isFinite(args[1])
    ) {
      return { x: args[0]!, y: args[1]! };
    }
    if (
      name === 'move' &&
      args.length >= 4 &&
      args.slice(0, 4).every((value) => Number.isFinite(value))
    ) {
      return { x: (args[0]! + args[2]!) / 2, y: (args[1]! + args[3]!) / 2 };
    }
  }
  return null;
}

function latinGlyphWidthWeight(glyph: string): number {
  if (/[ilIjtfr]/u.test(glyph)) return 0.6;
  if (/[mwMW]/u.test(glyph)) return 1.4;
  if (/[A-Z]/u.test(glyph)) return 1.1;
  return 1;
}

function latinFragmentWidthWeight(text: string): number | null {
  if (!/^[A-Za-z0-9'’.,!?;:-]+$/u.test(text)) return null;
  const punctuationWeight = /^[A-Za-z0-9]['’.,!?;:-]$/u.test(text) ? 0.5 : 0.25;
  return [...text].reduce(
    (width, glyph) =>
      width + (/['’.,!?;:-]/u.test(glyph) ? punctuationWeight : latinGlyphWidthWeight(glyph)),
    0,
  );
}

function normalizedLatinFragmentGap(
  previous: AssFragmentPart,
  current: AssFragmentPart,
): number | null {
  const previousWeight = latinFragmentWidthWeight(previous.text);
  const currentWeight = latinFragmentWidthWeight(current.text);
  const previousPosition = fragmentPosition(previous.cue);
  const currentPosition = fragmentPosition(current.cue);
  if (previousWeight === null || currentWeight === null || !previousPosition || !currentPosition) {
    return null;
  }
  const xDistance = currentPosition.x - previousPosition.x;
  const yDistance = Math.abs(currentPosition.y - previousPosition.y);
  if (yDistance <= 2 && xDistance <= 0) return null;

  // A wrapped authored line can return to the left on its next visual row. Preserve
  // that measured row transition as a separator without treating backwards movement
  // on the same row as a word gap.
  const positionDistance = yDistance <= 2 ? xDistance : Math.abs(xDistance) + yDistance;
  return positionDistance / ((previousWeight + currentWeight) / 2);
}

function commonLatinFragmentGap(values: readonly number[]): number {
  const sorted = [...values].sort((left, right) => left - right);
  // Romaji lines contain many short particles, so real word gaps can outnumber
  // within-word transitions. A lower quantile still represents ordinary glyph advance
  // while ignoring the narrowest character pair as an outlier.
  return sorted[Math.floor((sorted.length - 1) * 0.35)]!;
}

/**
 * Character-by-character typesetting often omits literal spaces because the authored
 * word gap exists only in each glyph's `\pos`. Estimate the normal adjacent-glyph
 * advance within that one line, then preserve only materially larger horizontal gaps.
 * Normalizing each gap by the neighboring fragment widths supports both single glyphs
 * and multi-character karaoke syllables without guessing from the text itself.
 */
function joinAssFragmentParts(parts: readonly AssFragmentPart[]): string {
  if (parts.some((part) => /\s/u.test(part.text))) {
    return parts
      .map((part) => part.text)
      .join('')
      .trim();
  }
  const normalizedGaps: number[] = [];
  for (let index = 1; index < parts.length; index += 1) {
    const gap = normalizedLatinFragmentGap(parts[index - 1]!, parts[index]!);
    if (gap !== null) normalizedGaps.push(gap);
  }
  const wordGapThreshold =
    normalizedGaps.length >= MIN_LATIN_POSITION_GAP_SAMPLES
      ? commonLatinFragmentGap(normalizedGaps) * LATIN_WORD_GAP_RATIO
      : Infinity;

  let text = parts[0]?.text ?? '';
  for (let index = 1; index < parts.length; index += 1) {
    const previous = parts[index - 1]!;
    const current = parts[index]!;
    const hasAuthoredSpace = /\s$/u.test(previous.text) || /^\s/u.test(current.text);
    const normalizedGap = normalizedLatinFragmentGap(previous, current);
    const hasPositionedWordGap = normalizedGap !== null && normalizedGap > wordGapThreshold;
    if (!hasAuthoredSpace && hasPositionedWordGap) {
      text += ' ';
    }
    text += current.text;
  }
  return text.trim();
}

function reconstructedAssFragmentLayout(
  parts: readonly AssFragmentPart[],
  owner: AnnotatedSubtitleCue,
): AssCueLayout | undefined {
  let positionedPartCount = 0;
  let minimumY = Infinity;
  let maximumY = -Infinity;
  for (const part of parts) {
    const layout = part.cue.assLayout;
    if (layout?.kind !== 'positioned') continue;
    positionedPartCount += 1;
    minimumY = Math.min(minimumY, layout.y);
    maximumY = Math.max(maximumY, layout.y);
  }

  if (
    positionedPartCount >= MIN_FRAGMENT_LINE_PARTS &&
    maximumY - minimumY > MAX_FRAGMENT_LINE_VERTICAL_SPAN
  ) {
    return { kind: 'fragment-grid', sourceOrder: owner.order };
  }
  return owner.assLayout;
}

interface AssFragmentTimingCluster {
  events: AnnotatedSubtitleCue[];
  minStartTime: number;
  maxStartTime: number;
  minEndTime: number;
  maxEndTime: number;
}

function addToFragmentTimingCluster(
  cluster: AssFragmentTimingCluster,
  cue: AnnotatedSubtitleCue,
): void {
  cluster.events.push(cue);
  cluster.minStartTime = Math.min(cluster.minStartTime, cue.startTime);
  cluster.maxStartTime = Math.max(cluster.maxStartTime, cue.startTime);
  cluster.minEndTime = Math.min(cluster.minEndTime, cue.endTime);
  cluster.maxEndTime = Math.max(cluster.maxEndTime, cue.endTime);
}

function fragmentTimingDistance(
  cluster: AssFragmentTimingCluster,
  cue: AnnotatedSubtitleCue,
): number {
  const nextMinStart = Math.min(cluster.minStartTime, cue.startTime);
  const nextMaxStart = Math.max(cluster.maxStartTime, cue.startTime);
  const nextMinEnd = Math.min(cluster.minEndTime, cue.endTime);
  const nextMaxEnd = Math.max(cluster.maxEndTime, cue.endTime);
  if (
    nextMaxStart - nextMinStart > MAX_FRAGMENT_LINE_TIMING_VARIANCE_SECONDS ||
    nextMaxEnd - nextMinEnd > MAX_FRAGMENT_LINE_TIMING_VARIANCE_SECONDS
  ) {
    return Infinity;
  }
  return (
    Math.abs(cue.startTime - (cluster.minStartTime + cluster.maxStartTime) / 2) +
    Math.abs(cue.endTime - (cluster.minEndTime + cluster.maxEndTime) / 2)
  );
}

function clusterAssFragmentEvents(
  events: readonly AnnotatedSubtitleCue[],
): AssFragmentTimingCluster[] {
  const clusters: AssFragmentTimingCluster[] = [];
  for (const cue of events) {
    let nearest: AssFragmentTimingCluster | null = null;
    let nearestDistance = Infinity;
    for (const cluster of clusters) {
      const distance = fragmentTimingDistance(cluster, cue);
      if (distance < nearestDistance) {
        nearest = cluster;
        nearestDistance = distance;
      }
    }
    if (nearest) {
      addToFragmentTimingCluster(nearest, cue);
    } else {
      clusters.push({
        events: [cue],
        minStartTime: cue.startTime,
        maxStartTime: cue.startTime,
        minEndTime: cue.endTime,
        maxEndTime: cue.endTime,
      });
    }
  }
  return clusters;
}

function decodeSingleAssFragment(cue: AnnotatedSubtitleCue): string | null {
  const visibleLines = decodeSubtitleCueText(cue.rawText)
    .split('\n')
    .filter((line) => line.trim().length > 0);
  return visibleLines.length === 1 ? visibleLines[0]! : null;
}

function reconstructAssFragmentLine(
  events: readonly AnnotatedSubtitleCue[],
): AnnotatedSubtitleCue | null {
  const hasRelaxedEvidence = hasRelaxedAssFragmentEvidence(events);
  const minimumEvents = hasRelaxedEvidence ? 2 : MIN_FRAGMENT_LINE_EVENTS;
  if (events.length < minimumEvents || !hasAssAnimationEvidence(events)) {
    return null;
  }

  const parts: AssFragmentPart[] = [];
  for (const cue of events) {
    const text = decodeSingleAssFragment(cue);
    if (text === null) {
      return null;
    }
    const compactText = compactAssMatchText(text);
    const isLayerCopy = parts.some(
      (part) =>
        compactAssMatchText(part.text) === compactText && isRepeatedFragmentCopy(part.cue, cue),
    );
    if (!isLayerCopy) {
      parts.push({ cue, text });
    }
  }

  const minimumParts = hasRelaxedEvidence ? 1 : MIN_FRAGMENT_LINE_PARTS;
  if (parts.length < minimumParts || (!hasRelaxedEvidence && parts.length === events.length)) {
    return null;
  }
  const lengths = parts
    .map((part) => compactAssMatchText(part.text).length)
    .sort((left, right) => left - right);
  if ((lengths[Math.floor(lengths.length / 2)] ?? Infinity) > MAX_FRAGMENT_MEDIAN_LENGTH) {
    return null;
  }

  const text = joinAssFragmentParts(parts);
  if (!text) {
    return null;
  }
  const owner = parts[0]!.cue;
  const animationStartTime = earliestStartTime(events);
  const animationEndTime = latestEndTime(events);
  return {
    ...owner,
    startTime: animationStartTime,
    endTime: animationEndTime,
    text,
    rawText: text,
    source: 'reconstructed-ass',
    animationStartTime,
    animationEndTime,
    assLayout: reconstructedAssFragmentLayout(parts, owner),
    overrides: [],
    overrideSignature: '',
  };
}

function recoverFragmentOnlyAssLines(dialogue: AnnotatedSubtitleCue[]): AnnotatedSubtitleCue[] {
  const groups = new Map<string, AnnotatedSubtitleCue[]>();
  for (const cue of dialogue) {
    if (cue.source !== undefined) {
      continue;
    }
    const key = assEventGroupKey(cue);
    const group = groups.get(key);
    if (group) {
      group.push(cue);
    } else {
      groups.set(key, [cue]);
    }
  }

  const recovered: AnnotatedSubtitleCue[] = [];
  const suppressed = new Set<AnnotatedSubtitleCue>();
  for (const events of groups.values()) {
    for (const cluster of clusterAssFragmentEvents(events)) {
      const line = reconstructAssFragmentLine(cluster.events);
      if (!line) {
        continue;
      }
      recovered.push(line);
      cluster.events.forEach((event) => suppressed.add(event));
    }
  }
  if (recovered.length === 0) {
    return dialogue;
  }
  return [...dialogue.filter((cue) => !suppressed.has(cue)), ...recovered].sort(
    (left, right) =>
      left.startTime - right.startTime || left.endTime - right.endTime || left.order - right.order,
  );
}

function groupConsecutiveAssFragments(events: readonly AnnotatedSubtitleCue[]): FragmentGroup[] {
  const groups: FragmentGroup[] = [];
  for (const event of events) {
    const text = compactCueMatchText(event);
    if (!text) {
      continue;
    }
    const previous = groups.at(-1);
    if (
      previous?.text === text &&
      previous.events.some((previousEvent) => isRepeatedFragmentCopy(previousEvent, event))
    ) {
      previous.events.push(event);
    } else {
      groups.push({ text, events: [event] });
    }
  }
  return groups;
}

function findCanonicalFragmentEvents(
  events: readonly AnnotatedSubtitleCue[],
  canonicalText: string,
): AnnotatedSubtitleCue[] {
  const groups = groupConsecutiveAssFragments(events);
  const matches = new Set<AnnotatedSubtitleCue>();

  for (let start = 0; start < groups.length; start += 1) {
    let combined = '';
    for (let end = start; end < groups.length; end += 1) {
      const group = groups[end]!;
      // A complete rendered copy cannot prove that the neighboring events are its
      // fragments. Exact full-line animation is handled separately for comments.
      if (group.text.length >= canonicalText.length) {
        break;
      }
      const next = combined + group.text;
      if (!canonicalText.startsWith(next)) {
        break;
      }
      combined = next;
      if (combined !== canonicalText) {
        continue;
      }
      for (let index = start; index <= end; index += 1) {
        for (const event of groups[index]!.events) {
          matches.add(event);
        }
      }
      start = end;
      break;
    }
  }

  return [...matches];
}

function matchingAssAnimationEvents(options: {
  candidate: AnnotatedSubtitleCue;
  group: AssEventGroupIndex;
  allowFullLineFrames: boolean;
}): AnnotatedSubtitleCue[] {
  const canonicalText = compactCueMatchText(options.candidate);
  // The group index already restricts to the candidate's style and name.
  const nearby = eventsOverlappingWindow(
    options.group,
    options.candidate.startTime - CANONICAL_MATCH_MARGIN_SECONDS,
    options.candidate.endTime + CANONICAL_MATCH_MARGIN_SECONDS,
  );
  const fragments = findCanonicalFragmentEvents(nearby, canonicalText);
  if (fragments.length >= MIN_CANONICAL_ANIMATION_EVENTS && hasAssAnimationEvidence(fragments)) {
    return fragments;
  }

  if (!options.allowFullLineFrames) {
    return [];
  }
  const fullLineFrames = nearby.filter((cue) => compactCueMatchText(cue) === canonicalText);
  return fullLineFrames.length >= MIN_CANONICAL_ANIMATION_EVENTS &&
    hasAssAnimationEvidence(fullLineFrames)
    ? fullLineFrames
    : [];
}

// Reductions rather than `Math.min(...events)`: one generated line can carry an
// unbounded number of events, and spreading them all as arguments risks the engine's
// argument-count limit.
function earliestStartTime(events: readonly AnnotatedSubtitleCue[], seed = Infinity): number {
  return events.reduce((earliest, event) => Math.min(earliest, event.startTime), seed);
}

function latestEndTime(events: readonly AnnotatedSubtitleCue[], seed = -Infinity): number {
  return events.reduce((latest, event) => Math.max(latest, event.endTime), seed);
}

function includeCanonicalBoundaryEvents(options: {
  candidate: AnnotatedSubtitleCue;
  group: AssEventGroupIndex;
  animationEvents: readonly AnnotatedSubtitleCue[];
}): AnnotatedSubtitleCue[] {
  const canonicalText = compactCueMatchText(options.candidate);
  const startTime = earliestStartTime(options.animationEvents);
  const endTime = latestEndTime(options.animationEvents);
  return eventsOverlappingWindow(
    options.group,
    startTime - CANONICAL_MATCH_MARGIN_SECONDS,
    endTime + CANONICAL_MATCH_MARGIN_SECONDS,
  ).filter((cue) => compactCueMatchText(cue) === canonicalText);
}

function recoverCanonicalAssEvents({
  dialogue,
  comments,
}: ParsedAssEvents): AnnotatedSubtitleCue[] {
  const recovered: AnnotatedSubtitleCue[] = [];
  const suppressed = new Set<AnnotatedSubtitleCue>();
  // A recovery is only as good as its owning event. When a later candidate proves that
  // an earlier candidate was itself a generated frame of its animation, the earlier
  // recovery is a duplicate of the same authored line and must be withdrawn.
  const recoveredByOwner = new Map<AnnotatedSubtitleCue, AnnotatedSubtitleCue>();
  const withdrawn = new Set<AnnotatedSubtitleCue>();
  const eventsByGroup = new Map<string, AnnotatedSubtitleCue[]>();
  for (const cue of dialogue) {
    const key = assEventGroupKey(cue);
    const group = eventsByGroup.get(key);
    if (group) {
      group.push(cue);
    } else {
      eventsByGroup.set(key, [cue]);
    }
  }
  const indexByGroup = new Map<string, AssEventGroupIndex>();
  for (const [key, events] of eventsByGroup) {
    indexByGroup.set(key, buildAssEventGroupIndex(events));
  }
  const emptyGroupIndex: AssEventGroupIndex = { byStart: [], prefixMaxEnd: [] };
  const candidates = [
    ...comments.map((cue) => ({ cue, kind: 'comment' as const })),
    ...dialogue
      .filter(
        (cue) =>
          compactCueMatchText(cue).length >= MIN_CANONICAL_DIALOGUE_TEXT_LENGTH &&
          hasAssAnimationEvidence([cue]),
      )
      .sort((left, right) => right.text.length - left.text.length || left.order - right.order)
      .map((cue) => ({ cue, kind: 'dialogue' as const })),
  ];

  for (const { cue: candidate, kind } of candidates) {
    if (candidate.endTime <= candidate.startTime || suppressed.has(candidate)) {
      continue;
    }
    const canonicalText = compactCueMatchText(candidate);
    if (!canonicalText) {
      continue;
    }

    const group = indexByGroup.get(assEventGroupKey(candidate)) ?? emptyGroupIndex;
    const animationEvents = matchingAssAnimationEvents({
      candidate,
      group,
      allowFullLineFrames: kind === 'comment',
    });
    if (animationEvents.length === 0) {
      continue;
    }

    const boundaryEvents = includeCanonicalBoundaryEvents({
      candidate,
      group,
      animationEvents,
    });
    const generatedEvents = [...new Set([...animationEvents, ...boundaryEvents])];
    const animationStartTime = earliestStartTime(generatedEvents, candidate.startTime);
    const animationEndTime = latestEndTime(generatedEvents, candidate.endTime);
    const startTime = kind === 'comment' ? candidate.startTime : animationStartTime;
    const endTime = kind === 'comment' ? candidate.endTime : animationEndTime;
    const recoveredCue: AnnotatedSubtitleCue = {
      ...candidate,
      startTime,
      endTime,
      animationStartTime,
      animationEndTime,
      source: 'canonical-ass',
    };
    recovered.push(recoveredCue);
    recoveredByOwner.set(candidate, recoveredCue);
    for (const event of generatedEvents) {
      suppressed.add(event);
      if (event === candidate) {
        continue;
      }
      const priorRecovery = recoveredByOwner.get(event);
      if (priorRecovery) {
        // No text is lost by withdrawing: a fragment claim means the withdrawn line is
        // a contiguous piece of this candidate's text, and a boundary claim means the
        // texts are equal, so the surviving canonical cue always contains it.
        withdrawn.add(priorRecovery);
      }
    }
  }

  const survivingRecovered = recovered.filter((cue) => !withdrawn.has(cue));
  if (survivingRecovered.length === 0) {
    return dialogue;
  }
  return [...dialogue.filter((cue) => !suppressed.has(cue)), ...survivingRecovered].sort(
    (a, b) => a.startTime - b.startTime || a.endTime - b.endTime || a.order - b.order,
  );
}

function parseAssCoordinate(value: string | undefined): number | null {
  if (!value?.trim()) return null;
  const coordinate = Number(value.trim());
  return Number.isFinite(coordinate) ? coordinate : null;
}

function buildAssCueLayout(
  overrides: readonly AssOverrideCommand[],
  sourceOrder: number,
): AssCueLayout {
  let y: number | null = null;
  for (const command of overrides) {
    if (command.animated) continue;
    const name = command.name.toLowerCase();
    const args = command.args.split(',');
    if (name === 'pos') {
      y = parseAssCoordinate(args[1]) ?? y;
      continue;
    }
    if (name !== 'move') continue;
    const startY = parseAssCoordinate(args[1]);
    const endY = parseAssCoordinate(args[3]);
    if (startY !== null && endY !== null) {
      y = (startY + endY) / 2;
    }
  }
  return y === null
    ? { kind: 'source-order', sourceOrder }
    : { kind: 'positioned', sourceOrder, y };
}

function parseAnnotatedAssEvents(content: string): ParsedAssEvents {
  const cues: AnnotatedSubtitleCue[] = [];
  const comments: AnnotatedSubtitleCue[] = [];
  const lines = content.split(/\r?\n/);
  let inEventsSection = false;
  let eventOrder = 0;
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
    // Event text can end in an authored space. Fragmented karaoke commonly uses that
    // space to retain word boundaries when its separately positioned events are joined
    // back into a line, so only remove indentation before slicing the event fields.
    const eventLine = line.trimStart();

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

    const eventPrefix = eventLine.startsWith(ASS_DIALOGUE_PREFIX)
      ? ASS_DIALOGUE_PREFIX
      : eventLine.startsWith(ASS_COMMENT_PREFIX)
        ? ASS_COMMENT_PREFIX
        : null;
    if (!eventPrefix) {
      continue;
    }

    if (fieldIndex.start < 0 || fieldIndex.end < 0 || fieldIndex.text < 0) {
      continue;
    }

    const fields = eventLine.slice(eventPrefix.length).split(',');
    if (
      fieldIndex.start >= fields.length ||
      fieldIndex.end >= fields.length ||
      fieldIndex.text >= fields.length
    ) {
      continue;
    }

    const startTime = parseAssTimestamp(fields[fieldIndex.start]!);
    const endTime = parseAssTimestamp(fields[fieldIndex.end]!);
    if (startTime === null || endTime === null || endTime <= startTime) {
      continue;
    }

    const rawText = fields.slice(fieldIndex.text).join(',');
    const text = sanitizeAssCueText(rawText);
    if (!text) {
      continue;
    }

    const effect = readField(fields, fieldIndex.effect);
    const layer = Number(readField(fields, fieldIndex.layer));
    const overrides = collectAssOverrideCommands(rawText);
    const cue: AnnotatedSubtitleCue = {
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
      order: eventOrder,
      assLayout: buildAssCueLayout(overrides, eventOrder),
    };
    eventOrder += 1;
    if (eventPrefix === ASS_COMMENT_PREFIX) {
      comments.push(cue);
    } else {
      cues.push(cue);
    }
  }

  return { dialogue: cues, comments };
}

function parseAnnotatedAssCues(content: string): AnnotatedSubtitleCue[] {
  return recoverFragmentOnlyAssLines(recoverCanonicalAssEvents(parseAnnotatedAssEvents(content)));
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
