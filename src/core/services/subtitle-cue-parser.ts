import {
  assOverrideSignature,
  assToPlainText,
  collectAssOverrideCommands,
  hasAssTemporalOverride,
  parseAssEffectField,
  removeAssControlDebrisLines,
  type AssEffectKind,
  type AssOverrideCommand,
} from './ass-text';
import { hasAssAnimationEvidence, mergeDuplicateCues } from './subtitle-cue-dedup';

/** Vertical third of the screen a cue is authored to occupy. */
export type AssVerticalBand = 'top' | 'middle' | 'bottom';

export type AssCueLayout =
  | { kind: 'positioned'; sourceOrder: number; y: number; verticalBand?: AssVerticalBand }
  | { kind: 'fragment-grid'; sourceOrder: number; verticalBand?: AssVerticalBand }
  | { kind: 'source-order'; sourceOrder: number; verticalBand?: AssVerticalBand };

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

// Anchor extraction walks every override command, and the canonical/fragment passes
// consult the same events once per candidate neighborhood, so heavy KFX tags make the
// uncached form quadratic in practice.
const placementAnchorCache = new WeakMap<AnnotatedSubtitleCue, Set<string>>();

function fragmentPlacementAnchors(event: AnnotatedSubtitleCue): Set<string> {
  let anchors = placementAnchorCache.get(event);
  if (anchors) {
    return anchors;
  }
  anchors = new Set<string>();
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
  placementAnchorCache.set(event, anchors);
  return anchors;
}

// Drop-shadow layer copies sit a few pixels off their base glyph, while even tightly
// kerned repeated glyphs in one line ("ii") measure 10px apart or more.
const LAYER_COPY_OFFSET_TOLERANCE_PX = 6;

// Copies chain only through genuine time overlap. Consecutive re-runs of one visual
// (chant bursts, countdown frames, jitter animation frames) abut or micro-overlap at
// frame seams, so the different-structure threshold sits above a frame seam while the
// phases of one effect (a highlight and the exit ghost it launches) overlap far longer.
const MIN_PHASE_OVERLAP_SECONDS = 0.04;

// Decoration is timed to the line it accompanies, but a lead-in echo can end exactly
// where the recovered line's first sung copy begins; a small tolerance keeps such
// flush decoration attached to its line.
const DECORATION_SPAN_TOLERANCE_SECONDS = 0.1;

// Guards for pathological event volumes. Real copy stacks and repaint chains stay in
// the hundreds; a same-text bucket or candidate sweep group in the thousands is a
// particle field, and the quadratic passes over it would stall the main process.
const MAX_COALESCE_BUCKET_EVENTS = 1500;
const MAX_SWEEP_GROUP_EVENTS = 4000;
// Fragment-layer collapsing repeatedly rescans the remaining parts after each match.
// Genuine authored lines stay far below this limit; larger groups are particle fields.
const MAX_FRAGMENT_COLLAPSE_PARTS = 1500;

/**
 * Sources behind a coalesced copy chain. Synthetic cues stand in for their sources
 * during fragment recovery, but suppression and copy-count evidence must reach the
 * original events, which are what the published dialogue list still holds.
 */
const coalescedSourceEvents = new WeakMap<AnnotatedSubtitleCue, AnnotatedSubtitleCue[]>();

function sourceEventsOf(cue: AnnotatedSubtitleCue): readonly AnnotatedSubtitleCue[] {
  return coalescedSourceEvents.get(cue) ?? [cue];
}

function sourceEventCount(events: readonly AnnotatedSubtitleCue[]): number {
  return events.reduce((count, event) => count + sourceEventsOf(event).length, 0);
}

// One representative point per placement command: the `\pos` point or the `\move`
// midpoint. Comparing raw `\move` endpoints cross-wise misreads a travel distance that
// matches the glyph advance as a layer copy of a neighboring same-letter glyph.
const anchorPointCache = new WeakMap<AnnotatedSubtitleCue, AssFragmentPosition[]>();

function fragmentAnchorPoints(event: AnnotatedSubtitleCue): AssFragmentPosition[] {
  let points = anchorPointCache.get(event);
  if (points) {
    return points;
  }
  points = [];
  for (const command of event.overrides) {
    const name = command.name.toLowerCase();
    const args = command.args.split(',').map((value) => Number(value.trim()));
    if (name === 'pos' && args.length >= 2 && args.slice(0, 2).every(Number.isFinite)) {
      points.push({ x: args[0]!, y: args[1]! });
    } else if (name === 'move' && args.length >= 4 && args.slice(0, 4).every(Number.isFinite)) {
      points.push({ x: (args[0]! + args[2]!) / 2, y: (args[1]! + args[3]!) / 2 });
    }
  }
  anchorPointCache.set(event, points);
  return points;
}

function isRepeatedFragmentCopy(
  previous: AnnotatedSubtitleCue,
  current: AnnotatedSubtitleCue,
): boolean {
  const previousAnchors = fragmentPlacementAnchors(previous);
  if ([...fragmentPlacementAnchors(current)].some((anchor) => previousAnchors.has(anchor))) {
    return true;
  }
  const previousPoints = fragmentAnchorPoints(previous);
  const nearbyAnchor = fragmentAnchorPoints(current).some((point) =>
    previousPoints.some(
      (previousPoint) =>
        Math.abs(point.x - previousPoint.x) <= LAYER_COPY_OFFSET_TOLERANCE_PX &&
        Math.abs(point.y - previousPoint.y) <= LAYER_COPY_OFFSET_TOLERANCE_PX,
    ),
  );
  if (nearbyAnchor) {
    return true;
  }
  return (
    previous.startTime === current.startTime &&
    previous.endTime === current.endTime &&
    previous.overrideSignature === current.overrideSignature
  );
}

// Coalescing keys on where a copy is anchored: the `\pos` point and both `\move`
// endpoints, since exit ghosts launch from the glyph anchor and entrance copies
// converge onto it.
function coalesceAnchorPoints(event: AnnotatedSubtitleCue): AssFragmentPosition[] {
  const points: AssFragmentPosition[] = [];
  for (const command of event.overrides) {
    const name = command.name.toLowerCase();
    const args = command.args.split(',').map((value) => Number(value.trim()));
    if (name === 'pos' && args.length >= 2 && args.slice(0, 2).every(Number.isFinite)) {
      points.push({ x: args[0]!, y: args[1]! });
    } else if (name === 'move' && args.length >= 4 && args.slice(0, 4).every(Number.isFinite)) {
      points.push({ x: args[0]!, y: args[1]! });
      points.push({ x: args[2]!, y: args[3]! });
    }
  }
  return points;
}

function shareCoalesceAnchor(
  left: readonly AssFragmentPosition[],
  right: readonly AssFragmentPosition[],
): boolean {
  return right.some((point) =>
    left.some(
      (other) =>
        Math.abs(point.x - other.x) <= LAYER_COPY_OFFSET_TOLERANCE_PX &&
        Math.abs(point.y - other.y) <= LAYER_COPY_OFFSET_TOLERANCE_PX,
    ),
  );
}

// The set of distinct command names, ignoring arguments and repetition. Two phases of
// one effect (pre-echo, highlight, hold) carry different command vocabularies; a re-run
// of the same visual (chant burst, countdown frame) repeats the same vocabulary with new
// argument values -- including a different number of animation keyframes, which is why
// repetition must not count.
const structuralSignatureCache = new WeakMap<AnnotatedSubtitleCue, string>();

function structuralOverrideSignature(cue: AnnotatedSubtitleCue): string {
  let signature = structuralSignatureCache.get(cue);
  if (signature === undefined) {
    const names = new Set(
      cue.overrides.map((command) => `${command.animated ? '~' : ''}${command.name.toLowerCase()}`),
    );
    signature = [...names].sort().join(',');
    structuralSignatureCache.set(cue, signature);
  }
  return signature;
}

// A transparent lead-in ends exactly where its glyph's first sung copy begins, so the
// echo-to-visible handoff must chain across a small seam.
const COALESCE_SEAM_TOLERANCE_SECONDS = 0.05;

/**
 * Two same-text copies chain when their windows genuinely overlap. Structurally
 * identical events are layers of one visual exactly when their windows substantially
 * coincide; a frame seam or few-millisecond overlap between identical structures is a
 * re-run (the next chant burst, the next countdown frame). Structurally different
 * events are phases of one effect and chain across any above-seam overlap. A bare seam
 * only joins a transparent echo to its visible phase: the lead-in before a highlight
 * belongs to its glyph, while two visible events that merely abut (the next burst of a
 * chant, an exit flash after a hold) are separate showings.
 */
function areTimeConnectedCopies(a: AnnotatedSubtitleCue, b: AnnotatedSubtitleCue): boolean {
  const overlap = Math.min(a.endTime, b.endTime) - Math.max(a.startTime, b.startTime);
  if (structuralOverrideSignature(a) === structuralOverrideSignature(b)) {
    const shorterDuration = Math.min(a.endTime - a.startTime, b.endTime - b.startTime);
    return overlap >= Math.max(MIN_PHASE_OVERLAP_SECONDS, shorterDuration / 2);
  }
  if (overlap >= MIN_PHASE_OVERLAP_SECONDS) {
    return true;
  }
  return (
    overlap >= -COALESCE_SEAM_TOLERANCE_SECONDS &&
    isTransparentFillEcho(a) !== isTransparentFillEcho(b)
  );
}

function buildCoalescedCopy(members: readonly AnnotatedSubtitleCue[]): AnnotatedSubtitleCue {
  const ordered = [...members].sort((left, right) => left.order - right.order);
  const representative =
    ordered.find((member) =>
      member.overrides.some((command) => !command.animated && command.name.toLowerCase() === 'pos'),
    ) ?? ordered[0]!;
  const synthetic: AnnotatedSubtitleCue = {
    ...representative,
    startTime: earliestStartTime(ordered),
    endTime: latestEndTime(ordered),
    order: ordered[0]!.order,
  };
  coalescedSourceEvents.set(synthetic, ordered);
  return synthetic;
}

/**
 * Generated glyph effects render one authored glyph as a stack of copies anchored at the
 * same point: a transparent pre-echo until the syllable is sung, a short highlight, an
 * exit ghost launched from the anchor, and a steady hold to the end of the line. Their
 * windows abut rather than coincide, so per-event timing says four unrelated fragments
 * while the anchor says one glyph. Merging each stack into a single presence spanning
 * the union window lets timing clusters see the authored line instead of its phases.
 */
function coalesceAssAnchorCopies(events: readonly AnnotatedSubtitleCue[]): AnnotatedSubtitleCue[] {
  const buckets = new Map<string, number[]>();
  const anchorPoints: (AssFragmentPosition[] | null)[] = events.map(() => null);
  events.forEach((event, index) => {
    const text = compactCueMatchText(event);
    if (!text) {
      return;
    }
    const points = coalesceAnchorPoints(event);
    if (points.length === 0) {
      return;
    }
    anchorPoints[index] = points;
    const bucket = buckets.get(text);
    if (bucket) {
      bucket.push(index);
    } else {
      buckets.set(text, [index]);
    }
  });
  for (const [text, bucket] of buckets) {
    if (bucket.length > MAX_COALESCE_BUCKET_EVENTS) {
      // Pathological same-text volume (a whole-episode particle field). Pairing would
      // stall the main process; uncoalesced events fall back to the burst/grid paths.
      buckets.delete(text);
    } else {
      bucket.sort((left, right) => events[left]!.startTime - events[right]!.startTime);
    }
  }

  const parent = events.map((_, index) => index);
  const findRoot = (index: number): number => {
    let root = index;
    while (parent[root] !== root) {
      root = parent[root]!;
    }
    while (parent[index] !== root) {
      const next = parent[index]!;
      parent[index] = root;
      index = next;
    }
    return root;
  };
  const union = (left: number, right: number): void => {
    parent[findRoot(left)] = findRoot(right);
  };

  for (const bucket of buckets.values()) {
    // Buckets are start-sorted; copies can only connect through time proximity, so the
    // backward scan stops once no earlier copy's window can still reach this one.
    const prefixMaxEnd: number[] = [];
    let maxEnd = -Infinity;
    for (const index of bucket) {
      maxEnd = Math.max(maxEnd, events[index]!.endTime);
      prefixMaxEnd.push(maxEnd);
    }
    for (let i = 1; i < bucket.length; i += 1) {
      const right = bucket[i]!;
      const reachableStart = events[right]!.startTime - COALESCE_SEAM_TOLERANCE_SECONDS;
      for (let j = i - 1; j >= 0 && prefixMaxEnd[j]! >= reachableStart; j -= 1) {
        const left = bucket[j]!;
        if (
          shareCoalesceAnchor(anchorPoints[left]!, anchorPoints[right]!) &&
          areTimeConnectedCopies(events[left]!, events[right]!)
        ) {
          union(left, right);
        }
      }
    }
  }

  const componentsByRoot = new Map<number, AnnotatedSubtitleCue[]>();
  events.forEach((event, index) => {
    const root = findRoot(index);
    const members = componentsByRoot.get(root);
    if (members) {
      members.push(event);
    } else {
      componentsByRoot.set(root, [event]);
    }
  });
  if (componentsByRoot.size === events.length) {
    return [...events];
  }
  return [...componentsByRoot.values()]
    .map((members) => (members.length === 1 ? members[0]! : buildCoalescedCopy(members)))
    .sort((left, right) => left.order - right.order);
}

function hasRelaxedAssFragmentEvidence(events: readonly AnnotatedSubtitleCue[]): boolean {
  if (
    sourceEventCount(events) < 2 ||
    !events.every((event) => fragmentPlacementAnchors(event).size > 0)
  ) {
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
  const hasPositionedLayerCopy =
    events.some((event) => sourceEventsOf(event).length > 1) ||
    events.some((event, index) =>
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
const LATIN_FRAGMENT_WORD_GAP_RATIO = 1.16;
const LATIN_GLYPH_WORD_GAP_RATIO = 1.4;
// Word-space advance beyond the width-predicted glyph advance, as a fraction of the
// line's common unit. Measured corpus extremes: widest within-word excess 0.32 (`pp`
// with tracking), narrowest word gap 0.40 (`s w` across a wide glyph). That margin only
// holds when the common unit is estimated from enough glyph pairs; a short single-word
// line (`Swelling`) skews the unit low and its ordinary advances read as word gaps.
const LATIN_GLYPH_WORD_EXCESS_RATIO = 0.36;
const MIN_LATIN_GLYPH_EXCESS_GAP_SAMPLES = 10;
// Multi-character syllable chunks average out proportional-font variation, so their
// advances track the width model far more closely than single glyphs do. Measured on a
// chunked lyric line, within-word excess stayed under 0.07 of the common unit while every
// word gap cleared 0.31, so a tighter margin separates them without splitting words.
const LATIN_CHUNK_WORD_EXCESS_RATIO = 0.2;
const MIN_LATIN_CHUNK_EXCESS_GAP_SAMPLES = 6;
const LATIN_TWO_GLYPH_WORD_NEXT_GAP_RATIO = 1.2;

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
  if (/[ilIj]/u.test(glyph)) return 0.6;
  if (/[tfr]/u.test(glyph)) return 0.8;
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

function isSingleLatinGlyphFragment(text: string): boolean {
  return [...text].filter((glyph) => /[A-Za-z0-9]/u.test(glyph)).length <= 1;
}

interface LatinFragmentGapMeasure {
  distance: number;
  meanWeight: number;
}

function latinFragmentGapMeasure(
  previous: AssFragmentPart,
  current: AssFragmentPart,
): LatinFragmentGapMeasure | null {
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
  const distance = yDistance <= 2 ? xDistance : Math.abs(xDistance) + yDistance;
  return { distance, meanWeight: (previousWeight + currentWeight) / 2 };
}

function normalizedLatinFragmentGap(
  previous: AssFragmentPart,
  current: AssFragmentPart,
): number | null {
  const measure = latinFragmentGapMeasure(previous, current);
  return measure === null ? null : measure.distance / measure.meanWeight;
}

function startsNewPositionedFragmentSequence(
  previous: AssFragmentPart,
  current: AssFragmentPart,
): boolean {
  const previousPosition = fragmentPosition(previous.cue);
  const currentPosition = fragmentPosition(current.cue);
  return Boolean(
    previousPosition &&
    currentPosition &&
    Math.abs(currentPosition.y - previousPosition.y) <= 2 &&
    currentPosition.x <= previousPosition.x &&
    current.cue.startTime > previous.cue.startTime,
  );
}

function commonLatinFragmentGap(values: readonly number[]): number {
  const sorted = [...values].sort((left, right) => left - right);
  // Romaji lines contain many short particles, so real word gaps can outnumber
  // within-word transitions. A lower quantile still represents ordinary glyph advance
  // while ignoring the narrowest character pair as an outlier.
  return sorted[Math.floor((sorted.length - 1) * 0.35)]!;
}

function isLikelyTwoGlyphCapitalizedWord(options: {
  parts: readonly AssFragmentPart[];
  index: number;
  gap: number;
  wordGapThreshold: number;
}): boolean {
  const first = options.parts[options.index - 1]!;
  const second = options.parts[options.index]!;
  if (!/^[A-Z]$/u.test(first.text) || !/^[a-z]$/u.test(second.text)) {
    return false;
  }

  const precedingGap =
    options.index > 1 ? normalizedLatinFragmentGap(options.parts[options.index - 2]!, first) : null;
  const following = options.parts[options.index + 1];
  const followingGap = following ? normalizedLatinFragmentGap(second, following) : null;
  const startsAtWordBoundary =
    options.index === 1 || (precedingGap !== null && precedingGap > options.wordGapThreshold);

  return (
    startsAtWordBoundary &&
    followingGap !== null &&
    followingGap > options.wordGapThreshold &&
    followingGap > options.gap * LATIN_TWO_GLYPH_WORD_NEXT_GAP_RATIO
  );
}

/**
 * Character-by-character typesetting often omits literal spaces because the authored
 * word gap exists only in each glyph's `\pos`. Estimate the normal adjacent-glyph
 * advance within that one line, then preserve only materially larger horizontal gaps.
 * Normalizing each gap by the neighboring fragment widths supports both single glyphs
 * and multi-character karaoke syllables without guessing from the text itself. Per-glyph
 * runs use a wider safety margin because proportional fonts vary more than syllable chunks.
 *
 * The ratio test alone under-detects a word gap next to a wide fragment (`waves within`
 * measured across `s`/`w`, or `Choices presumably` across two three-letter chunks, both
 * normalize to nearly a common advance), so a gap also counts as a word boundary when its
 * advance exceeds the width-predicted advance by a material fraction of the line's common
 * unit -- a word space adds a roughly constant extra distance no matter how wide its
 * neighbors are. Chunk runs use a tighter margin than per-glyph runs because their
 * advances deviate less from the width model.
 *
 * A line may mix both conventions: one fragment carrying a literal space while its
 * neighbors rely on position alone. Whitespace-bearing fragments have no width weight, so
 * they drop out of the estimate and their own boundary comes from the authored space,
 * leaving the surrounding positional gaps to be recovered normally.
 */
function joinAssFragmentParts(parts: readonly AssFragmentPart[]): string {
  const normalizedGaps: number[] = [];
  for (let index = 1; index < parts.length; index += 1) {
    const gap = normalizedLatinFragmentGap(parts[index - 1]!, parts[index]!);
    if (gap !== null) normalizedGaps.push(gap);
  }
  const isGlyphRun = parts.every((part) => isSingleLatinGlyphFragment(part.text));
  const commonGap =
    normalizedGaps.length >= MIN_LATIN_POSITION_GAP_SAMPLES
      ? commonLatinFragmentGap(normalizedGaps)
      : null;
  const wordGapThreshold =
    commonGap === null
      ? Infinity
      : commonGap * (isGlyphRun ? LATIN_GLYPH_WORD_GAP_RATIO : LATIN_FRAGMENT_WORD_GAP_RATIO);

  let text = parts[0]?.text ?? '';
  for (let index = 1; index < parts.length; index += 1) {
    const previous = parts[index - 1]!;
    const current = parts[index]!;
    const hasAuthoredSpace = /\s$/u.test(previous.text) || /^\s/u.test(current.text);
    const measure = latinFragmentGapMeasure(previous, current);
    const normalizedGap = measure === null ? null : measure.distance / measure.meanWeight;
    // A capital into lowercase is almost always a capitalized word's own first letters
    // (`S|miles`), and capitals overrun the width table too easily, so the excess rule
    // never fires there. A lone capital word like `I` is narrow enough for the ratio
    // test to catch its word gap on its own.
    const excessRatio = isGlyphRun ? LATIN_GLYPH_WORD_EXCESS_RATIO : LATIN_CHUNK_WORD_EXCESS_RATIO;
    const minimumExcessSamples = isGlyphRun
      ? MIN_LATIN_GLYPH_EXCESS_GAP_SAMPLES
      : MIN_LATIN_CHUNK_EXCESS_GAP_SAMPLES;
    const hasAdvanceExcess =
      commonGap !== null &&
      normalizedGaps.length >= minimumExcessSamples &&
      measure !== null &&
      !(/^[A-Z]$/u.test(previous.text) && /^[a-z]$/u.test(current.text)) &&
      measure.distance - measure.meanWeight * commonGap > excessRatio * commonGap;
    const hasPositionedWordGap =
      startsNewPositionedFragmentSequence(previous, current) ||
      (normalizedGap !== null &&
        (normalizedGap > wordGapThreshold || hasAdvanceExcess) &&
        !isLikelyTwoGlyphCapitalizedWord({
          parts,
          index,
          gap: normalizedGap,
          wordGapThreshold,
        }));
    if (!hasAuthoredSpace && hasPositionedWordGap) {
      text += ' ';
    }
    text += current.text;
  }
  return text.trim();
}

// A tall multi-part layout is only a visual grid when its parts read like tiling
// rather than prose: a couple of texts repeated across many fragments (sign walls),
// the same text re-shown at the same spot over time (countdown/animation frames),
// nothing but scattered single glyphs, or cells aligned into table columns. Wrapped
// lyric rows with repeated karaoke syllables and CC-style dialogue blocks (speaker
// labels plus a sentence) share the same tall geometry but stay publishable.
function looksLikeFragmentGridParts(parts: readonly AssFragmentPart[]): boolean {
  const positioned = parts
    .map((part) => ({
      text: part.text.trim(),
      layout: part.cue.assLayout,
      position: fragmentPosition(part.cue),
      startTime: part.cue.startTime,
    }))
    .filter((part) => part.text && part.layout?.kind === 'positioned');
  if (positioned.length === 0) return true;

  const uniqueTexts = new Set(positioned.map((part) => part.text));
  if (uniqueTexts.size * 3 <= positioned.length) return true;

  if (positioned.every((part) => [...part.text].length <= 1)) return true;

  const seenPlacements = new Map<string, number>();
  for (const part of positioned) {
    if (part.layout?.kind !== 'positioned' || !part.position) continue;
    const placement = `${part.text}@${Math.round(part.position.x)},${Math.round(part.position.y)}`;
    const earlierStart = seenPlacements.get(placement);
    if (earlierStart !== undefined && Math.abs(part.startTime - earlierStart) > 0.01) {
      return true;
    }
    seenPlacements.set(placement, part.startTime);
  }

  // Table cells align into columns: several x values each reused on multiple rows.
  // Requiring two such columns holding at least half the parts keeps a wrapped lyric
  // whose rows accidentally share one x coordinate out of the grid bucket.
  const columnRows = new Map<number, Set<number>>();
  for (const part of positioned) {
    if (!part.position) continue;
    const x = Math.round(part.position.x);
    const rows = columnRows.get(x) ?? new Set<number>();
    rows.add(Math.round(part.position.y));
    columnRows.set(x, rows);
  }
  let alignedColumns = 0;
  let alignedParts = 0;
  for (const part of positioned) {
    if (!part.position) continue;
    if ((columnRows.get(Math.round(part.position.x))?.size ?? 0) >= 2) alignedParts += 1;
  }
  for (const rows of columnRows.values()) {
    if (rows.size >= 2) alignedColumns += 1;
  }
  return alignedColumns >= 2 && alignedParts * 2 >= positioned.length;
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
    maximumY - minimumY > MAX_FRAGMENT_LINE_VERTICAL_SPAN &&
    looksLikeFragmentGridParts(parts)
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

interface FragmentInterval {
  startTime: number;
  endTime: number;
}

/** Event time ranges with repeated same-text, same-time layer copies collapsed. */
function distinctFragmentIntervals(events: readonly AnnotatedSubtitleCue[]): FragmentInterval[] {
  const intervals: FragmentInterval[] = [];
  const previousEvents: AnnotatedSubtitleCue[] = [];
  for (const event of events) {
    const compactText = compactCueMatchText(event);
    const isLayerCopy = previousEvents.some(
      (previous) =>
        compactCueMatchText(previous) === compactText &&
        previous.startTime === event.startTime &&
        previous.endTime === event.endTime &&
        isRepeatedFragmentCopy(previous, event),
    );
    previousEvents.push(event);
    if (isLayerCopy) continue;
    intervals.push({ startTime: event.startTime, endTime: event.endTime });
  }
  return intervals.sort((a, b) => a.startTime - b.startTime || a.endTime - b.endTime);
}

function intervalsNeverCoexist(intervals: readonly FragmentInterval[]): boolean {
  let latestEnd = -Infinity;
  for (const interval of intervals) {
    if (interval.startTime < latestEnd - 0.001) {
      return false;
    }
    latestEnd = Math.max(latestEnd, interval.endTime);
  }
  return true;
}

// A sweep progresses through the syllables of a lyric, so its events carry different
// texts. Sequential same-text repaints are one shaking line being redrawn, and its text
// must survive to the raw/burst path rather than be suppressed as decoration.
function hasMultipleFragmentTexts(events: readonly AnnotatedSubtitleCue[]): boolean {
  const first = events[0] ? compactCueMatchText(events[0]) : '';
  return events.some((event) => compactCueMatchText(event) !== first);
}

/**
 * A karaoke highlight sweep repaints one syllable at a time over an already-visible
 * lyric line: each event ends as the next begins, so the cluster's concatenated text is
 * never on screen as a whole. Publishing it would emit rolling partial copies of the
 * lyric ("to sou omo" beside "akenakute ii to sou omotteta"). Layer copies share one
 * placement and timing, so the test is whether any two distinct placements coexist.
 */
function isProgressiveHighlightSweep(events: readonly AnnotatedSubtitleCue[]): boolean {
  if (!hasMultipleFragmentTexts(events)) {
    return false;
  }
  const intervals = distinctFragmentIntervals(events);
  return intervals.length >= 2 && intervalsNeverCoexist(intervals);
}

/**
 * Timing clusters split a long sweep unevenly, leaving stragglers the per-cluster check
 * cannot judge: a two-event tail reconstructs on relaxed evidence, and a lone held
 * syllable stays raw and publishes as its own flickering cue. When an entire style group
 * reads as one chained repaint -- many short positioned animated fragments, no two ever
 * on screen together, transitions mostly back-to-back -- the whole group is highlight
 * decoration and none of it is publishable text. Independent one-off signs sharing a
 * style stay published: they are few, longer, or separated by real gaps.
 */
function isProgressiveHighlightSweepGroup(events: readonly AnnotatedSubtitleCue[]): boolean {
  if (
    events.length > MAX_SWEEP_GROUP_EVENTS ||
    sourceEventCount(events) < MIN_FRAGMENT_LINE_EVENTS ||
    !hasMultipleFragmentTexts(events) ||
    !events.every((event) => fragmentPlacementAnchors(event).size > 0) ||
    !hasAssAnimationEvidence(events)
  ) {
    return false;
  }
  const lengths = events
    .map((event) => compactCueMatchText(event).length)
    .sort((left, right) => left - right);
  if ((lengths[Math.floor(lengths.length / 2)] ?? Infinity) > MAX_FRAGMENT_MEDIAN_LENGTH) {
    return false;
  }
  const intervals = distinctFragmentIntervals(events);
  if (intervals.length < 2 || !intervalsNeverCoexist(intervals)) {
    return false;
  }
  let abutting = 0;
  for (let index = 1; index < intervals.length; index += 1) {
    if (Math.abs(intervals[index]!.startTime - intervals[index - 1]!.endTime) <= 0.1) {
      abutting += 1;
    }
  }
  return abutting * 2 >= intervals.length - 1;
}

function decodeSingleAssFragment(cue: AnnotatedSubtitleCue): string | null {
  const visibleLines = decodeSubtitleCueText(cue.rawText)
    .split('\n')
    .filter((line) => line.trim().length > 0);
  return visibleLines.length === 1 ? visibleLines[0]! : null;
}

/**
 * One cluster can hold the same authored text at two granularities: a whole-line event
 * and the per-glyph events that spell it (an assembly effect renders the line while its
 * glyph particles converge). Joining both doubles the line. A consecutive run of two or
 * more parts that concatenates to exactly another part's text is that part's fragment
 * layer; the whole part keeps the authored spacing, so the run is dropped. Single equal
 * parts are never dropped -- a repeated word in a lyric is real text, not a layer.
 *
 * Spelling alone is not proof: repeated digits after a thousands group also concatenate
 * to the earlier fragment's text while being real continuation. A duplicate layer sits
 * on top of its fragments, so the whole part's anchor must fall inside the run's
 * positional span; text that merely continues the line sits beyond it.
 */
function isWholePartOverItsRun(whole: AssFragmentPart, run: readonly AssFragmentPart[]): boolean {
  const wholePosition = fragmentPosition(whole.cue);
  if (!wholePosition) {
    return true;
  }
  const xs = run
    .map((part) => fragmentPosition(part.cue)?.x)
    .filter((x): x is number => Number.isFinite(x));
  if (xs.length === 0) {
    return true;
  }
  return wholePosition.x >= Math.min(...xs) && wholePosition.x <= Math.max(...xs);
}

function dropFragmentRunsCoveredByWholeParts(parts: AssFragmentPart[]): AssFragmentPart[] {
  if (parts.length > MAX_FRAGMENT_COLLAPSE_PARTS) {
    return parts;
  }
  const kept = [...parts];
  let changed = true;
  while (changed) {
    changed = false;
    const wholes = [...kept].sort(
      (left, right) =>
        compactAssMatchText(right.text).length - compactAssMatchText(left.text).length,
    );
    for (const whole of wholes) {
      const wholeText = compactAssMatchText(whole.text);
      if ([...wholeText].length < 2) {
        break;
      }
      for (let start = 0; start < kept.length && !changed; start += 1) {
        if (kept[start] === whole) {
          continue;
        }
        let combined = '';
        for (let end = start; end < kept.length; end += 1) {
          if (kept[end] === whole) {
            break;
          }
          combined += compactAssMatchText(kept[end]!.text);
          if (!wholeText.startsWith(combined)) {
            break;
          }
          if (combined === wholeText) {
            const run = kept.slice(start, end + 1);
            if (end > start && isWholePartOverItsRun(whole, run)) {
              kept.splice(start, end - start + 1);
              changed = true;
            }
            break;
          }
        }
      }
      if (changed) {
        break;
      }
    }
  }
  return kept;
}

function reconstructAssFragmentLine(
  events: readonly AnnotatedSubtitleCue[],
): AnnotatedSubtitleCue | null {
  const hasRelaxedEvidence = hasRelaxedAssFragmentEvidence(events);
  const minimumEvents = hasRelaxedEvidence ? 2 : MIN_FRAGMENT_LINE_EVENTS;
  // Coalesced copies stand in for their source events, so animation-volume thresholds
  // count sources: a merged four-phase glyph is still four generated events of evidence.
  if (sourceEventCount(events) < minimumEvents || !hasAssAnimationEvidence(events)) {
    return null;
  }

  const collected: AssFragmentPart[] = [];
  for (const cue of events) {
    const text = decodeSingleAssFragment(cue);
    if (text === null) {
      return null;
    }
    const compactText = compactAssMatchText(text);
    const isLayerCopy = collected.some(
      (part) =>
        compactAssMatchText(part.text) === compactText && isRepeatedFragmentCopy(part.cue, cue),
    );
    if (!isLayerCopy) {
      collected.push({ cue, text });
    }
  }
  const minimumParts = hasRelaxedEvidence ? 1 : MIN_FRAGMENT_LINE_PARTS;
  if (
    collected.length < minimumParts ||
    (!hasRelaxedEvidence && collected.length === sourceEventCount(events))
  ) {
    return null;
  }
  // A cluster whose every part is one identical glyph is a particle field, not a line.
  // Reconstructing it would also break the surviving particles' same-text chain that
  // burst deduplication collapses downstream.
  if (collected.length >= 2) {
    const firstText = compactAssMatchText(collected[0]!.text);
    if (
      [...firstText].length === 1 &&
      collected.every((part) => compactAssMatchText(part.text) === firstText)
    ) {
      return null;
    }
  }
  // The granularity gate judges the cluster as authored, so it runs before redundant
  // whole-vs-fragments layers collapse: a line plus its glyph swarm is fragment-sized
  // work even though only the whole-line part survives into the join.
  const lengths = collected
    .map((part) => compactAssMatchText(part.text).length)
    .sort((left, right) => left - right);
  if ((lengths[Math.floor(lengths.length / 2)] ?? Infinity) > MAX_FRAGMENT_MEDIAN_LENGTH) {
    return null;
  }
  const parts = dropFragmentRunsCoveredByWholeParts(collected);

  const text = joinAssFragmentParts(parts);
  if (!text) {
    return null;
  }
  const owner = parts[0]!.cue;
  const animationStartTime = earliestStartTime(events);
  const animationEndTime = latestEndTime(events);
  // Publish the window where the line reads as sung text. Transparent pre-echoes render
  // the upcoming line before its first syllable, and exit ghosts fade past the hold, so
  // the raw span makes consecutive lyrics overlap on screen. The line starts with its
  // first opaque copy and holds until the last statically anchored one ends; a line
  // placed entirely by `\move` has no hold phase and keeps its full visible span.
  const sources = events.flatMap((event) => [...sourceEventsOf(event)]);
  const visibleSources = sources.filter((source) => !isTransparentFillEcho(source));
  const displaySources = visibleSources.length > 0 ? visibleSources : sources;
  const heldSources = displaySources.filter((source) =>
    source.overrides.some((command) => !command.animated && command.name.toLowerCase() === 'pos'),
  );
  return {
    ...owner,
    startTime: earliestStartTime(displaySources),
    endTime: latestEndTime(heldSources.length > 0 ? heldSources : displaySources),
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

// `\fnSplit splat splodge` tokenizes as name `fnSplit` + args `splat splodge`, while
// `\fnArial` is all name and `\fn04b` is all args, so the font is both pieces rejoined.
function staticFontOverride(cue: AnnotatedSubtitleCue): string | null {
  let font: string | null = null;
  for (const command of cue.overrides) {
    if (command.animated || !command.name.toLowerCase().startsWith('fn')) continue;
    font = [command.name.slice(2), command.args].filter(Boolean).join(' ').trim().toLowerCase();
  }
  return font;
}

const MIN_TEXTURE_GLYPH_RUN = 8;
const MIN_TEXTURE_ALPHA_OVERRIDES = 6;
// Texture payloads switch secondary alpha at nearly every glyph. Authored text that a
// typesetter styles in syllable or word chunks measures two or more glyphs per
// override, so the seed test demands per-glyph density.
const MAX_TEXTURE_GLYPHS_PER_ALPHA_OVERRIDE = 1.5;
const MIN_TEXTURE_LAYER_ALPHA = 0xe0;
const MAX_TEXTURE_PAYLOAD_FONT_SIZE = 12;
const MIN_TEXTURE_PAYLOAD_LINES = 3;
const ASS_ALPHA_VALUE_PATTERN = /^&?H([0-9a-f]{1,2})&?$/iu;
const ASS_FONT_WEIGHT_SUFFIX_PATTERN =
  /\s+(?:black|bold|heavy|light|medium|regular|semibold|thin)$/u;

function hasStaticOverride(cue: AnnotatedSubtitleCue, expectedName: string): boolean {
  return cue.overrides.some(
    (command) => !command.animated && command.name.toLowerCase() === expectedName,
  );
}

function isRepeatedGlyphText(cue: AnnotatedSubtitleCue): boolean {
  const glyphs = [...compactCueMatchText(cue)];
  return glyphs.length > 0 && glyphs.every((glyph) => glyph === glyphs[0]);
}

function isClippedRepeatedGlyphFragment(cue: AnnotatedSubtitleCue): boolean {
  return (
    isRepeatedGlyphText(cue) && (hasStaticOverride(cue, 'clip') || hasStaticOverride(cue, 'iclip'))
  );
}

/**
 * Some ASS signs build image textures from clipped placeholder glyphs, optionally through
 * a texture font. A long clipped single-glyph run or frequent changing secondary alpha
 * tags identifies the effect without guessing from its visible text or font name.
 */
function isAssTextureSeed(cue: AnnotatedSubtitleCue): boolean {
  if (fragmentPosition(cue) === null) {
    return false;
  }

  // A truncated capture can leave tag debris after the placeholder run, so the seed
  // test looks for a long same-glyph run inside the clipped text rather than requiring
  // the whole event to be uniform.
  const glyphs = [...compactCueMatchText(cue)];
  if (hasStaticOverride(cue, 'clip') || hasStaticOverride(cue, 'iclip')) {
    let longestRun = 0;
    let run = 0;
    let previous = '';
    for (const glyph of glyphs) {
      run = glyph === previous ? run + 1 : 1;
      previous = glyph;
      longestRun = Math.max(longestRun, run);
    }
    if (longestRun >= MIN_TEXTURE_GLYPH_RUN) {
      return true;
    }
  }

  if (staticFontOverride(cue) === null) {
    return false;
  }

  const secondaryAlpha = cue.overrides.filter(
    (command) => !command.animated && command.name.toLowerCase() === '2a',
  );
  if (secondaryAlpha.length < MIN_TEXTURE_ALPHA_OVERRIDES) {
    return false;
  }
  // Real signs can alternate secondary alpha between words. Texture payloads switch it
  // per glyph, so sparse word-level styling must not seed a texture-font group.
  if (glyphs.length > secondaryAlpha.length * MAX_TEXTURE_GLYPHS_PER_ALPHA_OVERRIDE) {
    return false;
  }
  const alphaValues = secondaryAlpha.map((command) => command.args.toLowerCase());
  return new Set(alphaValues).size >= 2;
}

function staticGlobalAlpha(cue: AnnotatedSubtitleCue): number | null {
  return staticAlphaOverride(cue, 'alpha');
}

function staticAlphaOverride(cue: AnnotatedSubtitleCue, expectedName: string): number | null {
  let alpha: number | null = null;
  for (const command of cue.overrides) {
    if (command.animated || command.name.toLowerCase() !== expectedName) continue;
    const match = ASS_ALPHA_VALUE_PATTERN.exec(command.args.trim());
    const alphaValue = match?.[1];
    if (alphaValue !== undefined) {
      alpha = Number.parseInt(alphaValue, 16);
    }
  }
  return alpha;
}

function hasAnimatedAlphaOverride(cue: AnnotatedSubtitleCue): boolean {
  return cue.overrides.some((command) => {
    if (!command.animated) return false;
    const name = command.name.toLowerCase();
    return name === '1a' || name === 'alpha';
  });
}

// `\alpha&HFF&` blanks all four layers, but a later component override can turn one
// back on: chant overlays render entirely through `\4a&H00&` shadows. Any re-enabled
// layer means the event draws real text.
const ASS_COMPONENT_ALPHA_NAMES = ['1a', '2a', '3a', '4a'] as const;

function hasVisibleComponentAlpha(cue: AnnotatedSubtitleCue): boolean {
  return ASS_COMPONENT_ALPHA_NAMES.some((name) => {
    const value = staticAlphaOverride(cue, name);
    return value !== null && value < 0xff;
  });
}

/**
 * A glyph copy whose fill is statically fully transparent and never animated back in is
 * a glow or outline echo of the real glyph, not the text itself. A coalesced copy chain
 * counts only when every phase in it is such an echo; one opaque phase means the chain
 * carries the authored glyph.
 */
function isTransparentFillEcho(cue: AnnotatedSubtitleCue): boolean {
  return sourceEventsOf(cue).every(
    (source) =>
      (staticAlphaOverride(source, '1a') === 0xff ||
        staticAlphaOverride(source, 'alpha') === 0xff) &&
      !hasAnimatedAlphaOverride(source) &&
      !hasVisibleComponentAlpha(source),
  );
}

function staticFontSize(cue: AnnotatedSubtitleCue): number | null {
  let fontSize: number | null = null;
  for (const command of cue.overrides) {
    if (command.animated || command.name.toLowerCase() !== 'fs') continue;
    const value = Number(command.args.trim());
    if (Number.isFinite(value) && value > 0) {
      fontSize = value;
    }
  }
  return fontSize;
}

function isNearlyTransparentPositionedText(cue: AnnotatedSubtitleCue): boolean {
  const alpha = staticGlobalAlpha(cue);
  return (
    alpha !== null &&
    alpha >= MIN_TEXTURE_LAYER_ALPHA &&
    staticFontOverride(cue) !== null &&
    fragmentPosition(cue) !== null
  );
}

function hasAssTextureCandidateEvidence(cue: AnnotatedSubtitleCue): boolean {
  return (
    isClippedRepeatedGlyphFragment(cue) ||
    isNearlyTransparentPositionedText(cue) ||
    hasStaticOverride(cue, '2a')
  );
}

function textureFontFamilyKey(font: string): string {
  return font.replace(ASS_FONT_WEIGHT_SUFFIX_PATTERN, '');
}

function isTextureFontPayload(
  cue: AnnotatedSubtitleCue,
  textureFontFamilies: ReadonlySet<string>,
): boolean {
  const font = staticFontOverride(cue);
  const fontSize = staticFontSize(cue);
  const visibleLines = cue.text.split('\n').filter((line) => line.trim().length > 0);
  return (
    font !== null &&
    textureFontFamilies.has(textureFontFamilyKey(font)) &&
    fontSize !== null &&
    fontSize <= MAX_TEXTURE_PAYLOAD_FONT_SIZE &&
    staticGlobalAlpha(cue) !== null &&
    fragmentPosition(cue) !== null &&
    visibleLines.length >= MIN_TEXTURE_PAYLOAD_LINES
  );
}

// A sign translation can legitimately render faint text through one or two positioned
// events. Dozens of them sharing one window is a texture: near-invisible glyph strings
// laid out as pixels of an image, with no visible-text sibling to anchor them.
const MIN_TEXTURE_WALL_EVENTS = 6;

function textureWallGroupKey(cue: AnnotatedSubtitleCue): string {
  return `${cue.style}\0${cue.startTime}\0${cue.endTime}`;
}

/** Static zero scale or a degenerate static clip renders nothing, unless animation can
 * still bring the event into view (an entrance growing from `\fscx0`, a clip wipe).
 * Only an animated scale or clip reveals; `\t(...)` wrapping some other property leaves
 * the event invisible. Nested `\t(...)` needs no special case because the tags it
 * animates are recorded as animated in their own right. */
function isInvisiblyRenderedEvent(cue: AnnotatedSubtitleCue): boolean {
  let staticZeroScale = false;
  let staticZeroClip = false;
  let animatedReveal = false;
  for (const command of cue.overrides) {
    const name = command.name.toLowerCase();
    if (command.animated) {
      if (name === 'fscx' || name === 'fscy' || name === 'clip') {
        animatedReveal = true;
      }
      continue;
    }
    if (name === 'fscx' || name === 'fscy') {
      if (Number(command.args.trim()) === 0) {
        staticZeroScale = true;
      }
    } else if (name === 'clip') {
      const args = command.args.split(',').map((value) => Number(value.trim()));
      if (
        args.length >= 4 &&
        args.slice(0, 4).every(Number.isFinite) &&
        (args[0]! >= args[2]! || args[1]! >= args[3]!)
      ) {
        staticZeroClip = true;
      }
    }
  }
  return (staticZeroScale || staticZeroClip) && !animatedReveal;
}

function assFontTextureGroupKey(cue: AnnotatedSubtitleCue): string | null {
  const font = staticFontOverride(cue);
  return font === null ? null : `${cue.style}\0${cue.startTime}\0${cue.endTime}\0${font}`;
}

function assTextureTimingGroupKey(cue: AnnotatedSubtitleCue): string {
  return `${cue.style}\0${cue.startTime}\0${cue.endTime}`;
}

function removeAssFontTextureEvents(events: ParsedAssEvents): ParsedAssEvents {
  const seeds = events.dialogue.filter(isAssTextureSeed);
  const seedSet = new Set(seeds);
  // Short pieces can share the seeded font effect under another actor. Matching the
  // seed's style, timing, and font only narrows the candidates; each piece must still
  // carry structural texture evidence.
  const textureGroups = new Set(
    seeds.map(assFontTextureGroupKey).filter((key): key is string => key !== null),
  );
  const noFontTextureTimings = new Set(
    seeds
      .filter((seed) => staticFontOverride(seed) === null)
      .map((seed) => assTextureTimingGroupKey(seed)),
  );
  // Some signs switch actor and font between the texture mask and its payload. A nearly
  // transparent text event that overlaps a proven seed in the same style is another input
  // to that visual effect. Opaque authored text in the same sign remains publishable.
  const seedsByStyle = new Map<string, AnnotatedSubtitleCue[]>();
  for (const seed of seeds) {
    const styleSeeds = seedsByStyle.get(seed.style);
    if (styleSeeds) {
      styleSeeds.push(seed);
    } else {
      seedsByStyle.set(seed.style, [seed]);
    }
  }
  const seedIndexesByStyle = new Map(
    [...seedsByStyle].map(([style, styleSeeds]) => [style, buildAssEventGroupIndex(styleSeeds)]),
  );
  const isAssociatedWithTextureSeed = (cue: AnnotatedSubtitleCue): boolean => {
    const styleSeedIndex = seedIndexesByStyle.get(cue.style);
    return (
      styleSeedIndex !== undefined &&
      eventsOverlappingWindow(styleSeedIndex, cue.startTime, cue.endTime).length > 0
    );
  };
  const textureFontFamilies = new Set(
    [
      ...seeds,
      ...events.dialogue.filter(
        (cue) => isNearlyTransparentPositionedText(cue) && isAssociatedWithTextureSeed(cue),
      ),
      ...events.comments.filter(
        (cue) => isNearlyTransparentPositionedText(cue) && isAssociatedWithTextureSeed(cue),
      ),
    ]
      .map(staticFontOverride)
      .filter((font): font is string => font !== null)
      .map(textureFontFamilyKey),
  );

  // Wall membership additionally requires that no component alpha turns a layer back
  // on: `\alpha&HFF&` plus a visible `\4a` renders real text through its shadow, and a
  // sign typeset entirely from such layers must not read as a texture.
  const isTextureWallCandidate = (cue: AnnotatedSubtitleCue): boolean =>
    isNearlyTransparentPositionedText(cue) && !hasVisibleComponentAlpha(cue);
  const transparentWallCounts = new Map<string, number>();
  for (const cue of events.dialogue) {
    if (isTextureWallCandidate(cue)) {
      const key = textureWallGroupKey(cue);
      transparentWallCounts.set(key, (transparentWallCounts.get(key) ?? 0) + 1);
    }
  }

  return {
    dialogue: events.dialogue.filter((cue) => {
      if (isInvisiblyRenderedEvent(cue)) {
        return false;
      }
      if (seedSet.has(cue)) {
        return false;
      }
      if (
        isTextureWallCandidate(cue) &&
        (transparentWallCounts.get(textureWallGroupKey(cue)) ?? 0) >= MIN_TEXTURE_WALL_EVENTS
      ) {
        return false;
      }
      if (
        staticFontOverride(cue) === null &&
        noFontTextureTimings.has(assTextureTimingGroupKey(cue)) &&
        isClippedRepeatedGlyphFragment(cue)
      ) {
        return false;
      }
      const key = assFontTextureGroupKey(cue);
      if (key !== null && textureGroups.has(key) && hasAssTextureCandidateEvidence(cue)) {
        return false;
      }
      if (isTextureFontPayload(cue, textureFontFamilies)) {
        return false;
      }
      if (!isNearlyTransparentPositionedText(cue)) {
        return true;
      }
      return !isAssociatedWithTextureSeed(cue);
    }),
    comments: events.comments,
  };
}

/**
 * Generated lyric effects often layer decoration over the real syllables: single letters
 * positioned above each glyph, animated in, and rendered through a `\fn` override to a
 * symbol font where `a` draws as a sparkle rather than a letter. Reading them as text
 * corrupts the reconstructed line (`sotto mimi ni ateru to` gains a trailing `a z x`).
 * Within one style/name group, a font used only for scattered animated single glyphs --
 * while the group's actual text renders in another font -- marks those events as
 * decoration rather than dialogue.
 */
function decorativeGlyphEvents(events: readonly AnnotatedSubtitleCue[]): Set<AnnotatedSubtitleCue> {
  const byFont = new Map<string, AnnotatedSubtitleCue[]>();
  for (const cue of events) {
    const font = staticFontOverride(cue);
    if (font === null) continue;
    const group = byFont.get(font);
    if (group) {
      group.push(cue);
    } else {
      byFont.set(font, [cue]);
    }
  }

  const decorative = new Set<AnnotatedSubtitleCue>();
  for (const fontEvents of byFont.values()) {
    if (fontEvents.length * 2 >= events.length) continue;
    const allScatteredGlyphs = fontEvents.every(
      (cue) =>
        [...compactCueMatchText(cue)].length === 1 &&
        fragmentPosition(cue) !== null &&
        hasAssTemporalOverride(cue.overrides),
    );
    if (allScatteredGlyphs) {
      fontEvents.forEach((cue) => decorative.add(cue));
    }
  }
  return decorative;
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
  // The published dialogue list holds source events, so a suppressed coalesced copy
  // chain must suppress every event behind it.
  const suppress = (event: AnnotatedSubtitleCue): void => {
    for (const source of sourceEventsOf(event)) {
      suppressed.add(source);
    }
  };
  for (const events of groups.values()) {
    const units = coalesceAssAnchorCopies(events);
    const decorative = decorativeGlyphEvents(units);
    for (const unit of units) {
      if (
        !decorative.has(unit) &&
        fragmentPlacementAnchors(unit).size > 0 &&
        isTransparentFillEcho(unit)
      ) {
        decorative.add(unit);
      }
    }
    const lineEvents = decorative.size ? units.filter((event) => !decorative.has(event)) : units;
    if (isProgressiveHighlightSweepGroup(lineEvents)) {
      lineEvents.forEach(suppress);
      const spanStart = Math.min(...lineEvents.map((event) => event.startTime));
      const spanEnd = Math.max(...lineEvents.map((event) => event.endTime));
      for (const overlay of decorative) {
        if (
          overlay.startTime < spanEnd + DECORATION_SPAN_TOLERANCE_SECONDS &&
          overlay.endTime > spanStart - DECORATION_SPAN_TOLERANCE_SECONDS
        ) {
          suppress(overlay);
        }
      }
      continue;
    }
    for (const cluster of clusterAssFragmentEvents(lineEvents)) {
      const line = reconstructAssFragmentLine(cluster.events);
      if (!line) {
        continue;
      }
      // A sweep only re-highlights the lyric it decorates: hide its events without
      // publishing the reconstruction.
      if (isProgressiveHighlightSweep(cluster.events)) {
        cluster.events.forEach(suppress);
        continue;
      }
      recovered.push(line);
      cluster.events.forEach(suppress);
      // Decoration is timed to the line it overlays, so it disappears with the line's
      // full animation span. Decoration outside any recovered span stays published.
      const spanStart = line.animationStartTime ?? line.startTime;
      const spanEnd = line.animationEndTime ?? line.endTime;
      for (const overlay of decorative) {
        if (
          overlay.startTime < spanEnd + DECORATION_SPAN_TOLERANCE_SECONDS &&
          overlay.endTime > spanStart - DECORATION_SPAN_TOLERANCE_SECONDS
        ) {
          suppress(overlay);
        }
      }
    }
  }
  if (recovered.length === 0 && suppressed.size === 0) {
    return dialogue;
  }
  return [
    ...dialogue.filter((cue) => !suppressed.has(cue)),
    ...mergeAbuttingRecoveredLines(recovered),
  ].sort(
    (left, right) =>
      left.startTime - right.startTime || left.endTime - right.endTime || left.order - right.order,
  );
}

// One authored line can reconstruct twice from consecutive effect stages -- its steady
// glyphs, then an exit animation replaying the same text. Publishing both would show the
// line restarting, so identical recoveries that touch in time collapse into one span.
const RECOVERED_LINE_MERGE_GAP_SECONDS = 0.1;

function mergeAbuttingRecoveredLines(recovered: AnnotatedSubtitleCue[]): AnnotatedSubtitleCue[] {
  const byText = new Map<string, AnnotatedSubtitleCue[]>();
  for (const line of recovered) {
    const key = `${assEventGroupKey(line)}\0${compactAssMatchText(line.text)}`;
    const bucket = byText.get(key);
    if (bucket) {
      bucket.push(line);
    } else {
      byText.set(key, [line]);
    }
  }

  const merged: AnnotatedSubtitleCue[] = [];
  for (const bucket of byText.values()) {
    bucket.sort((left, right) => left.startTime - right.startTime || left.order - right.order);
    let current = bucket[0]!;
    for (let index = 1; index < bucket.length; index += 1) {
      const next = bucket[index]!;
      if (next.startTime <= current.endTime + RECOVERED_LINE_MERGE_GAP_SECONDS) {
        const endTime = Math.max(current.endTime, next.endTime);
        current = {
          ...current,
          endTime,
          animationEndTime: Math.max(
            current.animationEndTime ?? current.endTime,
            next.animationEndTime ?? next.endTime,
          ),
        };
      } else {
        merged.push(current);
        current = next;
      }
    }
    merged.push(current);
  }
  return merged;
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
          (cue.assLayout?.kind === 'source-order' || hasAssAnimationEvidence([cue])),
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

function bandFromNumpadAlignment(alignment: number): AssVerticalBand | null {
  if (alignment >= 7 && alignment <= 9) return 'top';
  if (alignment >= 4 && alignment <= 6) return 'middle';
  if (alignment >= 1 && alignment <= 3) return 'bottom';
  return null;
}

// SSA v4 alignment reuses the legacy `\a` codes: 1-3 bottom, +4 top, +8 middle.
function bandFromLegacyAlignment(alignment: number): AssVerticalBand | null {
  if (alignment >= 9 && alignment <= 11) return 'middle';
  if (alignment >= 5 && alignment <= 7) return 'top';
  if (alignment >= 1 && alignment <= 3) return 'bottom';
  return null;
}

interface AssPlacementContext {
  playResY: number | null;
  /** Lowercased style name -> vertical band from the style's Alignment column. */
  styleBands: Map<string, AssVerticalBand>;
}

const EMPTY_PLACEMENT_CONTEXT: AssPlacementContext = { playResY: null, styleBands: new Map() };

function parseAssPlacementContext(content: string): AssPlacementContext {
  const styleBands = new Map<string, AssVerticalBand>();
  let playResY: number | null = null;
  let section: 'info' | 'v4plus' | 'v4' | null = null;
  let alignmentIndex = -1;
  let nameIndex = -1;

  for (const line of content.split(/\r?\n/)) {
    const trimmed = line.trim();
    if (trimmed.startsWith('[') && trimmed.endsWith(']')) {
      const sectionName = trimmed.toLowerCase();
      section =
        sectionName === '[script info]'
          ? 'info'
          : sectionName === '[v4+ styles]'
            ? 'v4plus'
            : sectionName === '[v4 styles]'
              ? 'v4'
              : null;
      alignmentIndex = -1;
      nameIndex = -1;
      continue;
    }
    if (section === 'info') {
      const resMatch = trimmed.match(/^playresy\s*:\s*(\d+(?:\.\d+)?)\s*$/i);
      if (resMatch) playResY = Number(resMatch[1]);
      continue;
    }
    if (section !== 'v4plus' && section !== 'v4') continue;
    const separator = trimmed.indexOf(':');
    if (separator < 0) continue;
    const key = trimmed.slice(0, separator).trim().toLowerCase();
    const fields = trimmed.slice(separator + 1).split(',');
    if (key === 'format') {
      const names = fields.map((field) => field.trim().toLowerCase());
      alignmentIndex = names.indexOf('alignment');
      nameIndex = names.indexOf('name');
      continue;
    }
    if (key !== 'style' || alignmentIndex < 0 || nameIndex < 0) continue;
    const styleName = fields[nameIndex]?.trim().toLowerCase();
    const alignment = Number(fields[alignmentIndex]?.trim());
    if (!styleName || !Number.isFinite(alignment)) continue;
    const band =
      section === 'v4plus'
        ? bandFromNumpadAlignment(alignment)
        : bandFromLegacyAlignment(alignment);
    if (band) styleBands.set(styleName, band);
  }

  return { playResY, styleBands };
}

/**
 * Where on screen mpv will draw this event: an explicit `\pos`/`\move` coordinate when
 * the script declares its coordinate space, else an `\an`/`\a` override, else the
 * style's Alignment. Constant for the life of the event, which is what lets simultaneous
 * lines keep a stable stacking order in the overlay.
 */
function resolveVerticalBand(
  overrides: readonly AssOverrideCommand[],
  y: number | null,
  style: string,
  context: AssPlacementContext,
): AssVerticalBand | undefined {
  if (y !== null && context.playResY && context.playResY > 0) {
    const ratio = y / context.playResY;
    return ratio < 1 / 3 ? 'top' : ratio < 2 / 3 ? 'middle' : 'bottom';
  }
  for (const command of overrides) {
    if (command.animated) continue;
    const name = command.name.toLowerCase();
    if (name !== 'an' && name !== 'a') continue;
    const band =
      name === 'an'
        ? bandFromNumpadAlignment(Number(command.args))
        : bandFromLegacyAlignment(Number(command.args));
    if (band) return band;
  }
  return context.styleBands.get(style.trim().toLowerCase());
}

function parseAssCoordinate(value: string | undefined): number | null {
  if (!value?.trim()) return null;
  const coordinate = Number(value.trim());
  return Number.isFinite(coordinate) ? coordinate : null;
}

function buildAssCueLayout(
  overrides: readonly AssOverrideCommand[],
  sourceOrder: number,
  style: string,
  placement: AssPlacementContext,
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
  const verticalBand = resolveVerticalBand(overrides, y, style, placement);
  const base: AssCueLayout =
    y === null ? { kind: 'source-order', sourceOrder } : { kind: 'positioned', sourceOrder, y };
  return verticalBand ? { ...base, verticalBand } : base;
}

function parseAnnotatedAssEvents(content: string): ParsedAssEvents {
  const cues: AnnotatedSubtitleCue[] = [];
  const comments: AnnotatedSubtitleCue[] = [];
  const placement = content.includes('[')
    ? parseAssPlacementContext(content)
    : EMPTY_PLACEMENT_CONTEXT;
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
    const style = readField(fields, fieldIndex.style);
    const cue: AnnotatedSubtitleCue = {
      startTime,
      endTime,
      text,
      rawText,
      style,
      layer: Number.isFinite(layer) ? layer : 0,
      name: readField(fields, fieldIndex.name),
      effect,
      effectKind: parseAssEffectField(effect),
      overrides,
      overrideSignature: assOverrideSignature(overrides),
      order: eventOrder,
      assLayout: buildAssCueLayout(overrides, eventOrder, style, placement),
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
  const events = removeAssFontTextureEvents(parseAnnotatedAssEvents(content));
  return recoverFragmentOnlyAssLines(recoverCanonicalAssEvents(events));
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
