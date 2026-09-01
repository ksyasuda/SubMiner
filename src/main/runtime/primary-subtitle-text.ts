import type { AssVerticalBand, SubtitleCue } from '../../types';
import {
  removeAssControlDebrisLines,
  removeLiveGlyphFragmentLines,
} from '../../core/services/ass-text';

// Slack on top of each cue's recorded animation envelope, for time-pos observation
// staleness and small user sub-delay offsets. The envelope itself covers how far
// entrance/exit frames actually run past the authored timing.
const LIVE_CUE_EDGE_TOLERANCE_SECONDS = 1;

export interface ResolvedPrimarySubtitle {
  text: string;
  startTime: number;
  endTime: number;
  /** The parsed cues behind `text`, for consumers that record lines individually. */
  cues: SubtitleCue[];
}

function cuesUseAssSyntax(cues: readonly SubtitleCue[] | null | undefined): boolean {
  return (cues ?? []).some(
    (cue) =>
      cue.source === 'canonical-ass' ||
      cue.source === 'reconstructed-ass' ||
      cue.assLayout !== undefined,
  );
}

function animationSpan(cue: SubtitleCue): { start: number; end: number } {
  return {
    start: cue.animationStartTime ?? cue.startTime,
    end: cue.animationEndTime ?? cue.endTime,
  };
}

function nearbyCanonicalCues(
  cues: readonly SubtitleCue[] | null | undefined,
  currentTimeSec: number,
  includeFragmentGrids = false,
): SubtitleCue[] {
  return (cues ?? []).filter((cue) => {
    if (
      (cue.source !== 'canonical-ass' && cue.source !== 'reconstructed-ass') ||
      (!includeFragmentGrids && cue.assLayout?.kind === 'fragment-grid')
    ) {
      return false;
    }
    const span = animationSpan(cue);
    return (
      span.end >= currentTimeSec - LIVE_CUE_EDGE_TOLERANCE_SECONDS &&
      span.start <= currentTimeSec + LIVE_CUE_EDGE_TOLERANCE_SECONDS
    );
  });
}

function compactWhitespace(text: string): string {
  return text.normalize('NFKC').replace(/\s+/gu, '');
}

/**
 * Distinct simultaneous cues are separated by a blank line so the display layer can tell
 * a wrap inside one utterance from the boundary between two of them. Consumers that read
 * the text rather than display it fold these back to single breaks.
 */
const CUE_BOUNDARY = '\n\n';

const VERTICAL_BAND_RANK: Record<AssVerticalBand, number> = { top: 0, middle: 1, bottom: 2 };

/**
 * Stack simultaneous cues the way they sit on screen: mpv keeps a top-anchored lyric or
 * sign above bottom dialogue for its whole run, while cue-list order follows start time
 * and would swap the pair whenever one side is replaced mid-overlap. The band is
 * constant per event, so a line never changes rows while it is displayed.
 *
 * A cue whose placement could not be read -- an unknown style, a script with no styles
 * section -- sorts to the top. Dialogue is the case that reliably declares a bottom
 * alignment, so what is left unresolved is more often a sign or a song line, and keeping
 * the dialogue on the bottom row means the line worth reading stays where the eye
 * already is. Sort is stable, so cues sharing a rank keep their existing order.
 */
function orderCuesForDisplay(cues: readonly SubtitleCue[]): SubtitleCue[] {
  const rank = (cue: SubtitleCue): number =>
    VERTICAL_BAND_RANK[cue.assLayout?.verticalBand ?? 'top'];
  return [...cues].sort((a, b) => rank(a) - rank(b));
}

// ASS layers can encode the same visible spacing with ordinary, hard, or
// ideographic spaces. Matching and emission must use the same identity or each
// layer reappears as a copy.
function uniqueCueTextGroups(cues: readonly SubtitleCue[]): string[] {
  const groups: string[] = [];
  const seen = new Set<string>();
  for (const cue of cues) {
    const lines: string[] = [];
    for (const line of cue.text.split('\n')) {
      const compactText = compactWhitespace(line);
      if (!compactText || seen.has(compactText)) continue;
      seen.add(compactText);
      lines.push(line);
    }
    if (lines.length > 0) {
      groups.push(lines.join('\n'));
    }
  }
  return groups;
}

function compactLineSegments(text: string): string[] {
  return text.split('\n').map(compactWhitespace).filter(Boolean);
}

/**
 * Parsed cues have already collapsed exact ASS layers and animation runs. Trust that
 * cleaner view only when every live mpv line is accounted for by an active parsed cue.
 * This keeps unrelated concurrent dialogue on the live fallback while removing style
 * stacks where mpv repeats one full lyric for fill, border, blur, and shadow layers.
 */
function resolveActiveParsedPrimarySubtitle(options: {
  liveText: string;
  currentTimeSec: number;
  cues: readonly SubtitleCue[] | null | undefined;
}): ResolvedPrimarySubtitle | null {
  if (!Number.isFinite(options.currentTimeSec)) {
    return null;
  }

  const liveSegments = compactLineSegments(options.liveText);
  if (liveSegments.length === 0) {
    return null;
  }
  const liveSegmentSet = new Set(liveSegments);
  const selected = (options.cues ?? []).filter((cue) => {
    if (
      cue.startTime > options.currentTimeSec + LIVE_CUE_EDGE_TOLERANCE_SECONDS ||
      cue.endTime <= options.currentTimeSec - LIVE_CUE_EDGE_TOLERANCE_SECONDS
    ) {
      return false;
    }
    const cueSegments = compactLineSegments(cue.text);
    if (cueSegments.length === 0) return false;
    if (cue.source === 'canonical-ass' || cue.source === 'reconstructed-ass') {
      return liveSegments.some((segment) =>
        cueSegments.some((cueSegment) => cueSegment.includes(segment)),
      );
    }
    return cueSegments.every((segment) => liveSegmentSet.has(segment));
  });
  if (selected.length === 0) {
    return null;
  }

  const parsedSegments = selected.flatMap((cue) => {
    const recovered = cue.source === 'canonical-ass' || cue.source === 'reconstructed-ass';
    return [
      ...compactLineSegments(cue.text).map((segment) => ({ segment, recovered })),
      ...(cue.assFurigana ?? []).flatMap((text) =>
        compactLineSegments(text).map((segment) => ({ segment, recovered: false })),
      ),
    ];
  });
  if (
    !liveSegments.every((liveSegment) =>
      parsedSegments.some(({ segment, recovered }) =>
        recovered ? segment.includes(liveSegment) : segment === liveSegment,
      ),
    )
  ) {
    return null;
  }

  // A cue selected only through the edge tolerance on its end has already finished by
  // its published timing: a lyric whose exit ghosts linger into the next line. It still
  // explains those live fragments above, but must not re-surface beside cues that are
  // still running. The start side keeps the tolerance: mpv publishes the combined
  // sub-text the moment a joining line's first frame renders, while the observed
  // time-pos still sits just before that line's start, and the selection above already
  // required the cue's text to be on screen (#220). With every selected cue finished,
  // the edge cues remain the display fallback for stale time-pos readings.
  const unfinished = selected.filter((cue) => cue.endTime > options.currentTimeSec);
  const displayCues = unfinished.length > 0 ? unfinished : selected;

  // Dense sign grids still explain their raw mpv fragments, but are visual
  // typesetting rather than a publishable subtitle line.
  const groups = uniqueCueTextGroups(
    orderCuesForDisplay(displayCues.filter((cue) => cue.assLayout?.kind !== 'fragment-grid')),
  );
  return {
    text: groups.join(CUE_BOUNDARY),
    startTime: Math.min(...displayCues.map((cue) => cue.startTime)),
    endTime: Math.max(...displayCues.map((cue) => cue.endTime)),
    cues: displayCues,
  };
}

/**
 * mpv's `sub-text` renders each simultaneously active ASS event on its own line, so
 * while a generated animation plays every live line is a contiguous piece of the
 * authored text. A line that is not -- concurrent dialogue during an insert song, or a
 * fresh line starting just after the animation ended -- proves the live text is not this
 * animation, and substituting the canonical line would swallow real dialogue.
 */
function liveTextIsFromCues(liveText: string, cues: readonly SubtitleCue[]): boolean {
  const compactCues = cues.map((cue) => compactWhitespace(cue.text));
  const segments = liveText.split('\n').map(compactWhitespace).filter(Boolean);
  return (
    segments.length > 0 &&
    segments.every((segment) => compactCues.some((cueText) => cueText.includes(segment)))
  );
}

export function resolveCanonicalPrimarySubtitle(options: {
  liveText: string;
  currentTimeSec: number;
  cues: readonly SubtitleCue[] | null | undefined;
}): ResolvedPrimarySubtitle | null {
  if (!Number.isFinite(options.currentTimeSec)) {
    return null;
  }

  // Consecutive karaoke lines overlap: one line's exit frames are still on screen while
  // the next line's entrance frames appear. The fragment check therefore runs against
  // every canonical cue whose animation envelope reaches the current time, while only
  // the active (or single nearest) cue supplies the displayed text.
  const nearby = nearbyCanonicalCues(options.cues, options.currentTimeSec);
  const active = nearby.filter(
    (cue) => cue.startTime <= options.currentTimeSec && cue.endTime > options.currentTimeSec,
  );
  const liveSegments = options.liveText.split('\n').map(compactWhitespace).filter(Boolean);
  const selected =
    active.length > 0
      ? active
      : nearby
          // Between authored spans, proximity alone can pick the wrong neighbor: the
          // next line can sit closer while only the previous line's exit fragments are
          // on screen. Only cues that explain at least one live line may be selected.
          .filter((cue) => {
            const cueText = compactWhitespace(cue.text);
            return liveSegments.some((segment) => cueText.includes(segment));
          })
          .map((cue) => {
            const span = animationSpan(cue);
            const distance =
              options.currentTimeSec < span.start
                ? span.start - options.currentTimeSec
                : Math.max(0, options.currentTimeSec - span.end);
            return { cue, distance };
          })
          .sort((a, b) => a.distance - b.distance || a.cue.startTime - b.cue.startTime)
          .slice(0, 1)
          .map(({ cue }) => cue);
  if (selected.length === 0 || !liveTextIsFromCues(options.liveText, nearby)) {
    return null;
  }

  const groups = uniqueCueTextGroups(orderCuesForDisplay(selected));
  return {
    text: groups.join(CUE_BOUNDARY),
    startTime: Math.min(...selected.map((cue) => cue.startTime)),
    endTime: Math.max(...selected.map((cue) => cue.endTime)),
    cues: selected,
  };
}

/**
 * Live text with generated-animation fragment lines removed. Recording paths use this
 * when full canonical substitution declined -- concurrent dialogue during an insert
 * song: the dialogue is worth recording, the glyph fragments beside it are not. An
 * all-fragment visual grid becomes empty; other all-matched input remains unchanged as a
 * defensive fallback.
 */
export function stripCanonicalFragmentLines(options: {
  liveText: string;
  currentTimeSec: number;
  cues: readonly SubtitleCue[] | null | undefined;
}): string {
  if (!Number.isFinite(options.currentTimeSec)) {
    return removeLiveGlyphFragmentLines(options.liveText);
  }
  const nearby = nearbyCanonicalCues(options.cues, options.currentTimeSec, true);
  if (nearby.length === 0) {
    return removeLiveGlyphFragmentLines(options.liveText);
  }
  const compactCues = nearby.map((cue) => compactWhitespace(cue.text));
  const kept = options.liveText.split('\n').filter((line) => {
    const compact = compactWhitespace(line);
    return compact && !compactCues.some((cueText) => cueText.includes(compact));
  });
  if (kept.length > 0) return removeLiveGlyphFragmentLines(kept.join('\n'));
  if (nearby.some((cue) => cue.assLayout?.kind === 'fragment-grid')) return '';
  return removeLiveGlyphFragmentLines(options.liveText);
}

export function resolvePrimarySubtitleText(options: {
  liveText: string;
  currentTimeSec: number;
  cues: readonly SubtitleCue[] | null | undefined;
}): string {
  const liveText = cuesUseAssSyntax(options.cues)
    ? removeAssControlDebrisLines(options.liveText)
    : options.liveText;
  if (!liveText.trim()) {
    return liveText;
  }
  return (
    resolveCanonicalPrimarySubtitle({
      liveText,
      currentTimeSec: options.currentTimeSec,
      cues: options.cues,
    })?.text ??
    resolveActiveParsedPrimarySubtitle({ ...options, liveText })?.text ??
    removeLiveGlyphFragmentLines(liveText)
  );
}
