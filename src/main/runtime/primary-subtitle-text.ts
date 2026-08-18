import type { SubtitleCue } from '../../types';

// Slack on top of each cue's recorded animation envelope, for time-pos observation
// staleness and small user sub-delay offsets. The envelope itself covers how far
// entrance/exit frames actually run past the authored timing.
const CANONICAL_ANIMATION_EDGE_TOLERANCE_SECONDS = 1;

export interface ResolvedPrimarySubtitle {
  text: string;
  startTime: number;
  endTime: number;
  /** The canonical cues behind `text`, for consumers that record lines individually. */
  cues: SubtitleCue[];
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
): SubtitleCue[] {
  return (cues ?? []).filter((cue) => {
    if (cue.source !== 'canonical-ass') {
      return false;
    }
    const span = animationSpan(cue);
    return (
      span.end >= currentTimeSec - CANONICAL_ANIMATION_EDGE_TOLERANCE_SECONDS &&
      span.start <= currentTimeSec + CANONICAL_ANIMATION_EDGE_TOLERANCE_SECONDS
    );
  });
}

function compactWhitespace(text: string): string {
  return text.replace(/\s+/gu, '');
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

  const texts: string[] = [];
  const seen = new Set<string>();
  for (const cue of selected) {
    if (!seen.has(cue.text)) {
      seen.add(cue.text);
      texts.push(cue.text);
    }
  }
  return {
    text: texts.join('\n'),
    startTime: Math.min(...selected.map((cue) => cue.startTime)),
    endTime: Math.max(...selected.map((cue) => cue.endTime)),
    cues: selected,
  };
}

/**
 * Live text with generated-animation fragment lines removed. Recording paths use this
 * when full canonical substitution declined -- concurrent dialogue during an insert
 * song: the dialogue is worth recording, the glyph fragments beside it are not. Returns
 * the input unchanged when no canonical cue is near or nothing non-fragment remains.
 */
export function stripCanonicalFragmentLines(options: {
  liveText: string;
  currentTimeSec: number;
  cues: readonly SubtitleCue[] | null | undefined;
}): string {
  if (!Number.isFinite(options.currentTimeSec)) {
    return options.liveText;
  }
  const nearby = nearbyCanonicalCues(options.cues, options.currentTimeSec);
  if (nearby.length === 0) {
    return options.liveText;
  }
  const compactCues = nearby.map((cue) => compactWhitespace(cue.text));
  const kept = options.liveText.split('\n').filter((line) => {
    const compact = compactWhitespace(line);
    return compact && !compactCues.some((cueText) => cueText.includes(compact));
  });
  return kept.length > 0 ? kept.join('\n') : options.liveText;
}

export function resolvePrimarySubtitleText(options: {
  liveText: string;
  currentTimeSec: number;
  cues: readonly SubtitleCue[] | null | undefined;
}): string {
  if (!options.liveText.trim()) {
    return options.liveText;
  }
  return (
    resolveCanonicalPrimarySubtitle({
      liveText: options.liveText,
      currentTimeSec: options.currentTimeSec,
      cues: options.cues,
    })?.text ?? options.liveText
  );
}
