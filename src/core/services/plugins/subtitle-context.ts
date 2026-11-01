import type { PluginSubtitleContext, PluginSubtitleLine } from './types';

interface CueLike {
  text: string;
  startTime: number;
  endTime: number;
}

const MAX_CONTEXT_LINES = 20;

function normalizeText(text: string): string {
  return text.replace(/\s+/g, ' ').trim();
}

function clampCount(value: number | undefined, fallback: number): number {
  if (value === undefined || !Number.isFinite(value)) return fallback;
  return Math.min(MAX_CONTEXT_LINES, Math.max(0, Math.floor(value)));
}

function toLine(cue: CueLike): PluginSubtitleLine {
  return { text: cue.text.trim(), startTime: cue.startTime, endTime: cue.endTime };
}

/**
 * Locates the displayed line in the parsed cue list. Text decides, timing breaks ties: mpv's live
 * timing can be shifted by sub-delay, so the same text nearest the reported start wins.
 */
function findCurrentCueIndex(cues: readonly CueLike[], current: PluginSubtitleLine): number {
  const target = normalizeText(current.text);
  const anchor = current.startTime;
  let best = -1;
  let bestDistance = Number.POSITIVE_INFINITY;
  cues.forEach((cue, index) => {
    if (normalizeText(cue.text) !== target) return;
    const distance = anchor === null ? 0 : Math.abs(cue.startTime - anchor);
    if (distance < bestDistance) {
      best = index;
      bestDistance = distance;
    }
  });
  if (best >= 0 || anchor === null) return best;
  return cues.findIndex((cue) => anchor >= cue.startTime && anchor < cue.endTime);
}

/**
 * Builds the current line plus neighbors. Uses the parsed subtitle file when one is loaded and
 * falls back to recently displayed lines (no `after` context) for streams without a cue list.
 */
export function buildPluginSubtitleContext(input: {
  cues: readonly CueLike[];
  current: PluginSubtitleLine | null;
  recentLines: readonly PluginSubtitleLine[];
  before?: number;
  after?: number;
}): PluginSubtitleContext {
  const beforeCount = clampCount(input.before, 2);
  const afterCount = clampCount(input.after, 1);
  const current = input.current && normalizeText(input.current.text) ? input.current : null;
  if (!current) return { current: null, before: [], after: [] };

  const cues = input.cues
    .filter((cue) => normalizeText(cue.text).length > 0)
    .slice()
    .sort((a, b) => a.startTime - b.startTime || a.endTime - b.endTime);
  const index = findCurrentCueIndex(cues, current);
  if (index >= 0) {
    return {
      current,
      before: cues.slice(Math.max(0, index - beforeCount), index).map(toLine),
      after: cues.slice(index + 1, index + 1 + afterCount).map(toLine),
    };
  }

  const target = normalizeText(current.text);
  const recent = input.recentLines.filter((line) => normalizeText(line.text).length > 0);
  // The tracker usually already holds the current line as its newest entry.
  const history =
    recent.length > 0 && normalizeText(recent[recent.length - 1]!.text) === target
      ? recent.slice(0, -1)
      : recent;
  return {
    current,
    before: beforeCount > 0 ? history.slice(-beforeCount) : [],
    after: [],
  };
}
