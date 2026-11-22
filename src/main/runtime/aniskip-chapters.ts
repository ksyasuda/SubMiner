/** A skippable time range in seconds. */
export interface SkipInterval {
  start: number;
  end: number;
}

export interface ChapterSkipIntervals {
  intro: SkipInterval | null;
  outro: SkipInterval | null;
}

/** Prefix of the marker chapters SubMiner writes; never treated as the file's own chapters. */
export const ANISKIP_CHAPTER_PREFIX = 'AniSkip ';

interface Chapter {
  title: string;
  start: number;
  end: number;
}

const MIN_HEURISTIC_CHAPTERS = 3;
const MIN_SONG_SECONDS = 55;
const MAX_SONG_SECONDS = 150;
const TYPICAL_OPENING_SECONDS = 90;

function isSongLength(chapter: Chapter): boolean {
  const length = chapter.end - chapter.start;
  return length >= MIN_SONG_SECONDS && length <= MAX_SONG_SECONDS;
}

/**
 * Turns mpv's `chapter-list` (entries with `time` and `title`) into sorted
 * chapters with an end time (the next chapter's start, or the duration),
 * dropping SubMiner's own AniSkip marker chapters.
 */
function toChapters(chapterList: unknown, duration: number): Chapter[] {
  if (!Array.isArray(chapterList)) return [];
  const starts: Array<{ title: string; start: number }> = [];
  for (const entry of chapterList) {
    if (typeof entry !== 'object' || entry === null) continue;
    const { time, title } = entry as { time?: unknown; title?: unknown };
    if (typeof time !== 'number' || !Number.isFinite(time) || time < 0 || time >= duration) {
      continue;
    }
    const text = typeof title === 'string' ? title : '';
    if (text.startsWith(ANISKIP_CHAPTER_PREFIX)) continue;
    starts.push({ title: text, start: time });
  }
  starts.sort((a, b) => a.start - b.start);
  return starts.map((chapter, index) => ({
    ...chapter,
    end: starts[index + 1]?.start ?? duration,
  }));
}

function isNamedOpening(chapter: Chapter): boolean {
  const title = chapter.title.trim().toLowerCase();
  return (
    title.includes('opening') ||
    title === 'op' ||
    title.startsWith('op ') ||
    chapter.title.includes('オープニング')
  );
}

function isNamedEnding(chapter: Chapter, duration: number): boolean {
  const title = chapter.title.trim().toLowerCase();
  return (
    title.includes('ending') ||
    title === 'ed' ||
    title === 'outro' ||
    title === 'end credits' ||
    title === 'closing credits' ||
    (title === 'credits' && chapter.start > duration * 0.5) ||
    title.startsWith('ed ') ||
    chapter.title.includes('エンディング')
  );
}

function findOpening(chapters: Chapter[], duration: number): Chapter | null {
  const named = chapters.find(isNamedOpening);
  if (named) return named;
  if (chapters.length < MIN_HEURISTIC_CHAPTERS) return null;
  let best: Chapter | null = null;
  for (const chapter of chapters) {
    if (chapter.start >= duration * 0.3 || !isSongLength(chapter)) continue;
    const distance = Math.abs(chapter.end - chapter.start - TYPICAL_OPENING_SECONDS);
    if (!best || distance < Math.abs(best.end - best.start - TYPICAL_OPENING_SECONDS)) {
      best = chapter;
    }
  }
  return best;
}

function findEnding(chapters: Chapter[], duration: number): Chapter | null {
  const named = chapters.find((chapter) => isNamedEnding(chapter, duration));
  if (named) return named;
  if (chapters.length < MIN_HEURISTIC_CHAPTERS) return null;
  return (
    chapters.find((chapter) => chapter.start > duration * 0.7 && isSongLength(chapter)) ?? null
  );
}

function toInterval(chapter: Chapter | null): SkipInterval | null {
  return chapter && chapter.end > chapter.start ? { start: chapter.start, end: chapter.end } : null;
}

/**
 * Finds the opening and ending in a file's own chapters: by name first
 * ("Opening", "OP", "ED", "Credits", オープニング, ...), else by length and
 * position when the file has at least three chapters. Rules follow Telly's
 * `media.rs` opening()/ending().
 */
export function classifyChapterSkipIntervals(
  chapterList: unknown,
  duration: number,
): ChapterSkipIntervals {
  if (!Number.isFinite(duration) || duration <= 0) return { intro: null, outro: null };
  const chapters = toChapters(chapterList, duration);
  return {
    intro: toInterval(findOpening(chapters, duration)),
    outro: toInterval(findEnding(chapters, duration)),
  };
}
