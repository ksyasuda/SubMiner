/*
 * ASS/SSA text handling, split into two deliberately distinct contracts:
 *
 *   assToPlainText()             raw ASS event text -> plain text. Ingestion only.
 *   normalizePlainSubtitleText() already-decoded text -> display/lookup form.
 *
 * Subtitle text is decoded from ASS exactly once, at the point it enters the app: the
 * file cue parser does it for sidecar/embedded scripts, and mpv does it for live text
 * (`sub-text` is already run through mpv's own `ass_to_plaintext`). Everything
 * downstream -- renderer, timing tracker, tokenizer, tokenization cache keys -- gets
 * plain text and only normalizes whitespace, so no layer decodes the same string twice.
 *
 * assToPlainText mirrors mpv's `ass_to_plaintext` rather than inventing its own rules,
 * so a cue parsed from a file reads the same as the same line arriving live:
 *   - `{...}` override blocks are markup
 *   - `\pN ... \p0` runs are vector paths, not dialogue
 *   - `\N`, `\n` and `\h` are the only escapes; `\{`, `\}` and `\\` are NOT escapes,
 *     so `\{注\}` decodes to a lone backslash exactly as mpv renders it
 *   - an unclosed `{` is rendered verbatim instead of swallowing the rest of the line
 * Because the decoder never emits an escape or a closed brace, running it twice is a
 * no-op -- but downstream code should still use normalizePlainSubtitleText.
 */

/** What `\N` and `\n` become. */
export type AssLineBreak = '\n' | ' ';

// `\p<n>` with n > 0 switches libass into vector-drawing mode: everything until the
// next `\p0` is a path (`m 20 0 b 10 0 ...`), not dialogue. The negative lookahead keeps
// `\pos(...)` from being read as a drawing tag.
const ASS_DRAWING_SCALE_PATTERN = /\\p(?![a-zA-Z])(\d*)/g;

function readDrawingScale(block: string): number | null {
  ASS_DRAWING_SCALE_PATTERN.lastIndex = 0;
  let scale: number | null = null;
  let match: RegExpExecArray | null;
  // Drawing mode is whatever the last `\p` tag in this block set it to.
  while ((match = ASS_DRAWING_SCALE_PATTERN.exec(block)) !== null) {
    scale = match[1] ? Number(match[1]) : 0;
  }
  return scale;
}

/** Resolve `\N`, `\n` and `\h`. The only text-level escapes libass recognises. */
function resolveWhitespaceEscapes(text: string, lineBreak: AssLineBreak): string {
  return text.replace(/\\([Nnh])/g, (_match, escaped: string) =>
    escaped === 'h' ? ' ' : lineBreak,
  );
}

/** Strip `{...}` override blocks and the drawing runs they enable. */
function stripAssMarkup(raw: string): string {
  let out = '';
  let cursor = 0;
  let drawing = false;

  while (cursor < raw.length) {
    if (raw[cursor] !== '{') {
      if (!drawing) {
        out += raw[cursor];
      }
      cursor += 1;
      continue;
    }

    const close = raw.indexOf('}', cursor + 1);
    if (close === -1) {
      // mpv shows an unclosed `{` and everything after it. Guessing where the block was
      // meant to end can eat a whole line of dialogue.
      if (!drawing) {
        out += raw.slice(cursor);
      }
      break;
    }

    const scale = readDrawingScale(raw.slice(cursor, close + 1));
    if (scale !== null) {
      drawing = scale > 0;
    }
    cursor = close + 1;
  }

  return out;
}

/**
 * Decode a raw ASS/SSA event text field. Call this once, where the text enters the app;
 * downstream layers take the result as plain text.
 */
export function assToPlainText(text: string, lineBreak: AssLineBreak = '\n'): string {
  if (!text) return '';
  return resolveWhitespaceEscapes(stripAssMarkup(text.replace(/\r\n/g, '\n')), lineBreak);
}

export interface NormalizePlainSubtitleTextOptions {
  /** Fold every line break into a single space. */
  collapseLineBreaks?: boolean;
  trim?: boolean;
}

/**
 * Whitespace normalization for text that has already been decoded -- by mpv for live
 * subtitles, by the cue parser for files. Override blocks and drawing runs are none of
 * this function's business; a `{` that reaches here is literal text mpv chose to show.
 *
 * `\N`/`\n`/`\h` are still folded, because subtitle sources outside the ASS path (asbplayer
 * and other websocket clients) forward them raw and the display layer has to cope.
 */
export function normalizePlainSubtitleText(
  text: string,
  options: NormalizePlainSubtitleTextOptions = {},
): string {
  if (!text) return '';
  const { collapseLineBreaks = false, trim = true } = options;

  let normalized = resolveWhitespaceEscapes(
    text.replace(/\r\n/g, '\n'),
    collapseLineBreaks ? ' ' : '\n',
  );
  if (collapseLineBreaks) {
    normalized = normalized.replace(/\n/g, ' ').replace(/\s+/g, ' ');
  }

  return trim ? normalized.trim() : normalized;
}

/** The contents of each `{...}` block, without the braces. */
export function extractAssOverrideBlocks(text: string): string[] {
  const blocks: string[] = [];
  let cursor = 0;

  while (cursor < text.length) {
    const open = text.indexOf('{', cursor);
    if (open === -1) {
      break;
    }
    const close = text.indexOf('}', open + 1);
    if (close === -1) {
      break;
    }
    blocks.push(text.slice(open + 1, close));
    cursor = close + 1;
  }

  return blocks;
}

export interface AssOverrideCommand {
  /** Tag name without the backslash, e.g. `pos`, `kf`, `1c`. */
  name: string;
  /** Everything the tag was given, e.g. `960,1068` for `\pos(960,1068)`. */
  args: string;
  /** Nested inside a `\t(...)` argument, so its value is animated over the event. */
  animated: boolean;
}

const ASS_OVERRIDE_NAME_PATTERN = /[1-4]?[a-zA-Z]+/y;

function readCommandArgs(block: string, start: number): { args: string; next: number } {
  if (block[start] === '(') {
    let depth = 0;
    for (let i = start; i < block.length; i += 1) {
      if (block[i] === '(') depth += 1;
      else if (block[i] === ')') {
        depth -= 1;
        if (depth === 0) {
          return { args: block.slice(start + 1, i), next: i + 1 };
        }
      }
    }
    return { args: block.slice(start + 1), next: block.length };
  }

  const nextTag = block.indexOf('\\', start);
  const end = nextTag === -1 ? block.length : nextTag;
  return { args: block.slice(start, end), next: end };
}

function parseOverrideBlock(block: string, animated: boolean, into: AssOverrideCommand[]): void {
  let cursor = 0;

  while (cursor < block.length) {
    if (block[cursor] !== '\\') {
      cursor += 1;
      continue;
    }

    ASS_OVERRIDE_NAME_PATTERN.lastIndex = cursor + 1;
    const nameMatch = ASS_OVERRIDE_NAME_PATTERN.exec(block);
    if (!nameMatch) {
      cursor += 1;
      continue;
    }

    const name = nameMatch[0];
    const { args, next } = readCommandArgs(block, cursor + 1 + name.length);
    into.push({ name, args: args.trim(), animated });
    // `\t(0,500,\frz30)` animates whatever it wraps, so record the inner tags too.
    if (name === 't' && args.includes('\\')) {
      parseOverrideBlock(args, true, into);
    }
    cursor = next;
  }
}

/**
 * Override commands with their arguments, in source order. Only `{...}` blocks are
 * inspected, so a `\pos(...)` sitting in visible text is never mistaken for markup.
 */
export function collectAssOverrideCommands(text: string): AssOverrideCommand[] {
  const commands: AssOverrideCommand[] = [];
  for (const block of extractAssOverrideBlocks(text)) {
    parseOverrideBlock(block, false, commands);
  }
  return commands;
}

// Tags that are animated by definition: `\t` interpolates, `\move` travels, and the
// karaoke tags advance a highlight across the event's own duration. Everything else --
// `\pos`, `\clip`, `\frz`, `\blur`, `\fad` -- is a static value for the event, so its
// presence says nothing about whether neighbouring events form one animation.
const ASS_TEMPORAL_COMMANDS = new Set(['t', 'move', 'k', 'kf', 'ko', 'K']);

export function isAssTemporalCommand(name: string): boolean {
  return ASS_TEMPORAL_COMMANDS.has(name);
}

/** True when the event animates on its own, or animates a static tag through `\t(...)`. */
export function hasAssTemporalOverride(commands: readonly AssOverrideCommand[]): boolean {
  return commands.some((command) => command.animated || isAssTemporalCommand(command.name));
}

/**
 * Canonical form of an event's override values, for comparing consecutive events. Two
 * events with the same signature were typeset identically, so neither is a frame of an
 * animation the other belongs to.
 */
export function assOverrideSignature(commands: readonly AssOverrideCommand[]): string {
  return commands.map((command) => `${command.name}(${command.args})`).join('|');
}

export type AssEffectKind = 'none' | 'banner' | 'scroll' | 'karaoke' | 'other';

// The stock effects, matched exactly. Typesetting groups put their own template names in
// this column -- `scrolling-credit` is a static sign, not libass's `Scroll up` -- so a
// prefix match would hand out animation evidence to arbitrary custom effects.
const STOCK_ASS_EFFECTS = new Map<string, AssEffectKind>([
  ['banner', 'banner'],
  ['scroll up', 'scroll'],
  ['scroll down', 'scroll'],
  ['karaoke', 'karaoke'],
]);

/**
 * The event-level `Effect` column. The stock values (`Banner;...`, `Scroll up;...`,
 * `Scroll down;...`, `Karaoke`) all animate; anything else is a custom name and lands in
 * `other`.
 */
export function parseAssEffectField(raw: string): AssEffectKind {
  const value = raw.trim().toLowerCase();
  if (!value) return 'none';

  const name = value.split(';', 1)[0]!.trim();
  return STOCK_ASS_EFFECTS.get(name) ?? 'other';
}

const ANIMATED_ASS_EFFECT_KINDS = new Set<AssEffectKind>(['banner', 'scroll', 'karaoke']);

export function isAnimatedAssEffectKind(kind: AssEffectKind): boolean {
  return ANIMATED_ASS_EFFECT_KINDS.has(kind);
}
