/**
 * Character navigation for keyboard-only mode: instead of the word boxes, the cursor sits on
 * one character, and a lookup searches from it. That reaches words the subtitle split cuts in
 * two (もっとハ|キ|ハキ… hides ハキハキ from word navigation).
 */

/** The marked character, an element wrapping it inside its word span. */
export const KEYBOARD_SELECTED_CHAR_CLASS = 'keyboard-selected-char';
/** On the subtitle root while the character cursor is shown. */
export const KEYBOARD_CHARACTER_NAVIGATION_CLASS = 'keyboard-character-navigation';

/** A word span (by index among the line's words) and a character in it. */
export type CharacterCursorPosition = { wordIndex: number; charIndex: number };
export type CharacterCursorMoveResult = 'moved' | 'start-boundary' | 'end-boundary' | 'no-words';

/** Characters are counted as code points, so a surrogate pair is one character. */
export function countCharacters(text: string): number {
  return Array.from(text).length;
}

/** The position kept inside the line; `charIndex` past the end lands on the last character. */
export function clampCharacterCursor(
  wordLengths: readonly number[],
  position: CharacterCursorPosition,
): CharacterCursorPosition | null {
  if (wordLengths.length === 0) return null;
  const wordIndex = Math.min(Math.max(position.wordIndex, 0), wordLengths.length - 1);
  const lastChar = Math.max((wordLengths[wordIndex] ?? 0) - 1, 0);
  const charIndex = Math.min(Math.max(position.charIndex, 0), lastChar);
  return { wordIndex, charIndex };
}

/** One character left or right, crossing into the neighbouring word at its edges. */
export function moveCharacterCursor(
  wordLengths: readonly number[],
  position: CharacterCursorPosition,
  delta: -1 | 1,
): { position: CharacterCursorPosition; result: CharacterCursorMoveResult } {
  const start = clampCharacterCursor(wordLengths, position);
  if (!start) return { position, result: 'no-words' };

  let { wordIndex } = start;
  let charIndex = start.charIndex + delta;
  for (;;) {
    if (charIndex >= 0 && charIndex < (wordLengths[wordIndex] ?? 0)) {
      return { position: { wordIndex, charIndex }, result: 'moved' };
    }
    wordIndex += delta;
    if (wordIndex < 0) return { position: start, result: 'start-boundary' };
    if (wordIndex >= wordLengths.length) return { position: start, result: 'end-boundary' };
    // Empty words (an image only) have no character to stop on and are skipped.
    charIndex = delta > 0 ? 0 : (wordLengths[wordIndex] ?? 0) - 1;
  }
}

/** Marks and unmarks characters in the rendered subtitle. */
export type CharacterCursorPainter = {
  getLength: (word: HTMLElement) => number;
  /** Wraps one character of `word` in a marker element and returns it. */
  paint: (word: HTMLElement, charIndex: number) => HTMLElement | null;
  /** Removes every marker under `root`, restoring the plain text; true if there was one. */
  clear: (root: HTMLElement) => boolean;
  /** Range from the marker to the end of `lastWord`: the text a lookup searches. */
  rangeToEnd: (marker: HTMLElement, lastWord: HTMLElement) => Range;
};

function findCharacter(
  word: HTMLElement,
  charIndex: number,
): { node: Text; start: number; end: number } | null {
  const walker = document.createTreeWalker(word, NodeFilter.SHOW_TEXT);
  let remaining = charIndex;
  for (let node = walker.nextNode() as Text | null; node; node = walker.nextNode() as Text | null) {
    const chars = Array.from(node.data);
    if (remaining < chars.length) {
      const start = chars.slice(0, remaining).join('').length;
      return { node, start, end: start + (chars[remaining]?.length ?? 1) };
    }
    remaining -= chars.length;
  }
  return null;
}

export function createDomCharacterCursorPainter(): CharacterCursorPainter {
  return {
    getLength: (word) => countCharacters(word.textContent ?? ''),
    paint: (word, charIndex) => {
      const found = findCharacter(word, charIndex);
      if (!found) return null;
      const range = document.createRange();
      range.setStart(found.node, found.start);
      range.setEnd(found.node, found.end);
      const marker = document.createElement('span');
      marker.className = KEYBOARD_SELECTED_CHAR_CLASS;
      range.surroundContents(marker);
      return marker;
    },
    clear: (root) => {
      const markers = Array.from(
        root.querySelectorAll<HTMLElement>(`.${KEYBOARD_SELECTED_CHAR_CLASS}`),
      );
      for (const marker of markers) {
        const parent = marker.parentNode;
        if (!parent) continue;
        while (marker.firstChild) parent.insertBefore(marker.firstChild, marker);
        parent.removeChild(marker);
        parent.normalize();
      }
      return markers.length > 0;
    },
    rangeToEnd: (marker, lastWord) => {
      const range = document.createRange();
      range.setStart(marker, 0);
      range.setEnd(lastWord, lastWord.childNodes.length);
      return range;
    },
  };
}
