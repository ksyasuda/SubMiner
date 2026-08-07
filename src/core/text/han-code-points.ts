// Single source of truth for "this code point is a Han character", shared by
// the main-process character dictionary and the in-page Yomitan scan runtime.
// The two used to carry separate range lists, and they drifted: a name written
// with a supplementary-plane kanji could enter the generated dictionary while
// the scanner's greedy name pre-pass refused to probe the position.
//
// Ranges rather than \p{Script=Han}: the scan walk tests one code point per
// character of every subtitle line, where an integer compare beats building a
// string for a regex, and the script is injected as text into a page where a
// shared helper cannot be imported.
export const HAN_CODE_POINT_RANGES: ReadonlyArray<readonly [number, number]> = [
  [0x3400, 0x4dbf], // Extension A
  [0x4e00, 0x9fff], // CJK Unified Ideographs
  [0xf900, 0xfaff], // Compatibility Ideographs
  [0x20000, 0x2a6df], // Extension B
  [0x2a700, 0x2ebef], // Extensions C-F
  [0x2ebf0, 0x2ee5f], // Extension I
  [0x2f800, 0x2fa1f], // Compatibility Ideographs Supplement
  [0x30000, 0x3134f], // Extension G
  [0x31350, 0x323af], // Extension H
  [0x323b0, 0x33479], // Extension J (Unicode 17)
];

export function isHanCodePoint(codePoint: number): boolean {
  return HAN_CODE_POINT_RANGES.some(([start, end]) => codePoint >= start && codePoint <= end);
}

/** The same ranges as a regular expression character class body (needs the `u` flag). */
export const HAN_REGEXP_CLASS_BODY = HAN_CODE_POINT_RANGES.map(
  ([start, end]) => `\\u{${start.toString(16)}}-\\u{${end.toString(16)}}`,
).join('');
