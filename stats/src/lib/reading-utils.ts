function isHiragana(ch: string): boolean {
  const code = ch.charCodeAt(0);
  return code >= 0x3040 && code <= 0x309f;
}

function isKatakana(ch: string): boolean {
  const code = ch.charCodeAt(0);
  return code >= 0x30a0 && code <= 0x30ff;
}

function katakanaToHiragana(text: string): string {
  let result = '';
  for (const ch of text) {
    const code = ch.charCodeAt(0);
    if (code >= 0x30a1 && code <= 0x30f6) {
      result += String.fromCharCode(code - 0x60);
    } else {
      result += ch;
    }
  }
  return result;
}

/**
 * Reconstruct the full word reading from the surface form and the stored
 * (possibly partial) reading.
 *
 * MeCab/Yomitan sometimes stores only the kanji portion's reading. For example,
 * お前 (surface) with reading まえ — the stored reading covers only 前, missing
 * the leading お. This function walks through the surface form: hiragana/katakana
 * characters pass through as-is (converted to hiragana), and the remaining kanji
 * portion is filled in from the stored reading.
 */
export function fullReading(headword: string, storedReading: string): string {
  if (!storedReading || !headword) return storedReading || '';

  const reading = katakanaToHiragana(storedReading);

  const leadingKana: string[] = [];
  const trailingKana: string[] = [];
  const chars = [...headword];

  let i = 0;
  while (i < chars.length) {
    const ch = chars[i]!;
    if (!isHiragana(ch) && !isKatakana(ch)) break;
    leadingKana.push(katakanaToHiragana(ch));
    i++;
  }

  if (i === chars.length) {
    return reading;
  }

  let j = chars.length - 1;
  while (j > i) {
    const ch = chars[j]!;
    if (!isHiragana(ch) && !isKatakana(ch)) break;
    trailingKana.unshift(katakanaToHiragana(ch));
    j--;
  }

  // Strip matching trailing kana from the stored reading to get the core kanji reading
  let coreReading = reading;
  const trailStr = trailingKana.join('');
  if (trailStr && coreReading.endsWith(trailStr)) {
    coreReading = coreReading.slice(0, -trailStr.length);
  }

  // Strip matching leading kana from the stored reading if it already includes them
  const leadStr = leadingKana.join('');
  if (leadStr && coreReading.startsWith(leadStr)) {
    return reading;
  }

  return leadStr + coreReading + trailStr;
}
