export function normalizeDisplayText(text: string): string {
  return text.replace(/\r\n/g, '\n').replace(/\\N/g, '\n').replace(/\\n/g, '\n').trim();
}

const INVISIBLE_SEPARATOR_PATTERN = /[\u200b\u2060\ufeff]/g;

export function normalizeTokenizerInput(displayText: string): string {
  return displayText
    .replace(/\n/g, ' ')
    .replace(INVISIBLE_SEPARATOR_PATTERN, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}
