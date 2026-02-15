export function normalizeDisplayText(text: string): string {
  return text
    .replace(/\r\n/g, "\n")
    .replace(/\\N/g, "\n")
    .replace(/\\n/g, "\n")
    .trim();
}

export function normalizeTokenizerInput(displayText: string): string {
  return displayText
    .replace(/\n/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}
