type Segment = { text: string; reading?: string };

function readGroups(value: unknown): Segment[][] | null {
  if (!Array.isArray(value)) return null;
  const groups: Segment[][] = [];
  const rawGroups: unknown[] = value;
  for (const group of rawGroups) {
    if (!Array.isArray(group)) return null;
    const segments: Segment[] = [];
    const rawSegments: unknown[] = group;
    for (const segment of rawSegments) {
      if (
        typeof segment !== 'object' ||
        segment === null ||
        !('text' in segment) ||
        typeof segment.text !== 'string' ||
        ('reading' in segment &&
          segment.reading !== undefined &&
          typeof segment.reading !== 'string')
      )
        return null;
      segments.push({
        text: segment.text,
        reading:
          'reading' in segment && typeof segment.reading === 'string' ? segment.reading : undefined,
      });
    }
    groups.push(segments);
  }
  return groups;
}

function escapeText(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\[/g, '&#91;')
    .replace(/\]/g, '&#93;');
}

// Only accept a complete parse, so a lookup of just the headword cannot replace sentence context.
export function formatSentenceFurigana(
  text: string,
  results: unknown[] | null,
  highlightedText?: string,
): string | null {
  for (const result of results ?? []) {
    if (typeof result !== 'object' || result === null || !('content' in result)) continue;
    const groups = readGroups(result.content);
    if (
      !groups ||
      groups
        .flat()
        .map((segment) => segment.text)
        .join('') !== text
    )
      continue;
    const highlights: number[] = [];
    if (highlightedText) {
      let start = text.indexOf(highlightedText);
      while (start >= 0) {
        highlights.push(start);
        start = text.indexOf(highlightedText, start + highlightedText.length);
      }
    }
    let offset = 0;
    let bold = false;
    let output = '';
    for (const { text: surface, reading } of groups.flat()) {
      const start = offset;
      offset += surface.length;
      const highlighted = highlights.some(
        (position) => position < offset && position + (highlightedText?.length ?? 0) > start,
      );
      if (highlighted !== bold) output += highlighted ? '<b>' : '</b>';
      bold = highlighted;
      const escaped = escapeText(surface);
      output +=
        reading && reading !== surface && /[\p{Script=Han}々]/u.test(surface)
          ? ` ${escaped}[${escapeText(reading)}]`
          : escaped;
    }
    return output + (bold ? '</b>' : '');
  }
  return null;
}
