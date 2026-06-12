export function normalizeAnilistSearchQuery(query: string): string {
  const trimmed = query.trim().replace(/\s+/g, ' ');
  const withoutSeason = trimmed
    .replace(/\s*[\[(]\s*Season\s+0?\d+\s*[\])]\s*$/i, '')
    .replace(/\s*[-:]\s*Season\s+0?\d+\s*$/i, '')
    .replace(/\s+Season\s+0?\d+\s*$/i, '')
    .trim();
  return withoutSeason.length > 0 ? withoutSeason : trimmed;
}
