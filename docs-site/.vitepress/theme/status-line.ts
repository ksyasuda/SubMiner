export function formatStatusLineFilePath(routePath: string): string {
  if (routePath === '/') return 'index.md';
  return `${routePath.replace(/^\/|\/$/g, '')}.md`;
}

// Local calendar date as YYYY-MM-DD (toISOString would give the UTC date).
export function formatStatusLineDate(date: Date): string {
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${date.getFullYear()}-${month}-${day}`;
}
