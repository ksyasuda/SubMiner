export function formatStatusLineFilePath(routePath: string): string {
  if (routePath === '/') return 'index.md';
  return `${routePath.replace(/^\/|\/$/g, '')}.md`;
}
