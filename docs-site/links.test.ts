import { expect, test } from 'bun:test';
import { readdirSync, readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

const docsSiteDir = fileURLToPath(new URL('.', import.meta.url));

// Mirrors VitePress' heading slugifier (vitepress/dist/node, `rControl` + `rSpecial`).
// Note that a *run* of special characters collapses to a single `-`, so
// "KDE Plasma & other" becomes "kde-plasma-other", not "kde-plasma--other".
const rControl = new RegExp('[\\u0000-\\u001f]', 'g');
const rSpecial = /[\s~`!@#$%^&*()\-_+=[\]{}|\\;:"'“”‘’<>,.?/]+/g;

function slugify(heading: string): string {
  return heading
    .replace(rControl, '')
    .replace(rSpecial, '-')
    .replace(/-{2,}/g, '-')
    .replace(/^-+|-+$/g, '')
    .replace(/^(\d)/, '_$1')
    .toLowerCase();
}

const EXCLUDED_PAGES = new Set(['README.md']);
const UNLISTED_ROUTES = new Set(['/demos']);
const PUBLIC_PREFIXES = ['/assets/', '/screenshots/', '/config.example.jsonc', '/favicon'];

function loadPages(): Map<string, string> {
  const pages = new Map<string, string>();
  for (const file of readdirSync(docsSiteDir)) {
    if (!file.endsWith('.md') || EXCLUDED_PAGES.has(file)) continue;
    const route = `/${file.replace(/\.md$/, '')}`;
    pages.set(route, readFileSync(`${docsSiteDir}${file}`, 'utf8'));
  }
  return pages;
}

function anchorsFor(contents: string): Set<string> {
  const anchors = new Set<string>();
  for (const match of contents.matchAll(/^#{1,6}\s+(.+?)\s*$/gm)) {
    let heading = match[1]!;
    const explicitId = heading.match(/\{#([^}]+)\}\s*$/);
    if (explicitId) {
      anchors.add(explicitId[1]!);
      heading = heading.replace(/\{#[^}]+\}\s*$/, '');
    }
    anchors.add(slugify(heading.replace(/`/g, '')));
  }
  return anchors;
}

function resolveRoute(target: string, fromRoute: string): string {
  if (target === '') return fromRoute;
  if (target.startsWith('./')) return `/${target.slice(2).replace(/\.md$/, '')}`;
  const normalized = target.replace(/\.md$/, '').replace(/\/$/, '');
  return normalized === '' ? '/index' : normalized;
}

const pages = loadPages();
const anchors = new Map([...pages].map(([route, body]) => [route, anchorsFor(body)]));

test('every internal docs link resolves to an existing page', () => {
  const broken: string[] = [];

  for (const [route, body] of pages) {
    for (const match of body.matchAll(/\]\((\/[^)\s]*|\.\/[^)\s]*)\)/g)) {
      const link = match[1]!;
      const target = link.split('#')[0]!;
      if (PUBLIC_PREFIXES.some((prefix) => target.startsWith(prefix))) continue;

      const resolved = resolveRoute(target, route);
      if (resolved !== '/index' && !pages.has(resolved)) {
        broken.push(`${route.slice(1)}.md -> ${link}`);
      }
    }
  }

  expect(broken).toEqual([]);
});

test('every internal docs anchor matches a real heading slug', () => {
  const broken: string[] = [];

  for (const [route, body] of pages) {
    for (const match of body.matchAll(/\]\((\/[^)\s]*|\.\/[^)\s]*|#[^)\s]*)\)/g)) {
      const link = match[1]!;
      const hashIndex = link.indexOf('#');
      if (hashIndex < 0) continue;

      const target = link.slice(0, hashIndex);
      const anchor = link.slice(hashIndex + 1);
      if (PUBLIC_PREFIXES.some((prefix) => target.startsWith(prefix))) continue;

      const resolved = resolveRoute(target, route);
      const pageAnchors = anchors.get(resolved);
      if (!pageAnchors || pageAnchors.has(anchor)) continue;

      broken.push(`${route.slice(1)}.md -> ${link}`);
    }
  }

  expect(broken).toEqual([]);
});

test('slugify matches the VitePress cases these docs actually rely on', () => {
  // Regression guards for the anchors that were previously wrong.
  expect(slugify('N+1 Word Highlighting')).toBe('n-1-word-highlighting');
  expect(slugify('KDE Plasma & other Wayland compositors')).toBe(
    'kde-plasma-other-wayland-compositors',
  );
  expect(slugify('Proxy Mode Setup (Yomitan / Texthooker)')).toBe(
    'proxy-mode-setup-yomitan-texthooker',
  );
  expect(slugify('Kiku/Lapis Integration')).toBe('kiku-lapis-integration');
  expect(slugify('Secondary Subtitles')).toBe('secondary-subtitles');
  expect(slugify('2. Install SubMiner')).toBe('_2-install-subminer');
});

test('every docs page is reachable from the sidebar unless explicitly unlisted', async () => {
  const { default: config } = await import('./.vitepress/config');
  const sidebar = config.themeConfig?.sidebar as Array<{
    items?: Array<{ text: string; link?: string }>;
  }>;

  const linked = new Set<string>();
  for (const group of sidebar) {
    for (const item of group.items ?? []) {
      if (item.link) linked.add(item.link === '/' ? '/index' : item.link);
    }
  }

  const orphans = [...pages.keys()].filter(
    (route) => !linked.has(route) && !UNLISTED_ROUTES.has(route),
  );
  expect(orphans).toEqual([]);
});
