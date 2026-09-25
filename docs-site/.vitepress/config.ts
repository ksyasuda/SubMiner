import { spawnSync } from 'node:child_process';
import { existsSync } from 'node:fs';
import { join } from 'node:path';
import type { DefaultTheme, HeadConfig, TransformContext, UserConfig } from 'vitepress';

const DOCS_HOSTNAME = 'https://docs.subminer.moe';
const PLAUSIBLE_PROXY_HOSTNAME = 'https://worker.sudacode.com';
const PLAUSIBLE_SITE_SCRIPT_PATH = '/js/pa-h28Pn9ppgTJRmiSJlyPT6.js';
const PLAUSIBLE_ENDPOINT = `${PLAUSIBLE_PROXY_HOSTNAME}/api/event`;
const PLAUSIBLE_INIT_SCRIPT = [
  'window.plausible=window.plausible||function(){(plausible.q=plausible.q||[]).push(arguments)},plausible.init=plausible.init||function(i){plausible.o=i||{}};',
  `plausible.init({ endpoint: '${PLAUSIBLE_ENDPOINT}' });`,
].join('\n');

type DocsChannel = 'stable-root' | 'stable-archive' | 'main';

function optionalEnv(value: string | undefined): string | undefined {
  return value && value !== 'undefined' ? value : undefined;
}

const base = normalizeBase(optionalEnv(process.env.SUBMINER_DOCS_BASE) ?? '/');
const outDir = optionalEnv(process.env.SUBMINER_DOCS_OUT_DIR);
const docsSourceDir = optionalEnv(process.env.SUBMINER_DOCS_SOURCE_DIR) ?? process.cwd();
// The tracked `docs-site/` checkout, which stays a git working tree even when
// `docsSourceDir` points at an untracked release snapshot. Used for git lookups only.
const repoDocsDir = optionalEnv(process.env.SUBMINER_DOCS_REPO_DIR) ?? process.cwd();
const channel = normalizeChannel(optionalEnv(process.env.SUBMINER_DOCS_CHANNEL));
const docsVersion = optionalEnv(process.env.SUBMINER_DOCS_VERSION);

function normalizeBase(value: string): string {
  if (!value || value === '/') return '/';
  return `/${value.replace(/^\/+|\/+$/g, '')}/`;
}

function normalizeChannel(value: string | undefined): DocsChannel {
  if (value === 'main' || value === 'stable-archive') return value;
  return 'stable-root';
}

function withDocsBase(path: string): string {
  if (/^[a-z]+:\/\//i.test(path)) return path;
  const normalizedPath = path.startsWith('/') ? path : `/${path}`;
  if (base === '/') return normalizedPath;
  return `${base.replace(/\/$/, '')}${normalizedPath}`;
}

function pageToRoute(page: string): string | null {
  if (page === '404.md') return null;

  const route = page
    .replace(/(^|\/)index\.md$/, '')
    .replace(/\.md$/, '')
    .replace(/\/$/, '');
  return route ? `/${route}` : '/';
}

// Only the root channel is indexable. `main` and every /v/<version>/ archive are
// near-verbatim copies of it, so they own their URL via a self-referential canonical
// and are excluded from the index instead of being consolidated onto root. Uniform
// self-canonical plus noindex avoids mixing noindex with a cross-page canonical,
// which Google treats as a conflicting signal.
const isIndexableChannel = channel === 'stable-root';

function pageToCanonicalHref(page: string): string | null {
  const route = pageToRoute(page);
  if (!route) return null;

  if (!isIndexableChannel) {
    return `${DOCS_HOSTNAME}${canonicalRouteWithBase(route)}`;
  }

  return route === '/' ? `${DOCS_HOSTNAME}/` : `${DOCS_HOSTNAME}${route}`;
}

function canonicalRouteWithBase(route: string): string {
  const routeWithBase = withDocsBase(route);
  return route === '/' ? routeWithBase : routeWithBase.replace(/\/$/, '');
}

function transformPageHead({ page }: TransformContext): HeadConfig[] {
  const href = pageToCanonicalHref(page);
  const head: HeadConfig[] = href ? [['link', { rel: 'canonical', href }]] : [];

  // Crawlable so links still pass through, but out of the index: ~30 archived copies
  // of every page otherwise soak up the crawl budget the current docs need.
  if (!isIndexableChannel) {
    head.push(['meta', { name: 'robots', content: 'noindex,follow' }]);
  }

  return head;
}

function linkToPagePath(link: string): string | null {
  if (!link.startsWith('/') || link.startsWith('/v/') || link.startsWith('/main/')) {
    return null;
  }

  const withoutHash = link.split('#')[0] ?? '/';
  const withoutQuery = withoutHash.split('?')[0] ?? '/';
  const route = withoutQuery.replace(/^\/+|\/+$/g, '');
  return route ? `${route}.md` : 'index.md';
}

function hasPageForLink(link: string): boolean {
  const pagePath = linkToPagePath(link);
  if (!pagePath) return true;
  return existsSync(join(docsSourceDir, pagePath));
}

function filterNav(items: DefaultTheme.NavItem[]): DefaultTheme.NavItem[] {
  return items
    .map((item) => {
      if ('items' in item && item.items) {
        return { ...item, items: filterNav(item.items as DefaultTheme.NavItem[]) };
      }
      if ('link' in item && item.link && !hasPageForLink(item.link)) {
        return null;
      }
      return item;
    })
    .filter((item): item is DefaultTheme.NavItem => Boolean(item));
}

function filterSidebar(items: DefaultTheme.SidebarItem[]): DefaultTheme.SidebarItem[] {
  return items
    .map((item) => {
      const filteredChildren = item.items ? filterSidebar(item.items) : undefined;
      if (item.link && !hasPageForLink(item.link)) return null;
      if (item.items && filteredChildren?.length === 0 && !item.link) return null;
      return { ...item, items: filteredChildren };
    })
    .filter((item): item is DefaultTheme.SidebarItem => Boolean(item));
}

// Version navigation targets other builds (root, `/main/`, `/versions`), so it links to
// production by absolute URL: base-relative links would stay inside this build, and
// `target: '_self'` makes the VitePress router do a full page load. The list is
// deliberately static; the full release list lives on the root-only `/versions` page
// so frozen `/v/<version>/` archives never need a rebuild when a new tag ships.
const versionItems: DefaultTheme.NavItemWithLink[] = [
  { text: 'Latest stable', link: `${DOCS_HOSTNAME}/` },
  { text: 'main', link: `${DOCS_HOSTNAME}/main/` },
  { text: 'All versions', link: `${DOCS_HOSTNAME}/versions` },
].map((item) => ({ ...item, target: '_self', noIcon: true }));

function sitemapUrlToPage(url: string): string {
  const route = url.replace(/\.html$/, '').replace(/^\/+|\/+$/g, '');
  return route ? `${route}.md` : 'index.md';
}

// VitePress derives <lastmod> by running `git log` inside its source dir. Production
// builds point that at an untracked snapshot of the release tag, so the lookup comes
// back empty and the sitemap ships with no dates at all. Resolve it from the tracked
// checkout at the ref being built instead.
function lastModifiedFor(url: string): string | undefined {
  const ref = docsVersion && docsVersion !== 'main' ? docsVersion : 'HEAD';
  const result = spawnSync('git', ['log', '-1', '--format=%cI', ref, '--', sitemapUrlToPage(url)], {
    cwd: repoDocsDir,
    encoding: 'utf8',
  });

  return (result.status === 0 && result.stdout.trim()) || undefined;
}

// Only the root channel publishes a sitemap. Archived and `main` builds would emit
// their own copies listing the same canonical URLs, which just advertises the
// duplicate trees we are trying to keep out of the index.
const sitemap: UserConfig['sitemap'] = isIndexableChannel
  ? {
      hostname: DOCS_HOSTNAME,
      transformItems(items) {
        return items
          .filter((item) => item.url !== 'README' && item.url !== `${DOCS_HOSTNAME}/README`)
          .map((item) => ({ ...item, lastmod: item.lastmod ?? lastModifiedFor(item.url) }));
      },
    }
  : undefined;

const nav: DefaultTheme.NavItem[] = [
  { text: 'Home', link: '/' },
  { text: 'Get Started', link: '/installation' },
  { text: 'Mining', link: '/mining-workflow' },
  { text: 'Configuration', link: '/configuration' },
  { text: 'Changelog', link: '/changelog' },
  { text: 'Troubleshooting', link: '/troubleshooting' },
  { text: docsVersion ?? 'main', items: versionItems },
];

const sidebar: DefaultTheme.SidebarItem[] = [
  {
    text: 'Getting Started',
    items: [
      { text: 'Overview', link: '/' },
      { text: 'Installation', link: '/installation' },
      { text: 'Usage', link: '/usage' },
      { text: 'Mining Workflow', link: '/mining-workflow' },
      { text: 'Launcher Script', link: '/launcher-script' },
    ],
  },
  {
    text: 'Reference',
    items: [
      { text: 'Configuration', link: '/configuration' },
      { text: 'Keyboard Shortcuts', link: '/shortcuts' },
      { text: 'Subtitle Annotations', link: '/subtitle-annotations' },
      { text: 'Subtitle Sidebar', link: '/subtitle-sidebar' },
      { text: 'Immersion Tracking', link: '/immersion-tracking' },
      { text: 'Troubleshooting', link: '/troubleshooting' },
    ],
  },
  {
    text: 'Integrations',
    items: [
      { text: 'MPV Plugin', link: '/mpv-plugin' },
      { text: 'Anki', link: '/anki-integration' },
      { text: 'Jellyfin', link: '/jellyfin-integration' },
      { text: 'YouTube', link: '/youtube-integration' },
      { text: 'Jimaku', link: '/jimaku-integration' },
      { text: 'Subtitle Generation', link: '/subtitle-generation' },
      { text: 'TsukiHime', link: '/tsukihime-integration' },
      { text: 'AniList', link: '/anilist-integration' },
      { text: 'AniSkip', link: '/aniskip-integration' },
      { text: 'Character Dictionary', link: '/character-dictionary' },
    ],
  },
  {
    text: 'Development',
    items: [
      { text: 'Building & Testing', link: '/development' },
      { text: 'Architecture', link: '/architecture' },
      { text: 'IPC + Runtime Contracts', link: '/ipc-contracts' },
      { text: 'WebSocket + Texthooker API', link: '/websocket-texthooker-api' },
      { text: 'Changelog', link: '/changelog' },
    ],
  },
];

const config: UserConfig = {
  title: 'SubMiner Docs',
  description:
    'SubMiner: an MPV immersion-mining overlay with Yomitan and AnkiConnect integration.',
  base,
  ...(outDir ? { outDir } : {}),
  head: [
    ['link', { rel: 'preconnect', href: PLAUSIBLE_PROXY_HOSTNAME }],
    [
      'script',
      {
        async: '',
        src: `${PLAUSIBLE_PROXY_HOSTNAME}${PLAUSIBLE_SITE_SCRIPT_PATH}`,
      },
    ],
    ['script', {}, PLAUSIBLE_INIT_SCRIPT],
    ['link', { rel: 'icon', href: withDocsBase('/favicon.ico'), sizes: 'any' }],
    [
      'link',
      {
        rel: 'icon',
        type: 'image/png',
        href: withDocsBase('/favicon-32x32.png'),
        sizes: '32x32',
      },
    ],
    [
      'link',
      {
        rel: 'icon',
        type: 'image/png',
        href: withDocsBase('/favicon-16x16.png'),
        sizes: '16x16',
      },
    ],
    [
      'link',
      {
        rel: 'apple-touch-icon',
        href: withDocsBase('/apple-touch-icon.png'),
        sizes: '180x180',
      },
    ],
  ],
  appearance: 'dark',
  cleanUrls: true,
  metaChunk: true,
  sitemap,
  transformHead: transformPageHead,
  lastUpdated: true,
  srcExclude: ['subagents/**', 'README.md'],
  markdown: {
    theme: {
      light: 'catppuccin-latte',
      dark: 'catppuccin-macchiato',
    },
  },
  themeConfig: {
    logo: {
      light: '/assets/SubMiner.png',
      dark: '/assets/SubMiner.png',
    },
    siteTitle: 'SubMiner Docs',
    nav: filterNav(nav),
    sidebar: filterSidebar(sidebar),
    search: {
      provider: 'local',
    },
    footer: {
      message: 'Released under the GPL-3.0 License.',
      copyright: 'Copyright © 2026-present sudacode',
    },
    editLink: {
      pattern: 'https://github.com/ksyasuda/SubMiner/edit/main/docs-site/:path',
      text: 'Edit this page on GitHub',
    },
    outline: { level: [2, 3], label: 'On this page' },
    externalLinkIcon: true,
    docFooter: { prev: 'Previous', next: 'Next' },
    returnToTopLabel: 'Back to top',
    socialLinks: [{ icon: 'github', link: 'https://github.com/ksyasuda/SubMiner' }],
  },
};

export default config;
