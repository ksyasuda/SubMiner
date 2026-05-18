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

type VersionManifest = {
  latestStable: string;
  channels: Array<{ label: string; path: string }>;
  versions: Array<{ version: string; path: string }>;
};

const base = normalizeBase(process.env.SUBMINER_DOCS_BASE ?? '/');
const outDir = process.env.SUBMINER_DOCS_OUT_DIR;
const channel = normalizeChannel(process.env.SUBMINER_DOCS_CHANNEL);
const docsVersion = process.env.SUBMINER_DOCS_VERSION;
const latestStable = process.env.SUBMINER_DOCS_LATEST_STABLE ?? 'v0.14.0';
const versionManifest = parseVersionManifest(process.env.SUBMINER_DOCS_VERSION_MANIFEST);

function normalizeBase(value: string): string {
  if (!value || value === '/') return '/';
  return `/${value.replace(/^\/+|\/+$/g, '')}/`;
}

function normalizeChannel(value: string | undefined): DocsChannel {
  if (value === 'main' || value === 'stable-archive') return value;
  return 'stable-root';
}

function parseVersionManifest(value: string | undefined): VersionManifest {
  if (!value) {
    return {
      latestStable,
      channels: [
        { label: 'Latest stable', path: '/' },
        { label: 'main', path: '/main/' },
      ],
      versions: [{ version: latestStable, path: `/v/${latestStable.replace(/^v/, '')}/` }],
    };
  }

  return JSON.parse(value) as VersionManifest;
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

function pageToCanonicalHref(page: string): string | null {
  const route = pageToRoute(page);
  if (!route) return null;

  if (channel === 'main') {
    return `${DOCS_HOSTNAME}${canonicalRouteWithBase(route)}`;
  }

  if (channel === 'stable-archive' && docsVersion !== latestStable) {
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

  if (channel === 'main') {
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
  return existsSync(join(process.cwd(), pagePath));
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

const versionItems = [
  {
    text: `Latest stable (${versionManifest.latestStable})`,
    link: '/',
  },
  ...versionManifest.channels
    .filter((entry) => entry.label !== 'Latest stable')
    .map((entry) => ({ text: entry.label, link: entry.path })),
  ...versionManifest.versions.map((entry) => ({
    text: entry.version,
    link: entry.path,
  })),
];

const nav: DefaultTheme.NavItem[] = [
  { text: 'Home', link: '/' },
  { text: 'Get Started', link: '/installation' },
  { text: 'Mining', link: '/mining-workflow' },
  { text: 'Configuration', link: '/configuration' },
  { text: 'Changelog', link: '/changelog' },
  { text: 'Troubleshooting', link: '/troubleshooting' },
  { text: docsVersion ?? (channel === 'main' ? 'main' : latestStable), items: versionItems },
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
      { text: 'AniList', link: '/anilist-integration' },
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
  sitemap: {
    hostname: DOCS_HOSTNAME,
    transformItems(items) {
      return items.filter(
        (item) => item.url !== 'README' && item.url !== `${DOCS_HOSTNAME}/README`,
      );
    },
  },
  transformHead: transformPageHead,
  lastUpdated: true,
  srcExclude: ['subagents/**'],
  markdown: {
    theme: {
      light: 'catppuccin-latte',
      dark: 'catppuccin-macchiato',
    },
  },
  themeConfig: {
    logo: {
      light: withDocsBase('/assets/SubMiner.png'),
      dark: withDocsBase('/assets/SubMiner.png'),
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
