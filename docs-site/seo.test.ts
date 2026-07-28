import { expect, test } from 'bun:test';
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';
import type { TransformContext } from 'vitepress';
import docsConfig from './.vitepress/config';

const docsSiteDir = fileURLToPath(new URL('.', import.meta.url));

function makeTransformContext(page: string): TransformContext {
  return {
    page,
    siteConfig: {} as TransformContext['siteConfig'],
    siteData: {} as TransformContext['siteData'],
    pageData: {} as TransformContext['pageData'],
    title: 'SubMiner',
    description: 'SubMiner docs',
    head: [],
    content: '',
    assets: [],
  };
}

test('docs pages emit stable self-referential canonical URLs', async () => {
  const rootHead = await docsConfig.transformHead?.(makeTransformContext('index.md'));
  const usageHead = await docsConfig.transformHead?.(makeTransformContext('usage.md'));

  expect(rootHead).toContainEqual([
    'link',
    { rel: 'canonical', href: 'https://docs.subminer.moe/' },
  ]);
  expect(usageHead).toContainEqual([
    'link',
    { rel: 'canonical', href: 'https://docs.subminer.moe/usage' },
  ]);
  expect(JSON.stringify(rootHead).toLowerCase()).not.toContain('noindex');
});

test('main docs canonical uses /main/ and emits noindex', async () => {
  const previousChannel = process.env.SUBMINER_DOCS_CHANNEL;
  const previousBase = process.env.SUBMINER_DOCS_BASE;
  process.env.SUBMINER_DOCS_CHANNEL = 'main';
  process.env.SUBMINER_DOCS_BASE = '/main/';
  const { default: mainDocsConfig } = await import('./.vitepress/config?main-docs');

  const head = await mainDocsConfig.transformHead?.(makeTransformContext('usage.md'));
  const rootHead = await mainDocsConfig.transformHead?.(makeTransformContext('index.md'));

  expect(head).toContainEqual([
    'link',
    { rel: 'canonical', href: 'https://docs.subminer.moe/main/usage' },
  ]);
  expect(rootHead).toContainEqual([
    'link',
    { rel: 'canonical', href: 'https://docs.subminer.moe/main/' },
  ]);
  expect(head).toContainEqual(['meta', { name: 'robots', content: 'noindex,follow' }]);

  process.env.SUBMINER_DOCS_CHANNEL = previousChannel;
  process.env.SUBMINER_DOCS_BASE = previousBase;
});

test('latest stable archive canonical points to root equivalent', async () => {
  const previousChannel = process.env.SUBMINER_DOCS_CHANNEL;
  const previousBase = process.env.SUBMINER_DOCS_BASE;
  const previousVersion = process.env.SUBMINER_DOCS_VERSION;
  const previousLatest = process.env.SUBMINER_DOCS_LATEST_STABLE;
  process.env.SUBMINER_DOCS_CHANNEL = 'stable-archive';
  process.env.SUBMINER_DOCS_BASE = '/v/0.14.0/';
  process.env.SUBMINER_DOCS_VERSION = 'v0.14.0';
  process.env.SUBMINER_DOCS_LATEST_STABLE = 'v0.14.0';
  const { default: latestArchiveConfig } = await import('./.vitepress/config?latest-archive');

  const head = await latestArchiveConfig.transformHead?.(makeTransformContext('usage.md'));

  expect(head).toContainEqual([
    'link',
    { rel: 'canonical', href: 'https://docs.subminer.moe/usage' },
  ]);

  process.env.SUBMINER_DOCS_CHANNEL = previousChannel;
  process.env.SUBMINER_DOCS_BASE = previousBase;
  process.env.SUBMINER_DOCS_VERSION = previousVersion;
  process.env.SUBMINER_DOCS_LATEST_STABLE = previousLatest;
});

test('stable archive theme links stay on the selected version', async () => {
  const previousCwd = process.cwd();
  const previousChannel = process.env.SUBMINER_DOCS_CHANNEL;
  const previousBase = process.env.SUBMINER_DOCS_BASE;
  const previousVersion = process.env.SUBMINER_DOCS_VERSION;
  const previousLatest = process.env.SUBMINER_DOCS_LATEST_STABLE;
  const previousManifest = process.env.SUBMINER_DOCS_VERSION_MANIFEST;
  const previousVersionLinkOrigin = process.env.SUBMINER_DOCS_VERSION_LINK_ORIGIN;
  process.chdir(docsSiteDir);
  process.env.SUBMINER_DOCS_CHANNEL = 'stable-archive';
  process.env.SUBMINER_DOCS_BASE = '/v/0.12.0/';
  process.env.SUBMINER_DOCS_VERSION = 'v0.12.0';
  process.env.SUBMINER_DOCS_LATEST_STABLE = 'v0.14.0';
  process.env.SUBMINER_DOCS_VERSION_LINK_ORIGIN = 'production';
  process.env.SUBMINER_DOCS_VERSION_MANIFEST = JSON.stringify({
    latestStable: 'v0.14.0',
    channels: [
      { label: 'Latest stable', path: '/' },
      { label: 'main', path: '/main/' },
    ],
    versions: [
      { version: 'v0.14.0', path: '/v/0.14.0/' },
      { version: 'v0.12.0', path: '/v/0.12.0/' },
    ],
  });
  try {
    const { default: archiveConfig } = await import('./.vitepress/config?stable-archive-links');

    const nav = archiveConfig.themeConfig?.nav as Array<{
      text: string;
      link?: string;
      items?: Array<{ text: string; link: string }>;
    }>;
    const sidebar = archiveConfig.themeConfig?.sidebar as Array<{
      text: string;
      items?: Array<{ text: string; link: string }>;
    }>;
    const configurationNav = nav.find((item) => item.text === 'Configuration');
    const versionNav = nav.find((item) => item.text === 'v0.12.0');
    const referenceSidebar = sidebar.find((item) => item.text === 'Reference');
    const configurationSidebar = referenceSidebar?.items?.find(
      (item) => item.text === 'Configuration',
    );

    expect(configurationNav?.link).toBe('/configuration');
    expect(configurationSidebar?.link).toBe('/configuration');
    expect(versionNav?.items).toContainEqual({
      text: 'Latest stable (v0.14.0)',
      link: 'https://docs.subminer.moe/',
      target: '_self',
      noIcon: true,
    });
    expect(versionNav?.items).toContainEqual({
      text: 'main',
      link: 'https://docs.subminer.moe/main/',
      target: '_self',
      noIcon: true,
    });
    expect(versionNav?.items).toContainEqual({
      text: 'v0.14.0',
      link: 'https://docs.subminer.moe/v/0.14.0/',
      target: '_self',
      noIcon: true,
    });
    expect(versionNav?.items).toContainEqual({
      text: 'v0.12.0',
      link: 'https://docs.subminer.moe/v/0.12.0/',
      target: '_self',
      noIcon: true,
    });
    expect(archiveConfig.themeConfig?.logo).toEqual({
      light: '/assets/SubMiner.png',
      dark: '/assets/SubMiner.png',
    });
  } finally {
    process.chdir(previousCwd);
    process.env.SUBMINER_DOCS_CHANNEL = previousChannel;
    process.env.SUBMINER_DOCS_BASE = previousBase;
    process.env.SUBMINER_DOCS_VERSION = previousVersion;
    process.env.SUBMINER_DOCS_LATEST_STABLE = previousLatest;
    process.env.SUBMINER_DOCS_VERSION_MANIFEST = previousManifest;
    process.env.SUBMINER_DOCS_VERSION_LINK_ORIGIN = previousVersionLinkOrigin;
  }
});

test('local stable archive version links stay on the dev server', async () => {
  const previousCwd = process.cwd();
  const previousChannel = process.env.SUBMINER_DOCS_CHANNEL;
  const previousBase = process.env.SUBMINER_DOCS_BASE;
  const previousVersion = process.env.SUBMINER_DOCS_VERSION;
  const previousLatest = process.env.SUBMINER_DOCS_LATEST_STABLE;
  const previousManifest = process.env.SUBMINER_DOCS_VERSION_MANIFEST;
  const previousVersionLinkOrigin = process.env.SUBMINER_DOCS_VERSION_LINK_ORIGIN;
  process.chdir(docsSiteDir);
  process.env.SUBMINER_DOCS_CHANNEL = 'stable-archive';
  process.env.SUBMINER_DOCS_BASE = '/v/0.10.0/';
  process.env.SUBMINER_DOCS_VERSION = 'v0.10.0';
  process.env.SUBMINER_DOCS_LATEST_STABLE = 'v0.14.0';
  process.env.SUBMINER_DOCS_VERSION_LINK_ORIGIN = 'local';
  process.env.SUBMINER_DOCS_VERSION_MANIFEST = JSON.stringify({
    latestStable: 'v0.14.0',
    channels: [
      { label: 'Latest stable', path: '/' },
      { label: 'main', path: '/main/' },
    ],
    versions: [
      { version: 'v0.14.0', path: '/v/0.14.0/' },
      { version: 'v0.10.0', path: '/v/0.10.0/' },
    ],
  });
  try {
    const { default: archiveConfig } = await import('./.vitepress/config?local-archive-links');

    const nav = archiveConfig.themeConfig?.nav as Array<{
      text: string;
      items?: Array<{ text: string; link: string }>;
    }>;
    const versionNav = nav.find((item) => item.text === 'v0.10.0');

    expect(versionNav?.items).toContainEqual({
      text: 'Latest stable (v0.14.0)',
      link: '../../',
      target: '_self',
      noIcon: true,
    });
    expect(versionNav?.items).toContainEqual({
      text: 'main',
      link: '../../main/',
      target: '_self',
      noIcon: true,
    });
    expect(versionNav?.items).toContainEqual({
      text: 'v0.14.0',
      link: '../0.14.0/',
      target: '_self',
      noIcon: true,
    });
    expect(versionNav?.items).toContainEqual({
      text: 'v0.10.0',
      link: './',
      target: '_self',
      noIcon: true,
    });
  } finally {
    process.chdir(previousCwd);
    process.env.SUBMINER_DOCS_CHANNEL = previousChannel;
    process.env.SUBMINER_DOCS_BASE = previousBase;
    process.env.SUBMINER_DOCS_VERSION = previousVersion;
    process.env.SUBMINER_DOCS_LATEST_STABLE = previousLatest;
    process.env.SUBMINER_DOCS_VERSION_MANIFEST = previousManifest;
    process.env.SUBMINER_DOCS_VERSION_LINK_ORIGIN = previousVersionLinkOrigin;
  }
});

test('dev docs version links use local targets for version route testing', async () => {
  const previousCwd = process.cwd();
  const previousChannel = process.env.SUBMINER_DOCS_CHANNEL;
  const previousBase = process.env.SUBMINER_DOCS_BASE;
  const previousVersion = process.env.SUBMINER_DOCS_VERSION;
  const previousLatest = process.env.SUBMINER_DOCS_LATEST_STABLE;
  const previousManifest = process.env.SUBMINER_DOCS_VERSION_MANIFEST;
  const previousVersionLinkOrigin = process.env.SUBMINER_DOCS_VERSION_LINK_ORIGIN;
  process.chdir(docsSiteDir);
  delete process.env.SUBMINER_DOCS_CHANNEL;
  delete process.env.SUBMINER_DOCS_BASE;
  delete process.env.SUBMINER_DOCS_VERSION;
  // Set explicitly (like the sibling version-nav tests) so this assertion stays
  // pinned to the manifest under test instead of the config's fallback constant.
  process.env.SUBMINER_DOCS_LATEST_STABLE = 'v0.14.0';
  process.env.SUBMINER_DOCS_VERSION_LINK_ORIGIN = 'local';
  process.env.SUBMINER_DOCS_VERSION_MANIFEST = JSON.stringify({
    latestStable: 'v0.14.0',
    channels: [
      { label: 'Latest stable', path: '/' },
      { label: 'main', path: '/main/' },
    ],
    versions: [
      { version: 'v0.14.0', path: '/v/0.14.0/' },
      { version: 'v0.12.0', path: '/v/0.12.0/' },
      { version: 'v0.11.2', path: '/v/0.11.2/' },
    ],
  });
  try {
    const { default: devConfig } = await import('./.vitepress/config?dev-version-links');

    const nav = devConfig.themeConfig?.nav as Array<{
      text: string;
      items?: Array<{ text: string; link: string }>;
    }>;
    const versionNav = nav.find((item) => item.text === 'v0.14.0');

    expect(versionNav?.items).toContainEqual({
      text: 'Latest stable (v0.14.0)',
      link: '/',
      target: '_self',
      noIcon: true,
    });
    expect(versionNav?.items).toContainEqual({
      text: 'main',
      link: '/main/',
      target: '_self',
      noIcon: true,
    });
    expect(versionNav?.items).toContainEqual({
      text: 'v0.12.0',
      link: '/v/0.12.0/',
      target: '_self',
      noIcon: true,
    });
    expect(versionNav?.items?.map((item) => item.text)).toEqual([
      'Latest stable (v0.14.0)',
      'main',
      'v0.14.0',
      'v0.12.0',
      'v0.11.2',
    ]);
  } finally {
    process.chdir(previousCwd);
    process.env.SUBMINER_DOCS_CHANNEL = previousChannel;
    process.env.SUBMINER_DOCS_BASE = previousBase;
    process.env.SUBMINER_DOCS_VERSION = previousVersion;
    process.env.SUBMINER_DOCS_LATEST_STABLE = previousLatest;
    process.env.SUBMINER_DOCS_VERSION_MANIFEST = previousManifest;
    process.env.SUBMINER_DOCS_VERSION_LINK_ORIGIN = previousVersionLinkOrigin;
  }
});

test('dev server redirects unserved version routes to production docs', () => {
  let routeHandler:
    | ((req: { url?: string }, res: DevRedirectResponse, next: () => void) => void)
    | undefined;
  const fakeServer = {
    middlewares: {
      use(handler: typeof routeHandler) {
        routeHandler = handler;
      },
    },
  };
  const plugins = Array.isArray(docsConfig.vite?.plugins)
    ? docsConfig.vite.plugins
    : [docsConfig.vite?.plugins].filter(Boolean);
  const redirectPlugin = plugins.find(
    (plugin): plugin is { name: string; configureServer: (server: never) => void } =>
      Boolean(plugin) &&
      typeof plugin === 'object' &&
      'name' in plugin &&
      plugin.name === 'subminer-docs-local-version-redirects' &&
      'configureServer' in plugin,
  );
  expect(redirectPlugin).toBeDefined();
  redirectPlugin?.configureServer(fakeServer as never);

  const response = new DevRedirectResponse();
  let nextCalled = false;
  routeHandler?.({ url: '/v/0.14.0/?from=dev' }, response, () => {
    nextCalled = true;
  });

  expect(nextCalled).toBe(false);
  expect(response.statusCode).toBe(302);
  expect(response.headers.location).toBe('https://docs.subminer.moe/v/0.14.0/?from=dev');

  const rootResponse = new DevRedirectResponse();
  routeHandler?.({ url: '/configuration' }, rootResponse, () => {
    nextCalled = true;
  });
  expect(rootResponse.ended).toBe(false);
  expect(nextCalled).toBe(true);
});

test('dev server serves local archive files for local version links', async () => {
  const previousVersionLinkOrigin = process.env.SUBMINER_DOCS_VERSION_LINK_ORIGIN;
  const previousArchiveDir = process.env.SUBMINER_DOCS_LOCAL_ARCHIVE_DIR;
  const archiveDir = mkdtempSync(join(tmpdir(), 'subminer-docs-archive-'));
  mkdirSync(join(archiveDir, 'v/0.14.0'), { recursive: true });
  writeFileSync(join(archiveDir, 'v/0.14.0/index.html'), '<h1>local archive</h1>');
  process.env.SUBMINER_DOCS_VERSION_LINK_ORIGIN = 'local';
  process.env.SUBMINER_DOCS_LOCAL_ARCHIVE_DIR = archiveDir;
  try {
    const { default: localDevConfig } = await import(
      `./.vitepress/config?local-dev-redirects-${Date.now()}`
    );
    let routeHandler:
      | ((req: { url?: string }, res: DevRedirectResponse, next: () => void) => void)
      | undefined;
    const fakeServer = {
      middlewares: {
        use(handler: typeof routeHandler) {
          routeHandler = handler;
        },
      },
    };
    const plugins = Array.isArray(localDevConfig.vite?.plugins)
      ? localDevConfig.vite.plugins
      : [localDevConfig.vite?.plugins].filter(Boolean);
    const redirectPlugin = plugins.find(
      (plugin): plugin is { name: string; configureServer: (server: never) => void } =>
        Boolean(plugin) &&
        typeof plugin === 'object' &&
        'name' in plugin &&
        plugin.name === 'subminer-docs-local-version-redirects' &&
        'configureServer' in plugin,
    );
    redirectPlugin?.configureServer(fakeServer as never);

    const response = new DevRedirectResponse();
    let nextCalled = false;
    routeHandler?.({ url: '/v/0.14.0/?from=dev' }, response, () => {
      nextCalled = true;
    });

    expect(nextCalled).toBe(false);
    expect(response.statusCode).toBe(200);
    expect(response.headers['content-type']).toBe('text/html; charset=utf-8');
    expect(response.headers.location).toBeUndefined();
    expect(response.body).toBe('<h1>local archive</h1>');
  } finally {
    process.env.SUBMINER_DOCS_VERSION_LINK_ORIGIN = previousVersionLinkOrigin;
    process.env.SUBMINER_DOCS_LOCAL_ARCHIVE_DIR = previousArchiveDir;
    rmSync(archiveDir, { recursive: true, force: true });
  }
});

class DevRedirectResponse {
  statusCode = 200;
  headers: Record<string, string> = {};
  ended = false;
  body = '';

  setHeader(name: string, value: string) {
    this.headers[name.toLowerCase()] = value;
  }

  end(chunk?: string | Uint8Array) {
    if (chunk) {
      this.body = typeof chunk === 'string' ? chunk : new TextDecoder().decode(chunk);
    }
    this.ended = true;
  }
}

test('docs sitemap excludes duplicate README page from indexable URLs', async () => {
  const items = [{ url: '' }, { url: 'README' }, { url: 'usage' }];

  const transformedItems = await docsConfig.sitemap?.transformItems?.(items);

  expect(transformedItems?.map((item) => item.url)).toEqual(['', 'usage']);
});
