import { expect, test } from 'bun:test';
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
  expect(mainDocsConfig.sitemap).toBeUndefined();

  process.env.SUBMINER_DOCS_CHANNEL = previousChannel;
  process.env.SUBMINER_DOCS_BASE = previousBase;
});

test.each([
  ['v0.14.0', '/v/0.14.0/', 'https://docs.subminer.moe/v/0.14.0/usage'],
  ['v0.12.0', '/v/0.12.0/', 'https://docs.subminer.moe/v/0.12.0/usage'],
])(
  '%s archive keeps a self-referential canonical and stays out of the index',
  async (version, base, expectedCanonical) => {
    const previousChannel = process.env.SUBMINER_DOCS_CHANNEL;
    const previousBase = process.env.SUBMINER_DOCS_BASE;
    const previousVersion = process.env.SUBMINER_DOCS_VERSION;
    process.env.SUBMINER_DOCS_CHANNEL = 'stable-archive';
    process.env.SUBMINER_DOCS_BASE = base;
    process.env.SUBMINER_DOCS_VERSION = version;
    try {
      const { default: archiveConfig } = await import(`./.vitepress/config?archive-${version}`);

      const head = await archiveConfig.transformHead?.(makeTransformContext('usage.md'));

      expect(head).toContainEqual(['link', { rel: 'canonical', href: expectedCanonical }]);
      expect(head).toContainEqual(['meta', { name: 'robots', content: 'noindex,follow' }]);
      // A sitemap here would advertise the archive tree we just excluded.
      expect(archiveConfig.sitemap).toBeUndefined();
    } finally {
      process.env.SUBMINER_DOCS_CHANNEL = previousChannel;
      process.env.SUBMINER_DOCS_BASE = previousBase;
      process.env.SUBMINER_DOCS_VERSION = previousVersion;
    }
  },
);

test('archive nav keeps page links in-version and version links release-independent', async () => {
  const previousCwd = process.cwd();
  const previousChannel = process.env.SUBMINER_DOCS_CHANNEL;
  const previousBase = process.env.SUBMINER_DOCS_BASE;
  const previousVersion = process.env.SUBMINER_DOCS_VERSION;
  process.chdir(docsSiteDir);
  process.env.SUBMINER_DOCS_CHANNEL = 'stable-archive';
  process.env.SUBMINER_DOCS_BASE = '/v/0.12.0/';
  process.env.SUBMINER_DOCS_VERSION = 'v0.12.0';
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
    const configurationSidebar = sidebar
      .find((item) => item.text === 'Reference')
      ?.items?.find((item) => item.text === 'Configuration');

    expect(nav.find((item) => item.text === 'Configuration')?.link).toBe('/configuration');
    expect(configurationSidebar?.link).toBe('/configuration');
    // Frozen archives must not embed the release list, or every new tag would
    // invalidate them. They point at the root-only /versions page instead.
    expect(nav.find((item) => item.text === 'v0.12.0')?.items).toEqual([
      { text: 'Latest stable', link: 'https://docs.subminer.moe/', target: '_self', noIcon: true },
      { text: 'main', link: 'https://docs.subminer.moe/main/', target: '_self', noIcon: true },
      {
        text: 'All versions',
        link: 'https://docs.subminer.moe/versions',
        target: '_self',
        noIcon: true,
      },
    ]);
    expect(archiveConfig.themeConfig?.logo).toEqual({
      light: '/assets/SubMiner.png',
      dark: '/assets/SubMiner.png',
    });
  } finally {
    process.chdir(previousCwd);
    process.env.SUBMINER_DOCS_CHANNEL = previousChannel;
    process.env.SUBMINER_DOCS_BASE = previousBase;
    process.env.SUBMINER_DOCS_VERSION = previousVersion;
  }
});

test('docs sitemap excludes duplicate README page from indexable URLs', async () => {
  const items = [{ url: '' }, { url: 'README' }, { url: 'usage' }];

  const transformedItems = await docsConfig.sitemap?.transformItems?.(items);

  expect(transformedItems?.map((item) => item.url)).toEqual(['', 'usage']);
});

test('docs sitemap dates every URL from the tracked checkout', async () => {
  const previousRepoDir = process.env.SUBMINER_DOCS_REPO_DIR;
  // Production builds render from an untracked snapshot, so the date has to come from
  // the real checkout rather than VitePress's own srcDir git lookup.
  process.env.SUBMINER_DOCS_REPO_DIR = docsSiteDir;
  try {
    const { default: sitemapConfig } = await import('./.vitepress/config?sitemap-lastmod');

    const items = await sitemapConfig.sitemap?.transformItems?.([{ url: '' }, { url: 'usage' }]);

    expect(items).toHaveLength(2);
    for (const item of items ?? []) {
      expect(item.lastmod).toMatch(/^\d{4}-\d{2}-\d{2}T/);
    }
  } finally {
    process.env.SUBMINER_DOCS_REPO_DIR = previousRepoDir;
  }
});
