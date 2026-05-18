import { expect, test } from 'bun:test';
import type { TransformContext } from 'vitepress';
import docsConfig from './.vitepress/config';

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

test('docs sitemap excludes duplicate README page from indexable URLs', async () => {
  const items = [{ url: '' }, { url: 'README' }, { url: 'usage' }];

  const transformedItems = await docsConfig.sitemap?.transformItems?.(items);

  expect(transformedItems?.map((item) => item.url)).toEqual(['', 'usage']);
});
