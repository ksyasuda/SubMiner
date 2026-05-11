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

test('docs sitemap excludes duplicate README page from indexable URLs', async () => {
  const items = [{ url: '' }, { url: 'README' }, { url: 'usage' }];

  const transformedItems = await docsConfig.sitemap?.transformItems?.(items);

  expect(transformedItems?.map((item) => item.url)).toEqual(['', 'usage']);
});
