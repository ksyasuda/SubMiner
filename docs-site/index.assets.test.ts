import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';

const docsIndexPath = new URL('./index.md', import.meta.url);
const docsIndexContents = readFileSync(docsIndexPath, 'utf8');

test('docs demo media uses shared cache-busting asset version token', () => {
  expect(docsIndexContents).toMatch(/const demoAssetVersion = ['"][^'"]+['"]/);
  expect(docsIndexContents).toContain(
    ':poster="withBase(`/assets/minecard-poster.jpg?v=${demoAssetVersion}`)"',
  );
  expect(docsIndexContents).toContain(
    '<source :src="withBase(`/assets/minecard.webm?v=${demoAssetVersion}`)" type="video/webm" />',
  );
  expect(docsIndexContents).toContain(
    '<source :src="withBase(`/assets/minecard.mp4?v=${demoAssetVersion}`)" type="video/mp4" />',
  );
  expect(docsIndexContents).toContain(
    '<a :href="withBase(`/assets/minecard.webm?v=${demoAssetVersion}`)" target="_blank" rel="noreferrer">',
  );
  expect(docsIndexContents).toContain(
    '<img :src="withBase(`/assets/minecard.webp?v=${demoAssetVersion}`)" alt="SubMiner demo Animated fallback" style="width: 100%; height: auto;" />',
  );
});
